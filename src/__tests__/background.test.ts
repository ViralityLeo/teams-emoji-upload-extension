import * as slackmojisImport from "../slackmojisImport";
// Set up chrome mock before importing modules
const mockChrome = {
  action: { onClicked: { addListener: jest.fn() } },
  scripting: { executeScript: jest.fn().mockResolvedValue([]) },
  tabs: { query: jest.fn().mockResolvedValue([]) },
  runtime: {
    getURL: jest.fn((path: string) => `chrome-extension://test/${path}`),
    onMessage: {
      addListener: jest.fn(),
    },
    sendMessage: jest.fn(),
  },
  storage: {
    local: {
      get: jest.fn(),
      set: jest.fn(),
    },
  },
};

Object.defineProperty(global, "chrome", {
  value: mockChrome,
  writable: true,
  configurable: true,
});

// Now import the modules
import { ProcessResult, FileDetails } from "../types";
import {
  formatErrorMessage,
  handleFileProcessing,
  handleBrowseEmojis,
  handleSlackmojisImport,
} from "../background";

const backgroundListener =
  mockChrome.runtime.onMessage.addListener.mock.calls[0][0];
const toolbarListener =
  mockChrome.action.onClicked.addListener.mock.calls[0][0];

// Mock MsTeamsClient
jest.mock("../msTeams", () => {
  return jest.fn().mockImplementation(() => ({
    uploadFiles: jest.fn(),
  }));
});

describe("Background Script", () => {
  let MsTeamsClient: any;

  beforeEach(() => {
    jest.clearAllMocks();
    MsTeamsClient = require("../msTeams");
  });

  afterEach(() => {
    jest.clearAllMocks();
  });

  it("opens uploader inside the originating page", () => {
    expect(require("../../manifest.json").action.default_popup).toBeUndefined();
    toolbarListener({ id: 42 });
    expect(mockChrome.scripting.executeScript).toHaveBeenCalledWith({
      target: { tabId: 42 },
      func: expect.any(Function),
      args: ["chrome-extension://test/popup.html?tabId=42"],
    });
  });

  it("rejects Slackmojis messages from other pages", () => {
    const reply = jest.fn();
    backgroundListener(
      { action: "importSlackmoji" },
      { url: "https://evil.example" },
      reply,
    );
    expect(reply).toHaveBeenCalledWith({
      success: false,
      error: "Slackmojis request rejected",
    });
    expect(mockChrome.storage.local.get).not.toHaveBeenCalled();
    expect(MsTeamsClient).not.toHaveBeenCalled();
  });

  it("returns tenant IDs without exposing stored Teams tokens", async () => {
    const reply = jest.fn();
    mockChrome.storage.local.get.mockResolvedValue({
      teamsTenantSessions: [
        {
          permissionsId: "tenant",
          ic3: "secret-ic3",
          chatsvcagg: "secret-chat",
        },
      ],
    });
    backgroundListener(
      { action: "slackmojisTargets" },
      { url: "https://slackmojis.com/" },
      reply,
    );
    await new Promise((resolve) => setTimeout(resolve, 0));
    expect(reply).toHaveBeenCalledWith({
      success: true,
      tenantIds: ["tenant"],
    });
    expect(JSON.stringify(reply.mock.calls)).not.toContain("secret");
  });

  it("imports Slackmojis through existing deduplication upload flow", async () => {
    const tokens = { chatsvcagg: "chat", ic3: "ic3", permissionsId: "tenant" };
    mockChrome.storage.local.get.mockResolvedValue({
      teamsTenantSessions: [tokens],
    });
    const file = {
      name: "parrot.gif",
      size: 6,
      type: "image/gif",
      base64: "R0lGODlh",
    };
    const download = jest
      .spyOn(slackmojisImport, "downloadSlackmoji")
      .mockResolvedValue(file);
    const upload = jest
      .fn()
      .mockResolvedValue({ success: true, status: "Skipped existing image" });
    MsTeamsClient.mockImplementation(() => ({ uploadFiles: upload }));
    await expect(
      handleSlackmojisImport(
        "tenant",
        "https://emojis.slackmojis.com/emojis/images/1/2/parrot.gif",
        "parrot.gif",
      ),
    ).resolves.toEqual({ success: true, status: "Skipped existing image" });
    expect(upload).toHaveBeenCalledWith([file]);
    expect(MsTeamsClient).toHaveBeenCalledWith("ic3", "chat", "tenant");
    download.mockRestore();
  });

  it("rejects an unknown tenant before downloading Slackmojis", async () => {
    mockChrome.storage.local.get.mockResolvedValue({ teamsTenantSessions: [] });
    const download = jest.spyOn(slackmojisImport, "downloadSlackmoji");
    await expect(
      handleSlackmojisImport(
        "unknown",
        "https://emojis.slackmojis.com/emojis/images/1/2/parrot.gif",
        "parrot.gif",
      ),
    ).rejects.toThrow("signed-in Teams");
    expect(download).not.toHaveBeenCalled();
    download.mockRestore();
  });

  it("routes duplicate checks to the selected tenant", async () => {
    const checkDuplicates = jest
      .fn()
      .mockResolvedValue({ success: true, tenantId: "tenant", groups: [] });
    MsTeamsClient.mockImplementation(() => ({ checkDuplicates }));
    const result = await new Promise((resolve) => {
      expect(
        backgroundListener(
          {
            action: "checkEmojiDuplicates",
            tokens: { chatsvcagg: "chat", ic3: "ic3", permissionsId: "tenant" },
          },
          {},
          resolve,
        ),
      ).toBe(true);
    });
    expect(result).toMatchObject({ tenantId: "tenant" });
    expect(checkDuplicates).toHaveBeenCalledTimes(1);
    expect(MsTeamsClient).toHaveBeenCalledWith("ic3", "chat", "tenant");
  });

  it("routes tenant browsing through the shared background queue", async () => {
    const browse = jest
      .fn()
      .mockResolvedValue({ success: true, tenantId: "tenant", items: [] });
    MsTeamsClient.mockImplementation(() => ({ browseEmojis: browse }));
    await expect(
      handleBrowseEmojis(
        { chatsvcagg: "chat", ic3: "ic3", permissionsId: "tenant" },
        "cat",
        2,
      ),
    ).resolves.toMatchObject({ tenantId: "tenant" });
    expect(MsTeamsClient).toHaveBeenCalledWith("ic3", "chat", "tenant");
    expect(browse).toHaveBeenCalledWith("cat", 2, undefined);
  });

  it("rejects missing sessions and invalid pages before viewer requests", async () => {
    await expect(
      handleBrowseEmojis({ chatsvcagg: null, ic3: null, permissionsId: null }),
    ).rejects.toThrow("signed-in Teams");
    await expect(
      handleBrowseEmojis(
        { chatsvcagg: "chat", ic3: "ic3", permissionsId: "tenant" },
        "",
        -1,
      ),
    ).rejects.toThrow("Invalid emoji viewer request");
    expect(MsTeamsClient).not.toHaveBeenCalled();
  });

  describe("formatErrorMessage", () => {
    it("should format Error objects", () => {
      const error = new Error("Test error");
      expect(formatErrorMessage(error)).toBe("Test error");
    });

    it("should format JSON strings", () => {
      const jsonString = JSON.stringify({ message: "JSON error" });
      expect(formatErrorMessage(jsonString)).toBe(
        JSON.stringify({ message: "JSON error" }, null, 2),
      );
    });

    it("should handle non-JSON strings", () => {
      const plainString = "Plain error message";
      expect(formatErrorMessage(plainString)).toBe("Plain error message");
    });

    it("should format objects", () => {
      const errorObject = { code: 404, message: "Not found" };
      expect(formatErrorMessage(errorObject)).toBe(
        JSON.stringify(errorObject, null, 2),
      );
    });
  });

  describe("handleFileProcessing", () => {
    it("should process files successfully", async () => {
      // Setup mock implementation
      const mockUploadFiles = jest.fn().mockResolvedValue({
        success: true,
        status:
          "Your emojis have been uploaded - please allow 20 minutes for Teams to sync",
      });

      MsTeamsClient.mockImplementation(() => ({
        uploadFiles: mockUploadFiles,
      }));

      // Test data
      const mockFiles: FileDetails[] = [
        {
          name: "test.png",
          size: 1024,
          type: "image/png",
          base64: "dGVzdA==", // test in base64
        },
      ];

      const mockTokens = {
        chatsvcagg: "chat-token",
        ic3: "ic3-token",
        permissionsId: "permissions-id",
      };

      // Call the function
      const result = await handleFileProcessing(mockFiles, mockTokens);

      // Assertions
      expect(MsTeamsClient).toHaveBeenCalledWith(
        "ic3-token",
        "chat-token",
        "permissions-id",
      );
      expect(mockUploadFiles).toHaveBeenCalledWith(expect.any(Array));
      expect(result).toEqual({
        success: true,
        status:
          "Your emojis have been uploaded - please allow 20 minutes for Teams to sync",
      });
      expect(chrome.runtime.sendMessage).toHaveBeenCalledWith({
        type: "processUpdate",
        success: true,
        status:
          "Your emojis have been uploaded - please allow 20 minutes for Teams to sync",
        error: undefined,
      });
      // Verify storage state was saved correctly
      expect(chrome.storage.local.set).toHaveBeenCalledWith({
        processingState: {
          status:
            "Your emojis have been uploaded - please allow 20 minutes for Teams to sync",
          type: "success",
        },
      });
    });

    it("serializes concurrent batches", async () => {
      let finishFirst!: (value: ProcessResult) => void;
      const firstUpload = new Promise<ProcessResult>((resolve) => {
        finishFirst = resolve;
      });
      const upload = jest
        .fn()
        .mockReturnValueOnce(firstUpload)
        .mockResolvedValue({ success: true, status: "Done" });
      MsTeamsClient.mockImplementation(() => ({ uploadFiles: upload }));
      const tokens = {
        chatsvcagg: "chat",
        ic3: "ic3",
        permissionsId: "tenant",
      };
      const first = handleFileProcessing([], tokens);
      const second = handleFileProcessing([], tokens);
      await Promise.resolve();
      expect(upload).toHaveBeenCalledTimes(1);
      finishFirst({ success: true, status: "Done" });
      await Promise.all([first, second]);
      expect(upload).toHaveBeenCalledTimes(2);
    });

    it("should handle missing tokens", async () => {
      // Test data with missing tokens
      const mockFiles: FileDetails[] = [
        {
          name: "test.png",
          size: 1024,
          type: "image/png",
          base64: "dGVzdA==", // test in base64
        },
      ];

      const mockTokens = {
        chatsvcagg: null,
        ic3: null,
        permissionsId: null,
      };

      // Call the function
      const result = await handleFileProcessing(mockFiles, mockTokens);

      // Assertions
      expect(result).toEqual({
        success: false,
        error: "Could not find required tokens",
      });
      expect(chrome.runtime.sendMessage).toHaveBeenCalledWith({
        type: "processUpdate",
        success: false,
        error: "Could not find required tokens",
      });
      // Verify storage state was saved correctly
      expect(chrome.storage.local.set).toHaveBeenCalledWith({
        processingState: {
          status: "Could not find required tokens",
          type: "error",
        },
      });
    });

    it("should handle upload errors", async () => {
      // Setup mock implementation to throw an error
      const mockUploadFiles = jest
        .fn()
        .mockRejectedValue(new Error("Upload failed"));

      MsTeamsClient.mockImplementation(() => ({
        uploadFiles: mockUploadFiles,
      }));

      // Test data
      const mockFiles: FileDetails[] = [
        {
          name: "test.png",
          size: 1024,
          type: "image/png",
          base64: "dGVzdA==", // test in base64
        },
      ];

      const mockTokens = {
        chatsvcagg: "chat-token",
        ic3: "ic3-token",
        permissionsId: "permissions-id",
      };

      // Call the function
      const result = await handleFileProcessing(mockFiles, mockTokens);

      // Assertions
      expect(result).toEqual({
        success: false,
        error: "Upload failed",
      });
      expect(chrome.runtime.sendMessage).toHaveBeenCalledWith({
        type: "processUpdate",
        success: false,
        error: "Upload failed",
      });
      // Verify storage state was saved correctly
      expect(chrome.storage.local.set).toHaveBeenCalledWith({
        processingState: {
          status: "Upload failed",
          type: "error",
        },
      });
    });

    // Test that setting initial processing state works
    it("should set initial processing state on file processing", async () => {
      // Reset the mocks
      jest.clearAllMocks();

      // Create test data
      const mockFiles: FileDetails[] = [
        {
          name: "test.png",
          size: 1024,
          type: "image/png",
          base64: "dGVzdA==", // test in base64
        },
      ];

      const mockTokens = {
        chatsvcagg: "chat-token",
        ic3: "ic3-token",
        permissionsId: "permissions-id",
      };

      // Setup a mock for uploadFiles
      const mockUploadFiles = jest.fn().mockResolvedValue({
        success: true,
        status: "Test status",
      });

      MsTeamsClient.mockImplementation(() => ({
        uploadFiles: mockUploadFiles,
      }));

      // Call handleFileProcessing
      await handleFileProcessing(mockFiles, mockTokens);

      // Check if chrome.storage.local.set was called with the initial state
      // This verifies our state storage logic works properly
      expect(chrome.storage.local.set).toHaveBeenCalledWith(
        expect.objectContaining({
          processingState: expect.objectContaining({
            status: expect.any(String),
            type: expect.any(String),
          }),
        }),
      );
    });
  });
});
