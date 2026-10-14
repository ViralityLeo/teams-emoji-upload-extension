import MsTeamsClient from "../msTeams";
import { webcrypto } from "node:crypto";
import { FileDetails } from "../types";

Object.defineProperty(globalThis, "crypto", { value: webcrypto });
Object.defineProperty(globalThis, "chrome", {
  configurable: true,
  value: {
    storage: {
      local: {
        get: jest.fn().mockResolvedValue({}),
        set: jest.fn().mockResolvedValue(undefined),
      },
    },
  },
});

// Mock fetch
window.fetch = jest.fn();

describe("MsTeamsClient", () => {
  let client: MsTeamsClient;
  const mockIc3Token = "mock-ic3-token";
  const mockChatToken = "mock-chat-token";
  const mockPermissionsId = "mock-permissions-id";

  beforeEach(() => {
    client = new MsTeamsClient(mockIc3Token, mockChatToken, mockPermissionsId);
    (fetch as jest.Mock).mockClear();
  });

  describe("createObject", () => {
    it("should create an object and return its ID", async () => {
      const mockResponse = {
        id: "mock-document-id",
        headers: {
          get: jest.fn().mockReturnValue("mock-cv"),
        },
        ok: true,
        json: async () => ({ id: "mock-document-id" }),
      };
      (fetch as jest.Mock).mockResolvedValueOnce(mockResponse);

      // @ts-ignore - accessing private method for testing
      const result = await client.createObject();

      expect(result).toEqual({
        id: "mock-document-id",
        msCv: "mock-cv",
      });
      expect(fetch).toHaveBeenCalledWith(
        "https://as-prod.asyncgw.teams.microsoft.com/v1/objects/",
        expect.objectContaining({
          method: "POST",
          headers: expect.objectContaining({
            authorization: `Bearer ${mockIc3Token}`,
          }),
          body: expect.any(String),
        }),
      );
    });

    it("should throw an error if the request fails", async () => {
      (fetch as jest.Mock).mockResolvedValueOnce({
        ok: false,
        statusText: "Unauthorized",
      });

      // @ts-ignore - accessing private method for testing
      await expect(client.createObject()).rejects.toThrow(
        "Failed to create object: Unauthorized",
      );
    });
  });

  describe("uploadFileContent", () => {
    it("should upload file content successfully", async () => {
      (fetch as jest.Mock).mockResolvedValueOnce({
        ok: true,
      });

      const mockFile: FileDetails = {
        name: "test.png",
        size: 1024,
        type: "image/png",
        base64: "dGVzdA==", // test in base64
      };

      // @ts-ignore - accessing private method for testing
      await expect(
        client.uploadFileContent("mock-document-id", "mock-cv", mockFile),
      ).resolves.not.toThrow();

      expect(fetch).toHaveBeenCalledWith(
        "https://as-prod.asyncgw.teams.microsoft.com/v1/objects/mock-document-id/content/imgpsh",
        expect.objectContaining({
          method: "PUT",
          headers: expect.objectContaining({
            authorization: `Bearer ${mockIc3Token}`,
          }),
        }),
      );
    });

    it("should throw an error if the upload fails", async () => {
      const mockResponse = {
        ok: false,
        statusText: "Bad Request",
        text: async () => "Invalid file format",
      };
      (fetch as jest.Mock).mockResolvedValueOnce(mockResponse);

      const mockFile: FileDetails = {
        name: "test.png",
        size: 1024,
        type: "image/png",
        base64: "dGVzdA==", // test in base64
      };

      // @ts-ignore - accessing private method for testing
      await expect(
        client.uploadFileContent("mock-document-id", "mock-cv", mockFile),
      ).rejects.toThrow(
        "Failed to upload image: Bad Request - Invalid file format",
      );
    });
  });

  describe("sendMetadata", () => {
    it("should send metadata successfully", async () => {
      const mockResponse = {
        ok: true,
        json: async () => ({ success: true }),
      };
      (fetch as jest.Mock).mockResolvedValueOnce(mockResponse);

      // @ts-ignore - accessing private method for testing
      await expect(
        client.sendMetadata("mock-document-id", "mock-cv", ["test"]),
      ).resolves.not.toThrow();

      expect(fetch).toHaveBeenCalledWith(
        "https://teams.microsoft.com/api/csa/apac/api/v1/customemoji/metadata",
        expect.objectContaining({
          method: "POST",
          headers: expect.objectContaining({
            authorization: `Bearer ${mockChatToken}`,
          }),
          body: JSON.stringify({
            shortcuts: ["test"],
            documentId: "mock-document-id",
          }),
        }),
      );
    });

    it("should throw an error if sending metadata fails", async () => {
      (fetch as jest.Mock).mockResolvedValueOnce({
        ok: false,
        statusText: "Bad Request",
      });

      // @ts-ignore - accessing private method for testing
      await expect(
        client.sendMetadata("mock-document-id", "mock-cv", ["test"]),
      ).rejects.toThrow("Failed to create metadata: Bad Request");
    });
  });

  describe("uploadFile", () => {
    it("should upload a file successfully", async () => {
      // Mock the private methods
      const createObjectMock = jest
        .fn()
        .mockResolvedValue({ id: "mock-document-id", msCv: "mock-cv" });
      const uploadFileContentMock = jest.fn().mockResolvedValue(undefined);
      const sendMetadataMock = jest.fn().mockResolvedValue(undefined);

      // @ts-ignore - replacing private methods for testing
      client.createObject = createObjectMock;
      // @ts-ignore
      client.uploadFileContent = uploadFileContentMock;
      // @ts-ignore
      client.sendMetadata = sendMetadataMock;

      const mockFile: FileDetails = {
        name: "test.png",
        size: 1024,
        type: "image/png",
        base64: "dGVzdA==", // test in base64
      };

      await client.uploadFile(mockFile);

      expect(createObjectMock).toHaveBeenCalled();
      expect(uploadFileContentMock).toHaveBeenCalledWith(
        "mock-document-id",
        "mock-cv",
        mockFile,
      );
      expect(sendMetadataMock).toHaveBeenCalledWith(
        "mock-document-id",
        "mock-cv",
        ["test"],
      );
    });

    it("should handle errors during upload", async () => {
      const consoleSpy = jest
        .spyOn(console, "error")
        .mockImplementation(() => {});
      const error = new Error("Upload failed");
      // Mock createObject to throw an error
      // @ts-ignore - replacing private method for testing
      client.createObject = jest.fn().mockRejectedValue(error);

      const mockFile: FileDetails = {
        name: "test.png",
        size: 1024,
        type: "image/png",
        base64: "dGVzdA==",
      };

      await expect(client.uploadFile(mockFile)).rejects.toThrow(
        "Upload failed",
      );

      expect(consoleSpy).toHaveBeenCalledWith("Error uploading file:", error);
      consoleSpy.mockRestore();
    });
  });

  describe("uploadFiles", () => {
    it("should return an error if no files are provided", async () => {
      const result = await client.uploadFiles([]);
      expect(result).toEqual({
        success: false,
        error: "Please select files first",
      });
    });

    it("should reject a failed batch and stop uploading remaining files", async () => {
      const file: FileDetails = {
        name: "test.png",
        size: 4,
        type: "image/png",
        base64: "dGVzdA==",
      };
      jest.spyOn(client, "listEmojis").mockResolvedValue([]);
      const upload = jest
        .spyOn(client, "uploadFile")
        .mockResolvedValueOnce("uploaded-document")
        .mockRejectedValueOnce(new Error("Metadata rejected"));
      await expect(
        client.uploadFiles([
          file,
          { ...file, base64: "dGVzdDI=" },
          { ...file, base64: "dGVzdDM=" },
        ]),
      ).rejects.toThrow("Metadata rejected");
      expect(upload).toHaveBeenCalledTimes(2);
    });

    it("should upload multiple files successfully", async () => {
      jest.spyOn(client, "listEmojis").mockResolvedValue([]);
      const uploadFileSpy = jest
        .spyOn(client, "uploadFile")
        .mockResolvedValue("uploaded-document");

      const mockFiles: FileDetails[] = [
        {
          name: "test1.png",
          size: 1024,
          type: "image/png",
          base64: "dGVzdA==",
        },
        {
          name: "test2.png",
          size: 2048,
          type: "image/png",
          base64: "dGVzdDI=",
        },
      ];

      const result = await client.uploadFiles(mockFiles);

      expect(uploadFileSpy).toHaveBeenCalledTimes(2);
      expect(uploadFileSpy).toHaveBeenCalledWith(mockFiles[0], "test1");
      expect(uploadFileSpy).toHaveBeenCalledWith(mockFiles[1], "test2");
      expect(result).toEqual({
        success: true,
        counts: { uploaded: 2, skipped: 0, renamed: 0 },
        status:
          "Uploaded 2 emoji(s); skipped 0 identical image(s); renamed 0 name conflict(s).",
      });

      uploadFileSpy.mockRestore();
    });
  });
});

describe("existing emoji API checks", () => {
  let client: MsTeamsClient;
  beforeEach(() => {
    client = new MsTeamsClient("ic3", "chat", "tenant");
    (fetch as jest.Mock).mockReset();
  });

  it("reads full metadata without a continuation header", async () => {
    (fetch as jest.Mock).mockResolvedValue({
      ok: true,
      json: async () => ({
        continuationToken: "delta-token",
        categories: [
          {
            id: "customEmoji",
            emoticons: [
              {
                documentId: "doc",
                shortcuts: ["cat"],
                etag: "v1",
                isDeleted: false,
              },
            ],
          },
          { id: "standard", emoticons: [{ id: "ignore" }] },
        ],
      }),
    });
    expect(await client.listEmojis()).toEqual([
      { documentId: "doc", shortcuts: ["cat"], etag: "v1", isDeleted: false },
    ]);
    expect(fetch).toHaveBeenCalledWith(
      "https://teams.microsoft.com/api/csa/apac/api/v1/customemoji/metadata",
      {
        headers: { authorization: "Bearer chat" },
        cache: "no-store",
      },
    );
    expect(fetch).toHaveBeenCalledTimes(1);
  });

  it.each([
    {},
    { categories: [{ id: "customEmoji", emoticons: null }] },
    {
      categories: [
        {
          id: "customEmoji",
          emoticons: [{ documentId: "doc", shortcuts: "cat" }],
        },
      ],
    },
  ])(
    "rejects malformed metadata instead of assuming an empty library",
    async (data) => {
      (fetch as jest.Mock).mockResolvedValue({
        ok: true,
        json: async () => data,
      });
      await expect(client.listEmojis()).rejects.toThrow("Invalid Teams");
    },
  );

  it("fails closed on unauthorized metadata", async () => {
    (fetch as jest.Mock).mockResolvedValue({
      ok: false,
      statusText: "Unauthorized",
    });
    await expect(client.listEmojis()).rejects.toThrow(
      "Could not check existing emojis: Unauthorized",
    );
  });

  it("hashes downloaded original bytes and encodes document IDs", async () => {
    const bytes = new Uint8Array([116, 101, 115, 116]);
    (fetch as jest.Mock).mockResolvedValue({
      ok: true,
      arrayBuffer: async () => bytes.buffer,
      headers: { get: () => "image/png" },
    });
    expect(await client.existingImageHash("a/b")).toBe(
      "9f86d081884c7d659a2feaa0c55ad015a3bf4f1b2b0b822cd15d6c15b0f00a08",
    );
    expect(fetch).toHaveBeenCalledWith(
      "https://as-prod.asyncgw.teams.microsoft.com/v1/objects/a%2Fb/views/imgpsh_fullsize_anim",
      {
        headers: { authorization: "Bearer ic3" },
        credentials: "omit",
      },
    );
  });

  it("preserves GIF bytes from the animated full-size view", async () => {
    const bytes = new Uint8Array([71, 73, 70, 56, 57, 97]);
    (fetch as jest.Mock).mockResolvedValue({
      ok: true,
      status: 200,
      headers: { get: () => "image/gif" },
      arrayBuffer: async () => bytes.buffer,
    });
    expect(await client.readImage("gif")).toEqual({
      bytes,
      contentType: "image/gif",
    });
    expect((fetch as jest.Mock).mock.calls[0][0]).toContain(
      "/views/imgpsh_fullsize_anim",
    );
  });

  it("falls back to full-size for objects without an animated view", async () => {
    (fetch as jest.Mock)
      .mockResolvedValueOnce({ ok: false, status: 404 })
      .mockResolvedValueOnce({
        ok: true,
        status: 200,
        headers: { get: () => "image/png" },
        arrayBuffer: async () => new Uint8Array([1]).buffer,
      });
    expect((await client.readImage("static")).contentType).toBe("image/png");
    expect((fetch as jest.Mock).mock.calls[1][0]).toContain(
      "/views/imgpsh_fullsize",
    );
  });

  it("stops when original image content is inaccessible", async () => {
    (fetch as jest.Mock).mockResolvedValue({
      ok: false,
      statusText: "Forbidden",
    });
    await expect(client.existingImageHash("doc")).rejects.toThrow(
      "Could not check existing emoji image (Forbidden); upload stopped",
    );
  });

  it("reports the HTTP status when Teams omits status text", async () => {
    (fetch as jest.Mock).mockResolvedValue({
      ok: false,
      status: 404,
      statusText: "",
    });
    await expect(client.readImage("doc")).rejects.toThrow(
      "Could not check existing emoji image (404); upload stopped",
    );
  });

  it("rejects non-image responses even with HTTP 200", async () => {
    (fetch as jest.Mock).mockResolvedValue({
      ok: true,
      headers: { get: () => "application/json" },
    });
    await expect(client.existingImageHash("doc")).rejects.toThrow(
      "Teams did not return original image content",
    );
  });

  it("rejects explicit metadata failure even with HTTP 200", async () => {
    (fetch as jest.Mock).mockResolvedValue({
      ok: true,
      json: async () => ({ success: false }),
    });
    await expect(client.sendMetadata("doc", "cv", ["cat"])).rejects.toThrow(
      "Teams rejected emoji metadata",
    );
  });
});
