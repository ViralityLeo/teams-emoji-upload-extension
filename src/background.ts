import { showUploader } from "./uploaderModal";
import {
  downloadSlackmoji,
  isSlackmojisPage,
  slackmojisSessions,
} from "./slackmojisImport";
import { getExtensionApi } from "./extensionApi";
import MsTeamsClient from "./msTeams";
import { ProcessResult, FileDetails, TeamsTokens } from "./types";

getExtensionApi().action.onClicked.addListener((tab) => {
  if (tab.id === undefined) return;
  void getExtensionApi().scripting.executeScript({
    target: { tabId: tab.id },
    func: showUploader,
    args: [getExtensionApi().runtime.getURL(`popup.html?tabId=${tab.id}`)],
  });
});

getExtensionApi().runtime.onMessage.addListener(
  (message, sender, sendResponse) => {
    if (
      message.action === "slackmojisTargets" ||
      message.action === "importSlackmoji"
    ) {
      if (!isSlackmojisPage(sender.url)) {
        sendResponse({ success: false, error: "Slackmojis request rejected" });
        return;
      }
      (message.action === "slackmojisTargets"
        ? slackmojisSessions().then((sessions) => ({
            success: true,
            tenantIds: [
              ...new Set(sessions.map((session) => session.permissionsId!)),
            ],
          }))
        : handleSlackmojisImport(
            message.tenantId,
            message.imageUrl,
            message.name,
          )
      )
        .then(sendResponse)
        .catch((error) =>
          sendResponse({ success: false, error: formatErrorMessage(error) }),
        );
      return true;
    }
    if (message.action === "checkEmojiDuplicates") {
      enqueue(async () => {
        const tokens = message.tokens as TeamsTokens;
        if (!tokens?.ic3 || !tokens.chatsvcagg || !tokens.permissionsId)
          throw new Error("Open a signed-in Teams tab, then refresh tenants.");
        return new MsTeamsClient(
          tokens.ic3,
          tokens.chatsvcagg,
          tokens.permissionsId,
        ).checkDuplicates();
      })
        .then(sendResponse)
        .catch((error) =>
          sendResponse({ success: false, error: formatErrorMessage(error) }),
        );
      return true;
    }
    if (message.action === "browseEmojis") {
      handleBrowseEmojis(
        message.tokens,
        message.query,
        message.page,
        message.requestId,
      )
        .then(sendResponse)
        .catch((error) =>
          sendResponse({ success: false, error: formatErrorMessage(error) }),
        );
      return true;
    }
    if (message.action === "processFiles") {
      const files = JSON.parse(message.files) as FileDetails[];
      // Set initial processing state
      getExtensionApi().storage.local.set({
        processingState: { status: "Processing...", type: "processing" },
      });
      handleFileProcessing(files, message.tokens)
        .then(sendResponse)
        .catch((error) => {
          const errorMessage = formatErrorMessage(error);
          // Store error state
          getExtensionApi().storage.local.set({
            processingState: { status: errorMessage, type: "error" },
          });
          sendResponse({ success: false, error: errorMessage });
        });
      return true;
    }
  },
);

export function formatErrorMessage(error: any): string {
  // If it's a standard error object
  if (error instanceof Error) {
    return error.message;
  }

  // If it's a string that might be JSON
  if (typeof error === "string") {
    try {
      const parsed = JSON.parse(error);
      return JSON.stringify(parsed, null, 2);
    } catch {
      // Not JSON, return as is
      return error;
    }
  }

  // If it's already an object
  if (typeof error === "object") {
    try {
      return JSON.stringify(error, null, 2);
    } catch {
      return String(error);
    }
  }

  // Fallback
  return String(error);
}

// Serialize batches in this background context so two popups cannot race each other.
let processingQueue: Promise<void> = Promise.resolve();
export function handleFileProcessing(
  files: FileDetails[],
  tokens: Parameters<typeof processFiles>[1],
): Promise<ProcessResult> {
  return enqueue(() => processFiles(files, tokens));
}

function enqueue<T>(work: () => Promise<T>): Promise<T> {
  const task = processingQueue.then(work);
  processingQueue = task.then(
    () => undefined,
    () => undefined,
  );
  return task;
}

export function handleBrowseEmojis(
  tokens: TeamsTokens,
  query = "",
  page = 0,
  requestId?: string,
) {
  return enqueue(async () => {
    if (!tokens?.ic3 || !tokens.chatsvcagg || !tokens.permissionsId)
      throw new Error("Open a signed-in Teams tab, then reopen the extension.");
    if (
      typeof query !== "string" ||
      query.length > 200 ||
      !Number.isInteger(page) ||
      page < 0
    )
      throw new Error("Invalid emoji viewer request");
    return new MsTeamsClient(
      tokens.ic3,
      tokens.chatsvcagg,
      tokens.permissionsId,
    ).browseEmojis(
      query,
      page,
      requestId
        ? (update) => {
            void getExtensionApi()
              .runtime.sendMessage({
                type: "emojiPreviewUpdate",
                requestId,
                update,
              })
              .catch(() => undefined);
          }
        : undefined,
    );
  });
}

export async function handleSlackmojisImport(
  tenantId: string,
  imageUrl: string,
  name: string,
): Promise<ProcessResult> {
  const tokens = (await slackmojisSessions()).find(
    (session) => session.permissionsId === tenantId,
  );
  if (!tokens)
    throw new Error(
      "Open the extension on a signed-in Teams tab, then refresh Teams tenants here.",
    );
  const file = await downloadSlackmoji(imageUrl, name);
  return handleFileProcessing([file], tokens);
}

async function processFiles(
  files: FileDetails[],
  tokens: {
    chatsvcagg: string | null;
    ic3: string | null;
    permissionsId: string | null;
  },
): Promise<ProcessResult> {
  try {
    if (!tokens.chatsvcagg || !tokens.ic3 || !tokens.permissionsId) {
      const errorMsg = "Could not find required tokens";
      getExtensionApi().storage.local.set({
        processingState: { status: errorMsg, type: "error" },
      });
      throw new Error(errorMsg);
    }
    const teams = new MsTeamsClient(
      tokens.ic3,
      tokens.chatsvcagg,
      tokens.permissionsId,
    );
    const result = await teams.uploadFiles(files);

    // Store processing result state
    getExtensionApi().storage.local.set({
      processingState: {
        status: result.status || "",
        type: result.success ? "success" : "error",
      },
    });

    getExtensionApi().runtime.sendMessage({
      type: "processUpdate",
      success: result.success,
      status: result.status,
      error: result.error,
    });

    return result;
  } catch (error) {
    const errorMessage = formatErrorMessage(error);
    const errorResult = { success: false, error: errorMessage };

    // Store error state
    getExtensionApi().storage.local.set({
      processingState: { status: errorMessage, type: "error" },
    });

    getExtensionApi().runtime.sendMessage({
      type: "processUpdate",
      ...errorResult,
    });

    return errorResult;
  }
}
