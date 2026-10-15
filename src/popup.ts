import { clearSessionStorage } from "./emojiDedup";
import { getExtensionApi, isFirefox } from "./extensionApi";
import TokenStore from "./tokenStore";
import { ProcessResult, FileDetails } from "./types";

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") window.parent.postMessage("closeUploader", "*");
});

document.addEventListener("DOMContentLoaded", async () => {
  try {
    // Check for existing processing state
    const state = await getExtensionApi().storage.local.get<{
      processingState?: { status: string; type: string };
    }>("processingState");
    if (state.processingState) {
      updateStatus(
        state.processingState.status,
        state.processingState.type as
          | "ready"
          | "success"
          | "error"
          | "processing",
      );
    }

    const sessions = await TokenStore.refreshFromOpenTeamsTabs();
    updateStatus(
      sessions.length
        ? `Connected to ${sessions.length} Teams tenant(s).`
        : "Open a signed-in Teams tab, then reopen the extension.",
      sessions.length ? "ready" : "error",
    );
  } catch (err) {
    updateStatus(err instanceof Error ? err.message : String(err), "error");
  } finally {
    (
      document.getElementById("viewButton") as HTMLButtonElement | null
    )?.removeAttribute("disabled");
  }
});

function updateStatus(
  message: string,
  type: "ready" | "success" | "error" | "processing" = "ready",
) {
  const statusDiv = document.getElementById("status")!;
  const statusText = document.getElementById("statusText")!;
  const statusIcon = statusDiv.querySelector("i")!;

  statusDiv.className = "status"; // Reset class
  const statusClasses: { [key: string]: [string, string] } = {
    success: ["status-success", "fas fa-check-circle"],
    error: ["status-error", "fas fa-exclamation-circle"],
    processing: ["status-processing", "fas fa-spinner fa-spin"],
    ready: ["status-ready", "fas fa-circle-info"],
  };

  statusDiv.classList.add(statusClasses[type][0]);
  statusIcon.className = statusClasses[type][1];
  statusText.textContent = message;
}

async function handleFileProcessing() {
  try {
    const fileInput = document.getElementById("fileInput") as HTMLInputElement;
    const files = await Promise.all(
      Array.from(fileInput.files || []).map(async (file) => {
        const buffer = await file.arrayBuffer();
        const base64 = btoa(
          new Uint8Array(buffer).reduce(
            (data, byte) => data + String.fromCharCode(byte),
            "",
          ),
        );
        return {
          name: file.name,
          size: file.size,
          type: file.type,
          base64: base64,
        } as FileDetails;
      }),
    );
    if (files.length === 0) {
      updateStatus("Please select files first.", "error");
      return;
    }

    const filesJson = JSON.stringify(files);

    updateStatus("Processing...", "processing");

    const sessions = await TokenStore.refreshFromOpenTeamsTabs();
    if (sessions.length !== 1)
      throw new Error(
        sessions.length
          ? "Multiple Teams tenants found. Open the intended Teams account alone before uploading local files."
          : "Open a signed-in Teams tab first.",
      );
    const tokens = sessions[0];

    const result = await getExtensionApi().runtime.sendMessage({
      action: "processFiles",
      files: filesJson,
      tokens: tokens,
    });

    const { success, error, status } = result as ProcessResult;
    if (error) {
      updateStatus(error, "error");
    } else {
      updateStatus(status || "", success ? "success" : "processing");

      await clearTeamsData();

      await clearSessionStorage();

      await reloadTeamsTab();
    }
  } catch (error) {
    let errorMessage = "";

    if (error instanceof Error) {
      errorMessage = error.message;
    } else if (typeof error === "string") {
      try {
        const parsed = JSON.parse(error);
        errorMessage = JSON.stringify(parsed, null, 2);
      } catch {
        errorMessage = error;
      }
    } else {
      errorMessage = String(error);
    }

    updateStatus(errorMessage, "error");
  }
}

function updateSelectedFiles() {
  const fileInput = document.getElementById("fileInput") as HTMLInputElement;
  const selectedFilesDiv = document.getElementById("selectedFiles")!;
  if (fileInput.files && fileInput.files.length > 0) {
    const fileList = Array.from(fileInput.files)
      .map((file) => file.name)
      .join(", ");
    selectedFilesDiv.textContent = fileList;
  } else {
    selectedFilesDiv.textContent = "No files selected";
  }
}

document
  .getElementById("fileInput")
  ?.addEventListener("change", updateSelectedFiles);
document
  .getElementById("processButton")
  ?.addEventListener("click", () => handleFileProcessing());

async function clearTeamsData() {
  // Firefox refresh only: its cache removal cannot be scoped to Teams.
  if (isFirefox()) return;
  await getExtensionApi().browsingData.remove(
    {
      origins: ["https://teams.microsoft.com", "https://teams.cloud.microsoft"],
    },
    {
      cacheStorage: true,
      cookies: true,
      localStorage: true,
      serviceWorkers: true,
      indexedDB: true,
      cache: true,
      appcache: true,
    },
  );
}

async function refreshTeams() {
  try {
    updateStatus("Refreshing Teams...", "processing");

    await clearTeamsData();

    await clearSessionStorage();

    await reloadTeamsTab();
    updateStatus("Teams refreshed successfully!", "success");
  } catch (error) {
    updateStatus(`Refresh failed: ${error}`, "error");
  }
}

async function reloadTeamsTab() {
  const tabs = await getExtensionApi().tabs.query({
    url: ["https://teams.microsoft.com/*", "https://teams.cloud.microsoft/*"],
  });
  const sourceTabId = Number(new URLSearchParams(location.search).get("tabId"));
  const tab = tabs.find((tab) => tab.id === sourceTabId) || tabs[0];
  if (tab?.id !== undefined) await getExtensionApi().tabs.reload(tab.id);
}

document.getElementById("resetButton")?.addEventListener("click", refreshTeams);

getExtensionApi().runtime.onMessage.addListener(
  (message: { type: string; error: any; status: any; success: any }) => {
    if (message.type === "processUpdate") {
      const status = message.error || message.status;
      const type = message.success
        ? "success"
        : message.error
          ? "error"
          : "processing";

      // Update the UI
      updateStatus(status, type as "success" | "error" | "processing");

      // Ensure we store the state in case it was sent directly from background.ts
      // Background should be storing the state, but this is a good fallback
      getExtensionApi().storage.local.set({
        processingState: {
          status: status,
          type: type,
        },
      });
    }
  },
);

document.getElementById("viewButton")?.addEventListener("click", async () => {
  try {
    await getExtensionApi().tabs.create({
      url: getExtensionApi().runtime.getURL("viewer.html"),
    });
  } catch (error) {
    updateStatus(`Could not open emoji viewer: ${error}`, "error");
  }
});
