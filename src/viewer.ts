import TokenStore from "./tokenStore";
import { getExtensionApi } from "./extensionApi";
import {
  EmojiLibraryPage,
  EmojiPreview,
  EmojiDuplicates,
  TeamsTokens,
} from "./types";

export function renderLibrary(
  result: EmojiLibraryPage,
  grid: HTMLElement,
): void {
  grid.replaceChildren();
  for (const emoji of result.items) {
    const card = document.createElement("figure");
    card.className = "emoji-card";
    card.dataset.documentId = emoji.documentId;
    if (
      emoji.image &&
      /^data:(image\/(png|jpeg|gif|webp)|application\/octet-stream);base64,[A-Za-z0-9+/=]+$/.test(
        emoji.image,
      )
    ) {
      const image = document.createElement("img");
      image.src = emoji.image;
      image.alt = emoji.names[0] || "Custom emoji";
      image.addEventListener("error", () => {
        const message = document.createElement("div");
        message.className = "preview-error";
        message.textContent = "Preview unavailable";
        image.replaceWith(message);
      });
      card.append(image);
    } else if (emoji.loading) {
      const spinner = document.createElement("div");
      spinner.className = "preview-loading";
      spinner.setAttribute("role", "status");
      spinner.setAttribute("aria-label", "Loading preview");
      card.append(spinner);
    } else {
      const message = document.createElement("div");
      message.className = "preview-error";
      message.textContent = "Preview unavailable";
      message.title = emoji.error || "Unsupported image format";
      card.append(message);
    }
    const caption = document.createElement("figcaption");
    caption.textContent = emoji.names.join(", ") || "Unnamed emoji";
    card.append(caption);
    grid.append(card);
  }
}

export function renderDuplicates(
  result: EmojiDuplicates,
  container: HTMLElement,
): void {
  container.replaceChildren();
  const summary = document.createElement("p");
  summary.textContent = `${result.groups.length} duplicate group(s) · image hashes checked ${result.checked}/${result.total}${result.errors.length ? ` · ${result.errors.length} image(s) unchecked` : ""}.`;
  container.append(summary);
  const list = document.createElement("ul");
  for (const group of result.groups) {
    const item = document.createElement("li");
    item.textContent = `${group.kind === "name" ? `Same name (${group.value})` : "Same image hash"}: ${group.items.map((emoji) => `${emoji.names.join(", ") || "Unnamed"} [${emoji.documentId}]`).join(" · ")}`;
    list.append(item);
  }
  for (const error of result.errors) {
    const item = document.createElement("li");
    item.textContent = `Unchecked: ${error.names.join(", ")} — ${error.error}`;
    list.append(item);
  }
  container.append(list);
}

async function initializeViewer() {
  const select = document.getElementById("tenantSelect") as HTMLSelectElement;
  const search = document.getElementById("search") as HTMLInputElement;
  const grid = document.getElementById("grid")!;
  const status = document.getElementById("status")!;
  const previous = document.getElementById("previous") as HTMLButtonElement;
  const next = document.getElementById("next") as HTMLButtonElement;
  const searchButton = document.getElementById(
    "searchButton",
  ) as HTMLButtonElement;
  const refreshButton = document.getElementById(
    "refreshButton",
  ) as HTMLButtonElement;
  const duplicatesButton = document.getElementById(
    "duplicatesButton",
  ) as HTMLButtonElement;
  const duplicates = document.getElementById("duplicates")!;
  let sessions: TeamsTokens[] = [];
  let page = 0;
  let matched = 0;
  let pageSize = 24;
  let request = 0;
  let activeRequestId = "";
  const api = getExtensionApi();
  api.runtime.onMessage?.addListener((message: any) => {
    if (
      message.type !== "emojiPreviewUpdate" ||
      message.requestId !== activeRequestId
    )
      return;
    const update = message.update as EmojiLibraryPage | EmojiPreview;
    if ("items" in update) {
      renderLibrary(update, grid);
      status.textContent = `${update.matched} matching / ${update.total} tenant emojis · Loading previews…`;
    } else {
      const card = Array.from(grid.children).find(
        (element) =>
          (element as HTMLElement).dataset.documentId === update.documentId,
      );
      if (!card) return;
      const replacement = document.createElement("div");
      renderLibrary({ items: [update] } as EmojiLibraryPage, replacement);
      card.replaceWith(replacement.firstElementChild!);
    }
  });

  async function load(requestedPage: number) {
    const current = ++request;
    activeRequestId = crypto.randomUUID();
    const tokens = sessions.find(
      (session) => session.permissionsId === select.value,
    );
    grid.replaceChildren();
    duplicates.replaceChildren();
    status.textContent = "Loading emoji names…";
    [
      previous,
      next,
      searchButton,
      refreshButton,
      duplicatesButton,
      select,
    ].forEach((control) => (control.disabled = true));
    try {
      if (!tokens)
        throw new Error(
          "Open a signed-in Teams tab, then reopen the extension and its viewer.",
        );
      const response = (await getExtensionApi().runtime.sendMessage({
        action: "browseEmojis",
        requestId: activeRequestId,
        tokens,
        query: search.value.trim(),
        page: requestedPage,
      })) as EmojiLibraryPage | { success: false; error: string };
      if (current !== request) return;
      if (!response.success) throw new Error(response.error);
      if (response.tenantId !== tokens.permissionsId)
        throw new Error("Tenant changed; reopen the viewer.");
      page = response.page;
      pageSize = response.pageSize;
      matched = response.matched;
      if (!grid.childElementCount) renderLibrary(response, grid);
      document.getElementById("tenantNote")!.textContent =
        `Tenant: ${response.tenantId}`;
      document.getElementById("pageLabel")!.textContent =
        `Page ${page + 1} of ${Math.max(1, Math.ceil(matched / pageSize))}`;
      const unavailable = response.items.filter((emoji) => !emoji.image).length;
      status.textContent = `${response.matched} matching / ${response.total} tenant emojis${unavailable ? ` · ${unavailable} preview(s) unavailable` : ""}.`;
    } catch (error) {
      if (current !== request) return;
      matched = 0;
      status.textContent =
        error instanceof Error ? error.message : String(error);
    } finally {
      if (current === request) {
        [searchButton, refreshButton, duplicatesButton, select].forEach(
          (control) => (control.disabled = false),
        );
        previous.disabled = page === 0 || matched === 0;
        next.disabled = (page + 1) * pageSize >= matched;
      }
    }
  }

  duplicatesButton.addEventListener("click", async () => {
    const tokens = sessions.find(
      (session) => session.permissionsId === select.value,
    );
    if (!tokens) return;
    const current = request;
    const controls = [
      previous,
      next,
      searchButton,
      refreshButton,
      duplicatesButton,
      select,
    ];
    controls.forEach((control) => (control.disabled = true));
    duplicates.textContent =
      "Checking names and image hashes across this tenant…";
    try {
      const result = (await api.runtime.sendMessage({
        action: "checkEmojiDuplicates",
        tokens,
      })) as EmojiDuplicates | { success: false; error: string };
      if (current !== request) return;
      if (!result.success) throw new Error(result.error);
      if (result.tenantId !== select.value)
        throw new Error("Tenant changed; refresh the library.");
      renderDuplicates(result, duplicates);
    } catch (error) {
      if (current === request)
        duplicates.textContent =
          error instanceof Error ? error.message : String(error);
    } finally {
      if (current === request) {
        controls.forEach((control) => (control.disabled = false));
        previous.disabled = page === 0 || matched === 0;
        next.disabled = (page + 1) * pageSize >= matched;
      }
    }
  });

  document.getElementById("searchForm")!.addEventListener("submit", (event) => {
    event.preventDefault();
    void load(0);
  });
  refreshButton.addEventListener("click", () => void refreshSessions());
  select.addEventListener("change", () => void load(0));
  previous.addEventListener("click", () => void load(page - 1));
  next.addEventListener("click", () => void load(page + 1));
  async function refreshSessions() {
    activeRequestId = "";
    ++request;
    const previousTenant = select.value;
    select.replaceChildren();
    grid.replaceChildren();
    sessions = [];
    try {
      sessions = await TokenStore.refreshFromOpenTeamsTabs();
      for (const session of sessions) {
        const option = document.createElement("option");
        option.value = session.permissionsId!;
        option.textContent = session.permissionsId!;
        select.append(option);
      }
      if (sessions.some((session) => session.permissionsId === previousTenant))
        select.value = previousTenant;
      await load(0);
    } catch (error) {
      status.textContent =
        error instanceof Error ? error.message : String(error);
    }
  }
  await refreshSessions();
}

document.addEventListener("DOMContentLoaded", () => {
  if (document.getElementById("searchForm")) void initializeViewer();
});
