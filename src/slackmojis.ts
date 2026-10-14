import { getExtensionApi } from "./extensionApi";
import { ProcessResult } from "./types";

export interface Slackmoji {
  name: string;
  imageUrl: string;
}

export function emojiFromCard(card: Element): Slackmoji | null {
  const link = card.querySelector<HTMLAnchorElement>("a.downloader[download]");
  const image = link?.querySelector<HTMLImageElement>("img");
  const name = link?.getAttribute("download");
  const imageUrl = image?.getAttribute("src");
  if (!name || !imageUrl) return null;
  try {
    const url = new URL(imageUrl, "https://slackmojis.com");
    if (
      url.origin !== "https://emojis.slackmojis.com" ||
      !url.pathname.startsWith("/emojis/images/")
    )
      return null;
    return { name, imageUrl: url.href };
  } catch {
    return null;
  }
}

export function addSlackmojisButtons(
  root: ParentNode,
  add: (emoji: Slackmoji) => Promise<ProcessResult>,
): void {
  const cards = [
    ...(root instanceof Element && root.matches("li.emoji") ? [root] : []),
    ...Array.from(root.querySelectorAll("li.emoji")),
  ];
  for (const card of cards) {
    if (card.querySelector(".teams-slackmojis-button") || !emojiFromCard(card))
      continue;
    const button = document.createElement("button");
    button.type = "button";
    button.className = "teams-slackmojis-button";
    button.textContent = "＋";
    button.setAttribute("aria-label", "Add to Teams");
    button.title = "Import this emoji into the selected Teams tenant";
    button.addEventListener(
      "click",
      (event) => void importCardOnClick(event, button, card, add),
    );
    card.append(button);
  }
}

export async function importCardOnClick(
  event: Event,
  button: HTMLButtonElement,
  card: Element,
  add: (emoji: Slackmoji) => Promise<ProcessResult>,
): Promise<void> {
  // Page scripts may synthesize DOM clicks; only a user gesture can trigger an import.
  if (!event.isTrusted) return;
  event.preventDefault();
  event.stopPropagation();
  if (button.disabled) return;
  const emoji = emojiFromCard(card);
  if (!emoji) return;
  button.disabled = true;
  button.textContent = "";
  button.dataset.state = "loading";
  button.setAttribute("aria-label", "Adding to Teams");
  try {
    const result = await add(emoji);
    if (!result.success) throw new Error(result.error || "Teams import failed");
    const outcome = result.counts?.skipped
      ? "Duplicate"
      : result.counts?.renamed
        ? "Renamed"
        : "Added";
    button.dataset.state = outcome.toLowerCase();
    button.textContent =
      outcome === "Duplicate" ? "=" : outcome === "Renamed" ? "↪" : "✓";
    button.title = `${outcome}: ${result.status || "Import complete"}`;
    button.setAttribute("aria-label", button.title);
  } catch (error) {
    button.dataset.state = "error";
    button.textContent = "!";
    button.title = error instanceof Error ? error.message : String(error);
    button.setAttribute(
      "aria-label",
      `Import failed: ${button.title}. Reload page to retry.`,
    );
  }
}

export function initializeSlackmojis(): void {
  if (document.getElementById("teams-slackmojis-panel")) return;
  const panel = document.createElement("aside");
  panel.id = "teams-slackmojis-panel";
  panel.className = "teams-slackmojis-panel";
  const label = document.createElement("label");
  label.textContent = "Add emojis to Teams ";
  const select = document.createElement("select");
  select.setAttribute("aria-label", "Teams tenant for emoji imports");
  label.append(select);
  const refresh = document.createElement("button");
  refresh.type = "button";
  refresh.textContent = "Refresh tenants";
  const status = document.createElement("span");
  status.className = "teams-slackmojis-status";
  status.setAttribute("role", "status");
  status.setAttribute("aria-live", "polite");
  panel.append(label, refresh, status);
  document.body.prepend(panel);

  async function loadTargets() {
    refresh.disabled = true;
    try {
      const response = await getExtensionApi().runtime.sendMessage({
        action: "slackmojisTargets",
      });
      if (!response?.success || !Array.isArray(response.tenantIds))
        throw new Error(response?.error || "Could not load Teams tenants");
      const previous = select.value;
      const ids = response.tenantIds.filter(
        (id: unknown): id is string => typeof id === "string" && !!id,
      );
      select.replaceChildren();
      const placeholder = document.createElement("option");
      placeholder.value = "";
      placeholder.textContent = ids.length
        ? "Choose Teams tenant"
        : "No Teams session";
      select.append(placeholder);
      for (const id of ids) {
        const option = document.createElement("option");
        option.value = id;
        option.textContent = id;
        select.append(option);
      }
      select.value = ids.includes(previous)
        ? previous
        : ids.length === 1
          ? ids[0]
          : "";
      status.textContent = ids.length
        ? "Use the ＋ icon on an emoji. ✓ added · ↪ renamed · = duplicate."
        : "Open the extension on a signed-in Teams tab, then refresh tenants.";
    } catch (error) {
      select.replaceChildren();
      status.textContent =
        error instanceof Error ? error.message : String(error);
    } finally {
      refresh.disabled = false;
    }
  }
  refresh.addEventListener("click", () => void loadTargets());
  const add = async (emoji: Slackmoji): Promise<ProcessResult> => {
    const tenantId = select.value;
    try {
      if (!tenantId)
        throw new Error(
          "Choose a Teams tenant first. Open the extension on a signed-in Teams tab if none are listed.",
        );
      status.textContent = `Adding ${emoji.name} to tenant ${tenantId}…`;
      const result = (await getExtensionApi().runtime.sendMessage({
        action: "importSlackmoji",
        tenantId,
        ...emoji,
      })) as ProcessResult;
      if (!result?.success)
        throw new Error(result?.error || "Teams import failed");
      status.textContent = `${emoji.name}: ${result.status || "Import complete"} (Tenant ${tenantId})`;
      return result;
    } catch (error) {
      status.textContent =
        error instanceof Error ? error.message : String(error);
      throw error;
    }
  };
  addSlackmojisButtons(document, add);
  // Search/category results replace or append cards dynamically.
  new MutationObserver((changes) => {
    for (const change of changes)
      for (const node of Array.from(change.addedNodes)) {
        if (node instanceof Element) addSlackmojisButtons(node, add);
      }
  }).observe(document.body, { childList: true, subtree: true });
  void loadTargets();
}

if (location.origin === "https://slackmojis.com") initializeSlackmojis();
