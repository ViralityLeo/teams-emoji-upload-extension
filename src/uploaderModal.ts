// Self-contained: executeScript serializes this function into the active page.
export function showUploader(url: string) {
  const id = "teams-emoji-uploader-modal";
  if (document.getElementById(id)) return;
  const previousFocus = document.activeElement as HTMLElement | null;
  const host = document.createElement("div");
  host.id = id;
  host.style.cssText = "all:initial;position:fixed;inset:0;z-index:2147483647";
  const shadow = host.attachShadow({ mode: "closed" });
  const backdrop = document.createElement("div");
  backdrop.style.cssText =
    "position:absolute;inset:0;background:#0006;display:flex;align-items:flex-start;justify-content:flex-end;padding:16px";
  const panel = document.createElement("div");
  panel.style.cssText =
    "position:relative;max-width:100%;height:min(560px,90vh)";
  panel.setAttribute("role", "dialog");
  panel.setAttribute("aria-modal", "true");
  panel.setAttribute("aria-label", "Teams Emoji Uploader");
  const frame = document.createElement("iframe");
  frame.src = url;
  frame.title = "Teams Emoji Uploader";
  frame.style.cssText =
    "width:352px;max-width:100%;height:100%;border:0;border-radius:8px;background:#292929";
  const close = document.createElement("button");
  close.textContent = "×";
  close.setAttribute("aria-label", "Close uploader");
  close.style.cssText =
    "position:absolute;right:4px;top:4px;border:0;background:#292929;color:white;cursor:pointer;font-size:20px";
  function dismiss() {
    host.remove();
    window.removeEventListener("message", onMessage);
    document.removeEventListener("keydown", onKeydown, true);
    previousFocus?.focus();
  }
  function onMessage(event: MessageEvent) {
    if (event.source === frame.contentWindow && event.data === "closeUploader")
      dismiss();
  }
  function onKeydown(event: KeyboardEvent) {
    if (event.key === "Escape") dismiss();
  }
  backdrop.addEventListener("click", (event) => {
    if (event.target === backdrop) dismiss();
  });
  close.addEventListener("click", dismiss);
  window.addEventListener("message", onMessage);
  document.addEventListener("keydown", onKeydown, true);
  panel.append(frame, close);
  backdrop.append(panel);
  shadow.append(backdrop);
  document.documentElement.append(host);
  close.focus();
}
