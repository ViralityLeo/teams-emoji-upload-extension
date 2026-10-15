import { showUploader } from "../uploaderModal";

it("keeps modal on blur, dismisses on outside click or Escape", () => {
  const attachShadow = jest.spyOn(Element.prototype, "attachShadow");
  showUploader("chrome-extension://test/popup.html");
  const host = document.getElementById("teams-emoji-uploader-modal")!;
  const shadow = attachShadow.mock.results[0].value as ShadowRoot;
  window.dispatchEvent(new Event("blur"));
  document.dispatchEvent(new Event("visibilitychange"));
  shadow.querySelector("iframe")!.click();
  expect(host.isConnected).toBe(true);
  showUploader("chrome-extension://test/popup.html");
  expect(attachShadow).toHaveBeenCalledTimes(1);
  (shadow.firstElementChild as HTMLElement).click();
  expect(host.isConnected).toBe(false);
  showUploader("chrome-extension://test/popup.html");
  document.dispatchEvent(new KeyboardEvent("keydown", { key: "Escape" }));
  expect(document.getElementById("teams-emoji-uploader-modal")).toBeNull();
  attachShadow.mockRestore();
});
