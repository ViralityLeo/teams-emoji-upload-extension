// Firefox exposes Promise-based APIs through browser; Chrome keeps chrome.
export function getExtensionApi(): typeof chrome {
  return (
    (globalThis as typeof globalThis & { browser?: typeof chrome }).browser ||
    chrome
  );
}

export function isFirefox(): boolean {
  const api = getExtensionApi();
  return "getBrowserInfo" in api.runtime;
}
