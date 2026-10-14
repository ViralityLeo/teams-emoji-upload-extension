import { getExtensionApi, isFirefox } from "../extensionApi";

const globals = globalThis as typeof globalThis & { browser?: typeof chrome };

afterEach(() => {
  delete globals.browser;
});

test("uses Firefox Promise API when available", () => {
  const firefox = {
    runtime: { getBrowserInfo: jest.fn() },
  } as unknown as typeof chrome;
  globals.browser = firefox;
  expect(getExtensionApi()).toBe(firefox);
  expect(isFirefox()).toBe(true);
});

test("falls back to Chrome API", () => {
  const chromeApi = { runtime: {} } as typeof chrome;
  Object.defineProperty(globalThis, "chrome", {
    value: chromeApi,
    configurable: true,
  });
  expect(getExtensionApi()).toBe(chromeApi);
  expect(isFirefox()).toBe(false);
});
