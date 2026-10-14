import {
  addSlackmojisButtons,
  emojiFromCard,
  importCardOnClick,
  initializeSlackmojis,
} from "../slackmojis";
import {
  downloadSlackmoji,
  isSlackmojisPage,
  validateSlackmojisImage,
} from "../slackmojisImport";

const url =
  "https://emojis.slackmojis.com/emojis/images/123/7500/partyparrot.gif?123";
const cardMarkup = `<li class="emoji"><a class="downloader" download="partyparrot.gif" href="/emojis/7500-partyparrot/download"><div class="wrapper"><img src="${url}"></div><div class="name">:partyparrot:</div></a></li>`;

beforeEach(() => {
  document.body.innerHTML = `<ul class="emojis">${cardMarkup}</ul>`;
  globalThis.fetch = jest.fn();
});

test("recognizes live Slackmojis markup and keeps original GIF URL", () => {
  expect(emojiFromCard(document.querySelector("li")!)).toEqual({
    name: "partyparrot.gif",
    imageUrl: url,
  });
});

test("adds one shortcut per card, including newly appended search results", () => {
  const add = jest.fn();
  addSlackmojisButtons(document, add);
  addSlackmojisButtons(document, add);
  expect(document.querySelectorAll(".teams-slackmojis-button")).toHaveLength(1);
  const newCard = document.createElement("li");
  newCard.className = "emoji";
  newCard.innerHTML = document.querySelector("a")!.outerHTML;
  document.querySelector("ul")!.append(newCard);
  addSlackmojisButtons(newCard, add);
  expect(newCard.querySelectorAll(".teams-slackmojis-button")).toHaveLength(1);
});

test("click imports without triggering Slackmojis download; repeated clicks wait", async () => {
  let finish!: (value: {
    success: boolean;
    status: string;
    counts?: { uploaded: number; skipped: number; renamed: number };
  }) => void;
  const add = jest.fn(
    () =>
      new Promise<{ success: boolean; status: string }>((resolve) => {
        finish = resolve;
      }),
  );
  addSlackmojisButtons(document, add);
  const button = document.querySelector("button")!;
  const bubble = jest.fn();
  document.querySelector("li")!.addEventListener("click", bubble);
  const event = {
    isTrusted: true,
    preventDefault: jest.fn(),
    stopPropagation: jest.fn(),
  } as unknown as Event;
  void importCardOnClick(event, button, document.querySelector("li")!, add);
  void importCardOnClick(event, button, document.querySelector("li")!, add);
  expect(add).toHaveBeenCalledTimes(1);
  expect(bubble).not.toHaveBeenCalled();
  expect(button.disabled).toBe(true);
  finish({
    success: true,
    status: "Skipped existing image",
    counts: { uploaded: 0, skipped: 1, renamed: 0 },
  });
  await Promise.resolve();
  expect(button.textContent).toBe("=");
  expect(button.title).toBe("Duplicate: Skipped existing image");
  expect(button.disabled).toBe(true);
  await importCardOnClick(event, button, document.querySelector("li")!, add);
  expect(add).toHaveBeenCalledTimes(1);
});

test("failure locks shortcut until reload with error detail", async () => {
  const add = async () => ({ success: false, error: "Session expired" });
  addSlackmojisButtons(document, add);
  const button = document.querySelector("button")!;
  await importCardOnClick(
    {
      isTrusted: true,
      preventDefault: jest.fn(),
      stopPropagation: jest.fn(),
    } as unknown as Event,
    button,
    document.querySelector("li")!,
    add,
  );
  await Promise.resolve();
  expect(button.textContent).toBe("!");
  expect(button.title).toBe("Session expired");
  expect(button.disabled).toBe(true);
});

test("skips unrelated images and rejects untrusted request origins", () => {
  document.querySelector("img")!.src = "https://evil.example/image.png";
  addSlackmojisButtons(document, jest.fn());
  expect(document.querySelector("button")).toBeNull();
  expect(
    isSlackmojisPage("https://slackmojis.com/emojis/7500-partyparrot"),
  ).toBe(true);
  for (const origin of [
    undefined,
    "https://slackmojis.com.evil.example",
    "http://slackmojis.com",
    "https://emojis.slackmojis.com",
  ])
    expect(isSlackmojisPage(origin)).toBe(false);
});

test.each([
  "https://evil.example/emojis/images/x.gif",
  "https://emojis.slackmojis.com.evil.example/emojis/images/x.gif",
  "http://emojis.slackmojis.com/emojis/images/x.gif",
  "https://user:pass@emojis.slackmojis.com/emojis/images/x.gif",
  "https://emojis.slackmojis.com/other/x.gif",
  "https://emojis.slackmojis.com/emojis/images/x.svg",
])("rejects unsafe image URL %s", (imageUrl) => {
  expect(() => validateSlackmojisImage(imageUrl, "parrot.gif")).toThrow();
});

test("rejects path-like filenames", () => {
  expect(() => validateSlackmojisImage(url, "../image.gif")).toThrow();
});

function response(chunks: Uint8Array[], type = "image/gif", length?: string) {
  let index = 0;
  const cancel = jest.fn().mockResolvedValue(undefined);
  return {
    ok: true,
    headers: {
      get: (name: string) => (name === "content-type" ? type : length || null),
    },
    body: {
      getReader: () => ({
        read: async () =>
          index < chunks.length
            ? { done: false, value: chunks[index++] }
            : { done: true },
        cancel,
      }),
    },
    cancel,
  };
}

test("downloads original bytes without cookies or Teams authorization", async () => {
  const image = response([
    new Uint8Array([71, 73, 70]),
    new Uint8Array([56, 57, 97]),
  ]);
  (fetch as jest.Mock).mockResolvedValue(image);
  expect(await downloadSlackmoji(url, "partyparrot.gif")).toEqual({
    name: "partyparrot.gif",
    size: 6,
    type: "image/gif",
    base64: "R0lGODlh",
  });
  expect(fetch).toHaveBeenCalledWith(url, {
    credentials: "omit",
    redirect: "error",
  });
  expect(image.cancel).toHaveBeenCalledTimes(1);
});

test("rejects HTML responses instead of uploading them", async () => {
  (fetch as jest.Mock).mockResolvedValue(response([], "text/html"));
  await expect(downloadSlackmoji(url, "partyparrot.gif")).rejects.toThrow(
    "supported image",
  );
});

test("rejects oversized downloads even without Content-Length", async () => {
  const image = response([
    new Uint8Array(10 * 1024 * 1024),
    new Uint8Array([1]),
  ]);
  (fetch as jest.Mock).mockResolvedValue(image);
  await expect(downloadSlackmoji(url, "partyparrot.gif")).rejects.toThrow(
    "exceeds 10 MB",
  );
  expect(image.cancel).toHaveBeenCalledTimes(1);
});

test("synthetic page-script clicks cannot trigger uploads", () => {
  const add = jest.fn();
  addSlackmojisButtons(document, add);
  document.querySelector("button")!.click();
  expect(add).not.toHaveBeenCalled();
});

test("integration selects single tenant and decorates dynamic search results", async () => {
  const sendMessage = jest
    .fn()
    .mockResolvedValue({ success: true, tenantIds: ["tenant-one"] });
  Object.defineProperty(globalThis, "chrome", {
    configurable: true,
    value: { runtime: { sendMessage } },
  });
  initializeSlackmojis();
  initializeSlackmojis();
  await new Promise((resolve) => setTimeout(resolve, 0));
  expect(document.querySelectorAll("#teams-slackmojis-panel")).toHaveLength(1);
  expect(document.querySelector("select")!.value).toBe("tenant-one");
  expect(sendMessage).toHaveBeenCalledWith({ action: "slackmojisTargets" });
  expect(fetch).not.toHaveBeenCalled();
  document.querySelector("ul")!.insertAdjacentHTML("beforeend", cardMarkup);
  await new Promise((resolve) => setTimeout(resolve, 0));
  expect(document.querySelectorAll(".teams-slackmojis-button")).toHaveLength(2);
});

test.each([
  [{ uploaded: 1, skipped: 0, renamed: 0 }, "✓", "Added"],
  [{ uploaded: 1, skipped: 0, renamed: 1 }, "↪", "Renamed"],
] as const)(
  "completed import shows compact result and remains locked",
  async (counts, icon, label) => {
    const add = jest.fn(async () => ({
      success: true,
      counts: { ...counts },
      status: "Result detail",
    }));
    addSlackmojisButtons(document, add);
    const button = document.querySelector("button")!;
    const event = {
      isTrusted: true,
      preventDefault: jest.fn(),
      stopPropagation: jest.fn(),
    } as unknown as Event;
    await importCardOnClick(event, button, document.querySelector("li")!, add);
    await importCardOnClick(event, button, document.querySelector("li")!, add);
    expect(button.textContent).toBe(icon);
    expect(button.getAttribute("aria-label")).toContain(label);
    expect(button.disabled).toBe(true);
    expect(add).toHaveBeenCalledTimes(1);
  },
);
