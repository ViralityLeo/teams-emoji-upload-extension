import { webcrypto } from "node:crypto";
import MsTeamsClient from "../msTeams";
import {
  CACHE_PREFIX,
  clearSessionStorage,
  fileBytes,
  imageHash,
  loadCache,
  saveCache,
  uniqueName,
  ExistingEmoji,
} from "../emojiDedup";
import { FileDetails } from "../types";

Object.defineProperty(globalThis, "crypto", { value: webcrypto });
let stored: Record<string, unknown>;
const storage = {
  get: jest.fn(async (key: string | null) =>
    key === null ? stored : { [key]: stored[key] },
  ),
  set: jest.fn(async (items) => {
    Object.assign(stored, JSON.parse(JSON.stringify(items)));
  }),
  remove: jest.fn(async (keys: string[]) => {
    keys.forEach((key) => delete stored[key]);
  }),
};
Object.defineProperty(globalThis, "chrome", {
  configurable: true,
  value: { storage: { local: storage } },
});
const file = (name: string, text: string): FileDetails => ({
  name,
  size: text.length,
  type: "image/png",
  base64: btoa(text),
});
const emoji = (
  documentId: string,
  name: string,
  etag = "v1",
  isDeleted = false,
): ExistingEmoji => ({ documentId, shortcuts: [name], etag, isDeleted });
let client: MsTeamsClient;
let listing: jest.SpyInstance;
let download: jest.SpyInstance;
let upload: jest.SpyInstance;

beforeEach(() => {
  stored = {};
  jest.clearAllMocks();
  client = new MsTeamsClient("ic3", "chat", "tenant");
  listing = jest.spyOn(client, "listEmojis").mockResolvedValue([]);
  download = jest
    .spyOn(client, "existingImageHash")
    .mockImplementation(async (id) => imageHash(new TextEncoder().encode(id)));
  upload = jest
    .spyOn(client, "uploadFile")
    .mockImplementation(async (_, name) => `uploaded-${name}`);
});

// Node's TextEncoder is needed in jsdom; hashing still uses the actual Web Crypto API.
import { TextEncoder } from "node:util";

test("SHA-256 uses decoded bytes and rejects invalid base64", async () => {
  expect(await imageHash(fileBytes(file("x.png", "test")))).toBe(
    "9f86d081884c7d659a2feaa0c55ad015a3bf4f1b2b0b822cd15d6c15b0f00a08",
  );
  expect(() => fileBytes({ ...file("x.png", "test"), base64: "???" })).toThrow(
    "Invalid image data",
  );
});

test("suffixes names case-insensitively and keeps inner dots", () => {
  expect(
    uniqueName("Cat.party.png", new Set(["cat.party", "cat.party_2"])),
  ).toBe("Cat.party_3");
  expect(uniqueName(".png", new Set())).toBe("emoji");
});

test("skips remote and batch image duplicates, suffixes distinct images", async () => {
  listing.mockResolvedValue([emoji("old", "CAT"), emoji("other", "cat_2")]);
  const result = await client.uploadFiles([
    file("remote.png", "old"),
    file("cat.png", "new"),
    file("different.png", "new"),
    file("cat.gif", "another"),
  ]);
  expect(upload.mock.calls.map((args) => args[1])).toEqual(["cat_3", "cat_4"]);
  expect(result.status).toBe(
    "Uploaded 2 emoji(s); skipped 2 identical image(s); renamed 2 name conflict(s).",
  );
});

test("warm cache avoids image downloads but always refreshes metadata", async () => {
  listing.mockResolvedValue([emoji("old", "cat")]);
  await client.uploadFiles([file("x.png", "old")]);
  download.mockClear();
  await client.uploadFiles([file("x.png", "old")]);
  expect(listing).toHaveBeenCalledTimes(2);
  expect(download).not.toHaveBeenCalled();
  expect(upload).not.toHaveBeenCalled();
});

test("changed etag and expired hashes trigger downloads", async () => {
  listing.mockResolvedValue([emoji("old", "cat")]);
  await client.uploadFiles([file("x.png", "old")]);
  download.mockClear();
  listing.mockResolvedValue([emoji("old", "cat", "v2")]);
  await client.uploadFiles([file("x.png", "old")]);
  expect(download).toHaveBeenCalledTimes(1);
  const cache = stored[CACHE_PREFIX + "tenant"] as {
    entries: { checkedAt: number }[];
  };
  cache.entries[0].checkedAt -= 25 * 60 * 60 * 1000;
  await client.uploadFiles([file("x.png", "old")]);
  expect(download).toHaveBeenCalledTimes(2);
});

test("deleted emoji is neither duplicate nor reserved name", async () => {
  listing.mockResolvedValue([emoji("old", "cat")]);
  await client.uploadFiles([file("x.png", "old")]);
  listing.mockResolvedValue([emoji("old", "cat", "v1", true)]);
  await client.uploadFiles([file("cat.png", "old")]);
  expect(upload).toHaveBeenCalledWith(file("cat.png", "old"), "cat");
  expect(
    (await loadCache("tenant")).some((entry) => entry.documentId === "old"),
  ).toBe(false);
});

test("tenant caches do not bleed into another tenant", async () => {
  await client.uploadFiles([file("cat.png", "new")]);
  const other = new MsTeamsClient("ic3", "chat", "another-tenant");
  jest.spyOn(other, "listEmojis").mockResolvedValue([]);
  const otherUpload = jest
    .spyOn(other, "uploadFile")
    .mockResolvedValue("another-doc");
  await other.uploadFiles([file("cat.png", "new")]);
  expect(otherUpload).toHaveBeenCalledWith(file("cat.png", "new"), "cat");
});

test("recent uploads survive metadata lag, then expire", async () => {
  await client.uploadFiles([file("cat.png", "new")]);
  await client.uploadFiles([file("different.png", "new")]);
  expect(upload).toHaveBeenCalledTimes(1);
  const cache = stored[CACHE_PREFIX + "tenant"] as {
    entries: { uploadedAt: number }[];
  };
  cache.entries[0].uploadedAt -= 11 * 60 * 1000;
  await client.uploadFiles([file("cat.png", "new")]);
  expect(upload).toHaveBeenCalledTimes(2);
});

test("metadata/image failures stop uploads; failed upload never enters cache", async () => {
  listing.mockRejectedValueOnce(new Error("Unauthorized"));
  await expect(client.uploadFiles([file("cat.png", "new")])).rejects.toThrow(
    "Unauthorized",
  );
  listing.mockResolvedValue([emoji("old", "cat")]);
  download.mockRejectedValueOnce(new Error("Image unavailable"));
  await expect(client.uploadFiles([file("cat.png", "new")])).rejects.toThrow(
    "Image unavailable",
  );
  expect(upload).not.toHaveBeenCalled();
  listing.mockResolvedValue([]);
  upload.mockRejectedValueOnce(new Error("Upload failed"));
  await expect(client.uploadFiles([file("cat.png", "new")])).rejects.toThrow(
    "Upload failed",
  );
  expect(await loadCache("tenant")).toEqual([]);
});

test("session reset preserves cache and removes tokens and state", async () => {
  stored = {
    [CACHE_PREFIX + "tenant"]: { version: 1, entries: [] },
    secret: "token",
    processingState: {},
  };
  await clearSessionStorage();
  expect(Object.keys(stored)).toEqual([CACHE_PREFIX + "tenant"]);
});

test("malformed cache is ignored", async () => {
  stored[CACHE_PREFIX + "tenant"] = {
    version: 1,
    entries: [
      null,
      {},
      { documentId: "old", etag: "v1", hash: "bad", checkedAt: Date.now() },
    ],
  };
  expect(await loadCache("tenant")).toEqual([]);
});

test("cold indexing bounds downloads to four and checkpoints progress", async () => {
  listing.mockResolvedValue(
    Array.from({ length: 40 }, (_, i) => emoji(`doc${i}`, `cat${i}`)),
  );
  let active = 0;
  let maximum = 0;
  download.mockImplementation(async (id) => {
    maximum = Math.max(maximum, ++active);
    const hash = await imageHash(new TextEncoder().encode(id));
    active--;
    return hash;
  });
  await client.uploadFiles([file("existing.png", "doc0")]);
  expect(maximum).toBeLessThanOrEqual(4);
  expect(download).toHaveBeenCalledTimes(40);
  expect(storage.set.mock.calls.length).toBeGreaterThanOrEqual(2);
  expect(await loadCache("tenant")).toHaveLength(40);
});

test("viewer downloads originals once and warms upload deduplication cache", async () => {
  listing.mockResolvedValue([
    emoji("old", "cat"),
    emoji("gone", "deleted", "v1", true),
  ]);
  const read = jest.spyOn(client, "readImage").mockResolvedValue({
    bytes: new Uint8Array([116, 101, 115, 116]),
    contentType: "image/png",
  });
  const result = await client.browseEmojis("CAT", 0);
  expect(result).toMatchObject({
    tenantId: "tenant",
    total: 1,
    matched: 1,
    page: 0,
  });
  expect(result.items).toEqual([
    {
      documentId: "old",
      names: ["cat"],
      image: "data:image/png;base64,dGVzdA==",
    },
  ]);
  expect(read).toHaveBeenCalledTimes(1);
  download.mockClear();
  await client.uploadFiles([file("different.png", "test")]);
  expect(download).not.toHaveBeenCalled();
  expect(upload).not.toHaveBeenCalled();
});

test("viewer pages and bounds image requests without downloading whole tenant", async () => {
  listing.mockResolvedValue(
    Array.from({ length: 30 }, (_, i) => emoji(`doc${i}`, `cat${i}`)),
  );
  const read = jest.spyOn(client, "readImage").mockResolvedValue({
    bytes: new Uint8Array([1]),
    contentType: "image/png",
  });
  const result = await client.browseEmojis("", 1);
  expect(result.items).toHaveLength(6);
  expect(result.items[0].documentId).toBe("doc24");
  expect(read).toHaveBeenCalledTimes(6);
  const last = await client.browseEmojis("", 999);
  expect(last.page).toBe(1);
});

test("unavailable previews retain names and never become valid cached hashes", async () => {
  listing.mockResolvedValue([emoji("old", "cat")]);
  jest.spyOn(client, "readImage").mockRejectedValue(new Error("Forbidden"));
  const result = await client.browseEmojis("", 0);
  expect(result.items).toEqual([
    { documentId: "old", names: ["cat"], error: "Forbidden" },
  ]);
  expect(await loadCache("tenant")).toEqual([]);
});

test("viewer cache and previews remain separate per tenant", async () => {
  listing.mockResolvedValue([emoji("old", "cat")]);
  jest.spyOn(client, "readImage").mockResolvedValue({
    bytes: new Uint8Array([1]),
    contentType: "image/png",
  });
  await client.browseEmojis("", 0);
  expect(await loadCache("tenant")).toHaveLength(1);
  expect(await loadCache("another-tenant")).toEqual([]);
  expect(JSON.stringify(stored)).not.toContain("base64");
});

test("viewer streams names first and fast previews before slow downloads finish", async () => {
  listing.mockResolvedValue([
    emoji("slow", "slow-cat"),
    emoji("fast", "fast-cat"),
  ]);
  let finishSlow!: (value: { bytes: Uint8Array; contentType: string }) => void;
  const image = { bytes: new Uint8Array([1]), contentType: "image/png" };
  jest.spyOn(client, "readImage").mockImplementation((id) =>
    id === "slow"
      ? new Promise((resolve) => {
          finishSlow = resolve;
        })
      : Promise.resolve(image),
  );
  let fastReady!: () => void;
  const ready = new Promise<void>((resolve) => {
    fastReady = resolve;
  });
  const updates = jest.fn((update) => {
    if ("documentId" in update && update.documentId === "fast") fastReady();
  });
  const task = client.browseEmojis("", 0, updates);
  await ready;
  expect(updates.mock.calls[0][0].items).toEqual([
    { documentId: "slow", names: ["slow-cat"], loading: true },
    { documentId: "fast", names: ["fast-cat"], loading: true },
  ]);
  expect(updates.mock.calls[1][0]).toMatchObject({
    documentId: "fast",
    image: expect.any(String),
  });
  expect(updates).toHaveBeenCalledTimes(2);
  finishSlow(image);
  await task;
  expect(updates.mock.calls[2][0]).toMatchObject({
    documentId: "slow",
    image: expect.any(String),
  });
  expect(await loadCache("tenant")).toHaveLength(2);
});

test("duplicate scan groups names case-insensitively and hashes across entire tenant", async () => {
  listing.mockResolvedValue([
    emoji("a", "CAT"),
    emoji("b", "cat"),
    emoji("c", "other"),
    emoji("gone", "cat", "v1", true),
  ]);
  download.mockImplementation(async (id) =>
    imageHash(new TextEncoder().encode(id === "c" ? "different" : "same")),
  );
  const result = await client.checkDuplicates();
  expect(result).toMatchObject({ total: 3, checked: 3, errors: [] });
  expect(
    result.groups.map((group) => [
      group.kind,
      group.items.map((item) => item.documentId).sort(),
    ]),
  ).toEqual([
    ["name", ["a", "b"]],
    ["image", ["a", "b"]],
  ]);
  download.mockClear();
  await client.checkDuplicates();
  expect(download).not.toHaveBeenCalled();
});

test("duplicate scan reports unreadable images as unchecked and ignores repeated aliases", async () => {
  listing.mockResolvedValue([
    { ...emoji("a", "cat"), shortcuts: ["cat", "CAT"] },
    emoji("b", "dog"),
  ]);
  download.mockRejectedValue(new Error("Forbidden"));
  const result = await client.checkDuplicates();
  expect(result.groups).toEqual([]);
  expect(result.checked).toBe(0);
  expect(result.errors).toHaveLength(2);
});

test("duplicate check replaces fresh upload and legacy hashes with Teams-download hashes", async () => {
  const downloadedHash = await imageHash(
    new TextEncoder().encode("Teams bytes"),
  );
  await saveCache("tenant", [
    {
      documentId: "upload",
      etag: "v1",
      hash: await imageHash(new TextEncoder().encode("uploaded bytes")),
      checkedAt: Date.now(),
      source: "upload",
      uploadedAt: Date.now(),
      name: "meow_beret-gun",
    },
    {
      documentId: "legacy",
      etag: "v1",
      hash: await imageHash(new TextEncoder().encode("unknown bytes")),
      checkedAt: Date.now(),
    },
    {
      documentId: "preview",
      etag: "v1",
      hash: downloadedHash,
      checkedAt: Date.now(),
      source: "download",
    },
  ]);
  listing.mockResolvedValue([
    emoji("upload", "meow_beret-gun"),
    emoji("legacy", "meow-beret-gun"),
    emoji("preview", "another-name"),
  ]);
  download.mockResolvedValue(downloadedHash);
  const result = await client.checkDuplicates();
  expect(download).toHaveBeenCalledTimes(2);
  expect(download).toHaveBeenCalledWith("upload");
  expect(download).toHaveBeenCalledWith("legacy");
  expect(result.groups).toHaveLength(1);
  expect(result.groups[0].kind).toBe("image");
  expect(result.groups[0].items).toHaveLength(3);
  expect(
    (await loadCache("tenant")).every((entry) => entry.source === "download"),
  ).toBe(true);
  download.mockClear();
  await client.checkDuplicates();
  expect(download).not.toHaveBeenCalled();
});
