import { getExtensionApi } from "./extensionApi";
import { FileDetails } from "./types";

export const CACHE_PREFIX = "emojiHashCache:v2:";
const HASH_TTL = 24 * 60 * 60 * 1000;
const UPLOAD_GRACE = 10 * 60 * 1000;

export interface ExistingEmoji {
  documentId: string;
  shortcuts: string[];
  etag: string;
  isDeleted: boolean;
}

export interface CachedEmoji {
  documentId: string;
  etag: string;
  hash: string;
  checkedAt: number;
  source?: "upload" | "download";
  uploadedAt?: number;
  name?: string;
}

export async function imageHash(bytes: Uint8Array): Promise<string> {
  const digest = await crypto.subtle.digest(
    "SHA-256",
    new Uint8Array(bytes).buffer,
  );
  return Array.from(new Uint8Array(digest), (byte) =>
    byte.toString(16).padStart(2, "0"),
  ).join("");
}

export function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  for (let offset = 0; offset < bytes.length; offset += 8192) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 8192));
  }
  return btoa(binary);
}

export function fileBytes(file: FileDetails): Uint8Array {
  if (
    !file.base64 ||
    !/^(?:[A-Za-z0-9+/]{4})*(?:[A-Za-z0-9+/]{2}==|[A-Za-z0-9+/]{3}=)?$/.test(
      file.base64,
    )
  ) {
    throw new Error(`Invalid image data: ${file.name}`);
  }
  return Uint8Array.from(atob(file.base64), (char) => char.charCodeAt(0));
}

export function uniqueName(filename: string, names: Set<string>): string {
  const base = filename.replace(/\.[^.]*$/, "").trim() || "emoji";
  let name = base;
  for (let suffix = 2; names.has(name.toLowerCase()); suffix++)
    name = `${base}_${suffix}`;
  return name;
}

export async function loadCache(tenantId: string): Promise<CachedEmoji[]> {
  const key = CACHE_PREFIX + tenantId;
  const stored = (await getExtensionApi().storage.local.get(key))[key] as
    | { version?: unknown; entries?: unknown }
    | undefined;
  if (stored?.version !== 1 || !Array.isArray(stored.entries)) return [];
  return stored.entries.filter(
    (entry: CachedEmoji) =>
      entry &&
      typeof entry.documentId === "string" &&
      typeof entry.etag === "string" &&
      typeof entry.hash === "string" &&
      /^[a-f0-9]{64}$/.test(entry.hash) &&
      (entry.source === undefined ||
        entry.source === "upload" ||
        entry.source === "download") &&
      Number.isFinite(entry.checkedAt) &&
      entry.checkedAt <= Date.now() &&
      (entry.uploadedAt === undefined ||
        (Number.isFinite(entry.uploadedAt) &&
          entry.uploadedAt <= Date.now())) &&
      (entry.name === undefined || typeof entry.name === "string"),
  );
}

export async function saveCache(
  tenantId: string,
  entries: CachedEmoji[],
): Promise<void> {
  await getExtensionApi().storage.local.set({
    [CACHE_PREFIX + tenantId]: { version: 1, entries },
  });
}

export function reusableHash(
  entry: CachedEmoji | undefined,
  emoji: ExistingEmoji,
): boolean {
  return (
    !!entry &&
    Date.now() - entry.checkedAt < HASH_TTL &&
    (entry.etag === emoji.etag || !!entry.uploadedAt)
  );
}

export function pendingUploads(
  cache: CachedEmoji[],
  metadata: ExistingEmoji[],
): CachedEmoji[] {
  const known = new Set(metadata.map((emoji) => emoji.documentId));
  // ponytail: short grace for Teams metadata lag; full snapshot wins after ten minutes.
  return cache.filter(
    (entry) =>
      entry.uploadedAt &&
      entry.name &&
      Date.now() - entry.uploadedAt < UPLOAD_GRACE &&
      !known.has(entry.documentId),
  );
}

export async function clearSessionStorage(): Promise<void> {
  const storage = getExtensionApi().storage.local;
  const items = await storage.get(null);
  await storage.remove(
    Object.keys(items).filter((key) => !key.startsWith(CACHE_PREFIX)),
  );
}
