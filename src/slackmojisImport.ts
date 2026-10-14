import TokenStore from "./tokenStore";
import { bytesToBase64 } from "./emojiDedup";
import { FileDetails, TeamsTokens } from "./types";

const MAX_IMAGE_BYTES = 10 * 1024 * 1024;

export function isSlackmojisPage(url: string | undefined): boolean {
  try {
    return !!url && new URL(url).origin === "https://slackmojis.com";
  } catch {
    return false;
  }
}

export function validateSlackmojisImage(url: string, name: string): URL {
  const image = new URL(url);
  if (
    image.origin !== "https://emojis.slackmojis.com" ||
    image.username ||
    image.password ||
    !image.pathname.startsWith("/emojis/images/") ||
    !/\.(png|gif|jpe?g|webp)$/i.test(image.pathname) ||
    typeof name !== "string" ||
    !name ||
    name.length > 200 ||
    /[\/\\\u0000-\u001f]/.test(name)
  ) {
    throw new Error("Invalid Slackmojis image");
  }
  return image;
}

export async function slackmojisSessions(): Promise<TeamsTokens[]> {
  return TokenStore.refreshFromOpenTeamsTabs();
}

export async function downloadSlackmoji(
  url: string,
  name: string,
): Promise<FileDetails> {
  const image = validateSlackmojisImage(url, name);
  // Third-party download receives neither Teams credentials nor site cookies.
  const response = await fetch(image.href, {
    credentials: "omit",
    redirect: "error",
  });
  if (!response.ok)
    throw new Error(`Slackmojis download failed: ${response.statusText}`);
  const contentType = response.headers
    .get("content-type")
    ?.split(";")[0]
    .trim()
    .toLowerCase();
  const extensions: Record<string, string> = {
    "image/png": "png",
    "image/gif": "gif",
    "image/jpeg": "jpg",
    "image/webp": "webp",
  };
  if (!contentType || !extensions[contentType])
    throw new Error("Slackmojis did not return a supported image");
  const length = Number(response.headers.get("content-length"));
  if (length > MAX_IMAGE_BYTES) {
    await response.body?.cancel();
    throw new Error("Slackmojis image exceeds 10 MB");
  }
  // Limit actual bytes too: Content-Length may be missing or incorrect.
  if (!response.body) throw new Error("Slackmojis returned an empty image");
  const reader = response.body.getReader();
  const chunks: Uint8Array[] = [];
  let size = 0;
  try {
    while (true) {
      const { value, done } = await reader.read();
      if (done) break;
      size += value.length;
      if (size > MAX_IMAGE_BYTES)
        throw new Error("Slackmojis image exceeds 10 MB");
      chunks.push(value);
    }
  } finally {
    await reader.cancel();
  }
  if (!size) throw new Error("Slackmojis returned an empty image");
  const bytes = new Uint8Array(size);
  let offset = 0;
  for (const chunk of chunks) {
    bytes.set(chunk, offset);
    offset += chunk.length;
  }
  const stem = name.replace(/\.[^.]*$/, "");
  return {
    name: `${stem}.${extensions[contentType]}`,
    size,
    type: contentType,
    base64: bytesToBase64(bytes),
  };
}
