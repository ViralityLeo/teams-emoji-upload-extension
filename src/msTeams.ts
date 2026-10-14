import {
  CachedEmoji,
  ExistingEmoji,
  fileBytes,
  bytesToBase64,
  imageHash,
  loadCache,
  pendingUploads,
  reusableHash,
  saveCache,
  uniqueName,
} from "./emojiDedup";
import {
  TeamsUploadResponse,
  TeamsMetadataResponse,
  FileDetails,
  EmojiLibraryPage,
  EmojiPreview,
  EmojiDuplicates,
  ProcessResult,
} from "./types";

class MsTeamsClient {
  private _ic3Token: string;
  private _chatToken: string;
  private _permissionsId: string;

  constructor(ic3Token: string, chatToken: string, permissionsId: string) {
    this._ic3Token = ic3Token;
    this._chatToken = chatToken;
    this._permissionsId = permissionsId;
  }

  public async createObject(): Promise<{ id: string; msCv: string }> {
    const response = await fetch(
      "https://as-prod.asyncgw.teams.microsoft.com/v1/objects/",
      {
        method: "POST",
        headers: {
          authorization: `Bearer ${this._ic3Token}`,
          "content-type": "application/json",
          accept: "application/json",
          "x-ms-client-version": "1415/25021400912",
          origin: "https://teams.microsoft.com",
          referer: "https://teams.microsoft.com/",
        },
        mode: "cors",
        body: JSON.stringify({
          type: "pish/image",
          permissions: {
            ["*:tid:" + this._permissionsId]: ["read"],
          },
          sharingMode: "Unknown",
        }),
      },
    );

    if (!response.ok) {
      throw new Error(`Failed to create object: ${response.statusText}`);
    }

    const msCv = response.headers.get("ms-cv") || "";
    const data = (await response.json()) as TeamsUploadResponse;
    return { id: data.id, msCv };
  }

  public async uploadFileContent(
    documentId: string,
    msCv: string,
    file: FileDetails,
  ): Promise<void> {
    const buf = fileBytes(file);
    const response = await fetch(
      `https://as-prod.asyncgw.teams.microsoft.com/v1/objects/${documentId}/content/imgpsh`,
      {
        method: "PUT",
        headers: {
          authorization: `Bearer ${this._ic3Token}`,
          "content-type": "application/octet-stream",
          "x-ms-client-version": "1415/25021400912",
          DNT: "1",
          origin: "https://teams.microsoft.com",
          referer: "https://teams.microsoft.com/",
          "x-ms-migration": "True",
          "x-ms-test-user": "False",
          "ms-cv": msCv,
        },
        body: buf,
      },
    );

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(
        `Failed to upload image: ${response.statusText} - ${errorText}`,
      );
    }
  }

  public async sendMetadata(
    documentId: string,
    msCv: string,
    shortcuts: string[],
  ): Promise<void> {
    const response = await fetch(
      "https://teams.microsoft.com/api/csa/apac/api/v1/customemoji/metadata",
      {
        method: "POST",
        headers: {
          authorization: `Bearer ${this._chatToken}`,
          "content-type": "application/json",
          accept: "application/json",
          "x-ms-client-version": "1415/25021400912",
          "x-ms-client-type": "web",
          "ms-cv": msCv,
        },
        body: JSON.stringify({
          shortcuts: shortcuts,
          documentId: documentId,
        }),
      },
    );

    if (!response.ok) {
      throw new Error(`Failed to create metadata: ${response.statusText}`);
    }

    const metadataData = (await response.json()) as TeamsMetadataResponse;
    if (metadataData.success === false)
      throw new Error("Teams rejected emoji metadata");
  }

  public async uploadFile(
    file: FileDetails,
    shortcut = uniqueName(file.name, new Set()),
  ): Promise<string> {
    try {
      const { id: documentId, msCv } = await this.createObject();
      console.log("Document ID:", documentId);
      await this.uploadFileContent(documentId, msCv, file);
      await this.sendMetadata(documentId, msCv, [shortcut]);
      return documentId;
    } catch (error) {
      console.error("Error uploading file:", error);
      throw error;
    }
  }

  public async listEmojis(): Promise<ExistingEmoji[]> {
    // No x-ms-continuation header: fetch a full snapshot, not a delta.
    const response = await fetch(
      "https://teams.microsoft.com/api/csa/apac/api/v1/customemoji/metadata",
      {
        headers: { authorization: `Bearer ${this._chatToken}` },
        cache: "no-store",
      },
    );
    if (!response.ok)
      throw new Error(
        `Could not check existing emojis: ${response.statusText}`,
      );
    const data = await response.json();
    if (!Array.isArray(data.categories))
      throw new Error("Invalid Teams emoji metadata; upload stopped");
    const emojis: ExistingEmoji[] = [];
    for (const category of data.categories) {
      if (category.id !== "customEmoji") continue;
      if (!Array.isArray(category.emoticons))
        throw new Error("Invalid Teams emoji list; upload stopped");
      for (const emoji of category.emoticons) {
        if (
          typeof emoji.documentId !== "string" ||
          !emoji.documentId ||
          !Array.isArray(emoji.shortcuts) ||
          !emoji.shortcuts.every((name: unknown) => typeof name === "string") ||
          (emoji.etag !== undefined && typeof emoji.etag !== "string") ||
          (emoji.isDeleted !== undefined &&
            typeof emoji.isDeleted !== "boolean")
        ) {
          throw new Error("Invalid Teams emoji entry; upload stopped");
        }
        emojis.push({
          documentId: emoji.documentId,
          shortcuts: emoji.shortcuts,
          etag: emoji.etag || "",
          isDeleted: !!emoji.isDeleted,
        });
      }
    }
    return emojis;
  }

  public async readImage(
    documentId: string,
  ): Promise<{ bytes: Uint8Array; contentType: string }> {
    // Ask for animation-preserving bytes before falling back for static-only objects.
    const base = `https://as-prod.asyncgw.teams.microsoft.com/v1/objects/${encodeURIComponent(documentId)}/views/`;
    const options: RequestInit = {
      headers: { authorization: `Bearer ${this._ic3Token}` },
      credentials: "omit",
    };
    let response = await fetch(base + "imgpsh_fullsize_anim", options);
    if (response.status === 404)
      response = await fetch(base + "imgpsh_fullsize", options);
    if (!response.ok)
      throw new Error(
        `Could not check existing emoji image (${response.statusText || response.status}); upload stopped`,
      );
    const contentType = response.headers
      .get("content-type")
      ?.split(";")[0]
      .trim()
      .toLowerCase();
    if (
      !contentType ||
      (!contentType.startsWith("image/") &&
        contentType !== "application/octet-stream")
    ) {
      throw new Error(
        "Teams did not return original image content; upload stopped",
      );
    }
    const bytes = new Uint8Array(await response.arrayBuffer());
    if (!bytes.length)
      throw new Error("Teams returned an empty image; upload stopped");
    return { bytes, contentType };
  }

  public async existingImageHash(documentId: string): Promise<string> {
    return imageHash((await this.readImage(documentId)).bytes);
  }

  public async browseEmojis(
    query: string,
    page: number,
    onUpdate?: (update: EmojiLibraryPage | EmojiPreview) => void,
  ): Promise<EmojiLibraryPage> {
    const metadata = await this.listEmojis();
    const live = metadata.filter((emoji) => !emoji.isDeleted);
    const matches = live.filter((emoji) =>
      emoji.shortcuts.some((name) =>
        name.toLowerCase().includes(query.toLowerCase()),
      ),
    );
    const pageSize = 24;
    const currentPage = Math.min(
      page,
      Math.max(0, Math.ceil(matches.length / pageSize) - 1),
    );
    const selected = matches.slice(
      currentPage * pageSize,
      (currentPage + 1) * pageSize,
    );
    onUpdate?.({
      success: true,
      tenantId: this._permissionsId,
      total: live.length,
      matched: matches.length,
      page: currentPage,
      pageSize,
      items: selected.map((emoji) => ({
        documentId: emoji.documentId,
        names: emoji.shortcuts,
        loading: true,
      })),
    });
    const cache = await loadCache(this._permissionsId);
    const liveIds = new Set(live.map((emoji) => emoji.documentId));
    const entries = new Map(
      cache
        .filter((entry) => liveIds.has(entry.documentId))
        .concat(pendingUploads(cache, metadata))
        .map((entry) => [entry.documentId, entry]),
    );
    const items: EmojiPreview[] = [];
    for (let offset = 0; offset < selected.length; offset += 4) {
      items.push(
        ...(await Promise.all(
          selected
            .slice(offset, offset + 4)
            .map(async (emoji): Promise<EmojiPreview> => {
              try {
                const { bytes, contentType } = await this.readImage(
                  emoji.documentId,
                );
                const hash = await imageHash(bytes);
                entries.set(emoji.documentId, {
                  documentId: emoji.documentId,
                  etag: emoji.etag,
                  hash,
                  checkedAt: Date.now(),
                  source: "download",
                });
                const preview: EmojiPreview = {
                  documentId: emoji.documentId,
                  names: emoji.shortcuts,
                  image: `data:${contentType};base64,${bytesToBase64(bytes)}`,
                };
                onUpdate?.(preview);
                return preview;
              } catch (error) {
                // Keep names visible when a preview cannot be read; do not cache failed downloads.
                const preview: EmojiPreview = {
                  documentId: emoji.documentId,
                  names: emoji.shortcuts,
                  error: error instanceof Error ? error.message : String(error),
                };
                onUpdate?.(preview);
                return preview;
              }
            }),
        )),
      );
      await saveCache(this._permissionsId, Array.from(entries.values()));
    }
    await saveCache(this._permissionsId, Array.from(entries.values()));
    return {
      success: true,
      tenantId: this._permissionsId,
      total: live.length,
      matched: matches.length,
      page: currentPage,
      pageSize,
      items,
    };
  }

  public async checkDuplicates(): Promise<EmojiDuplicates> {
    const metadata = await this.listEmojis();
    const live = metadata.filter((emoji) => !emoji.isDeleted);
    const cache = await loadCache(this._permissionsId);
    const cached = new Map(cache.map((entry) => [entry.documentId, entry]));
    const entries: CachedEmoji[] = pendingUploads(cache, metadata);
    const names = new Map<string, ExistingEmoji[]>();
    const hashes = new Map<string, ExistingEmoji[]>();
    const errors: EmojiDuplicates["errors"] = [];
    let checked = 0;
    for (const emoji of live) {
      for (const name of new Set(
        emoji.shortcuts.map((name) => name.toLowerCase()),
      )) {
        names.set(name, [...(names.get(name) || []), emoji]);
      }
    }
    for (let offset = 0; offset < live.length; offset += 4) {
      await Promise.all(
        live.slice(offset, offset + 4).map(async (emoji) => {
          try {
            const previous = cached.get(emoji.documentId);
            const reuse =
              previous?.source === "download" && reusableHash(previous, emoji);
            const hash = reuse
              ? previous!.hash
              : await this.existingImageHash(emoji.documentId);
            checked++;
            hashes.set(hash, [...(hashes.get(hash) || []), emoji]);
            entries.push({
              documentId: emoji.documentId,
              etag: emoji.etag,
              hash,
              checkedAt: reuse ? previous!.checkedAt : Date.now(),
              source: "download",
            });
          } catch (error) {
            errors.push({
              names: emoji.shortcuts,
              error: error instanceof Error ? error.message : String(error),
            });
          }
        }),
      );
      await saveCache(this._permissionsId, entries);
    }
    const groups: EmojiDuplicates["groups"] = [];
    for (const [kind, map] of [
      ["name", names],
      ["image", hashes],
    ] as const) {
      for (const [value, emojis] of map) {
        if (emojis.length > 1)
          groups.push({
            kind,
            value,
            items: emojis.map((emoji) => ({
              documentId: emoji.documentId,
              names: emoji.shortcuts,
            })),
          });
      }
    }
    return {
      success: true,
      tenantId: this._permissionsId,
      total: live.length,
      checked,
      groups,
      errors,
    };
  }

  public async uploadFiles(files: FileDetails[]): Promise<ProcessResult> {
    if (files.length === 0)
      return { success: false, error: "Please select files first" };
    const candidates = await Promise.all(
      files.map(async (file) => ({
        file,
        hash: await imageHash(fileBytes(file)),
      })),
    );
    const metadata = await this.listEmojis();
    const cache = await loadCache(this._permissionsId);
    const cached = new Map(cache.map((entry) => [entry.documentId, entry]));
    const entries: CachedEmoji[] = pendingUploads(cache, metadata);
    const names = new Set(entries.map((entry) => entry.name!.toLowerCase()));
    const hashes = new Set(entries.map((entry) => entry.hash));
    const live = metadata.filter((emoji) => !emoji.isDeleted);
    live.forEach((emoji) =>
      emoji.shortcuts.forEach((name) => names.add(name.toLowerCase())),
    );
    const warming = new Map(cached);
    let downloads = 0;
    for (let offset = 0; offset < live.length; offset += 4) {
      const chunk = await Promise.all(
        live.slice(offset, offset + 4).map(async (emoji) => {
          const previous = cached.get(emoji.documentId);
          const reuse = reusableHash(previous, emoji);
          const hash = reuse
            ? previous!.hash
            : await this.existingImageHash(emoji.documentId);
          if (!reuse) downloads++;
          return {
            documentId: emoji.documentId,
            etag: emoji.etag,
            hash,
            checkedAt: reuse ? previous!.checkedAt : Date.now(),
            source: reuse ? previous!.source : ("download" as const),
          };
        }),
      );
      entries.push(...chunk);
      chunk.forEach((entry) => {
        hashes.add(entry.hash);
        warming.set(entry.documentId, entry);
      });
      if (downloads >= 32) {
        await saveCache(this._permissionsId, Array.from(warming.values()));
        downloads = 0;
      }
    }
    await saveCache(this._permissionsId, entries);
    let uploaded = 0;
    let skipped = 0;
    let renamed = 0;
    for (const { file, hash } of candidates) {
      if (hashes.has(hash)) {
        skipped++;
        continue;
      }
      const base = uniqueName(file.name, new Set());
      const name = uniqueName(file.name, names);
      const documentId = await this.uploadFile(file, name);
      uploaded++;
      if (name !== base) renamed++;
      names.add(name.toLowerCase());
      hashes.add(hash);
      entries.push({
        documentId,
        etag: "",
        hash,
        checkedAt: Date.now(),
        source: "upload",
        uploadedAt: Date.now(),
        name,
      });
      await saveCache(this._permissionsId, entries);
    }
    return {
      success: true,
      counts: { uploaded, skipped, renamed },
      status: `Uploaded ${uploaded} emoji(s); skipped ${skipped} identical image(s); renamed ${renamed} name conflict(s).`,
    };
  }
}

export default MsTeamsClient;
