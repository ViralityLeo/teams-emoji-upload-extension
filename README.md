# 🎨 Teams Emoji Upload Extension

> Effortlessly bulk upload custom emojis to your Microsoft Teams organization with just a few clicks!

![Screenshot](./readme/screenshot.png)

## ✨ Features

- Bulk upload multiple emoji files simultaneously
- Works via Teams in the browser

## 📋 Prerequisites

- Active Microsoft Teams tab in Chrome or Firefox 142+

## 🚀 Installation

### From Chrome Web Store

Download [Teams Emoji Uploader](https://chromewebstore.google.com/detail/teams-emoji-uploader/mlajagdepghhbclefnmcdnjfhfdmoofo) from the Chrome Web Store

### From Source

1. Clone this repository
2. Install dependencies and build:

```console
npm install
npm run build
npm test
```

3. Open Chrome and navigate to `chrome://extensions`
4. Enable "Developer Mode" in the top right
5. Click "Load Unpacked" and select this repository folder

![Unpack Instructions](./readme/unpack.png)

> 💡 **Tip**: After making changes, rebuild the project and refresh the extension in Chrome

## 🛠️ Development

### Formatting Code

```console
npm run format
```

### Running Tests

```console
npm test
```

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📄 License

[MIT License](LICENSE)

## Firefox

Build the Firefox variant from the same source:

```console
npm ci
npm run build:firefox
npm test
```

Open `about:debugging#/runtime/this-firefox`, click **Load Temporary Add-on**, and select `firefox/manifest.json`. Temporary installation lasts until Firefox restarts. For distribution, run `npm run zip:firefox` and submit `firefox-extension.zip` to Mozilla for signing (listed or unlisted).

Open a signed-in Teams tab before opening the extension. Test a single emoji, then a batch; verify the emojis appear in Teams. Upload failure stops the batch and preserves tokens for retry. Earlier files may already have uploaded: avoid retrying those files.

Firefox's refresh button reloads Teams and clears extension storage; it does not erase Teams cookies or site data. If emojis remain missing, allow Teams time to sync. Chrome retains its existing site-data reset behavior.

The Firefox manifest declares authentication information and website content: Teams tokens and selected emoji images/names are sent to Microsoft Teams to perform uploads. No analytics service is used. The default Firefox add-on ID is `teams-emoji-uploader@extensions.local`; keep it stable after signing.

## Duplicate handling and cache

Before each batch, the extension fetches a fresh full custom emoji list. Existing names are compared case-insensitively. A different image with the same filename-derived name gets the first free suffix: `cat`, `cat_2`, `cat_3`. Only the final filename extension is removed.

Images are compared using SHA-256 of file bytes, regardless of filename. Identical images already in Teams or earlier in the batch are skipped. Different encodings, compression, metadata, and GIF frame data count as different files; this is not visual similarity detection.

First use downloads existing original images, with at most four concurrent requests. Tenant-scoped hashes are saved in extension local storage; later runs reuse them while metadata is unchanged, for up to 24 hours. Deleted entries are removed. Recent confirmed uploads remain reserved for ten minutes while Teams catches up. Upload and refresh preserve this cache while clearing session tokens/state. Removing the extension clears its cache.

The popup reports uploaded, skipped, and renamed counts. If Teams metadata or an original image cannot be checked, the batch stops before uploads. A failed upload stops the remaining batch; successfully registered earlier uploads are cached immediately. Batches from this extension are serialized. Another user can still add an emoji during a batch; a resulting server conflict is reported as an error.

These checks use unofficial Teams endpoints. The full snapshot behavior (omit `x-ms-continuation`) is described in [this endpoint investigation](https://devblog.lac.co.jp/entry/20240801). Images are downloaded through `/views/imgpsh_fullsize_anim` to preserve GIF animation, falling back to `/views/imgpsh_fullsize` when the animated view is unavailable (404); `/content/imgpsh` is used only for uploads. Static full-size reads were verified in the local Firefox tenant; the animated variant still needs live GIF verification. Duplicate checks compare the downloaded bytes exactly; inaccessible content causes a clear error rather than uploading unchecked duplicates.

## Tenant emoji viewer

Keep a signed-in Teams tab open on `teams.microsoft.com` or `teams.cloud.microsoft`, then click **Browse tenant emojis**. The viewer opens in an extension tab, with a tenant selector for complete token pairs captured from that Teams page. Search names, refresh, or page through 24 emojis at a time. Tenant IDs identify each library; switching tenants clears the previous grid before loading.

Cards and names appear as soon as metadata arrives. Each card shows a loading spinner until its preview arrives; failed previews keep the name and show an error. Preview pixels are not persisted locally (Firefox may reuse its normal HTTP cache). Each preview reads original image bytes and stores their SHA-256 hash in that tenant's existing deduplication cache. Browsing warms the cache progressively, without downloading the whole library upfront. Preview image data stays in the viewer; only hashes are persisted. Missing previews keep their emoji names visible and are not cached as successful reads. Uploads and viewer cache writes share the same background queue.

Use **Check duplicates** to scan all live emojis in the selected tenant. Results group case-insensitive duplicate names and identical downloaded-image SHA-256 hashes, including document IDs. Duplicate scans reuse only fresh hashes explicitly recorded from Teams downloads. Upload-origin and older unlabelled hashes are downloaded again before comparison. Unreadable images are reported as unchecked; no emojis are deleted.

## Slackmojis shortcuts

1. Keep a signed-in Teams tab open on `teams.microsoft.com` or `teams.cloud.microsoft`. Sessions are discovered automatically when opening the extension/viewer or refreshing Slackmojis tenants.
2. Visit [Slackmojis](https://slackmojis.com/), then choose a Teams tenant in the small integration bar. A single available tenant is selected automatically.
3. Click the **＋** icon at the top right of an emoji. The original PNG/GIF/JPEG/WebP is imported directly, preserving GIF animation. Duplicate-image skipping, name suffixes, and the tenant cache use the same upload flow as local files.

Shortcuts also appear on individual emoji pages and dynamically loaded search/category results. The bar reports upload/skip results; errors leave the shortcut retryable. If sessions expire or no tenant is listed, reopen the extension on Teams and click **Refresh tenants** on Slackmojis. Imports do not reset or reload Teams.

The content script receives only tenant IDs and import results. Teams tokens stay in extension storage/background and are never sent to Slackmojis. Image downloads are restricted to `https://emojis.slackmojis.com/emojis/images/`, omit cookies, reject redirects, and have a 10 MB limit. Reload the extension and Slackmojis after rebuilding; the browser may ask to approve the new Slackmojis site access.

Session discovery reads both localStorage and sessionStorage from open Teams tabs, including Firefox container tabs. Popup, library, and Slackmojis use the same capture routine. A signed-in page with an unsupported/encrypted token cache shows a specific error instead of silently returning no tenants; encrypted caches are not decrypted by this extension.

Slackmojis import icons show a spinner while importing, then **✓** (added), **↪** (renamed), **=** (duplicate skipped), or **!** (failed). Hover for details. Every attempted import stays disabled until the page reloads.
