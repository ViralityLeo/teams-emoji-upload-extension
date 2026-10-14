# Firefox source build for Mozilla review

This archive contains the matching TypeScript source, static assets, tests,
build configuration and npm lockfile. Dependencies and generated bundles are excluded.

## Build environment

Produced on macOS (Darwin ARM64), Node.js 22.18.0 and npm 10.9.3.
The build also runs on Linux; no macOS-specific tools are required.
Install Node.js 22 from https://nodejs.org/en/download and use npm 10.
All dependencies come from the public npm registry and are pinned by package-lock.json.

From the extracted archive's root:

```sh
npm ci
npm run build:firefox
```

The complete unsigned Firefox extension is in `firefox/`.
Webpack compiles/minifies TypeScript into `firefox/dist/` and
`scripts/prepare-firefox.cjs` produces the Firefox manifest and copies static assets.
To create the submission ZIP on macOS/Linux (requires the standard zip command):

```sh
npm run zip:firefox
```

The result is `firefox-extension.zip`. Optional validation: `npm test -- --runInBand`.
No credentials or Teams account are required to build or run the unit tests.
For live use, sign into Teams in Firefox. The extension reads Teams tokens from
that page's browser storage and sends them only to the Teams APIs for the user's
requested operations. Slackmojis imports fetch public images; tenant sessions
and duplicate-check hashes are stored locally in extension storage.
