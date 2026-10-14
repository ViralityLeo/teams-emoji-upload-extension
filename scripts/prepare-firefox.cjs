const fs = require("node:fs");
const manifest = JSON.parse(fs.readFileSync("manifest.json", "utf8"));
manifest.background = { scripts: ["dist/background.js"] };
manifest.permissions = manifest.permissions.filter(
  (permission) => permission !== "browsingData",
);
manifest.content_scripts = manifest.content_scripts.filter((script) =>
  script.matches.includes("https://slackmojis.com/*"),
);
delete manifest.web_accessible_resources;
delete manifest.content_security_policy.sandbox;
manifest.browser_specific_settings = {
  gecko: {
    id: "teams-emoji-uploader@extensions.local",
    strict_min_version: "142.0",
    data_collection_permissions: {
      required: ["authenticationInfo", "websiteContent"],
    },
  },
};
fs.writeFileSync(
  "firefox/manifest.json",
  JSON.stringify(manifest, null, 2) + "\n",
);
fs.copyFileSync("popup.html", "firefox/popup.html");
fs.copyFileSync("viewer.html", "firefox/viewer.html");
fs.copyFileSync("slackmojis.css", "firefox/slackmojis.css");
fs.cpSync("icons", "firefox/icons", { recursive: true });
