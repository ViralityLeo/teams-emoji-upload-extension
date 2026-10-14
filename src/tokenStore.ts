import { TeamsTokens } from "./types";
import { getExtensionApi } from "./extensionApi";
class TokenStore {
  chatsvcagg: string | null = null;
  ic3: string | null = null;
  permissionsId: string | null = null;

  static async collectTokensFromStorage(): Promise<TokenStore> {
    const tokens = new TokenStore();

    try {
      const items = await getExtensionApi().storage.local.get(null); // Get all items

      for (const [key, value] of Object.entries(items)) {
        try {
          const obj = typeof value === "string" ? JSON.parse(value) : value;

          if (key.includes("chatsvcagg")) {
            tokens.chatsvcagg = obj.secret;
          }
          if (key.includes("ic3.teams.office")) {
            tokens.ic3 = obj.secret;
            tokens.permissionsId = obj.realm;
          }
          if (tokens.chatsvcagg && tokens.ic3) {
            break;
          }
        } catch (e) {
          console.error(`Failed to parse storage item: ${key}`, e);
        }
      }
    } catch (e) {
      console.error("Failed to get tokens from storage:", e);
    }

    return tokens;
  }

  static tenantSessions(items: Record<string, unknown>): TeamsTokens[] {
    const sessions = new Map<string, TeamsTokens>();
    for (const [key, value] of Object.entries(items)) {
      if (!value || typeof value !== "object") continue;
      const token = value as { secret?: unknown; realm?: unknown };
      if (
        typeof token.secret !== "string" ||
        !token.secret ||
        typeof token.realm !== "string" ||
        !token.realm
      )
        continue;
      const session = sessions.get(token.realm) || {
        chatsvcagg: null,
        ic3: null,
        permissionsId: token.realm,
      };
      if (key.includes("chatsvcagg")) session.chatsvcagg = token.secret;
      if (key.includes("ic3.teams.office")) session.ic3 = token.secret;
      sessions.set(token.realm, session);
    }
    return Array.from(sessions.values()).filter(
      (session) => session.chatsvcagg && session.ic3,
    );
  }

  static async refreshFromOpenTeamsTabs(): Promise<TeamsTokens[]> {
    const api = getExtensionApi();
    const stored = await api.storage.local.get("teamsTenantSessions");
    const sessions = new Map<string, TeamsTokens>();
    if (Array.isArray(stored.teamsTenantSessions)) {
      for (const session of stored.teamsTenantSessions as TeamsTokens[]) {
        if (
          session &&
          typeof session.permissionsId === "string" &&
          typeof session.ic3 === "string" &&
          typeof session.chatsvcagg === "string"
        )
          sessions.set(session.permissionsId, session);
      }
    }
    const tabs = await api.tabs.query({
      url: ["https://teams.microsoft.com/*", "https://teams.cloud.microsoft/*"],
    });
    let matchingItems = 0;
    let captureErrors = 0;
    for (const tab of tabs) {
      if (tab.id === undefined) continue;
      try {
        const results = await api.scripting.executeScript({
          target: { tabId: tab.id },
          func: () => {
            const items: Record<string, unknown> = {};
            for (const storage of [localStorage, sessionStorage]) {
              for (let i = 0; i < storage.length; i++) {
                const key = storage.key(i);
                if (
                  !key ||
                  (!key.includes("chatsvcagg") &&
                    !key.includes("ic3.teams.office"))
                )
                  continue;
                try {
                  items[key] = JSON.parse(storage.getItem(key) || "null");
                } catch {
                  /* Ignore non-JSON entries. */
                }
              }
            }
            return items;
          },
        });
        const items = results[0]?.result;
        if (!items || typeof items !== "object") continue;
        matchingItems += Object.keys(items).length;
        for (const session of TokenStore.tenantSessions(items))
          sessions.set(session.permissionsId!, session);
      } catch {
        captureErrors++;
      }
    }
    const result = Array.from(sessions.values());
    await api.storage.local.set({ teamsTenantSessions: result });
    if (!result.length && tabs.length) {
      throw new Error(
        captureErrors
          ? "Could not read the Teams tab. Reload Teams and allow the extension on its site."
          : matchingItems
            ? "Teams login cache found, but no readable token pair. Its storage format may have changed."
            : "Teams is open, but required login tokens are not available yet. Wait for Teams to finish loading, then refresh tenants.",
      );
    }
    return result;
  }

  static collectTokensFromLocalStorage(): TokenStore {
    const tokens = new TokenStore();

    try {
      for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (!key) continue;

        try {
          const value = localStorage.getItem(key);
          if (!value) continue;

          const obj = JSON.parse(value);
          if (key.includes("chatsvcagg")) {
            tokens.chatsvcagg = obj.secret;
          }
          if (key.includes("ic3.teams.office")) {
            tokens.ic3 = obj.secret;
            tokens.permissionsId = obj.realm;
          }
          if (tokens.chatsvcagg && tokens.ic3) {
            break;
          }
        } catch (e) {
          console.error(`Failed to parse localStorage item: ${key}`, e);
        }
      }
    } catch (e) {
      console.error("Error accessing localStorage:", e);
    }

    return tokens;
  }

  static captureAndStoreTokens(): void {
    try {
      const tokens = TokenStore.collectTokensFromLocalStorage();
      if (tokens.chatsvcagg && tokens.ic3) {
        // Store the tokens in Chrome storage for later use
        getExtensionApi().storage.local.set(
          {
            teams_chatsvcagg: { secret: tokens.chatsvcagg },
            teams_ic3: { secret: tokens.ic3, realm: tokens.permissionsId },
          },
          () => {
            console.log("Teams tokens saved to Chrome storage");
          },
        );
      }
    } catch (e) {
      console.error("Failed to capture and store tokens:", e);
    }
  }
}

export default TokenStore;
