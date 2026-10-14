import { readFileSync } from "node:fs";
import { renderLibrary, renderDuplicates } from "../viewer";
import TokenStore from "../tokenStore";
import { EmojiLibraryPage } from "../types";

const result: EmojiLibraryPage = {
  success: true,
  tenantId: "tenant",
  total: 1,
  matched: 1,
  page: 0,
  pageSize: 24,
  items: [],
};

test("renders untrusted emoji names as text and rejects remote preview URLs", () => {
  const grid = document.createElement("div");
  renderLibrary(
    {
      ...result,
      items: [
        {
          documentId: "doc",
          names: ["<script>alert(1)</script>"],
          image: "https://evil.example/image",
        },
      ],
    },
    grid,
  );
  expect(grid.querySelector("script")).toBeNull();
  expect(grid.querySelector("img")).toBeNull();
  expect(grid.querySelector("figcaption")?.textContent).toBe(
    "<script>alert(1)</script>",
  );
  expect(grid.textContent).toContain("Preview unavailable");
});

test("renders image and replaces broken preview with accessible fallback", () => {
  const grid = document.createElement("div");
  renderLibrary(
    {
      ...result,
      items: [
        {
          documentId: "doc",
          names: ["cat"],
          image: "data:image/png;base64,dGVzdA==",
        },
      ],
    },
    grid,
  );
  const image = grid.querySelector("img")!;
  expect(image.alt).toBe("cat");
  image.dispatchEvent(new Event("error"));
  expect(grid.querySelector("img")).toBeNull();
  expect(grid.textContent).toContain("Preview unavailable");
  renderLibrary({ ...result, items: [] }, grid);
  expect(grid.childNodes).toHaveLength(0);
});

test("pairs captured tokens by tenant rather than mixing accounts", () => {
  expect(
    TokenStore.tenantSessions({
      "chatsvcagg-one": { secret: "chat-one", realm: "one" },
      "ic3.teams.office-two": { secret: "ic3-two", realm: "two" },
      "ic3.teams.office-one": { secret: "ic3-one", realm: "one" },
      "chatsvcagg-bad": { secret: "bad" },
    }),
  ).toEqual([{ chatsvcagg: "chat-one", ic3: "ic3-one", permissionsId: "one" }]);
});

test("viewer switches tenants without retaining previous previews", async () => {
  document.body.innerHTML = readFileSync("viewer.html", "utf8").replace(
    /<script[\s\S]*?<\/script>/g,
    "",
  );
  const sessions = ["one", "two"].map((permissionsId) => ({
    permissionsId,
    ic3: "ic3-" + permissionsId,
    chatsvcagg: "chat-" + permissionsId,
  }));
  let completeSecond!: (value: EmojiLibraryPage) => void;
  const sendMessage = jest
    .fn()
    .mockResolvedValueOnce({
      ...result,
      tenantId: "one",
      items: [{ documentId: "one-doc", names: ["one-emoji"] }],
    })
    .mockImplementationOnce(
      () =>
        new Promise<EmojiLibraryPage>((resolve) => {
          completeSecond = resolve;
        }),
    );
  let onUpdate!: (message: any) => void;
  Object.defineProperty(globalThis, "chrome", {
    configurable: true,
    value: {
      tabs: { query: async () => [] },
      storage: {
        local: {
          get: async () => ({ teamsTenantSessions: sessions }),
          set: async () => undefined,
        },
      },
      runtime: {
        sendMessage,
        onMessage: {
          addListener: (listener: typeof onUpdate) => {
            onUpdate = listener;
          },
        },
      },
    },
  });
  document.dispatchEvent(new Event("DOMContentLoaded"));
  await new Promise((resolve) => setTimeout(resolve, 0));
  expect(document.getElementById("tenantNote")?.textContent).toBe(
    "Tenant: one",
  );
  expect(document.getElementById("grid")?.textContent).toContain("one-emoji");
  const select = document.getElementById("tenantSelect") as HTMLSelectElement;
  select.value = "two";
  select.dispatchEvent(new Event("change"));
  expect(document.getElementById("grid")?.childNodes).toHaveLength(0);
  expect(sendMessage.mock.calls[1][0].tokens.permissionsId).toBe("two");
  const requestId = sendMessage.mock.calls[1][0].requestId;
  onUpdate({
    type: "emojiPreviewUpdate",
    requestId,
    update: {
      ...result,
      tenantId: "two",
      items: [{ documentId: "two-doc", names: ["two-emoji"], loading: true }],
    },
  });
  expect(document.querySelector(".preview-loading")).not.toBeNull();
  onUpdate({
    type: "emojiPreviewUpdate",
    requestId: sendMessage.mock.calls[0][0].requestId,
    update: {
      documentId: "two-doc",
      names: ["wrong-tenant"],
      image: "data:image/png;base64,AQ==",
    },
  });
  expect(document.querySelector(".preview-loading")).not.toBeNull();
  onUpdate({
    type: "emojiPreviewUpdate",
    requestId,
    update: {
      documentId: "two-doc",
      names: ["two-emoji"],
      image: "data:image/png;base64,AQ==",
    },
  });
  expect(document.querySelector(".preview-loading")).toBeNull();
  expect(document.querySelector("#grid img")).not.toBeNull();
  completeSecond({
    ...result,
    tenantId: "two",
    items: [{ documentId: "two-doc", names: ["two-emoji"] }],
  });
  await new Promise((resolve) => setTimeout(resolve, 0));
  expect(document.getElementById("grid")?.textContent).toContain("two-emoji");
  expect(document.getElementById("grid")?.textContent).not.toContain(
    "one-emoji",
  );
  expect(document.getElementById("tenantNote")?.textContent).toBe(
    "Tenant: two",
  );
  document.body.replaceChildren();
});

test("shows accessible per-card spinner with name before preview arrives", () => {
  const grid = document.createElement("div");
  renderLibrary(
    {
      ...result,
      items: [{ documentId: "pending", names: ["cat"], loading: true }],
    },
    grid,
  );
  expect(grid.querySelector('[aria-label="Loading preview"]')).not.toBeNull();
  expect(grid.querySelector("figcaption")?.textContent).toBe("cat");
  expect(grid.textContent).not.toContain("Preview unavailable");
});

test("duplicate report keeps tenant names inert and clearly reports incomplete checks", () => {
  const container = document.createElement("section");
  renderDuplicates(
    {
      success: true,
      tenantId: "t",
      total: 3,
      checked: 2,
      groups: [
        {
          kind: "name",
          value: "cat",
          items: [
            { documentId: "a", names: ["<script>bad</script>"] },
            { documentId: "b", names: ["cat"] },
          ],
        },
      ],
      errors: [{ names: ["dog"], error: "Forbidden" }],
    },
    container,
  );
  expect(container.querySelector("script")).toBeNull();
  expect(container.textContent).toContain("1 image(s) unchecked");
  expect(container.textContent).toContain("Same name (cat)");
  expect(readFileSync("viewer.html", "utf8")).not.toMatch(
    /Deletion permission|id="permission"/,
  );
});
