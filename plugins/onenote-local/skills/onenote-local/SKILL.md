---
name: onenote-local
description: Use the local Microsoft OneNote desktop MCP bridge to list notebooks, search pages, read page text, open pages, create pages, append plain text, or update fields and sections on an existing Kanban card.
---

# OneNote Local

Use this skill when the user wants the agent to control Microsoft OneNote installed on the same Windows machine.

## Boundary

- This plugin controls the local OneNote desktop app through Windows COM automation.
- It does not use OneNote Web, Microsoft Graph, browser selectors, cloud sync status, or account-level APIs.
- It intentionally avoids delete or replace-page tools. Page writes are limited to creating a plain page, creating and filling a Kanban card from the section template, appending a new plain-text outline block, updating supported scalar fields, or updating bullet items beneath a supported Kanban heading.
- Use `onenote_create_kanban_page` for CrossDetect task cards. Do not pass `bodyText` through `onenote_create_page` for a templated `tasks` page because OneNote keeps the section template and adds `bodyText` as a separate outline block.
- `onenote_update_kanban_fields` updates only the supported scalar labels and requires every requested label to occur exactly once in one outline.
- `onenote_update_kanban_section` is not a generic table editor. It targets exactly one supported heading (`work`, `completion`, `prerequisites`, or `risks`), preserves the containing outline, and uses a required `entryId` as an idempotency marker.
- Treat page IDs and section IDs as OneNote-generated opaque strings. Get them from `onenote_list_hierarchy` or `onenote_search_pages`; do not invent them.

## Common Flow

1. Confirm that `onenote_status`, `onenote_list_hierarchy`, `onenote_create_page`, `onenote_create_kanban_page`, `onenote_read_page`, and any requested write tool are callable. OpenCode may prefix MCP tool names with the server name; use the attached tools with these suffixes. If they are absent, report that the MCP tools are not attached to the current task; do not classify this as a OneNote COM failure or bypass the MCP server with a direct script call.
2. Call `onenote_status` to verify that OneNote COM automation is available.
3. Call `onenote_list_hierarchy` with `scope: "notebooks"` to see notebooks.
4. For a notebook or section group, call `onenote_list_hierarchy` with its `startNodeId` and `scope: "sections"`.
5. For a section, call `onenote_list_hierarchy` with its `startNodeId` and `scope: "pages"`.
6. Before creating a page, check the target section for an exact-title duplicate.
7. Use `onenote_read_page` to inspect page text before editing.
   - Start with `includeXml: false`. Request raw XML only when plain text cannot identify a required link or structure.
   - A `page-id` embedded in a `onenote:` hyperlink is not guaranteed to be a callable local COM page ID. If `onenote_read_page` returns HRESULT `0x80042005`, search the exact page title with `onenote_search_pages` and retry with the returned local ID.
   - When an exact-title search returns more than one page, prefer the result in the expected notebook or section, then the most recently modified result. If the result remains ambiguous, read the candidates and report the selection basis.
8. Use `onenote_append_text` only when the user asked to add content.
9. For a CrossDetect task/Kanban card, use `onenote_create_kanban_page` instead of `onenote_create_page` with `bodyText`. Verify the exact title, updated scalar fields, and expected section items.
10. Use `onenote_create_page` only for a plain new page. If `bodyText` is supplied, expect it to become a separate outline block in addition to any section template.
11. Use `onenote_update_kanban_fields` for `assignee`, `workType`, `priority`, `effort`, `createdDate`, or `dueDate`. Supply only fields the user asked to change.
12. Use `onenote_update_kanban_section` only after confirming the target heading. Choose `append` to preserve existing items or `replace` to replace only that section's child items. Supply a stable `entryId` and one or more plain-text `items`.
13. Read the page again after either Kanban update and verify the requested values or items. A section response with `duplicate: true` is a successful idempotent no-op, not a new write.

## Dashboard and linked pages

- Find Dashboard tables by their visible labels before interpreting rows.
- Resolve a linked page through its exact title when the hyperlink ID is not readable through COM. Do not treat `0x80042005` from one hyperlink ID as a OneNote availability failure.
- The server serializes OneNote COM commands. Batched reads may still complete one at a time; wait for the queued calls within the configured timeout instead of retrying them in parallel.
- Prefer plain-text reads for collection work. Use raw XML only as a targeted fallback because Dashboard XML can be very large.

## Safety

- Before writing, confirm the target page or section from a fresh hierarchy/search result when possible.
- Do not retry a write automatically after a timeout or partial result because the first attempt may already have created a page.
- Do not use a Kanban update tool if the page lacks the expected label or heading, or has more than one matching target.
- Report that OneNote local changes may still need OneNote's own sync to reach OneDrive or other devices.
- If OneNote COM reports that it is unavailable, ask the user to install or open the OneNote desktop app and retry.
