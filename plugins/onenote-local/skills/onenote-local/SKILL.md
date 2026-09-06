---
name: onenote-local
description: Use the local Microsoft OneNote desktop MCP bridge to list notebooks, search pages, read page text, open pages, create pages, append plain text, or update fields and sections on an existing Kanban card.
---

# OneNote Local

Use this skill when the user wants Codex to control Microsoft OneNote installed on the same Windows machine.

## Boundary

- This plugin controls the local OneNote desktop app through Windows COM automation.
- It does not use OneNote Web, Microsoft Graph, browser selectors, cloud sync status, or account-level APIs.
- It intentionally avoids delete or replace-page tools. Page writes are limited to creating a page, appending a new plain-text outline block, updating supported scalar fields, or updating bullet items beneath a supported Kanban heading.
- `onenote_update_kanban_fields` updates only the supported scalar labels and requires every requested label to occur exactly once in one outline.
- `onenote_update_kanban_section` is not a generic table editor. It targets exactly one supported heading (`work`, `completion`, `prerequisites`, or `risks`), preserves the containing outline, and uses a required `entryId` as an idempotency marker.
- Treat page IDs and section IDs as OneNote-generated opaque strings. Get them from `onenote_list_hierarchy` or `onenote_search_pages`; do not invent them.

## Common Flow

1. Confirm that `onenote_status`, `onenote_list_hierarchy`, `onenote_create_page`, `onenote_read_page`, and any requested write tool are callable. If they are absent, report that the MCP tools are not attached to the current task; do not classify this as a OneNote COM failure or bypass the MCP server with a direct script call.
2. Call `onenote_status` to verify that OneNote COM automation is available.
3. Call `onenote_list_hierarchy` with `scope: "notebooks"` to see notebooks.
4. For a notebook or section group, call `onenote_list_hierarchy` with its `startNodeId` and `scope: "sections"`.
5. For a section, call `onenote_list_hierarchy` with its `startNodeId` and `scope: "pages"`.
6. Before creating a page, check the target section for an exact-title duplicate.
7. Use `onenote_read_page` to inspect page text before editing.
8. Use `onenote_append_text` only when the user asked to add content.
9. Use `onenote_create_page` only when the user asked for a new page, then read the returned `pageId` and verify the exact title plus a distinctive requested body marker. Allow additional content inherited from the section's default page template.
10. Use `onenote_update_kanban_fields` for `assignee`, `workType`, `priority`, `effort`, `createdDate`, or `dueDate`. Supply only fields the user asked to change.
11. Use `onenote_update_kanban_section` only after confirming the target heading. Choose `append` to preserve existing items or `replace` to replace only that section's child items. Supply a stable `entryId` and one or more plain-text `items`.
12. Read the page again after either Kanban update and verify the requested values or items. A section response with `duplicate: true` is a successful idempotent no-op, not a new write.

## Safety

- Before writing, confirm the target page or section from a fresh hierarchy/search result when possible.
- Do not retry a write automatically after a timeout or partial result because the first attempt may already have created a page.
- Do not use a Kanban update tool if the page lacks the expected label or heading, or has more than one matching target.
- Report that OneNote local changes may still need OneNote's own sync to reach OneDrive or other devices.
- If OneNote COM reports that it is unavailable, ask the user to install or open the OneNote desktop app and retry.
