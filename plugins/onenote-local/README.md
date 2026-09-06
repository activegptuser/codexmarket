# OneNote Local

OneNote Local is a personal Codex plugin for controlling Microsoft OneNote desktop on Windows.

It provides one MCP server, `onenote-local`, with tools to:

- check local OneNote automation availability
- list notebooks, sections, section groups, and pages
- search page titles
- read plain text from a page
- open a page in OneNote
- create a page in a section
- create and fill a Kanban card page without adding a separate body text block
- append plain text to a page
- update scalar fields on an existing Kanban card
- append or replace idempotent bullet items in a supported Kanban section

The bridge uses Windows PowerShell and the `OneNote.Application` COM object. It does not require Microsoft Graph, browser automation, or external npm packages.

## Notes

- The plugin must run on Windows with Microsoft OneNote desktop installed.
- Page and section IDs are OneNote-generated opaque values.
- Write operations are intentionally limited to create, append, and targeted Kanban field or section updates. There is no delete, generic table editor, or full-page replacement tool.
- Use `onenote_create_kanban_page` for CrossDetect task cards. It creates the page from the section template, fills supported fields and sections, and does not accept `bodyText`.
- Use `onenote_create_page` with `bodyText` only for plain pages where a separate text block is acceptable.
- `onenote_update_kanban_fields` updates only `assignee`, `workType`, `priority`, `effort`, `createdDate`, and `dueDate` on the CrossDetect card template.
- `onenote_update_kanban_section` targets `work`, `completion`, `prerequisites`, or `risks`; it supports `append` and `replace`, preserves the surrounding outline, and requires an `entryId` to prevent duplicate writes.
- OneNote cloud sync is owned by OneNote itself; a successful local write does not prove remote sync completion.
