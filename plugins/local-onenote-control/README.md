# Local OneNote Control

Local OneNote Control is a personal Codex plugin for controlling Microsoft OneNote desktop on Windows.

It provides one MCP server, `local-onenote-control`, with tools to:

- check local OneNote automation availability
- list notebooks, sections, section groups, and pages
- search page titles
- read plain text from a page
- open a page in OneNote
- create a page in a section
- append plain text to a page

The bridge uses Windows PowerShell and the `OneNote.Application` COM object. It does not require Microsoft Graph, browser automation, or external npm packages.

## Notes

- The plugin must run on Windows with Microsoft OneNote desktop installed.
- Page and section IDs are OneNote-generated opaque values.
- Write operations are intentionally limited to create and append. There is no delete or full-page replacement tool.
- OneNote cloud sync is owned by OneNote itself; a successful local write does not prove remote sync completion.
