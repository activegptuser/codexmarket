#!/usr/bin/env node
import { spawn } from "node:child_process";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const SERVER_NAME = "local-onenote-control";
const SERVER_VERSION = "0.1.0";
const PROTOCOL_VERSION = "2024-11-05";
const scriptDir = dirname(fileURLToPath(import.meta.url));
const oneNoteScript = resolve(scriptDir, "onenote-com.ps1");
const powershellPath = "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe";
const commandTimeoutMs = Number.parseInt(process.env.LOCAL_ONENOTE_TIMEOUT_MS ?? "45000", 10);

const tools = [
  {
    name: "onenote_status",
    description: "Check whether local Microsoft OneNote COM automation is available.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      properties: {}
    }
  },
  {
    name: "onenote_list_hierarchy",
    description: "List local OneNote notebooks, sections, section groups, and pages.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      properties: {
        startNodeId: {
          type: "string",
          description: "Optional OneNote node ID to list below. Omit for all notebooks."
        },
        scope: {
          type: "string",
          enum: ["notebooks", "sections", "pages", "children", "self"],
          description: "Hierarchy scope. Use pages for page lists under a section."
        }
      }
    }
  },
  {
    name: "onenote_search_pages",
    description: "Search local OneNote page titles by case-insensitive text.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["query"],
      properties: {
        query: {
          type: "string",
          minLength: 1,
          description: "Text to match against page titles."
        },
        maxResults: {
          type: "integer",
          minimum: 1,
          maximum: 100,
          description: "Maximum pages to return."
        }
      }
    }
  },
  {
    name: "onenote_read_page",
    description: "Read plain text from a local OneNote page by page ID.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["pageId"],
      properties: {
        pageId: {
          type: "string",
          minLength: 1,
          description: "OneNote page ID."
        },
        includeXml: {
          type: "boolean",
          description: "Include raw OneNote XML for troubleshooting."
        }
      }
    }
  },
  {
    name: "onenote_open_page",
    description: "Open a local OneNote page in the desktop app by page ID.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["pageId"],
      properties: {
        pageId: {
          type: "string",
          minLength: 1,
          description: "OneNote page ID."
        }
      }
    }
  },
  {
    name: "onenote_create_page",
    description: "Create a new page in a local OneNote section.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["sectionId", "title"],
      properties: {
        sectionId: {
          type: "string",
          minLength: 1,
          description: "OneNote section ID."
        },
        title: {
          type: "string",
          minLength: 1,
          description: "New page title."
        },
        bodyText: {
          type: "string",
          description: "Optional plain text body to place on the page."
        },
        openAfterCreate: {
          type: "boolean",
          description: "Open the page after creating it."
        }
      }
    }
  },
  {
    name: "onenote_append_text",
    description: "Append plain text as a new outline block on a local OneNote page.",
    inputSchema: {
      type: "object",
      additionalProperties: false,
      required: ["pageId", "text"],
      properties: {
        pageId: {
          type: "string",
          minLength: 1,
          description: "OneNote page ID."
        },
        text: {
          type: "string",
          minLength: 1,
          description: "Plain text to append."
        },
        openAfterAppend: {
          type: "boolean",
          description: "Open the page after appending."
        }
      }
    }
  }
];

const operationByTool = {
  onenote_status: "status",
  onenote_list_hierarchy: "list_hierarchy",
  onenote_search_pages: "search_pages",
  onenote_read_page: "read_page",
  onenote_open_page: "open_page",
  onenote_create_page: "create_page",
  onenote_append_text: "append_text"
};

function send(message) {
  process.stdout.write(`${JSON.stringify(message)}\n`);
}

function result(id, value) {
  send({ jsonrpc: "2.0", id, result: value });
}

function error(id, code, message, data) {
  send({ jsonrpc: "2.0", id, error: { code, message, ...(data ? { data } : {}) } });
}

function asToolContent(payload, isError = false) {
  return {
    content: [
      {
        type: "text",
        text: typeof payload === "string" ? payload : JSON.stringify(payload, null, 2)
      }
    ],
    ...(isError ? { isError: true } : {})
  };
}

function runOneNote(operation, args) {
  return new Promise((resolvePromise, reject) => {
    let settled = false;
    const child = spawn(
      powershellPath,
      [
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        oneNoteScript,
        "-Operation",
        operation,
        "-InputJson",
        JSON.stringify(args ?? {})
      ],
      {
        windowsHide: true,
        stdio: ["ignore", "pipe", "pipe"]
      }
    );

    const timeout = setTimeout(() => {
      if (settled) {
        return;
      }

      settled = true;
      child.kill();
      reject(new Error(`OneNote command timed out after ${commandTimeoutMs} ms. Open OneNote desktop, clear any dialogs, and try again.`));
    }, commandTimeoutMs);

    let stdout = "";
    let stderr = "";
    child.stdout.setEncoding("utf8");
    child.stderr.setEncoding("utf8");
    child.stdout.on("data", (chunk) => {
      stdout += chunk;
    });
    child.stderr.on("data", (chunk) => {
      stderr += chunk;
    });
    child.on("error", (spawnError) => {
      if (settled) {
        return;
      }

      settled = true;
      clearTimeout(timeout);
      reject(spawnError);
    });
    child.on("close", (code) => {
      if (settled) {
        return;
      }

      settled = true;
      clearTimeout(timeout);
      const trimmedStdout = stdout.trim();
      if (code !== 0) {
        reject(new Error((stderr || trimmedStdout || `OneNote command failed with exit code ${code}`).trim()));
        return;
      }

      if (!trimmedStdout) {
        resolvePromise({});
        return;
      }

      try {
        resolvePromise(JSON.parse(trimmedStdout));
      } catch {
        resolvePromise({ output: trimmedStdout });
      }
    });
  });
}

async function handleRequest(message) {
  const { id, method, params } = message;

  try {
    if (method === "initialize") {
      result(id, {
        protocolVersion: params?.protocolVersion ?? PROTOCOL_VERSION,
        capabilities: { tools: {} },
        serverInfo: { name: SERVER_NAME, version: SERVER_VERSION }
      });
      return;
    }

    if (method === "notifications/initialized") {
      return;
    }

    if (method === "tools/list") {
      result(id, { tools });
      return;
    }

    if (method === "tools/call") {
      const toolName = params?.name;
      const operation = operationByTool[toolName];
      if (!operation) {
        result(id, asToolContent(`Unknown tool: ${toolName}`, true));
        return;
      }

      try {
        const payload = await runOneNote(operation, params?.arguments ?? {});
        result(id, asToolContent(payload));
      } catch (toolError) {
        result(id, asToolContent({ error: String(toolError?.message ?? toolError) }, true));
      }
      return;
    }

    error(id, -32601, `Unsupported method: ${method}`);
  } catch (requestError) {
    error(id, -32603, String(requestError?.message ?? requestError));
  }
}

let buffer = "";
process.stdin.setEncoding("utf8");
process.stdin.on("data", (chunk) => {
  buffer += chunk;
  let newlineIndex = buffer.indexOf("\n");

  while (newlineIndex >= 0) {
    const line = buffer.slice(0, newlineIndex).trim();
    buffer = buffer.slice(newlineIndex + 1);
    if (line) {
      try {
        void handleRequest(JSON.parse(line));
      } catch (parseError) {
        error(null, -32700, String(parseError?.message ?? parseError));
      }
    }
    newlineIndex = buffer.indexOf("\n");
  }
});
