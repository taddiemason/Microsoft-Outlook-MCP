import { GraphError } from "../graph.js";

/** The content shape MCP tools must return. */
export interface ToolResult {
  content: { type: "text"; text: string }[];
  isError?: boolean;
  [key: string]: unknown;
}

/** Wrap a successful result as pretty-printed JSON text content. */
export function ok(data: unknown): ToolResult {
  return {
    content: [{ type: "text", text: JSON.stringify(data, null, 2) }],
  };
}

/** Wrap an error as an MCP tool error result (never throws to the transport). */
export function fail(err: unknown): ToolResult {
  let text: string;
  if (err instanceof GraphError) {
    text = `Microsoft Graph error (HTTP ${err.status}): ${err.message}`;
    if (err.status === 401) {
      text +=
        "\nThe session may have expired or lacks the required scope. " +
        "Re-run `npm run login`.";
    }
  } else if (err instanceof Error) {
    text = err.message;
  } else if (typeof err === "string") {
    text = err;
  } else {
    text = JSON.stringify(err);
  }
  return {
    content: [{ type: "text", text }],
    isError: true,
  };
}
