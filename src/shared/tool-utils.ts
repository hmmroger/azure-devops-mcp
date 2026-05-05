/* eslint-disable header/header */
import type { TranslatorConfigObject } from "node-html-markdown";
import type { Visitor } from "node-html-markdown/dist/visitor.js";
import { ShapeOutput } from "@modelcontextprotocol/sdk/server/zod-compat";
import { CallToolResult, ToolAnnotations } from "@modelcontextprotocol/sdk/types";
import { NodeHtmlMarkdown } from "node-html-markdown";
import { WebApi } from "azure-devops-node-api";
import type { ZodRawShape } from "zod";

const MIN_TABLE_SEPARATOR_COUNT = 3;
const htmlToMarkdownConverter = new NodeHtmlMarkdown(
  {
    bulletMarker: "-",
    useInlineLinks: true,
  },
  { ...getTableCustomTranslator() }
);

export type ToolHandler<InputArgs extends ZodRawShape> = (args: ShapeOutput<InputArgs>, extra: unknown) => Promise<CallToolResult>;

export interface McpToolConfig<InputArgs extends ZodRawShape> {
  name: string;
  description: string;
  inputSchema: InputArgs;
  annotations?: ToolAnnotations;
  handler: ToolHandler<InputArgs>;
}

const GUID_REGEX = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export const isGuid = (value: string): boolean => GUID_REGEX.test(value);

const projectIdCache = new Map<string, string>();

const getOrgFromServerUrl = (serverUrl: string): string => {
  try {
    return new URL(serverUrl).pathname.replace(/^\/+|\/+$/g, "");
  } catch {
    return serverUrl;
  }
};

/**
 * Resolves a project name or ID to its GUID. Short-circuits when the input is already
 * a GUID; otherwise calls the Core API and memoizes by `{org}-{name}` so subsequent
 * calls in the same process are free. The org is derived from `connection.serverUrl`,
 * so switching orgs naturally produces a fresh cache namespace.
 *
 * Use this in write paths (POST/PATCH) before passing `project` to the SDK. The
 * underlying typed-rest-client follows ADO's name → ID redirect on POST as a GET,
 * which drops the body and surfaces as a silent 404.
 */
export const resolveProjectId = async (connection: WebApi, project: string): Promise<string> => {
  if (isGuid(project)) {
    return project;
  }

  const org = getOrgFromServerUrl(connection.serverUrl);
  const cacheKey = `${org}-${project}`;
  const cached = projectIdCache.get(cacheKey);
  if (cached) {
    return cached;
  }

  const coreApi = await connection.getCoreApi();
  const found = await coreApi.getProject(project);
  if (!found?.id) {
    throw new Error(`Project '${project}' not found.`);
  }

  projectIdCache.set(cacheKey, found.id);
  return found.id;
};

/**
 * Creates a tool result object with text content.
 *
 * @param texts - Array of strings to be joined with newlines
 * @param isError - Optional flag indicating if this is an error result
 * @returns Object with content array containing text and optional isError flag
 */
export const textToolResult = (texts: string[], isError?: boolean) => {
  const text = texts.join("\n");
  return {
    content: [
      {
        type: "text" as const,
        text,
      },
    ],
    isError,
  };
};

/**
 * Creates an error tool result from an exception or error object.
 *
 * @param error - The error object or unknown error to extract message from
 * @param fallbackMessage - Message to use if no error message can be extracted
 * @returns Error tool result object with isError flag set to true
 */
export const getErrorToolResult = (error: unknown, fallbackMessage: string) => {
  const err = error as {
    message?: string;
    statusCode?: number;
    result?: { message?: string; value?: { Message?: string } };
  };

  const baseMessage = err?.message || fallbackMessage;
  const lines: string[] = [baseMessage];
  if (typeof err?.statusCode === "number") {
    lines.push(`Status code: ${err.statusCode}`);
  }

  const innerMessage = err?.result?.message ?? err?.result?.value?.Message;
  if (innerMessage && innerMessage !== baseMessage) {
    lines.push(`Details: ${innerMessage}`);
  }

  return textToolResult(lines, true);
};

/**
 * Converts an HTML string to markdown. Returns an empty string for falsy input.
 * Use to render rich-text fields (work item descriptions, PR comments, wiki HTML)
 * in agent-friendly text.
 */
export const htmlToMarkdown = (html: string | undefined | null): string => {
  if (!html) {
    return "";
  }
  return htmlToMarkdownConverter.translate(html);
};

function getTableCustomTranslator(maxSeparatorCount?: number): TranslatorConfigObject {
  maxSeparatorCount = maxSeparatorCount ?? MIN_TABLE_SEPARATOR_COUNT;
  // max separator count can't be smaller than min
  maxSeparatorCount = Math.max(MIN_TABLE_SEPARATOR_COUNT, maxSeparatorCount);

  return {
    table: ({ visitor }: { visitor: Visitor }) => ({
      surroundingNewlines: 2,
      childTranslators: visitor.instance.tableTranslators,
      postprocess: ({ content, nodeMetadata, node }) => {
        // Split into lines and filter out empty lines
        const lines = content.split("\n").filter((line) => line.trim());
        if (lines.length < 1) {
          return "RemoveNode";
        }

        // Process each line to extract column data and track max content length per column
        const rows: string[][] = [];
        let maxCols = 0;
        const colMaxLen = new Map<number, number>();

        for (const line of lines) {
          // Remove leading/trailing pipes and split by pipe
          const cleanLine = line.replace(/^\|\s*/, "").replace(/\s*\|$/, "");
          const cols = cleanLine.split("|").map((col) => col.trim());
          rows.push(cols);
          maxCols = Math.max(maxCols, cols.length);
          for (let i = 0; i < cols.length; i++) {
            colMaxLen.set(i, Math.max(colMaxLen.get(i) ?? MIN_TABLE_SEPARATOR_COUNT, cols[i].length));
          }
        }

        // Rebuild table with minimal separators
        let res = "";
        const caption = nodeMetadata.get(node)?.tableMeta?.caption;
        if (caption) {
          res += caption + "\n";
        }

        rows.forEach((cols: string[], rowNumber: number) => {
          res += "| ";
          for (let i = 0; i < maxCols; i++) {
            const cellContent = cols[i] || "";
            const padLen = Math.min(colMaxLen.get(i) ?? MIN_TABLE_SEPARATOR_COUNT, maxSeparatorCount);
            res += cellContent + " ".repeat(Math.max(0, padLen - cellContent.length)) + " |";
            if (i < maxCols - 1) {
              res += " ";
            }
          }

          res += "\n";

          // Add separator row after header with dash count based on column content width
          if (rowNumber === 0) {
            res += "|";
            for (let i = 0; i < maxCols; i++) {
              const dashes = Math.max(MIN_TABLE_SEPARATOR_COUNT, Math.min(colMaxLen.get(i) ?? MIN_TABLE_SEPARATOR_COUNT, maxSeparatorCount));
              res += " " + "-".repeat(dashes) + " |";
            }

            res += "\n";
          }
        });

        return res;
      },
    }),
  };
}
