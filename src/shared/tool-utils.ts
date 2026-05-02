/* eslint-disable header/header */
import type { TranslatorConfigObject } from "node-html-markdown";
import type { Visitor } from "node-html-markdown/dist/visitor.js";
import { ShapeOutput } from "@modelcontextprotocol/sdk/server/zod-compat";
import { CallToolResult, ToolAnnotations } from "@modelcontextprotocol/sdk/types";
import { NodeHtmlMarkdown } from "node-html-markdown";
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
  const exceptionError = (error as Error).message;
  const errorMessage = exceptionError ? exceptionError : fallbackMessage;
  return textToolResult([errorMessage], true);
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
