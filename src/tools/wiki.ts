// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { AccessToken } from "@azure/identity";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { WebApi } from "azure-devops-node-api";
import { z, ZodRawShape } from "zod";
import { WikiPage, WikiPageDetail, WikiPagesBatchRequest, WikiType, WikiV2 } from "azure-devops-node-api/interfaces/WikiInterfaces.js";
import { getErrorToolResult, McpToolConfig, textToolResult, ToolHandler } from "../shared/tool-utils.js";
import { apiVersion } from "../utils.js";

const WIKI_TOOLS = {
  list_azure_devops_wikis: "list_azure_devops_wikis",
  get_azure_devops_wiki: "get_azure_devops_wiki",
  list_azure_devops_wiki_pages: "list_azure_devops_wiki_pages",
  get_azure_devops_wiki_page: "get_azure_devops_wiki_page",
  get_azure_devops_wiki_page_by_url: "get_azure_devops_wiki_page_by_url",
  get_azure_devops_wiki_page_content: "get_azure_devops_wiki_page_content",
  create_or_update_azure_devops_wiki_page: "create_or_update_azure_devops_wiki_page",
};

type ToolConfigType =
  | ReturnType<typeof listWikis>
  | ReturnType<typeof getWiki>
  | ReturnType<typeof listWikiPages>
  | ReturnType<typeof getWikiPage>
  | ReturnType<typeof getWikiPageByUrl>
  | ReturnType<typeof getWikiPageContent>
  | ReturnType<typeof createOrUpdateWikiPage>;

function formatWikiSummary(wiki: WikiV2): string {
  if (!wiki.name) {
    throw new Error("Wiki missing required 'name' field.");
  }

  const attrs: string[] = [];
  if (wiki.id) {
    attrs.push(`ID: ${wiki.id}`);
  }

  if (wiki.type !== undefined) {
    attrs.push(`type: ${WikiType[wiki.type] ?? wiki.type}`);
  }

  if (wiki.projectId) {
    attrs.push(`projectId: ${wiki.projectId}`);
  }

  if (wiki.repositoryId) {
    attrs.push(`repositoryId: ${wiki.repositoryId}`);
  }

  if (wiki.mappedPath) {
    attrs.push(`mappedPath: ${wiki.mappedPath}`);
  }

  if (wiki.isDisabled) {
    attrs.push(`disabled: true`);
  }

  return attrs.length > 0 ? `${wiki.name} (${attrs.join(", ")})` : wiki.name;
}

function formatWikiSummaryLines(wikis: WikiV2[]): string[] {
  return wikis.map((w) => `- ${formatWikiSummary(w)}`);
}

function formatWikiDetailLines(wiki: WikiV2): string[] {
  if (!wiki.name) {
    throw new Error("Wiki missing required 'name' field.");
  }
  const lines: string[] = [`Wiki '${wiki.name}':`];
  if (wiki.id) {
    lines.push(`- ID: ${wiki.id}`);
  }
  if (wiki.type !== undefined) {
    lines.push(`- type: ${WikiType[wiki.type] ?? wiki.type}`);
  }
  if (wiki.projectId) {
    lines.push(`- projectId: ${wiki.projectId}`);
  }
  if (wiki.repositoryId) {
    lines.push(`- repositoryId: ${wiki.repositoryId}`);
  }
  if (wiki.mappedPath) {
    lines.push(`- mappedPath: ${wiki.mappedPath}`);
  }
  if (wiki.isDisabled) {
    lines.push(`- disabled: true`);
  }
  return lines;
}

function formatWikiPageDetail(page: WikiPageDetail): string {
  if (!page.path) {
    throw new Error("Wiki page missing required 'path' field.");
  }

  return `ID: ${page.id} Path: ${page.path}`;
}

function formatWikiPageDetailLines(pages: WikiPageDetail[]): string[] {
  return pages.map((p) => `- ${formatWikiPageDetail(p)}`);
}

function formatWikiSubPage(page: WikiPage): string {
  if (!page.path) {
    throw new Error("Wiki page missing required 'path' field.");
  }

  const attrs: string[] = [];
  if (page.id) {
    attrs.push(`ID: ${page.id}`);
  }

  if (page.isParentPage) {
    attrs.push(`isParentPage: true`);
  }

  return `Path: ${page.path}${attrs.length > 0 ? ` (${attrs.join(", ")})` : ""}`;
}

function formatWikiPageLines(page: WikiPage, recursionLevel?: string): string[] {
  if (!page.path) {
    throw new Error("Wiki page missing required 'path' field.");
  }

  const effectiveRecursion = recursionLevel ?? "None";
  const attrs: string[] = [];
  if (page.id !== undefined) {
    attrs.push(`ID: ${page.id}`);
  }

  if (page.gitItemPath) {
    attrs.push(`gitItemPath: ${page.gitItemPath}`);
  }

  if (page.isParentPage) {
    attrs.push(`isParentPage: true`);
  }

  if (page.isNonConformant) {
    attrs.push(`isNonConformant: true`);
  }

  if (page.order !== undefined) {
    attrs.push(`order: ${page.order}`);
  }

  attrs.push(`recursionLevel: ${effectiveRecursion}`);
  const subpagesFetched = effectiveRecursion !== "None";
  if (subpagesFetched && page.subPages !== undefined) {
    attrs.push(`subPages: ${page.subPages.length}`);
  }

  const header = `Wiki page path: '${page.path}' (${attrs.join(", ")}):`;
  const lines: string[] = [header];

  if (!subpagesFetched) {
    lines.push(`(subpages not fetched at recursionLevel=None — use 'OneLevel' or 'Full' to include children)`);
  } else if (page.subPages && page.subPages.length > 0) {
    lines.push("Subpages (path-based endpoint omits subpage IDs — call 'get_azure_devops_wiki_page' on a subpage path to resolve its ID):");
    for (const sub of page.subPages) {
      lines.push(`- ${formatWikiSubPage(sub)}`);
    }
  }

  return lines;
}

function formatTokenPagination(returned: number, top: number, continuationToken: string | undefined): string {
  const more = continuationToken ? `yes (continuationToken=${continuationToken})` : "no";
  return `[Returned: ${returned} | Top: ${top} | More: ${more}]`;
}

async function fetchWikiPagesBatch(
  tokenProvider: () => Promise<AccessToken>,
  connectionProvider: () => Promise<WebApi>,
  userAgentProvider: () => string,
  project: string,
  wikiIdentifier: string,
  body: WikiPagesBatchRequest
): Promise<{ pages: WikiPageDetail[]; continuationToken: string | undefined }> {
  const connection = await connectionProvider();
  const accessToken = await tokenProvider();
  const baseUrl = connection.serverUrl.replace(/\/$/, "");
  const params = new URLSearchParams({ "api-version": apiVersion });
  const restUrl = `${baseUrl}/${encodeURIComponent(project)}/_apis/wiki/wikis/${encodeURIComponent(wikiIdentifier)}/pagesbatch?${params.toString()}`;

  const response = await fetch(restUrl, {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${accessToken.token}`,
      "Content-Type": "application/json",
      "User-Agent": userAgentProvider(),
    },
    body: JSON.stringify(body),
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Failed to list wiki pages (${response.status}): ${errorText}`);
  }

  const json = await response.json();
  const pages: WikiPageDetail[] = json?.value ?? (Array.isArray(json) ? json : []);
  const continuationToken = response.headers.get("x-ms-continuationtoken") ?? undefined;
  return { pages, continuationToken };
}

export function configureWikiTools(
  server: McpServer,
  tokenProvider: () => Promise<AccessToken>,
  connectionProvider: () => Promise<WebApi>,
  userAgentProvider: () => string
): McpToolConfig<ZodRawShape>[] {
  const wikiTools: ToolConfigType[] = [];
  wikiTools.push(listWikis(connectionProvider));
  wikiTools.push(getWiki(connectionProvider));
  wikiTools.push(listWikiPages(tokenProvider, connectionProvider, userAgentProvider));
  wikiTools.push(getWikiPage(tokenProvider, connectionProvider, userAgentProvider));
  wikiTools.push(getWikiPageByUrl(tokenProvider, connectionProvider, userAgentProvider));
  wikiTools.push(getWikiPageContent(connectionProvider));
  wikiTools.push(createOrUpdateWikiPage(tokenProvider, connectionProvider, userAgentProvider));
  return wikiTools as unknown as McpToolConfig<ZodRawShape>[];
}

function listWikis(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    project: z.string().optional().describe("Project name or ID. If omitted, returns all wikis in the organization."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ project }) => {
    try {
      const connection = await connectionProvider();
      const wikiApi = await connection.getWikiApi();
      const wikis = (await wikiApi.getAllWikis(project)) ?? [];

      const scope = project ? ` in project '${project}'` : " in the organization";
      if (wikis.length === 0) {
        return textToolResult([`No wikis found${scope}.`]);
      }

      const header = `Wikis${scope} (count: ${wikis.length}):`;
      return textToolResult([header, ...formatWikiSummaryLines(wikis)]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to fetch wikis.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: WIKI_TOOLS.list_azure_devops_wikis,
    description: "List wikis in an organization or project.",
    inputSchema,
    handler,
  };

  return config;
}

function getWiki(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    wikiIdentifier: z.string().describe("The unique identifier or name of the wiki."),
    project: z.string().optional().describe("Project name or ID. If omitted, the default project is used."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ wikiIdentifier, project }) => {
    try {
      const connection = await connectionProvider();
      const wikiApi = await connection.getWikiApi();
      const wiki = await wikiApi.getWiki(wikiIdentifier, project);

      if (!wiki) {
        return textToolResult([`Wiki '${wikiIdentifier}' not found.`], true);
      }

      return textToolResult(formatWikiDetailLines(wiki));
    } catch (error) {
      return getErrorToolResult(error, "Failed to fetch wiki.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: WIKI_TOOLS.get_azure_devops_wiki,
    description: "Get a wiki by identifier or name.",
    inputSchema,
    handler,
  };

  return config;
}

function listWikiPages(tokenProvider: () => Promise<AccessToken>, connectionProvider: () => Promise<WebApi>, userAgentProvider: () => string) {
  const inputSchema = {
    wikiIdentifier: z.string().describe("The unique identifier or name of the wiki."),
    project: z.string().describe("Project name or ID where the wiki is located."),
    top: z.number().optional().describe("Maximum number of pages to return. Defaults to 25."),
    continuationToken: z.string().optional().describe("Token returned from a prior call to fetch the next batch."),
    pageViewsForDays: z.number().optional().describe("If set, include page view counts over the last N days."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ wikiIdentifier, project, top, continuationToken, pageViewsForDays }) => {
    try {
      const effectiveTop = top ?? 25;
      const { pages, continuationToken: nextToken } = await fetchWikiPagesBatch(tokenProvider, connectionProvider, userAgentProvider, project, wikiIdentifier, {
        top: effectiveTop,
        continuationToken,
        pageViewsForDays,
      });

      if (pages.length === 0) {
        return textToolResult([`No pages found in wiki '${wikiIdentifier}' (project '${project}').`]);
      }

      const header = `Pages in wiki '${wikiIdentifier}' (project '${project}'):`;
      const pagination = formatTokenPagination(pages.length, effectiveTop, nextToken);
      return textToolResult([header, pagination, ...formatWikiPageDetailLines(pages)]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to fetch wiki pages.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: WIKI_TOOLS.list_azure_devops_wiki_pages,
    description: "List pages in a wiki. Pass continuationToken from a prior call to fetch the next batch.",
    inputSchema,
    handler,
  };

  return config;
}

function getWikiPage(tokenProvider: () => Promise<AccessToken>, connectionProvider: () => Promise<WebApi>, userAgentProvider: () => string) {
  const inputSchema = {
    wikiIdentifier: z.string().describe("The unique identifier or name of the wiki."),
    project: z.string().describe("Project name or ID where the wiki is located."),
    path: z.string().describe("Path of the wiki page (e.g. '/Home' or '/Documentation/Setup')."),
    recursionLevel: z
      .enum(["None", "OneLevel", "OneLevelPlusNestedEmptyFolders", "Full"])
      .optional()
      .describe("'None' returns only the page, 'OneLevel' includes direct children, 'Full' includes all descendants."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ wikiIdentifier, project, path, recursionLevel }) => {
    try {
      const connection = await connectionProvider();
      const accessToken = await tokenProvider();

      const normalizedPath = path.startsWith("/") ? path : `/${path}`;
      const baseUrl = connection.serverUrl.replace(/\/$/, "");
      const params = new URLSearchParams({
        "path": normalizedPath,
        "api-version": apiVersion,
      });

      recursionLevel = recursionLevel ?? "OneLevel";
      params.append("recursionLevel", recursionLevel);

      const url = `${baseUrl}/${encodeURIComponent(project)}/_apis/wiki/wikis/${encodeURIComponent(wikiIdentifier)}/pages?${params.toString()}`;
      const response = await fetch(url, {
        headers: {
          "Authorization": `Bearer ${accessToken.token}`,
          "User-Agent": userAgentProvider(),
        },
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Failed to get wiki page (${response.status}): ${errorText}`);
      }

      const page = (await response.json()) as WikiPage;
      return textToolResult(formatWikiPageLines(page, recursionLevel));
    } catch (error) {
      return getErrorToolResult(error, "Failed to fetch wiki page metadata.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: WIKI_TOOLS.get_azure_devops_wiki_page,
    description: "Get wiki page metadata (no content) by path. Use recursionLevel to include subpages.",
    inputSchema,
    handler,
  };

  return config;
}

function getWikiPageContent(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    wikiIdentifier: z.string().describe("The unique identifier or name of the wiki."),
    project: z.string().describe("Project name or ID where the wiki is located."),
    pageId: z.number().describe("Numeric page ID. Obtain via 'list_azure_devops_wiki_pages' or 'get_azure_devops_wiki_page'."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ wikiIdentifier, project, pageId }) => {
    try {
      const connection = await connectionProvider();
      const wikiApi = await connection.getWikiApi();
      const stream = await wikiApi.getPageByIdText(project, wikiIdentifier, pageId, undefined, true);

      if (!stream) {
        return textToolResult([`No content returned for page ID ${pageId} in wiki '${wikiIdentifier}'.`], true);
      }

      const content = await streamToString(stream);
      const header = `Wiki page content (wiki: ${wikiIdentifier}, pageId: ${pageId}):`;
      return textToolResult([header, content]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to fetch wiki page content.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: WIKI_TOOLS.get_azure_devops_wiki_page_content,
    description: "Get wiki page markdown content by page ID. Use 'get_azure_devops_wiki_page' (by path) or 'list_azure_devops_wiki_pages' to discover the ID.",
    inputSchema,
    handler,
  };

  return config;
}

function createOrUpdateWikiPage(tokenProvider: () => Promise<AccessToken>, connectionProvider: () => Promise<WebApi>, userAgentProvider: () => string) {
  const inputSchema = {
    wikiIdentifier: z.string().describe("Unique identifier or name of the wiki."),
    path: z.string().describe("Path of the wiki page (e.g. '/Home' or '/Documentation/Setup')."),
    content: z.string().describe("Markdown content of the wiki page."),
    project: z.string().optional().describe("Project name or ID. If omitted, the default project is used."),
    etag: z.string().optional().describe("ETag for editing an existing page. Will be fetched automatically if not provided."),
    branch: z.string().optional().describe("Wiki repository branch. Defaults to 'wikiMaster'."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ wikiIdentifier, path, content, project, etag, branch }) => {
    try {
      const effectiveBranch = branch ?? "wikiMaster";
      const connection = await connectionProvider();
      const accessToken = await tokenProvider();

      const normalizedPath = path.startsWith("/") ? path : `/${path}`;
      const encodedPath = encodeURIComponent(normalizedPath);
      const baseUrl = connection.serverUrl;
      const projectParam = project ?? "";
      const url = `${baseUrl}/${encodeURIComponent(projectParam)}/_apis/wiki/wikis/${encodeURIComponent(wikiIdentifier)}/pages?path=${encodedPath}&versionDescriptor.versionType=branch&versionDescriptor.version=${encodeURIComponent(effectiveBranch)}&api-version=${apiVersion}`;
      const userAgent = userAgentProvider();

      const createResponse = await fetch(url, {
        method: "PUT",
        headers: {
          "Authorization": `Bearer ${accessToken.token}`,
          "Content-Type": "application/json",
          "User-Agent": userAgent,
        },
        body: JSON.stringify({ content }),
      });

      if (createResponse.ok) {
        const result = (await createResponse.json()) as WikiPage;
        return textToolResult([`Created wiki page '${normalizedPath}' (page ID: ${result.id ?? "unknown"}).`]);
      }

      const isPageExists = createResponse.status === 409 || createResponse.status === 500;
      if (!isPageExists) {
        const errorText = await createResponse.text();
        throw new Error(`Failed to create page (${createResponse.status}): ${errorText}`);
      }

      let currentEtag = etag;
      if (!currentEtag) {
        const getResponse = await fetch(url, {
          method: "GET",
          headers: {
            "Authorization": `Bearer ${accessToken.token}`,
            "User-Agent": userAgent,
          },
        });

        if (getResponse.ok) {
          currentEtag = getResponse.headers.get("etag") ?? getResponse.headers.get("ETag") ?? undefined;
          if (!currentEtag) {
            const pageData = await getResponse.json();
            currentEtag = pageData?.eTag;
          }
        }
      }

      if (!currentEtag) {
        throw new Error("Could not retrieve ETag for existing page.");
      }

      const updateResponse = await fetch(url, {
        method: "PUT",
        headers: {
          "Authorization": `Bearer ${accessToken.token}`,
          "Content-Type": "application/json",
          "User-Agent": userAgent,
          "If-Match": currentEtag,
        },
        body: JSON.stringify({ content }),
      });

      if (!updateResponse.ok) {
        const errorText = await updateResponse.text();
        throw new Error(`Failed to update page (${updateResponse.status}): ${errorText}`);
      }

      const result = (await updateResponse.json()) as WikiPage;
      return textToolResult([`Updated wiki page '${normalizedPath}' (page ID: ${result.id ?? "unknown"}).`]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to create or update wiki page.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: WIKI_TOOLS.create_or_update_azure_devops_wiki_page,
    description: "Create a new wiki page or update an existing one. Falls back to update with ETag if the page already exists.",
    inputSchema,
    handler,
  };

  return config;
}

function getWikiPageByUrl(tokenProvider: () => Promise<AccessToken>, connectionProvider: () => Promise<WebApi>, userAgentProvider: () => string) {
  const inputSchema = {
    url: z.string().describe("Full Azure DevOps wiki page URL. Supports query-string form '/_wiki/wikis/{wikiId}?pagePath=%2FMy%20Page' and path-id form '/_wiki/wikis/{wikiId}/{pageId}/Title'."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ url }) => {
    try {
      const parsed = parseWikiUrl(url);
      if (!parsed.ok) {
        return textToolResult([parsed.error], true);
      }

      const connection = await connectionProvider();
      const accessToken = await tokenProvider();
      const baseUrl = connection.serverUrl.replace(/\/$/, "");
      const projectSegment = encodeURIComponent(parsed.project);
      const wikiSegment = encodeURIComponent(parsed.wikiIdentifier);

      let restUrl: string;
      if (parsed.pageId !== undefined) {
        const params = new URLSearchParams({ "api-version": apiVersion });
        restUrl = `${baseUrl}/${projectSegment}/_apis/wiki/wikis/${wikiSegment}/pages/${parsed.pageId}?${params.toString()}`;
      } else {
        const params = new URLSearchParams({
          "path": parsed.pagePath ?? "/",
          "api-version": apiVersion,
        });
        restUrl = `${baseUrl}/${projectSegment}/_apis/wiki/wikis/${wikiSegment}/pages?${params.toString()}`;
      }

      const response = await fetch(restUrl, {
        headers: {
          "Authorization": `Bearer ${accessToken.token}`,
          "User-Agent": userAgentProvider(),
        },
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Failed to resolve wiki page (${response.status}): ${errorText}`);
      }

      const page = (await response.json()) as WikiPage;
      const header = `Resolved wiki URL → project '${parsed.project}', wiki '${parsed.wikiIdentifier}':`;
      return textToolResult([header, ...formatWikiPageLines(page, "None")]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to resolve wiki page by URL.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: WIKI_TOOLS.get_azure_devops_wiki_page_by_url,
    description: "Resolve an Azure DevOps wiki page URL to page metadata. Use the returned page ID with 'get_azure_devops_wiki_page_content' to fetch markdown.",
    inputSchema,
    handler,
  };

  return config;
}

type ParsedWikiUrl = { ok: true; project: string; wikiIdentifier: string; pagePath?: string; pageId?: number } | { ok: false; error: string };

// Parses Azure DevOps wiki page URLs.
// Supported examples:
//  - https://dev.azure.com/org/project/_wiki/wikis/wikiIdentifier?wikiVersion=GBmain&pagePath=%2FHome
//  - https://dev.azure.com/org/project/_wiki/wikis/wikiIdentifier/123/Title-Of-Page
function parseWikiUrl(url: string): ParsedWikiUrl {
  try {
    const u = new URL(url);
    const segments = u.pathname.split("/").filter(Boolean);
    const idx = segments.findIndex((s) => s === "_wiki");
    if (idx < 1 || segments[idx + 1] !== "wikis") {
      return { ok: false, error: "URL does not match expected wiki pattern (missing /_wiki/wikis/ segment)." };
    }
    const project = segments[idx - 1];
    const wikiIdentifier = segments[idx + 2];
    if (!project || !wikiIdentifier) {
      return { ok: false, error: "Could not extract project or wikiIdentifier from URL." };
    }

    const pagePathParam = u.searchParams.get("pagePath");
    if (pagePathParam) {
      let decoded = decodeURIComponent(pagePathParam);
      if (!decoded.startsWith("/")) {
        decoded = "/" + decoded;
      }
      return { ok: true, project, wikiIdentifier, pagePath: decoded };
    }

    const afterWiki = segments.slice(idx + 3);
    if (afterWiki.length >= 1) {
      const maybeId = parseInt(afterWiki[0], 10);
      if (!isNaN(maybeId)) {
        return { ok: true, project, wikiIdentifier, pageId: maybeId };
      }
    }

    return { ok: true, project, wikiIdentifier, pagePath: "/" };
  } catch {
    return { ok: false, error: "Invalid URL format." };
  }
}

function streamToString(stream: NodeJS.ReadableStream): Promise<string> {
  return new Promise((resolve, reject) => {
    let data = "";
    stream.setEncoding("utf8");
    stream.on("data", (chunk) => (data += chunk));
    stream.on("end", () => resolve(data));
    stream.on("error", reject);
  });
}

export { WIKI_TOOLS };
