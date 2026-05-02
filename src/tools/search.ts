// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { z, ZodRawShape } from "zod";
import { apiVersion } from "../utils.js";
import { AzureDevOpsClientManager } from "azure-client.js";
import { formatPagination } from "../shared/format-utils.js";
import { getErrorToolResult, McpToolConfig, textToolResult, ToolHandler } from "../shared/tool-utils.js";

export const SEARCH_TOOLS = {
  search_azure_devops_code: "search_azure_devops_code",
  search_azure_devops_wikis: "search_azure_devops_wikis",
  search_azure_devops_work_items: "search_azure_devops_work_items",
};

type ToolConfigType = ReturnType<typeof searchCode> | ReturnType<typeof searchWikis> | ReturnType<typeof searchWorkItems>;

interface CodeSearchFilters {
  Project?: string[];
  Repository?: string[];
  Path?: string[];
  Branch?: string[];
  CodeElement?: string[];
}

interface CodeSearchRequest {
  searchText: string;
  $skip: number;
  $top: number;
  filters?: CodeSearchFilters;
  includeFacets?: boolean;
}

interface WikiSearchFilters {
  Project?: string[];
  Wiki?: string[];
}

interface WikiSearchRequest {
  searchText: string;
  $skip: number;
  $top: number;
  filters?: WikiSearchFilters;
  includeFacets?: boolean;
}

interface WorkItemSearchFilters {
  "System.TeamProject"?: string[];
  "System.AreaPath"?: string[];
  "System.WorkItemType"?: string[];
  "System.State"?: string[];
  "System.AssignedTo"?: string[];
}

interface WorkItemSearchRequest {
  searchText: string;
  $skip: number;
  $top: number;
  filters?: WorkItemSearchFilters;
  includeFacets?: boolean;
}

interface CodeSearchResultItem {
  fileName?: string;
  path?: string;
  matches?: { content?: unknown[]; fileName?: unknown[] };
  project?: { id?: string; name?: string };
  repository?: { id?: string; name?: string };
  versions?: { branchName?: string; changeId?: string }[];
  contentId?: string;
}

interface WikiSearchResultItem {
  fileName?: string;
  path?: string;
  hits?: { fieldReferenceName?: string; highlights?: string[] }[];
  project?: { id?: string; name?: string };
  wiki?: { id?: string; name?: string; mappedPath?: string; version?: string };
}

interface WorkItemSearchResultItem {
  fields?: Record<string, unknown>;
  hits?: { fieldReferenceName?: string; highlights?: string[] }[];
  project?: { id?: string; name?: string };
}

interface SearchResponse<T> {
  count?: number;
  results?: T[];
}

function buildFilterDescription(filters: Record<string, string | string[] | undefined>): string {
  const parts: string[] = [];
  for (const [key, value] of Object.entries(filters)) {
    if (Array.isArray(value) && value.length > 0) {
      parts.push(`${key}=${value.join("|")}`);
    } else if (typeof value === "string" && value.length > 0) {
      parts.push(`${key}='${value}'`);
    }
  }
  return parts.length > 0 ? ` (filters: ${parts.join(", ")})` : "";
}

function formatCodeSearchResult(result: CodeSearchResultItem): string[] {
  const lines: string[] = [];
  const fileName = result.fileName ?? "(no fileName)";
  const path = result.path ?? "(no path)";
  const projectName = result.project?.name;
  const repoName = result.repository?.name;
  const branch = result.versions?.[0]?.branchName;
  const changeId = result.versions?.[0]?.changeId;

  const attrs: string[] = [];
  if (projectName) {
    attrs.push(`project: ${projectName}`);
  }
  if (repoName) {
    attrs.push(`repo: ${repoName}`);
  }
  if (branch) {
    attrs.push(`branch: ${branch}`);
  }
  if (changeId) {
    attrs.push(`changeId: ${changeId}`);
  }
  if (result.contentId) {
    attrs.push(`contentId: ${result.contentId}`);
  }

  lines.push(`- ${fileName} at ${path}${attrs.length > 0 ? ` (${attrs.join(", ")})` : ""}`);

  const contentMatches = result.matches?.content?.length ?? 0;
  const fileNameMatches = result.matches?.fileName?.length ?? 0;
  if (contentMatches > 0 || fileNameMatches > 0) {
    lines.push(`  matches: ${contentMatches} content, ${fileNameMatches} fileName`);
  }
  return lines;
}

function formatHighlightsLine(hits: { fieldReferenceName?: string; highlights?: string[] }[] | undefined): string | undefined {
  if (!hits || hits.length === 0) {
    return undefined;
  }
  const fields = hits.map((h) => h.fieldReferenceName).filter((n): n is string => !!n);
  if (fields.length === 0) {
    return undefined;
  }
  return `  hit fields: ${fields.join(", ")}`;
}

function formatWikiSearchResult(result: WikiSearchResultItem): string[] {
  const lines: string[] = [];
  const fileName = result.fileName ?? "(no fileName)";
  const path = result.path ?? "(no path)";
  const projectName = result.project?.name;
  const wikiName = result.wiki?.name;
  const wikiId = result.wiki?.id;
  const mappedPath = result.wiki?.mappedPath;

  const attrs: string[] = [];
  if (projectName) {
    attrs.push(`project: ${projectName}`);
  }
  if (wikiName) {
    attrs.push(`wiki: ${wikiName}`);
  }
  if (wikiId) {
    attrs.push(`wikiId: ${wikiId}`);
  }
  if (mappedPath) {
    attrs.push(`mappedPath: ${mappedPath}`);
  }

  lines.push(`- ${fileName} at ${path}${attrs.length > 0 ? ` (${attrs.join(", ")})` : ""}`);
  const hitsLine = formatHighlightsLine(result.hits);
  if (hitsLine) {
    lines.push(hitsLine);
  }
  return lines;
}

function formatWorkItemSearchResult(result: WorkItemSearchResultItem): string[] {
  const fields = result.fields ?? {};
  const id = typeof fields["System.Id"] === "number" || typeof fields["System.Id"] === "string" ? fields["System.Id"] : undefined;
  const type = typeof fields["System.WorkItemType"] === "string" ? (fields["System.WorkItemType"] as string) : undefined;
  const title = typeof fields["System.Title"] === "string" ? (fields["System.Title"] as string) : "(no title)";
  const state = typeof fields["System.State"] === "string" ? (fields["System.State"] as string) : undefined;
  const assignedToRaw = fields["System.AssignedTo"];
  let assignedTo: string | undefined;
  if (typeof assignedToRaw === "string") {
    assignedTo = assignedToRaw;
  }
  const projectName = result.project?.name;

  const attrs: string[] = [];
  if (state) {
    attrs.push(`state: ${state}`);
  }
  if (projectName) {
    attrs.push(`project: ${projectName}`);
  }
  if (assignedTo) {
    attrs.push(`assignedTo: ${assignedTo}`);
  }

  const idLabel = id !== undefined ? `#${id}` : "#?";
  const typeLabel = type ?? "(no type)";
  const lines: string[] = [`- ${idLabel} ${typeLabel} "${title}"${attrs.length > 0 ? ` (${attrs.join(", ")})` : ""}`];
  const hitsLine = formatHighlightsLine(result.hits);
  if (hitsLine) {
    lines.push(hitsLine);
  }
  return lines;
}

async function postSearch<T>(adoManager: AzureDevOpsClientManager, endpoint: string, body: unknown, errorPrefix: string): Promise<SearchResponse<T>> {
  const accessToken = await adoManager.getToken();
  const orgName = await adoManager.getOrgName();
  const url = `https://almsearch.dev.azure.com/${orgName}/_apis/search/${endpoint}?api-version=${apiVersion}`;

  const response = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${accessToken.token}`,
      "User-Agent": adoManager.getUserAgent(),
    },
    body: JSON.stringify(body),
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`${errorPrefix} (${response.status}): ${errorText}`);
  }

  return (await response.json()) as SearchResponse<T>;
}

export function configureSearchTools(adoManager: AzureDevOpsClientManager): McpToolConfig<ZodRawShape>[] {
  const searchTools: ToolConfigType[] = [];
  searchTools.push(searchCode(adoManager));
  searchTools.push(searchWikis(adoManager));
  searchTools.push(searchWorkItems(adoManager));
  return searchTools as unknown as McpToolConfig<ZodRawShape>[];
}

function searchCode(adoManager: AzureDevOpsClientManager) {
  const inputSchema = {
    searchText: z.string().describe("Text to search for in code."),
    skip: z.number().optional().describe("Number of results to skip. Defaults to 0."),
    top: z.number().optional().describe("Maximum results to return. Defaults to 5."),
    projectId: z.string().uuid().optional().describe("Project ID (GUID) to filter to. Discover via 'list_azure_devops_projects'."),
    repositoryId: z.string().uuid().optional().describe("Repository ID (GUID) to filter to. Discover via 'get_azure_devops_repositories'. Pairs with projectId."),
    path: z.string().optional().describe("Path prefix filter (e.g. '/src'). Requires projectId and repositoryId."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ searchText, skip, top, projectId, repositoryId, path }) => {
    try {
      const effectiveSkip = skip ?? 0;
      const effectiveTop = top ?? 5;
      const connection = await adoManager.getClient();

      let projectName: string | undefined;
      let repoName: string | undefined;

      if (projectId) {
        const coreApi = await connection.getCoreApi();
        const project = await coreApi.getProject(projectId);
        if (!project?.name) {
          return textToolResult([`Project '${projectId}' not found.`], true);
        }
        projectName = project.name;
      }

      if (repositoryId) {
        const gitApi = await connection.getGitApi();
        const repo = await gitApi.getRepository(repositoryId);
        if (!repo?.name) {
          return textToolResult([`Repository '${repositoryId}' not found.`], true);
        }
        repoName = repo.name;
        if (!projectName && repo.project?.name) {
          projectName = repo.project.name;
        }
      }

      if (repoName && !projectName) {
        return textToolResult(["projectId is required when repositoryId is used."], true);
      }

      if (path && (!repoName || !projectName)) {
        return textToolResult(["Both projectId and repositoryId are required when path is used."], true);
      }

      const request: CodeSearchRequest = { searchText, $skip: effectiveSkip, $top: effectiveTop };
      if (projectName || repoName || path) {
        request.filters = {
          Project: projectName ? [projectName] : undefined,
          Repository: repoName ? [repoName] : undefined,
          Path: path ? [path] : undefined,
        };
      }

      const json = await postSearch<CodeSearchResultItem>(adoManager, "codesearchresults", request, "Azure DevOps Code Search API error");
      const results = json.results ?? [];
      const total = json.count;
      const filterDesc = buildFilterDescription({ project: projectName, repo: repoName, path });

      if (results.length === 0) {
        return textToolResult([`No code search results for '${searchText}'${filterDesc}.`]);
      }

      const header = `Code search '${searchText}'${filterDesc}:`;
      const pagination = formatPagination(results.length, effectiveTop, effectiveSkip, total);
      const lines: string[] = [header, pagination];
      for (const r of results) {
        lines.push(...formatCodeSearchResult(r));
      }
      return textToolResult(lines);
    } catch (error) {
      return getErrorToolResult(error, "Code search failed.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: SEARCH_TOOLS.search_azure_devops_code,
    description: "Search across Azure DevOps code. Returns file paths and match counts; fetch content via 'get_azure_devops_item_content_by_commit'.",
    inputSchema,
    handler,
  };

  return config;
}

function searchWikis(adoManager: AzureDevOpsClientManager) {
  const inputSchema = {
    searchText: z.string().describe("Text to search for in wikis."),
    skip: z.number().optional().describe("Number of results to skip. Defaults to 0."),
    top: z.number().optional().describe("Maximum results to return. Defaults to 10."),
    projectFilter: z.array(z.string()).optional().describe("Project names to restrict the search to."),
    wikiFilter: z.array(z.string()).optional().describe("Wiki names to restrict the search to."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ searchText, skip, top, projectFilter, wikiFilter }) => {
    try {
      const effectiveSkip = skip ?? 0;
      const effectiveTop = top ?? 10;

      const request: WikiSearchRequest = { searchText, $skip: effectiveSkip, $top: effectiveTop };
      if ((projectFilter && projectFilter.length > 0) || (wikiFilter && wikiFilter.length > 0)) {
        request.filters = {
          Project: projectFilter && projectFilter.length > 0 ? projectFilter : undefined,
          Wiki: wikiFilter && wikiFilter.length > 0 ? wikiFilter : undefined,
        };
      }

      const json = await postSearch<WikiSearchResultItem>(adoManager, "wikisearchresults", request, "Azure DevOps Wiki Search API error");
      const results = json.results ?? [];
      const total = json.count;
      const filterDesc = buildFilterDescription({ project: projectFilter, wiki: wikiFilter });

      if (results.length === 0) {
        return textToolResult([`No wiki search results for '${searchText}'${filterDesc}.`]);
      }

      const header = `Wiki search '${searchText}'${filterDesc}:`;
      const pagination = formatPagination(results.length, effectiveTop, effectiveSkip, total);
      const lines: string[] = [header, pagination];
      for (const r of results) {
        lines.push(...formatWikiSearchResult(r));
      }
      return textToolResult(lines);
    } catch (error) {
      return getErrorToolResult(error, "Wiki search failed.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: SEARCH_TOOLS.search_azure_devops_wikis,
    description: "Search across Azure DevOps wikis. Returns page paths and hit fields; fetch content via 'get_azure_devops_wiki_page_content'.",
    inputSchema,
    handler,
  };

  return config;
}

function searchWorkItems(adoManager: AzureDevOpsClientManager) {
  const inputSchema = {
    searchText: z.string().describe("Text to search for in work items."),
    skip: z.number().optional().describe("Number of results to skip. Defaults to 0."),
    top: z.number().optional().describe("Maximum results to return. Defaults to 10."),
    projectFilter: z.array(z.string()).optional().describe("System.TeamProject values to restrict to."),
    areaPathFilter: z.array(z.string()).optional().describe("System.AreaPath values to restrict to."),
    workItemTypeFilter: z.array(z.string()).optional().describe("Work item types (e.g. 'Bug', 'Task', 'User Story')."),
    stateFilter: z.array(z.string()).optional().describe("Work item states to restrict to."),
    assignedToFilter: z.array(z.string()).optional().describe("AssignedTo values to restrict to."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ searchText, skip, top, projectFilter, areaPathFilter, workItemTypeFilter, stateFilter, assignedToFilter }) => {
    try {
      const effectiveSkip = skip ?? 0;
      const effectiveTop = top ?? 10;

      const filters: WorkItemSearchFilters = {};
      if (projectFilter && projectFilter.length > 0) {
        filters["System.TeamProject"] = projectFilter;
      }
      if (areaPathFilter && areaPathFilter.length > 0) {
        filters["System.AreaPath"] = areaPathFilter;
      }
      if (workItemTypeFilter && workItemTypeFilter.length > 0) {
        filters["System.WorkItemType"] = workItemTypeFilter;
      }
      if (stateFilter && stateFilter.length > 0) {
        filters["System.State"] = stateFilter;
      }
      if (assignedToFilter && assignedToFilter.length > 0) {
        filters["System.AssignedTo"] = assignedToFilter;
      }

      const request: WorkItemSearchRequest = { searchText, $skip: effectiveSkip, $top: effectiveTop };
      if (Object.keys(filters).length > 0) {
        request.filters = filters;
      }

      const json = await postSearch<WorkItemSearchResultItem>(adoManager, "workitemsearchresults", request, "Azure DevOps Work Item Search API error");
      const results = json.results ?? [];
      const total = json.count;
      const filterDesc = buildFilterDescription({
        project: projectFilter,
        areaPath: areaPathFilter,
        type: workItemTypeFilter,
        state: stateFilter,
        assignedTo: assignedToFilter,
      });

      if (results.length === 0) {
        return textToolResult([`No work item search results for '${searchText}'${filterDesc}.`]);
      }

      const header = `Work item search '${searchText}'${filterDesc}:`;
      const pagination = formatPagination(results.length, effectiveTop, effectiveSkip, total);
      const lines: string[] = [header, pagination];
      for (const r of results) {
        lines.push(...formatWorkItemSearchResult(r));
      }
      return textToolResult(lines);
    } catch (error) {
      return getErrorToolResult(error, "Work item search failed.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: SEARCH_TOOLS.search_azure_devops_work_items,
    description: "Search across Azure DevOps work items. Returns id/type/title/state plus hit fields; fetch full data via 'get_azure_devops_work_items_batch_by_ids'.",
    inputSchema,
    handler,
  };

  return config;
}
