// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { IGitApi } from "azure-devops-node-api/GitApi.js";
import { z } from "zod";
import { validate } from "uuid";
import { apiVersion } from "../utils.js";
import { VersionControlRecursionType } from "azure-devops-node-api/interfaces/GitInterfaces.js";
import { GitItem } from "azure-devops-node-api/interfaces/GitInterfaces.js";
import { AzureDevOpsClientManager } from "azure-client.js";

export interface CodeSearchFilters {
  Project?: string[];
  Repository?: string[];
  Path?: string[];
  Branch?: string[];
  CodeElement?: string[];
}

export interface CodeSearchRequest {
  searchText: string;
  $skip: number;
  $top: number;
  filters?: CodeSearchFilters;
  includeFacets?: boolean;
}

export const SEARCH_TOOLS = {
  search_azure_devops_code: "search_azure_devops_code",
  search_azure_devops_wiki: "search_azure_devops_wiki",
  search_azure_devops_workitem: "search_azure_devops_workitem",
};

export async function performCodeSearch(searchRequest: CodeSearchRequest, adoManager: AzureDevOpsClientManager, projectFilter?: string, repoFilter?: string, pathFilter?: string): Promise<string> {
  const accessToken = await adoManager.getToken();
  const connection = await adoManager.getClient();
  const orgName = await adoManager.getOrgName();
  const gitApi = await connection.getGitApi();
  const url = `https://almsearch.dev.azure.com/${orgName}/_apis/search/codesearchresults?api-version=${apiVersion}`;

  // ID check
  if (projectFilter && validate(projectFilter)) {
    const coreApi = await connection.getCoreApi();
    const project = await coreApi.getProject(projectFilter);
    projectFilter = project.name;
  }

  if (repoFilter && validate(repoFilter)) {
    const repo = await gitApi.getRepository(repoFilter);
    repoFilter = repo.name;
    // auto-recover project constraint
    if (repo.project) {
      projectFilter = repo.project.name;
    }
  }

  if (repoFilter && !projectFilter) {
    throw new Error(`project filter is reuiqred when repository filter is used.`);
  } else if (pathFilter && (!repoFilter || !projectFilter)) {
    throw new Error(`Both project and repository filter are reuiqred when path filter is used.`);
  }

  if (repoFilter || projectFilter || pathFilter) {
    searchRequest.filters = {
      Project: projectFilter ? [projectFilter] : undefined,
      Repository: repoFilter ? [repoFilter] : undefined,
      Path: pathFilter ? [pathFilter] : undefined,
    };
  }

  const response = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${accessToken.token}`,
      "User-Agent": adoManager.getUserAgent(),
    },
    body: JSON.stringify(searchRequest),
  });

  if (!response.ok) {
    throw new Error(`Azure DevOps Code Search API error: ${response.status} ${response.statusText}`);
  }

  const resultText = await response.text();
  const resultJson = JSON.parse(resultText) as { results?: SearchResult[] };

  const topResults: SearchResult[] = Array.isArray(resultJson.results) ? resultJson.results.slice(0, Math.min(searchRequest.$top, resultJson.results.length)) : [];

  const combinedResults = await fetchCombinedResults(topResults, gitApi);

  return resultText + JSON.stringify(combinedResults);
}

export function configureSearchTools(server: McpServer, adoManager: AzureDevOpsClientManager, disabledTools: Set<string>) {
  /**
   * CODE SEARCH
   * Get the code search results for a given search text.
   */
  if (!disabledTools.has(SEARCH_TOOLS.search_azure_devops_code)) {
    server.tool(
      SEARCH_TOOLS.search_azure_devops_code,
      "Get the code search results for a given search text.",
      {
        searchText: z.string().describe("Search text to find in code"),
        skip: z.number().default(0).describe("Number of results to skip (for pagination)"),
        top: z.number().default(5).describe("Number of results to return (for pagination)"),
        projectName: z.string().optional().describe("Filter search results in this project name."),
        repositoryName: z.string().optional().describe("Filter search results in this repository name."),
        path: z.string().optional().describe("Filter search results under this item path."),
      },
      async ({ searchText, skip, top, projectName, repositoryName, path }) => {
        try {
          const result = await performCodeSearch({ searchText, $skip: skip, $top: top }, adoManager, projectName, repositoryName, path);
          return {
            content: [{ type: "text", text: result }],
          };
        } catch (error) {
          const message = (error as Error).message || "Code search failed.";
          return {
            content: [{ type: "text", text: message }],
            isError: true,
          };
        }
      }
    );
  }

  /**
   * WIKI SEARCH
   * Get wiki search results for a given search text.
   */
  if (!disabledTools.has(SEARCH_TOOLS.search_azure_devops_wiki)) {
    server.tool(
      SEARCH_TOOLS.search_azure_devops_wiki,
      "Get wiki search results for a given search text.",
      {
        searchRequest: z
          .object({
            searchText: z.string().describe("Search text to find in wikis"),
            $skip: z.number().default(0).describe("Number of results to skip (for pagination)"),
            $top: z.number().default(10).describe("Number of results to return (for pagination)"),
            filters: z
              .object({
                Project: z.array(z.string()).optional().describe("Filter in these projects"),
                Wiki: z.array(z.string()).optional().describe("Filter in these wiki names"),
              })
              .partial()
              .optional()
              .describe("Filters to apply to the search text"),
            includeFacets: z.boolean().optional(),
          })
          .strict(),
      },
      async ({ searchRequest }) => {
        const accessToken = await adoManager.getToken();
        await adoManager.getClient(); // for org name check
        const orgName = await adoManager.getOrgName();
        const url = `https://almsearch.dev.azure.com/${orgName}/_apis/search/wikisearchresults?api-version=${apiVersion}`;

        const response = await fetch(url, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            "Authorization": `Bearer ${accessToken.token}`,
            "User-Agent": adoManager.getUserAgent(),
          },
          body: JSON.stringify(searchRequest),
        });

        if (!response.ok) {
          throw new Error(`Azure DevOps Wiki Search API error: ${response.status} ${response.statusText}`);
        }

        const result = await response.text();
        return {
          content: [{ type: "text", text: result }],
        };
      }
    );
  }

  /**
   * WORK ITEM SEARCH
   * Get work item search results for a given search text.
   */
  if (!disabledTools.has(SEARCH_TOOLS.search_azure_devops_workitem)) {
    server.tool(
      SEARCH_TOOLS.search_azure_devops_workitem,
      "Get work item search results for a given search text.",
      {
        searchRequest: z
          .object({
            searchText: z.string().describe("Search text to find in work items"),
            $skip: z.number().default(0).describe("Number of results to skip for pagination"),
            $top: z.number().default(10).describe("Number of results to return"),
            filters: z
              .object({
                "System.TeamProject": z.array(z.string()).optional().describe("Filter by team project"),
                "System.AreaPath": z.array(z.string()).optional().describe("Filter by area path"),
                "System.WorkItemType": z.array(z.string()).optional().describe("Filter by work item type like Bug, Task, User Story"),
                "System.State": z.array(z.string()).optional().describe("Filter by state"),
                "System.AssignedTo": z.array(z.string()).optional().describe("Filter by assigned to"),
              })
              .partial()
              .optional(),
            includeFacets: z.boolean().optional(),
          })
          .strict(),
      },
      async ({ searchRequest }) => {
        const accessToken = await adoManager.getToken();
        await adoManager.getClient(); // for org name check
        const orgName = await adoManager.getOrgName();
        const url = `https://almsearch.dev.azure.com/${orgName}/_apis/search/workitemsearchresults?api-version=${apiVersion}`;

        const response = await fetch(url, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            "Authorization": `Bearer ${accessToken.token}`,
            "User-Agent": adoManager.getUserAgent(),
          },
          body: JSON.stringify(searchRequest),
        });

        if (!response.ok) {
          throw new Error(`Azure DevOps Work Item Search API error: ${response.status} ${response.statusText}`);
        }

        const result = await response.text();
        return {
          content: [{ type: "text", text: result }],
        };
      }
    );
  }
}

interface SearchResult {
  project?: { id?: string };
  repository?: { id?: string };
  path?: string;
  versions?: { changeId?: string }[];
  [key: string]: unknown;
}

type CombinedResult = { gitItem: GitItem } | { error: string };

async function fetchCombinedResults(topSearchResults: SearchResult[], gitApi: IGitApi): Promise<CombinedResult[]> {
  const combinedResults: CombinedResult[] = [];
  for (const searchResult of topSearchResults) {
    try {
      const projectId = searchResult.project?.id;
      const repositoryId = searchResult.repository?.id;
      const filePath = searchResult.path;
      const changeId = Array.isArray(searchResult.versions) && searchResult.versions.length > 0 ? searchResult.versions[0].changeId : undefined;
      if (!projectId || !repositoryId || !filePath || !changeId) {
        combinedResults.push({
          error: `Missing projectId, repositoryId, filePath, or changeId in the result: ${JSON.stringify(searchResult)}`,
        });
        continue;
      }

      const versionDescriptor = changeId ? { version: changeId, versionType: 2, versionOptions: 0 } : undefined;

      const item = await gitApi.getItem(
        repositoryId,
        filePath,
        projectId,
        undefined,
        VersionControlRecursionType.None,
        true, // includeContentMetadata
        false, // latestProcessedChange
        false, // download
        versionDescriptor,
        true, // includeContent
        true, // resolveLfs
        true // sanitize
      );
      combinedResults.push({
        gitItem: item,
      });
    } catch (err) {
      combinedResults.push({
        error: err instanceof Error ? err.message : String(err),
      });
    }
  }
  return combinedResults;
}
