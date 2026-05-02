// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { AccessToken } from "@azure/identity";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { WebApi } from "azure-devops-node-api";
import { z, ZodRawShape } from "zod";
import { apiVersion } from "../utils.js";
import { formatIdentityLines, formatPagination, formatProjectLines, formatTeamLines } from "../shared/format-utils.js";
import { getErrorToolResult, McpToolConfig, textToolResult, ToolHandler } from "../shared/tool-utils.js";

import type { ProjectInfo } from "azure-devops-node-api/interfaces/CoreInterfaces.js";
import { IdentityBase } from "azure-devops-node-api/interfaces/IdentitiesInterfaces.js";

const CORE_TOOLS = {
  list_azure_devops_teams_by_project: "list_azure_devops_teams_by_project",
  list_azure_devops_projects: "list_azure_devops_projects",
  get_azure_devops_identity_ids: "get_azure_devops_identity_ids",
};

type ToolConfigType = ReturnType<typeof listTeamsByProject> | ReturnType<typeof listProjects> | ReturnType<typeof getIdentityIds>;

function filterProjectsByName(projects: ProjectInfo[], projectNameFilter: string): ProjectInfo[] {
  const lowerCaseFilter = projectNameFilter.toLowerCase();
  return projects.filter((project) => project.name?.toLowerCase().includes(lowerCaseFilter));
}

export function configureCoreTools(server: McpServer, tokenProvider: () => Promise<AccessToken>, connectionProvider: () => Promise<WebApi>): McpToolConfig<ZodRawShape>[] {
  const coreTools: ToolConfigType[] = [];
  coreTools.push(listTeamsByProject(connectionProvider));
  coreTools.push(listProjects(connectionProvider));
  coreTools.push(getIdentityIds(tokenProvider, connectionProvider));
  return coreTools as unknown as McpToolConfig<ZodRawShape>[];
}

function listTeamsByProject(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    project: z.string().describe("The name or ID of the Azure DevOps project."),
    mine: z.boolean().optional().describe("If true, only return teams that the authenticated user is a member of."),
    top: z.number().optional().describe("The maximum number of teams to return. Defaults to 100."),
    skip: z.number().optional().describe("The number of teams to skip for pagination. Defaults to 0."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ project, mine, top, skip }) => {
    try {
      const effectiveTop = top ?? 100;
      const effectiveSkip = skip ?? 0;
      const connection = await connectionProvider();
      const coreApi = await connection.getCoreApi();
      const teams = (await coreApi.getTeams(project, mine, effectiveTop, effectiveSkip, false)) ?? [];

      if (teams.length === 0) {
        return textToolResult([`No teams found in project '${project}'.`]);
      }

      const header = `Teams in project '${project}':`;
      const pagination = formatPagination(teams.length, effectiveTop, effectiveSkip);
      return textToolResult([header, pagination, ...formatTeamLines(teams)]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to fetch project teams.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: CORE_TOOLS.list_azure_devops_teams_by_project,
    description: "List teams in an Azure DevOps project. Set 'mine' to restrict to the authenticated user's teams; use 'top'/'skip' to page.",
    inputSchema,
    handler,
  };

  return config;
}

function listProjects(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    stateFilter: z.enum(["all", "wellFormed", "createPending", "deleted"]).default("wellFormed").describe("Filter projects by their state. Defaults to 'wellFormed'."),
    top: z.number().optional().describe("The maximum number of projects to return. Defaults to 100."),
    skip: z.number().optional().describe("The number of projects to skip for pagination. Defaults to 0."),
    continuationToken: z.number().optional().describe("Continuation token for pagination. Used to fetch the next set of results if available."),
    projectNameFilter: z.string().optional().describe("Filter projects by name. Supports partial matches."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ stateFilter, top, skip, continuationToken, projectNameFilter }) => {
    try {
      const effectiveTop = top ?? 100;
      const effectiveSkip = skip ?? 0;
      const connection = await connectionProvider();
      const coreApi = await connection.getCoreApi();
      const projects = await coreApi.getProjects(stateFilter, effectiveTop, effectiveSkip, continuationToken, false);

      const filteredProjects = projects ? (projectNameFilter ? filterProjectsByName(projects, projectNameFilter) : projects) : [];
      const filterSuffix = projectNameFilter ? ` matching '${projectNameFilter}'` : "";

      if (filteredProjects.length === 0) {
        return textToolResult([`No projects found in your Azure DevOps organization${filterSuffix} (state: ${stateFilter}).`]);
      }

      const header = `Projects in your Azure DevOps organization${filterSuffix} (state: ${stateFilter}):`;
      const pagination = formatPagination(filteredProjects.length, effectiveTop, effectiveSkip);
      return textToolResult([header, pagination, ...formatProjectLines(filteredProjects)]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to fetch projects.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: CORE_TOOLS.list_azure_devops_projects,
    description: "List projects in the current Azure DevOps organization. Filter by lifecycle state (default 'wellFormed') or partial name match; use 'top'/'skip' to page.",
    inputSchema,
    handler,
  };

  return config;
}

function getIdentityIds(tokenProvider: () => Promise<AccessToken>, connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    searchFilter: z.string().describe("Search filter (unique name, display name, or email) to retrieve identity IDs for."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ searchFilter }) => {
    try {
      const token = await tokenProvider();
      const connection = await connectionProvider();
      const orgName = connection.serverUrl.split("/")[3];
      const baseUrl = `https://vssps.dev.azure.com/${orgName}/_apis/identities`;

      const params = new URLSearchParams({
        "api-version": apiVersion,
        "searchFilter": "General",
        "filterValue": searchFilter,
      });

      const response = await fetch(`${baseUrl}?${params}`, {
        headers: {
          "Authorization": `Bearer ${token.token}`,
          "Content-Type": "application/json",
        },
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`HTTP ${response.status}: ${errorText}`);
      }

      const identities = await response.json();
      const values: IdentityBase[] = identities?.value ?? [];

      if (values.length === 0) {
        return textToolResult([`No identities found matching '${searchFilter}'.`]);
      }

      const header = `Found ${values.length} identities matching '${searchFilter}':`;
      return textToolResult([header, ...formatIdentityLines(values)]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to fetch identities.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: CORE_TOOLS.get_azure_devops_identity_ids,
    description: "Resolve Azure DevOps identity IDs from a unique name, display name, or email. Pass IDs to other tools as reviewer, assignee, or permission target inputs.",
    inputSchema,
    handler,
  };

  return config;
}

export { CORE_TOOLS };
