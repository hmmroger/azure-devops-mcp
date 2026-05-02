// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { AccessToken } from "@azure/identity";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { WebApi } from "azure-devops-node-api";
import { z, ZodRawShape } from "zod";
import { TimeFrame, TeamSettingsIteration } from "azure-devops-node-api/interfaces/WorkInterfaces.js";
import { TreeStructureGroup, WorkItemClassificationNode } from "azure-devops-node-api/interfaces/WorkItemTrackingInterfaces.js";
import { getErrorToolResult, McpToolConfig, textToolResult, ToolHandler } from "../shared/tool-utils.js";

const WORK_TOOLS = {
  list_azure_devops_team_iterations: "list_azure_devops_team_iterations",
  create_azure_devops_iterations: "create_azure_devops_iterations",
  assign_azure_devops_iterations_to_team: "assign_azure_devops_iterations_to_team",
};

type ToolConfigType = ReturnType<typeof listTeamIterations> | ReturnType<typeof createIterations> | ReturnType<typeof assignIterationsToTeam>;

function toIsoDate(value: Date | string | undefined): string | undefined {
  if (!value) {
    return undefined;
  }
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) {
    return undefined;
  }
  return date.toISOString().slice(0, 10);
}

function formatTeamIteration(iteration: TeamSettingsIteration): string {
  if (!iteration.name) {
    throw new Error("Team iteration missing required 'name' field.");
  }
  const attrs: string[] = [];
  if (iteration.id) {
    attrs.push(`ID: ${iteration.id}`);
  }
  if (iteration.path) {
    attrs.push(`path: ${iteration.path}`);
  }
  const start = toIsoDate(iteration.attributes?.startDate);
  if (start) {
    attrs.push(`start: ${start}`);
  }
  const finish = toIsoDate(iteration.attributes?.finishDate);
  if (finish) {
    attrs.push(`finish: ${finish}`);
  }
  const timeFrame = iteration.attributes?.timeFrame;
  if (timeFrame !== undefined) {
    attrs.push(`timeFrame: ${TimeFrame[timeFrame] ?? timeFrame}`);
  }
  return attrs.length > 0 ? `${iteration.name} (${attrs.join(", ")})` : iteration.name;
}

function formatTeamIterationLines(iterations: TeamSettingsIteration[]): string[] {
  return iterations.map((i) => `- ${formatTeamIteration(i)}`);
}

function formatClassificationNode(node: WorkItemClassificationNode): string {
  if (!node.name) {
    throw new Error("Classification node missing required 'name' field.");
  }
  const attrs: string[] = [];
  if (node.identifier) {
    attrs.push(`identifier: ${node.identifier}`);
  }
  if (node.id !== undefined) {
    attrs.push(`ID: ${node.id}`);
  }
  if (node.path) {
    attrs.push(`path: ${node.path}`);
  }
  const start = toIsoDate(node.attributes?.startDate);
  if (start) {
    attrs.push(`start: ${start}`);
  }
  const finish = toIsoDate(node.attributes?.finishDate);
  if (finish) {
    attrs.push(`finish: ${finish}`);
  }
  return attrs.length > 0 ? `${node.name} (${attrs.join(", ")})` : node.name;
}

function formatClassificationNodeLines(nodes: WorkItemClassificationNode[]): string[] {
  return nodes.map((n) => `- ${formatClassificationNode(n)}`);
}

export function configureWorkTools(server: McpServer, tokenProvider: () => Promise<AccessToken>, connectionProvider: () => Promise<WebApi>): McpToolConfig<ZodRawShape>[] {
  const workTools: ToolConfigType[] = [];
  workTools.push(listTeamIterations(connectionProvider));
  workTools.push(createIterations(connectionProvider));
  workTools.push(assignIterationsToTeam(connectionProvider));
  return workTools as unknown as McpToolConfig<ZodRawShape>[];
}

function listTeamIterations(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    project: z.string().describe("The name or ID of the Azure DevOps project."),
    team: z.string().describe("The name or ID of the Azure DevOps team."),
    timeframe: z.enum(["current"]).optional().describe("If set to 'current', limits results to the active iteration."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ project, team, timeframe }) => {
    try {
      const connection = await connectionProvider();
      const workApi = await connection.getWorkApi();
      const iterations = (await workApi.getTeamIterations({ project, team }, timeframe)) ?? [];

      const scope = timeframe ? ` (timeframe: ${timeframe})` : "";
      if (iterations.length === 0) {
        return textToolResult([`No iterations found for team '${team}' in project '${project}'${scope}.`]);
      }

      const header = `Iterations for team '${team}' in project '${project}'${scope}:`;
      return textToolResult([header, ...formatTeamIterationLines(iterations)]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to fetch team iterations.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: WORK_TOOLS.list_azure_devops_team_iterations,
    description: "List iterations assigned to a team. Use 'timeframe=current' to limit to the active sprint.",
    inputSchema,
    handler,
  };

  return config;
}

function createIterations(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    project: z.string().describe("The name or ID of the Azure DevOps project."),
    iterations: z
      .array(
        z.object({
          iterationName: z.string().describe("Name of the iteration to create."),
          startDate: z.string().optional().describe("Start date in ISO format (e.g. '2026-01-01T00:00:00Z')."),
          finishDate: z.string().optional().describe("Finish date in ISO format (e.g. '2026-01-14T23:59:59Z')."),
        })
      )
      .describe("Iterations to create under the project's iteration tree."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ project, iterations }) => {
    try {
      const connection = await connectionProvider();
      const workItemTrackingApi = await connection.getWorkItemTrackingApi();
      const created: WorkItemClassificationNode[] = [];

      for (const { iterationName, startDate, finishDate } of iterations) {
        const node = await workItemTrackingApi.createOrUpdateClassificationNode(
          {
            name: iterationName,
            attributes: {
              startDate: startDate ? new Date(startDate) : undefined,
              finishDate: finishDate ? new Date(finishDate) : undefined,
            },
          },
          project,
          TreeStructureGroup.Iterations
        );

        if (node) {
          created.push(node);
        }
      }

      if (created.length === 0) {
        return textToolResult([`No iterations were created in project '${project}'.`], true);
      }

      const header = `Created ${created.length} iteration${created.length === 1 ? "" : "s"} in project '${project}':`;
      return textToolResult([header, ...formatClassificationNodeLines(created)]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to create iterations.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: WORK_TOOLS.create_azure_devops_iterations,
    description: "Create iterations in a project's iteration tree. Optional ISO start/finish dates set the iteration window.",
    inputSchema,
    handler,
  };

  return config;
}

function assignIterationsToTeam(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    project: z.string().describe("The name or ID of the Azure DevOps project."),
    team: z.string().describe("The name or ID of the Azure DevOps team."),
    iterations: z
      .array(
        z.object({
          identifier: z.string().describe("Identifier (GUID) of an existing project iteration."),
          path: z.string().describe("Iteration path, e.g. 'Project\\Iteration\\Sprint 1'."),
        })
      )
      .describe("Iterations to assign to the team's backlog."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ project, team, iterations }) => {
    try {
      const connection = await connectionProvider();
      const workApi = await connection.getWorkApi();
      const teamContext = { project, team };
      const assigned: TeamSettingsIteration[] = [];

      for (const { identifier, path } of iterations) {
        const assignment = await workApi.postTeamIteration({ id: identifier, path }, teamContext);
        if (assignment) {
          assigned.push(assignment);
        }
      }

      if (assigned.length === 0) {
        return textToolResult([`No iterations were assigned to team '${team}' in project '${project}'.`], true);
      }

      const header = `Assigned ${assigned.length} iteration${assigned.length === 1 ? "" : "s"} to team '${team}' in project '${project}':`;
      return textToolResult([header, ...formatTeamIterationLines(assigned)]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to assign iterations to team.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: WORK_TOOLS.assign_azure_devops_iterations_to_team,
    description: "Assign existing project iterations to a team's backlog by iteration identifier and path.",
    inputSchema,
    handler,
  };

  return config;
}

export { WORK_TOOLS };
