// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { configureCoreTools, CORE_TOOLS } from "./tools/core.js";
import { configureWorkTools, WORK_TOOLS } from "./tools/work.js";
import { configureRepoTools, REPO_TOOLS } from "./tools/repos.js";
import { configureWorkItemTools, WORKITEM_TOOLS } from "./tools/workitems.js";
import { configureWikiTools, WIKI_TOOLS } from "./tools/wiki.js";
import { configureSearchTools, SEARCH_TOOLS } from "./tools/search.js";
import { AzureDevOpsClientManager } from "azure-client.js";
import { z, ZodRawShape } from "zod";
import { getErrorToolResult, McpToolConfig, textToolResult, ToolHandler } from "./shared/tool-utils.js";

const TOOLS_CATEGORY_MAP = new Map<string, object>([
  ["category_core", CORE_TOOLS],
  ["category_work", WORK_TOOLS],
  ["category_repo", REPO_TOOLS],
  ["category_workitem", WORKITEM_TOOLS],
  ["category_wiki", WIKI_TOOLS],
  ["category_search", SEARCH_TOOLS],
]);

function expandToolsAndCategories(entries: string): string[] {
  return entries
    .split(",")
    .map((entry) => entry.trim())
    .filter((entry) => entry.length > 0)
    .flatMap((entry) => {
      const category = TOOLS_CATEGORY_MAP.get(entry);
      if (category) {
        return Object.values(category) as string[];
      }
      return [entry];
    });
}

export async function configureAllTools(server: McpServer, adoManager: AzureDevOpsClientManager) {
  const disabledTools = new Set<string>(expandToolsAndCategories(process.env.ADO_MCP_DISABLED_TOOLS ?? ""));
  for (const tool of expandToolsAndCategories(process.env.ADO_MCP_ENABLED_TOOLS ?? "")) {
    disabledTools.delete(tool);
  }

  const tokenProvider = () => adoManager.getToken();
  const connectionProvider = adoManager.getClientFactory();
  const userAgentProvider = () => adoManager.getUserAgent();

  const tools: McpToolConfig<ZodRawShape>[] = [];
  tools.push(changeAzureDevOpsOrg(adoManager) as unknown as McpToolConfig<ZodRawShape>);
  tools.push(getAzureDevOpsOrg(adoManager) as unknown as McpToolConfig<ZodRawShape>);

  const coreTools = configureCoreTools(server, tokenProvider, connectionProvider);
  tools.push(...coreTools);

  const repoTools = configureRepoTools(server, tokenProvider, connectionProvider);
  tools.push(...repoTools);

  const workTools = configureWorkTools(server, tokenProvider, connectionProvider);
  tools.push(...workTools);

  const wikiTools = configureWikiTools(server, tokenProvider, connectionProvider, userAgentProvider);
  tools.push(...wikiTools);

  const workItemTools = configureWorkItemTools(server, tokenProvider, connectionProvider, userAgentProvider);
  tools.push(...workItemTools);

  const searchTools = configureSearchTools(adoManager);
  tools.push(...searchTools);

  for (const tool of tools) {
    if (disabledTools.has(tool.name)) {
      continue;
    }

    server.registerTool(tool.name, { description: tool.description, inputSchema: tool.inputSchema, annotations: tool.annotations }, tool.handler);
  }
}

function changeAzureDevOpsOrg(adoManager: AzureDevOpsClientManager) {
  const inputSchema = {
    organization: z.string().describe("Azure DevOps organization name."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ organization }) => {
    try {
      await adoManager.setOrgName(organization);
      return textToolResult([`Organization set to: ${organization}`]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to change organization.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: "change_azure_devops_org",
    description: "Change current Azure DevOps organization name.",
    inputSchema,
    handler,
  };

  return config;
}

function getAzureDevOpsOrg(adoManager: AzureDevOpsClientManager) {
  const inputSchema = {};
  const handler: ToolHandler<typeof inputSchema> = async () => {
    try {
      const name = await adoManager.getOrgName();
      return textToolResult([name ? `Current organization: ${name}` : "No Azure DevOps organization set. Use 'change_azure_devops_org' to set one."]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to get organization.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: "get_azure_devops_org",
    description: "Get current Azure DevOps organization name.",
    inputSchema,
    handler,
  };

  return config;
}
