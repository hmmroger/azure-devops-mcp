// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { AccessToken } from "@azure/identity";
import { WebApi } from "azure-devops-node-api";

import { configureCoreTools, CORE_TOOLS } from "./tools/core.js";
import { configureWorkTools, WORK_TOOLS } from "./tools/work.js";
import { configureBuildTools, BUILD_TOOLS } from "./tools/builds.js";
import { configureRepoTools, REPO_TOOLS } from "./tools/repos.js";
import { configureWorkItemTools, WORKITEM_TOOLS } from "./tools/workitems.js";
import { configureReleaseTools, RELEASE_TOOLS } from "./tools/releases.js";
import { configureWikiTools, WIKI_TOOLS } from "./tools/wiki.js";
import { configureTestPlanTools, Test_Plan_Tools } from "./tools/testplans.js";
import { configureSearchTools, SEARCH_TOOLS } from "./tools/search.js";

const TOOLS_CATEGORY_MAP = new Map<string, object>([
  ["category_core", CORE_TOOLS],
  ["category_work", WORK_TOOLS],
  ["category_build", BUILD_TOOLS],
  ["category_repo", REPO_TOOLS],
  ["category_workitem", WORKITEM_TOOLS],
  ["category_release", RELEASE_TOOLS],
  ["category_wiki", WIKI_TOOLS],
  ["category_testplan", Test_Plan_Tools],
  ["category_search", SEARCH_TOOLS],
]);

const DEFAULT_ENABLED_TOOLS = [
  CORE_TOOLS.list_azure_devops_projects,
  CORE_TOOLS.get_azure_devops_identity_ids,
  SEARCH_TOOLS.search_azure_devops_code,
  REPO_TOOLS.get_azure_devops_repositories,
  REPO_TOOLS.get_azure_devops_pull_request_by_id,
  REPO_TOOLS.get_azure_devops_changes_by_commit,
  REPO_TOOLS.get_azure_devops_content_by_objectid,
  REPO_TOOLS.get_azure_devops_item_content_by_commit,
  REPO_TOOLS.list_azure_devops_pull_request_threads,
  REPO_TOOLS.list_azure_devops_pull_request_comments_by_thread,
];

async function configureAllTools(server: McpServer, tokenProvider: () => Promise<AccessToken>, connectionProvider: () => Promise<WebApi>, userAgentProvider: () => string) {
  // disable all tools by default except DEFAULT_ENABLED_TOOLS and ADO_MCP_ENABLED_TOOLS
  const disabledTools = new Set<string>(
    Array.from(TOOLS_CATEGORY_MAP.values())
      .map((tools) => Object.values(tools))
      .flatMap((toolName) => toolName)
  );
  DEFAULT_ENABLED_TOOLS.concat(
    (process.env.ADO_MCP_ENABLED_TOOLS || "")
      .split(",")
      .map((tool) => {
        tool = tool.trim();
        const category = TOOLS_CATEGORY_MAP.get(tool);
        if (category) {
          return Object.values(category);
        }

        return [tool];
      })
      .flatMap((tools) => tools)
  ).forEach((tool) => disabledTools.delete(tool));

  configureCoreTools(server, tokenProvider, connectionProvider, disabledTools);
  configureWorkTools(server, tokenProvider, connectionProvider, disabledTools);
  configureBuildTools(server, tokenProvider, connectionProvider, disabledTools);
  configureRepoTools(server, tokenProvider, connectionProvider, disabledTools);
  configureWorkItemTools(server, tokenProvider, connectionProvider, userAgentProvider, disabledTools);
  configureReleaseTools(server, tokenProvider, connectionProvider, disabledTools);
  configureWikiTools(server, tokenProvider, connectionProvider, disabledTools);
  configureTestPlanTools(server, tokenProvider, connectionProvider, disabledTools);
  configureSearchTools(server, tokenProvider, connectionProvider, userAgentProvider, disabledTools);
}

export { configureAllTools };
