// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { AccessToken } from "@azure/identity";
import { WebApi } from "azure-devops-node-api";

import { configureCoreTools } from "./tools/core.js";
import { configureWorkTools, WORK_TOOLS } from "./tools/work.js";
import { configureBuildTools, BUILD_TOOLS } from "./tools/builds.js";
import { configureRepoTools, REPO_TOOLS } from "./tools/repos.js";
import { configureWorkItemTools, WORKITEM_TOOLS } from "./tools/workitems.js";
import { configureReleaseTools, RELEASE_TOOLS } from "./tools/releases.js";
import { configureWikiTools, WIKI_TOOLS } from "./tools/wiki.js";
import { configureTestPlanTools, Test_Plan_Tools } from "./tools/testplans.js";
import { configureSearchTools } from "./tools/search.js";

const TOOLS_CATEGORY_MAP = new Map<string, object>([
  ["category_work", WORK_TOOLS],
  ["category_build", BUILD_TOOLS],
  ["category_repo", REPO_TOOLS],
  ["category_workitem", WORKITEM_TOOLS],
  ["category_release", RELEASE_TOOLS],
  ["category_wiki", WIKI_TOOLS],
  ["category_testplan", Test_Plan_Tools],
]);

function configureAllTools(server: McpServer, tokenProvider: () => Promise<AccessToken>, connectionProvider: () => Promise<WebApi>, userAgentProvider: () => string) {
  const disabledTools = new Set<string>(
    (process.env.ADO_MCP_DISABLED_TOOLS || "")
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
  );

  configureCoreTools(server, tokenProvider, connectionProvider);
  configureWorkTools(server, tokenProvider, connectionProvider, disabledTools);
  configureBuildTools(server, tokenProvider, connectionProvider, disabledTools);
  configureRepoTools(server, tokenProvider, connectionProvider, disabledTools);
  configureWorkItemTools(server, tokenProvider, connectionProvider, userAgentProvider, disabledTools);
  configureReleaseTools(server, tokenProvider, connectionProvider, disabledTools);
  configureWikiTools(server, tokenProvider, connectionProvider, disabledTools);
  configureTestPlanTools(server, tokenProvider, connectionProvider, disabledTools);
  configureSearchTools(server, tokenProvider, connectionProvider, userAgentProvider);
}

export { configureAllTools };
