// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { AccessToken } from "@azure/identity";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { WebApi } from "azure-devops-node-api";
import { z, ZodRawShape } from "zod";
import {
  Comment,
  CommentList,
  QueryExpand,
  QueryHierarchyItem,
  WorkItem,
  WorkItemExpand,
  WorkItemLink,
  WorkItemQueryResult,
  WorkItemReference,
  WorkItemRelation,
  WorkItemType,
} from "azure-devops-node-api/interfaces/WorkItemTrackingInterfaces.js";
import { BacklogLevelConfiguration, BacklogLevelWorkItems, BacklogType, IterationWorkItems, PredefinedQuery } from "azure-devops-node-api/interfaces/WorkInterfaces.js";
import { getErrorToolResult, htmlToMarkdown, McpToolConfig, resolveProjectId, textToolResult, ToolHandler } from "../shared/tool-utils.js";
import { batchApiVersion } from "../utils.js";

const WORKITEM_TOOLS = {
  list_my_azure_devops_work_items: "list_my_azure_devops_work_items",
  list_azure_devops_backlogs: "list_azure_devops_backlogs",
  list_azure_devops_backlog_work_items: "list_azure_devops_backlog_work_items",
  get_azure_devops_work_item: "get_azure_devops_work_item",
  get_azure_devops_work_items_batch_by_ids: "get_azure_devops_work_items_batch_by_ids",
  update_azure_devops_work_item: "update_azure_devops_work_item",
  create_azure_devops_work_item: "create_azure_devops_work_item",
  list_azure_devops_work_item_comments: "list_azure_devops_work_item_comments",
  list_azure_devops_iteration_work_items: "list_azure_devops_iteration_work_items",
  add_azure_devops_work_item_comment: "add_azure_devops_work_item_comment",
  add_azure_devops_child_work_items: "add_azure_devops_child_work_items",
  link_azure_devops_work_item_to_pull_request: "link_azure_devops_work_item_to_pull_request",
  get_azure_devops_work_item_type: "get_azure_devops_work_item_type",
  get_azure_devops_query: "get_azure_devops_query",
  get_azure_devops_query_results_by_id: "get_azure_devops_query_results_by_id",
  update_azure_devops_work_items_batch: "update_azure_devops_work_items_batch",
  close_and_link_azure_devops_work_item_duplicates: "close_and_link_azure_devops_work_item_duplicates",
  link_azure_devops_work_items_batch: "link_azure_devops_work_items_batch",
};

type ToolConfigType =
  | ReturnType<typeof listMyWorkItems>
  | ReturnType<typeof listBacklogs>
  | ReturnType<typeof listBacklogWorkItems>
  | ReturnType<typeof getWorkItem>
  | ReturnType<typeof getWorkItemsBatchByIds>
  | ReturnType<typeof updateWorkItem>
  | ReturnType<typeof createWorkItem>
  | ReturnType<typeof listWorkItemComments>
  | ReturnType<typeof listIterationWorkItems>
  | ReturnType<typeof addWorkItemComment>
  | ReturnType<typeof addChildWorkItems>
  | ReturnType<typeof linkWorkItemToPullRequest>
  | ReturnType<typeof getWorkItemType>
  | ReturnType<typeof getQuery>
  | ReturnType<typeof getQueryResultsById>
  | ReturnType<typeof updateWorkItemsBatch>
  | ReturnType<typeof closeAndLinkWorkItemDuplicates>
  | ReturnType<typeof linkWorkItemsBatch>;

type LinkTypeName = "parent" | "child" | "duplicate" | "duplicate of" | "related" | "successor" | "predecessor" | "tested by" | "tests";

function getLinkTypeRefName(name: LinkTypeName): string {
  switch (name) {
    case "parent":
      return "System.LinkTypes.Hierarchy-Reverse";
    case "child":
      return "System.LinkTypes.Hierarchy-Forward";
    case "duplicate":
      return "System.LinkTypes.Duplicate-Forward";
    case "duplicate of":
      return "System.LinkTypes.Duplicate-Reverse";
    case "related":
      return "System.LinkTypes.Related";
    case "successor":
      return "System.LinkTypes.Dependency-Forward";
    case "predecessor":
      return "System.LinkTypes.Dependency-Reverse";
    case "tested by":
      return "Microsoft.VSTS.Common.TestedBy-Forward";
    case "tests":
      return "Microsoft.VSTS.Common.TestedBy-Reverse";
  }
}

interface WorkItemFieldsView {
  type?: string;
  title?: string;
  state?: string;
  assignedTo?: string;
  areaPath?: string;
  iterationPath?: string;
  tags?: string;
  description?: string;
  parent?: number;
}

function readIdentityField(value: unknown): string | undefined {
  if (typeof value === "string") {
    return value;
  }
  if (value && typeof value === "object") {
    const obj = value as { displayName?: string; uniqueName?: string };
    if (obj.displayName && obj.uniqueName) {
      return `${obj.displayName} <${obj.uniqueName}>`;
    }
    if (obj.displayName) {
      return obj.displayName;
    }
    if (obj.uniqueName) {
      return obj.uniqueName;
    }
  }
  return undefined;
}

function readWorkItemFields(fields: Record<string, unknown> | undefined): WorkItemFieldsView {
  const f = fields ?? {};
  const parentRaw = f["System.Parent"];
  return {
    type: typeof f["System.WorkItemType"] === "string" ? (f["System.WorkItemType"] as string) : undefined,
    title: typeof f["System.Title"] === "string" ? (f["System.Title"] as string) : undefined,
    state: typeof f["System.State"] === "string" ? (f["System.State"] as string) : undefined,
    assignedTo: readIdentityField(f["System.AssignedTo"]),
    areaPath: typeof f["System.AreaPath"] === "string" ? (f["System.AreaPath"] as string) : undefined,
    iterationPath: typeof f["System.IterationPath"] === "string" ? (f["System.IterationPath"] as string) : undefined,
    tags: typeof f["System.Tags"] === "string" ? (f["System.Tags"] as string) : undefined,
    description: typeof f["System.Description"] === "string" ? (f["System.Description"] as string) : undefined,
    parent: typeof parentRaw === "number" ? parentRaw : undefined,
  };
}

function formatWorkItemSummary(item: WorkItem): string {
  if (item.id === undefined) {
    throw new Error("Work item missing required 'id' field.");
  }
  const f = readWorkItemFields(item.fields);
  const title = f.title ?? "(no title)";
  const attrs: string[] = [];
  if (f.type) {
    attrs.push(`type: ${f.type}`);
  }
  if (f.state) {
    attrs.push(`state: ${f.state}`);
  }
  if (f.assignedTo) {
    attrs.push(`assignedTo: ${f.assignedTo}`);
  }
  return attrs.length > 0 ? `#${item.id} ${title} (${attrs.join(", ")})` : `#${item.id} ${title}`;
}

function formatWorkItemSummaryLines(items: WorkItem[]): string[] {
  return items.map((item) => `- ${formatWorkItemSummary(item)}`);
}

function formatRelationLine(rel: WorkItemRelation): string {
  const parts: string[] = [];
  if (rel.rel) {
    parts.push(rel.rel);
  }
  if (rel.url) {
    parts.push(`→ ${rel.url}`);
  }
  const name = rel.attributes?.name;
  if (typeof name === "string" && name.length > 0) {
    parts.push(`(${name})`);
  }
  return parts.join(" ");
}

function formatWorkItemDetailLines(item: WorkItem): string[] {
  if (item.id === undefined) {
    throw new Error("Work item missing required 'id' field.");
  }
  const f = readWorkItemFields(item.fields);

  const headerAttrs: string[] = [];
  if (f.type) {
    headerAttrs.push(`type: ${f.type}`);
  }
  if (f.state) {
    headerAttrs.push(`state: ${f.state}`);
  }
  if (item.rev !== undefined) {
    headerAttrs.push(`rev: ${item.rev}`);
  }

  const lines: string[] = [];
  lines.push(headerAttrs.length > 0 ? `Work item #${item.id} (${headerAttrs.join(", ")}):` : `Work item #${item.id}:`);
  if (f.title) {
    lines.push(`- title: ${f.title}`);
  }
  if (f.assignedTo) {
    lines.push(`- assignedTo: ${f.assignedTo}`);
  }
  if (f.areaPath) {
    lines.push(`- areaPath: ${f.areaPath}`);
  }
  if (f.iterationPath) {
    lines.push(`- iterationPath: ${f.iterationPath}`);
  }
  if (f.tags) {
    lines.push(`- tags: ${f.tags}`);
  }
  if (f.parent !== undefined) {
    lines.push(`- parent: #${f.parent}`);
  }
  if (item.url) {
    lines.push(`- url: ${item.url}`);
  }
  if (f.description) {
    lines.push(`- description:`);
    lines.push(htmlToMarkdown(f.description));
  }

  if (item.relations && item.relations.length > 0) {
    lines.push(`Relations (count: ${item.relations.length}):`);
    for (const rel of item.relations) {
      lines.push(`- ${formatRelationLine(rel)}`);
    }
  }

  return lines;
}

function formatBacklogLine(backlog: BacklogLevelConfiguration): string {
  if (!backlog.name) {
    throw new Error("Backlog missing required 'name' field.");
  }
  const attrs: string[] = [];
  if (backlog.id) {
    attrs.push(`ID: ${backlog.id}`);
  }
  if (backlog.type !== undefined) {
    attrs.push(`type: ${BacklogType[backlog.type] ?? backlog.type}`);
  }
  if (backlog.rank !== undefined) {
    attrs.push(`rank: ${backlog.rank}`);
  }
  const witNames = (backlog.workItemTypes ?? []).map((w) => w.name).filter((n): n is string => !!n);
  if (witNames.length > 0) {
    attrs.push(`workItemTypes: ${witNames.join(", ")}`);
  }
  return attrs.length > 0 ? `${backlog.name} (${attrs.join(", ")})` : backlog.name;
}

function formatBacklogLines(backlogs: BacklogLevelConfiguration[]): string[] {
  return backlogs.map((b) => `- ${formatBacklogLine(b)}`);
}

function formatWorkItemLinkLine(link: WorkItemLink): string {
  const targetId = link.target?.id ?? "?";
  const sourceId = link.source?.id;
  if (sourceId === undefined || sourceId === null) {
    return `#${targetId}`;
  }
  const rel = link.rel ?? "(no rel)";
  return `#${targetId} (parent: #${sourceId}, rel: ${rel})`;
}

function formatWorkItemReferenceLine(ref: WorkItemReference): string {
  const id = ref.id ?? "?";
  return ref.url ? `#${id} (link: ${ref.url})` : `#${id}`;
}

function formatCommentBlock(comment: Comment): string[] {
  if (comment.id === undefined) {
    throw new Error("Comment missing required 'id' field.");
  }
  const author = comment.createdBy?.displayName ?? "(unknown author)";
  const dateStr = comment.createdDate ? new Date(comment.createdDate).toISOString() : "(unknown date)";
  const lines: string[] = [`Comment #${comment.id} by ${author} on ${dateStr}:`];
  if (comment.text) {
    lines.push(comment.text);
  }
  return lines;
}

function formatCommentListLines(list: CommentList): string[] {
  const comments = list.comments ?? [];
  const lines: string[] = [];
  for (const c of comments) {
    lines.push(...formatCommentBlock(c));
  }
  return lines;
}

function formatWorkItemTypeLines(witype: WorkItemType): string[] {
  if (!witype.name) {
    throw new Error("Work item type missing required 'name' field.");
  }
  const lines: string[] = [`Work item type '${witype.name}':`];
  if (witype.referenceName) {
    lines.push(`- referenceName: ${witype.referenceName}`);
  }
  if (witype.description) {
    lines.push(`- description: ${witype.description}`);
  }
  if (witype.color) {
    lines.push(`- color: ${witype.color}`);
  }
  if (witype.isDisabled) {
    lines.push(`- disabled: true`);
  }
  const states = witype.states ?? [];
  if (states.length > 0) {
    const stateNames = states.map((s) => s.name).filter((n): n is string => !!n);
    lines.push(`- states (${stateNames.length}): ${stateNames.join(", ")}`);
  }
  const fields = witype.fields ?? [];
  if (fields.length > 0) {
    const fieldNames = fields.map((f) => f.referenceName).filter((n): n is string => !!n);
    lines.push(`- fields (${fieldNames.length}): ${fieldNames.join(", ")}`);
  }
  return lines;
}

function formatQueryLines(query: QueryHierarchyItem): string[] {
  if (!query.name) {
    throw new Error("Query missing required 'name' field.");
  }
  const headerAttrs: string[] = [];
  if (query.id) {
    headerAttrs.push(`ID: ${query.id}`);
  }
  if (query.path) {
    headerAttrs.push(`path: ${query.path}`);
  }
  if (query.queryType !== undefined) {
    headerAttrs.push(`queryType: ${query.queryType}`);
  }
  if (query.isFolder) {
    headerAttrs.push(`isFolder: true`);
  }

  const lines: string[] = [`Query '${query.name}' (${headerAttrs.join(", ")}):`];
  const columnNames = (query.columns ?? []).map((c) => c.referenceName).filter((n): n is string => !!n);
  if (columnNames.length > 0) {
    lines.push(`- columns: ${columnNames.join(", ")}`);
  }
  if (query.wiql) {
    lines.push(`- WIQL:`);
    lines.push(query.wiql);
  }
  return lines;
}

function formatQueryResultLines(result: WorkItemQueryResult): string[] {
  const items = result.workItems ?? [];
  const relations = result.workItemRelations ?? [];
  const headerParts: string[] = [];
  if (result.queryType !== undefined) {
    headerParts.push(`queryType: ${result.queryType}`);
  }
  if (result.queryResultType !== undefined) {
    headerParts.push(`resultType: ${result.queryResultType}`);
  }
  if (result.asOf) {
    headerParts.push(`asOf: ${new Date(result.asOf).toISOString()}`);
  }

  const lines: string[] = [];
  if (items.length > 0) {
    lines.push(`Query results (count: ${items.length}${headerParts.length > 0 ? `, ${headerParts.join(", ")}` : ""}):`);
    for (const ref of items) {
      lines.push(`- ${formatWorkItemReferenceLine(ref)}`);
    }
  } else if (relations.length > 0) {
    lines.push(`Query results (relations: ${relations.length}${headerParts.length > 0 ? `, ${headerParts.join(", ")}` : ""}):`);
    for (const link of relations) {
      lines.push(`- ${formatWorkItemLinkLine(link)}`);
    }
  } else {
    lines.push(`Query returned no results${headerParts.length > 0 ? ` (${headerParts.join(", ")})` : ""}.`);
  }
  return lines;
}

interface BatchSubResponse {
  code?: number;
  body?: unknown;
}

function parseBatchValue(json: unknown): BatchSubResponse[] {
  if (json && typeof json === "object" && "value" in json) {
    const value = (json as { value: unknown }).value;
    if (Array.isArray(value)) {
      return value as BatchSubResponse[];
    }
  }
  if (Array.isArray(json)) {
    return json as BatchSubResponse[];
  }
  return [];
}

function extractBodyId(body: unknown): number | undefined {
  if (body && typeof body === "object" && "id" in body) {
    const id = (body as { id?: unknown }).id;
    if (typeof id === "number") {
      return id;
    }
  }
  return undefined;
}

function extractBodyMessage(body: unknown): string {
  if (typeof body === "string") {
    return body;
  }
  if (body && typeof body === "object") {
    const obj = body as { message?: unknown; value?: { Message?: unknown } };
    if (typeof obj.message === "string") {
      return obj.message;
    }
    const innerMessage = obj.value?.Message;
    if (typeof innerMessage === "string") {
      return innerMessage;
    }
  }
  return "(no detail)";
}

function formatBatchResultLines(action: string, results: BatchSubResponse[]): string[] {
  const successes: number[] = [];
  const failures: { id?: number; code: number; message: string }[] = [];
  for (const r of results) {
    const code = r.code ?? 0;
    if (code >= 200 && code < 300) {
      const id = extractBodyId(r.body);
      if (id !== undefined) {
        successes.push(id);
      }
    } else {
      failures.push({ id: extractBodyId(r.body), code, message: extractBodyMessage(r.body) });
    }
  }

  const lines: string[] = [`${action}: ${successes.length} succeeded, ${failures.length} failed (total: ${results.length}).`];
  if (successes.length > 0) {
    lines.push(`Succeeded: ${successes.map((id) => `#${id}`).join(", ")}`);
  }
  if (failures.length > 0) {
    lines.push(`Failed:`);
    for (const f of failures) {
      const idLabel = f.id !== undefined ? `#${f.id}` : "(no id)";
      lines.push(`- ${idLabel} (HTTP ${f.code}): ${f.message}`);
    }
  }
  return lines;
}

async function postBatch(connection: WebApi, accessToken: AccessToken, userAgent: string, body: unknown): Promise<unknown> {
  const response = await fetch(`${connection.serverUrl}/_apis/wit/$batch?api-version=${batchApiVersion}`, {
    method: "PATCH",
    headers: {
      "Authorization": `Bearer ${accessToken.token}`,
      "Content-Type": "application/json",
      "User-Agent": userAgent,
    },
    body: JSON.stringify(body),
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Work item batch failed (${response.status}): ${errorText}`);
  }

  return response.json();
}

export function configureWorkItemTools(
  server: McpServer,
  tokenProvider: () => Promise<AccessToken>,
  connectionProvider: () => Promise<WebApi>,
  userAgentProvider: () => string
): McpToolConfig<ZodRawShape>[] {
  const workItemTools: ToolConfigType[] = [];
  workItemTools.push(listMyWorkItems(connectionProvider));
  workItemTools.push(listBacklogs(connectionProvider));
  workItemTools.push(listBacklogWorkItems(connectionProvider));
  workItemTools.push(getWorkItem(connectionProvider));
  workItemTools.push(getWorkItemsBatchByIds(connectionProvider));
  workItemTools.push(updateWorkItem(connectionProvider));
  workItemTools.push(createWorkItem(connectionProvider));
  workItemTools.push(listWorkItemComments(connectionProvider));
  workItemTools.push(listIterationWorkItems(connectionProvider));
  workItemTools.push(addWorkItemComment(connectionProvider));
  workItemTools.push(addChildWorkItems(tokenProvider, connectionProvider, userAgentProvider));
  workItemTools.push(linkWorkItemToPullRequest(connectionProvider));
  workItemTools.push(getWorkItemType(connectionProvider));
  workItemTools.push(getQuery(connectionProvider));
  workItemTools.push(getQueryResultsById(connectionProvider));
  workItemTools.push(updateWorkItemsBatch(tokenProvider, connectionProvider, userAgentProvider));
  workItemTools.push(closeAndLinkWorkItemDuplicates(tokenProvider, connectionProvider, userAgentProvider));
  workItemTools.push(linkWorkItemsBatch(tokenProvider, connectionProvider, userAgentProvider));
  return workItemTools as unknown as McpToolConfig<ZodRawShape>[];
}

function listMyWorkItems(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    project: z.string().describe("The name or ID of the Azure DevOps project."),
    type: z.enum(["assignedtome", "myactivity"]).optional().describe("Predefined query name. Defaults to 'assignedtome'."),
    top: z.number().optional().describe("Maximum number of work items to return. Defaults to 50."),
    includeCompleted: z.boolean().optional().describe("Include completed work items. Defaults to false."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ project, type, top, includeCompleted }) => {
    try {
      const effectiveType = type ?? "assignedtome";
      const effectiveTop = top ?? 50;
      const effectiveCompleted = includeCompleted ?? false;
      const connection = await connectionProvider();
      const workApi = await connection.getWorkApi();
      const queryResult: PredefinedQuery = await workApi.getPredefinedQueryResults(project, effectiveType, effectiveTop, effectiveCompleted);

      const items = queryResult?.results ?? [];
      if (items.length === 0) {
        return textToolResult([`No work items found for query '${effectiveType}' in project '${project}'.`]);
      }

      const header = `My work items (query: ${effectiveType}, project: '${project}', count: ${items.length}, hasMore: ${queryResult.hasMore ?? false}):`;
      return textToolResult([header, ...formatWorkItemSummaryLines(items)]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to fetch my work items.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: WORKITEM_TOOLS.list_my_azure_devops_work_items,
    description: "List work items relevant to the authenticated user. 'assignedtome' returns assigned items; 'myactivity' returns recently changed.",
    inputSchema,
    handler,
  };

  return config;
}

function listBacklogs(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    project: z.string().describe("The name or ID of the Azure DevOps project."),
    team: z.string().describe("The name or ID of the Azure DevOps team."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ project, team }) => {
    try {
      const connection = await connectionProvider();
      const workApi = await connection.getWorkApi();
      const backlogs = (await workApi.getBacklogs({ project, team })) ?? [];

      if (backlogs.length === 0) {
        return textToolResult([`No backlogs found for team '${team}' in project '${project}'.`]);
      }

      const header = `Backlogs for team '${team}' in project '${project}' (count: ${backlogs.length}):`;
      return textToolResult([header, ...formatBacklogLines(backlogs)]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to fetch backlogs.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: WORKITEM_TOOLS.list_azure_devops_backlogs,
    description: "List backlog levels (Epics, Features, Stories, Tasks) for a team.",
    inputSchema,
    handler,
  };

  return config;
}

function listBacklogWorkItems(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    project: z.string().describe("The name or ID of the Azure DevOps project."),
    team: z.string().describe("The name or ID of the Azure DevOps team."),
    backlogId: z.string().describe("Backlog ID (from list_azure_devops_backlogs)."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ project, team, backlogId }) => {
    try {
      const connection = await connectionProvider();
      const workApi = await connection.getWorkApi();
      const result: BacklogLevelWorkItems = await workApi.getBacklogLevelWorkItems({ project, team }, backlogId);
      const links = result?.workItems ?? [];

      if (links.length === 0) {
        return textToolResult([`No work items in backlog '${backlogId}' for team '${team}' in project '${project}'.`]);
      }

      const header = `Work items in backlog '${backlogId}' for team '${team}' in project '${project}' (count: ${links.length}):`;
      const lines = links.map((l) => `- ${formatWorkItemLinkLine(l)}`);
      return textToolResult([header, ...lines]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to fetch backlog work items.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: WORKITEM_TOOLS.list_azure_devops_backlog_work_items,
    description: "List work item links in a backlog level. Returns IDs and parent/child relations only — call 'get_azure_devops_work_items_batch_by_ids' for titles and state.",
    inputSchema,
    handler,
  };

  return config;
}

function getWorkItem(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    id: z.number().describe("Work item ID."),
    project: z.string().describe("The name or ID of the Azure DevOps project."),
    fields: z.array(z.string()).optional().describe("Specific fields to return. If omitted, returns the default field set."),
    asOf: z.string().datetime().optional().describe("Retrieve the work item as of a specific time. ISO 8601 (e.g. '2026-01-15T00:00:00Z')."),
    expand: z.enum(["all", "fields", "links", "none", "relations"]).optional().describe("Expansion mode for related data. Defaults to 'none'."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ id, project, fields, asOf, expand }) => {
    try {
      const connection = await connectionProvider();
      const workItemApi = await connection.getWorkItemTrackingApi();
      const asOfDate = asOf ? new Date(asOf) : undefined;
      const workItem = await workItemApi.getWorkItem(id, fields, asOfDate, expand as unknown as WorkItemExpand, project);

      if (!workItem) {
        return textToolResult([`Work item #${id} not found in project '${project}'.`], true);
      }

      return textToolResult(formatWorkItemDetailLines(workItem));
    } catch (error) {
      return getErrorToolResult(error, "Failed to fetch work item.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: WORKITEM_TOOLS.get_azure_devops_work_item,
    description: "Get a single work item by ID. Use 'expand=relations' to include parent/child/PR links.",
    inputSchema,
    handler,
  };

  return config;
}

function getWorkItemsBatchByIds(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    project: z.string().describe("The name or ID of the Azure DevOps project."),
    ids: z.array(z.number()).describe("Work item IDs to retrieve in one call."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ project, ids }) => {
    try {
      const connection = await connectionProvider();
      const workItemApi = await connection.getWorkItemTrackingApi();
      const fields = ["System.Id", "System.WorkItemType", "System.Title", "System.State", "System.Parent", "System.Tags"];
      const items = (await workItemApi.getWorkItemsBatch({ ids, fields }, project)) ?? [];

      if (items.length === 0) {
        return textToolResult([`No work items found for the requested IDs in project '${project}'.`]);
      }

      const foundIds = new Set(items.map((i) => i.id).filter((id): id is number => id !== undefined));
      const missing = ids.filter((id) => !foundIds.has(id));

      const header = `Retrieved ${items.length} of ${ids.length} work items in project '${project}':`;
      const lines: string[] = [header, ...formatWorkItemSummaryLines(items)];
      if (missing.length > 0) {
        lines.push(`Missing: ${missing.map((id) => `#${id}`).join(", ")}`);
      }
      return textToolResult(lines);
    } catch (error) {
      return getErrorToolResult(error, "Failed to fetch work items batch.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: WORKITEM_TOOLS.get_azure_devops_work_items_batch_by_ids,
    description: "Fetch a batch of work items by IDs (compact: id, type, title, state, parent, tags).",
    inputSchema,
    handler,
  };

  return config;
}

function updateWorkItem(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    id: z.number().describe("Work item ID to update."),
    updates: z
      .array(
        z.object({
          op: z.enum(["add", "replace", "remove"]).optional().describe("JSON Patch operation. Defaults to 'add'."),
          path: z.string().describe("Field path, e.g. '/fields/System.Title'."),
          value: z.string().describe("New value for the field. Required for 'add'/'replace'."),
        })
      )
      .describe("Field updates to apply."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ id, updates }) => {
    try {
      const connection = await connectionProvider();
      const workItemApi = await connection.getWorkItemTrackingApi();
      const document = updates.map((u) => ({ op: u.op ?? "add", path: u.path, value: u.value }));
      const updated = await workItemApi.updateWorkItem(null, document, id);

      if (!updated) {
        return textToolResult([`Work item #${id} update returned no result.`], true);
      }

      return textToolResult([`Updated work item #${id}.`, ...formatWorkItemDetailLines(updated)]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to update work item.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: WORKITEM_TOOLS.update_azure_devops_work_item,
    description: "Update fields on a work item using JSON Patch operations.",
    inputSchema,
    handler,
  };

  return config;
}

function createWorkItem(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    project: z.string().describe("The name or ID of the Azure DevOps project."),
    workItemType: z.string().describe("Work item type, e.g. 'Task', 'Bug', 'User Story'."),
    fields: z
      .array(
        z.object({
          name: z.string().describe("Field reference name, e.g. 'System.Title'."),
          value: z.string().describe("Field value."),
          format: z.enum(["Html", "Markdown"]).optional().describe("Format for large text fields. Defaults to 'Markdown'."),
        })
      )
      .describe("Fields to set on the new work item."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ project, workItemType, fields }) => {
    try {
      const connection = await connectionProvider();
      const projectId = await resolveProjectId(connection, project);
      const workItemApi = await connection.getWorkItemTrackingApi();

      const document: { op: string; path: string; value: string }[] = fields.map(({ name, value }) => ({
        op: "add",
        path: `/fields/${name}`,
        value,
      }));

      for (const { name, value, format } of fields) {
        if (format !== "Html" && value.length > 50) {
          document.push({ op: "add", path: `/multilineFieldsFormat/${name}`, value: "Markdown" });
        }
      }

      const created = await workItemApi.createWorkItem(null, document, projectId, workItemType);

      if (!created) {
        return textToolResult([`Failed to create work item of type '${workItemType}' in project '${project}'.`], true);
      }

      return textToolResult([`Created work item #${created.id} (type: ${workItemType}) in project '${project}'.`, ...formatWorkItemDetailLines(created)]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to create work item.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: WORKITEM_TOOLS.create_azure_devops_work_item,
    description: "Create a work item of the given type with initial field values.",
    inputSchema,
    handler,
  };

  return config;
}

function listWorkItemComments(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    project: z.string().describe("The name or ID of the Azure DevOps project."),
    workItemId: z.number().describe("Work item ID to fetch comments for."),
    top: z.number().optional().describe("Maximum number of comments. Defaults to 50."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ project, workItemId, top }) => {
    try {
      const effectiveTop = top ?? 50;
      const connection = await connectionProvider();
      const workItemApi = await connection.getWorkItemTrackingApi();
      const list: CommentList = await workItemApi.getComments(project, workItemId, effectiveTop);

      const comments = list?.comments ?? [];
      if (comments.length === 0) {
        return textToolResult([`No comments on work item #${workItemId}.`]);
      }

      const total = list.totalCount ?? comments.length;
      const header = `Comments on work item #${workItemId} (showing ${comments.length} of ${total}):`;
      return textToolResult([header, ...formatCommentListLines(list)]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to fetch work item comments.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: WORKITEM_TOOLS.list_azure_devops_work_item_comments,
    description: "List comments on a work item.",
    inputSchema,
    handler,
  };

  return config;
}

function listIterationWorkItems(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    project: z.string().describe("The name or ID of the Azure DevOps project."),
    team: z.string().optional().describe("Team name or ID. Defaults to the project's default team."),
    iterationId: z.string().describe("Iteration ID."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ project, team, iterationId }) => {
    try {
      const connection = await connectionProvider();
      const workApi = await connection.getWorkApi();
      const result: IterationWorkItems = await workApi.getIterationWorkItems({ project, team }, iterationId);
      const relations = result?.workItemRelations ?? [];

      const teamLabel = team ? `team '${team}'` : "default team";
      if (relations.length === 0) {
        return textToolResult([`No work items in iteration '${iterationId}' for ${teamLabel} in project '${project}'.`]);
      }

      const header = `Work items in iteration '${iterationId}' for ${teamLabel} in project '${project}' (count: ${relations.length}):`;
      const lines = relations.map((r) => `- ${formatWorkItemLinkLine(r)}`);
      return textToolResult([header, ...lines]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to fetch iteration work items.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: WORKITEM_TOOLS.list_azure_devops_iteration_work_items,
    description: "List work item links in an iteration. Returns IDs and parent links — fetch titles via 'get_azure_devops_work_items_batch_by_ids'.",
    inputSchema,
    handler,
  };

  return config;
}

function addWorkItemComment(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    project: z.string().describe("The name or ID of the Azure DevOps project."),
    workItemId: z.number().describe("Work item ID to comment on."),
    comment: z.string().describe("Comment text."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ project, workItemId, comment }) => {
    try {
      const connection = await connectionProvider();
      const projectId = await resolveProjectId(connection, project);
      const workItemApi = await connection.getWorkItemTrackingApi();
      const created: Comment = await workItemApi.addComment({ text: comment }, projectId, workItemId);

      if (!created) {
        return textToolResult([`Failed to add comment to work item #${workItemId}.`], true);
      }

      return textToolResult([`Added comment to work item #${workItemId}.`, ...formatCommentBlock(created)]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to add work item comment.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: WORKITEM_TOOLS.add_azure_devops_work_item_comment,
    description: "Add a comment to a work item.",
    inputSchema,
    handler,
  };

  return config;
}

function addChildWorkItems(tokenProvider: () => Promise<AccessToken>, connectionProvider: () => Promise<WebApi>, userAgentProvider: () => string) {
  const inputSchema = {
    parentId: z.number().describe("Parent work item ID."),
    project: z.string().describe("The name or ID of the Azure DevOps project."),
    workItemType: z.string().describe("Type of child work items to create, e.g. 'Task'."),
    items: z
      .array(
        z.object({
          title: z.string().describe("Title of the child work item."),
          description: z.string().describe("Description of the child work item."),
          format: z.enum(["Markdown", "Html"]).optional().describe("Format for description. Defaults to 'Html'."),
          areaPath: z.string().optional().describe("Optional area path."),
          iterationPath: z.string().optional().describe("Optional iteration path."),
        })
      )
      .describe("Up to 50 children to create under the parent."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ parentId, project, workItemType, items }) => {
    try {
      if (items.length > 50) {
        return textToolResult([`A maximum of 50 child work items can be created in a single call (got ${items.length}).`], true);
      }

      const connection = await connectionProvider();
      const projectId = await resolveProjectId(connection, project);
      const accessToken = await tokenProvider();

      const body = items.map((item, index) => {
        const ops: { op: string; path: string; value: unknown }[] = [
          { op: "add", path: "/id", value: `-${index + 1}` },
          { op: "add", path: "/fields/System.Title", value: item.title },
          { op: "add", path: "/fields/System.Description", value: item.description },
          { op: "add", path: "/fields/Microsoft.VSTS.TCM.ReproSteps", value: item.description },
          {
            op: "add",
            path: "/relations/-",
            value: {
              rel: "System.LinkTypes.Hierarchy-Reverse",
              url: `${connection.serverUrl}/${projectId}/_apis/wit/workItems/${parentId}`,
            },
          },
        ];

        if (item.areaPath && item.areaPath.trim().length > 0) {
          ops.push({ op: "add", path: "/fields/System.AreaPath", value: item.areaPath });
        }
        if (item.iterationPath && item.iterationPath.trim().length > 0) {
          ops.push({ op: "add", path: "/fields/System.IterationPath", value: item.iterationPath });
        }
        if (item.format === "Markdown") {
          ops.push({ op: "add", path: "/multilineFieldsFormat/System.Description", value: "Markdown" });
          ops.push({ op: "add", path: "/multilineFieldsFormat/Microsoft.VSTS.TCM.ReproSteps", value: "Markdown" });
        }

        return {
          method: "PATCH",
          uri: `/${projectId}/_apis/wit/workitems/$${workItemType}?api-version=${batchApiVersion}`,
          headers: { "Content-Type": "application/json-patch+json" },
          body: ops,
        };
      });

      const result = await postBatch(connection, accessToken, userAgentProvider(), body);
      const subResponses = parseBatchValue(result);
      const action = `Add ${items.length} child '${workItemType}' under #${parentId}`;
      return textToolResult(formatBatchResultLines(action, subResponses));
    } catch (error) {
      return getErrorToolResult(error, "Failed to create child work items.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: WORKITEM_TOOLS.add_azure_devops_child_work_items,
    description: "Create up to 50 child work items under a parent in one batch call.",
    inputSchema,
    handler,
  };

  return config;
}

function linkWorkItemToPullRequest(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    project: z.string().describe("The name or ID of the Azure DevOps project."),
    repositoryId: z.string().describe("Repository ID (use the GUID, not the repo name)."),
    pullRequestId: z.number().describe("Pull request ID."),
    workItemId: z.number().describe("Work item ID to link."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ project, repositoryId, pullRequestId, workItemId }) => {
    try {
      const connection = await connectionProvider();
      const projectId = await resolveProjectId(connection, project);
      const workItemTrackingApi = await connection.getWorkItemTrackingApi();

      const artifactPathValue = `${projectId}/${repositoryId}/${pullRequestId}`;
      const vstfsUrl = `vstfs:///Git/PullRequestId/${encodeURIComponent(artifactPathValue)}`;
      const patchDocument = [
        {
          op: "add",
          path: "/relations/-",
          value: {
            rel: "ArtifactLink",
            url: vstfsUrl,
            attributes: { name: "Pull Request" },
          },
        },
      ];

      const updated = await workItemTrackingApi.updateWorkItem({}, patchDocument, workItemId, projectId);
      if (!updated) {
        return textToolResult([`Failed to link work item #${workItemId} to pull request #${pullRequestId}.`], true);
      }

      return textToolResult([`Linked work item #${workItemId} to pull request #${pullRequestId} in repository ${repositoryId} (rel: ArtifactLink, name: 'Pull Request').`]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to link work item to pull request.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: WORKITEM_TOOLS.link_azure_devops_work_item_to_pull_request,
    description: "Link a work item to an existing pull request via an ArtifactLink relation.",
    inputSchema,
    handler,
  };

  return config;
}

function getWorkItemType(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    project: z.string().describe("The name or ID of the Azure DevOps project."),
    workItemType: z.string().describe("Work item type name, e.g. 'Bug'."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ project, workItemType }) => {
    try {
      const connection = await connectionProvider();
      const workItemApi = await connection.getWorkItemTrackingApi();
      const witype: WorkItemType = await workItemApi.getWorkItemType(project, workItemType);

      if (!witype) {
        return textToolResult([`Work item type '${workItemType}' not found in project '${project}'.`], true);
      }

      return textToolResult(formatWorkItemTypeLines(witype));
    } catch (error) {
      return getErrorToolResult(error, "Failed to fetch work item type.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: WORKITEM_TOOLS.get_azure_devops_work_item_type,
    description: "Get the schema (states, fields, color) for a specific work item type.",
    inputSchema,
    handler,
  };

  return config;
}

function getQuery(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    project: z.string().describe("The name or ID of the Azure DevOps project."),
    query: z.string().describe("Query ID (GUID) or query path."),
    expand: z.enum(["all", "clauses", "minimal", "none", "wiql"]).optional().describe("Expansion mode. Use 'wiql' to include the query text. Defaults to 'none'."),
    depth: z.number().optional().describe("How many child query levels to expand. Defaults to 0."),
    includeDeleted: z.boolean().optional().describe("Include deleted items. Defaults to false."),
    useIsoDateFormat: z.boolean().optional().describe("Use ISO date format. Defaults to false."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ project, query, expand, depth, includeDeleted, useIsoDateFormat }) => {
    try {
      const effectiveDepth = depth ?? 0;
      const effectiveIncludeDeleted = includeDeleted ?? false;
      const effectiveIso = useIsoDateFormat ?? false;
      const connection = await connectionProvider();
      const workItemApi = await connection.getWorkItemTrackingApi();
      const queryDetails: QueryHierarchyItem = await workItemApi.getQuery(project, query, expand as unknown as QueryExpand, effectiveDepth, effectiveIncludeDeleted, effectiveIso);

      if (!queryDetails) {
        return textToolResult([`Query '${query}' not found in project '${project}'.`], true);
      }

      return textToolResult(formatQueryLines(queryDetails));
    } catch (error) {
      return getErrorToolResult(error, "Failed to fetch query.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: WORKITEM_TOOLS.get_azure_devops_query,
    description: "Get a saved query by ID or path. Use 'expand=wiql' to include the query text.",
    inputSchema,
    handler,
  };

  return config;
}

function getQueryResultsById(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    id: z.string().describe("Query ID (GUID)."),
    project: z.string().optional().describe("Project name or ID. Defaults to the current project."),
    team: z.string().optional().describe("Team name or ID. Defaults to the project's default team."),
    timePrecision: z.boolean().optional().describe("Include time precision in results. Defaults to false."),
    top: z.number().optional().describe("Maximum number of results. Defaults to 50."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ id, project, team, timePrecision, top }) => {
    try {
      const effectiveTop = top ?? 50;
      const connection = await connectionProvider();
      const workItemApi = await connection.getWorkItemTrackingApi();
      const result: WorkItemQueryResult = await workItemApi.queryById(id, { project, team }, timePrecision, effectiveTop);
      return textToolResult(formatQueryResultLines(result));
    } catch (error) {
      return getErrorToolResult(error, "Failed to execute query.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: WORKITEM_TOOLS.get_azure_devops_query_results_by_id,
    description: "Run a saved query by ID and return matching work item references. Fetch full data via 'get_azure_devops_work_items_batch_by_ids'.",
    inputSchema,
    handler,
  };

  return config;
}

function updateWorkItemsBatch(tokenProvider: () => Promise<AccessToken>, connectionProvider: () => Promise<WebApi>, userAgentProvider: () => string) {
  const inputSchema = {
    updates: z
      .array(
        z.object({
          op: z.enum(["add", "replace", "remove"]).optional().describe("JSON Patch operation. Defaults to 'add'."),
          id: z.number().describe("Work item ID to update."),
          path: z.string().describe("Field path, e.g. '/fields/System.Title'."),
          value: z.string().describe("New value. Required for 'add'/'replace'."),
          format: z.enum(["Html", "Markdown"]).optional().describe("Format for large text fields. Defaults to 'Html'."),
        })
      )
      .describe("Updates grouped by work item. Multiple entries per ID are merged into one PATCH call per work item."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ updates }) => {
    try {
      const connection = await connectionProvider();
      const accessToken = await tokenProvider();

      const uniqueIds = Array.from(new Set(updates.map((u) => u.id)));
      const body = uniqueIds.map((id) => {
        const itemUpdates = updates.filter((u) => u.id === id);
        const operations: { op: string; path: string; value: string }[] = itemUpdates.map(({ op, path, value }) => ({ op: op ?? "add", path, value }));

        for (const { path, value, format } of itemUpdates) {
          if (format === "Markdown" && value && value.length > 50) {
            operations.push({ op: "add", path: `/multilineFieldsFormat${path.replace("/fields", "")}`, value: "Markdown" });
          }
        }

        return {
          method: "PATCH",
          uri: `/_apis/wit/workitems/${id}?api-version=${batchApiVersion}`,
          headers: { "Content-Type": "application/json-patch+json" },
          body: operations,
        };
      });

      const result = await postBatch(connection, accessToken, userAgentProvider(), body);
      const subResponses = parseBatchValue(result);
      return textToolResult(formatBatchResultLines(`Update ${uniqueIds.length} work items`, subResponses));
    } catch (error) {
      return getErrorToolResult(error, "Failed to update work items batch.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: WORKITEM_TOOLS.update_azure_devops_work_items_batch,
    description: "Update multiple work items in one batch call. Updates are grouped by work item ID.",
    inputSchema,
    handler,
  };

  return config;
}

function linkWorkItemsBatch(tokenProvider: () => Promise<AccessToken>, connectionProvider: () => Promise<WebApi>, userAgentProvider: () => string) {
  const inputSchema = {
    project: z.string().describe("The name or ID of the Azure DevOps project."),
    updates: z
      .array(
        z.object({
          id: z.number().describe("Work item ID that will receive the link."),
          linkToId: z.number().describe("Target work item ID to link to."),
          type: z.enum(["parent", "child", "duplicate", "duplicate of", "related", "successor", "predecessor", "tested by", "tests"]).optional().describe("Link type name. Defaults to 'related'."),
          comment: z.string().optional().describe("Optional comment to attach to the link."),
        })
      )
      .describe("Links to create. Multiple entries with the same id are merged into one PATCH call."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ project, updates }) => {
    try {
      const connection = await connectionProvider();
      const projectId = await resolveProjectId(connection, project);
      const accessToken = await tokenProvider();
      const orgUrl = connection.serverUrl;

      const uniqueIds = Array.from(new Set(updates.map((u) => u.id)));
      const body = uniqueIds.map((id) => ({
        method: "PATCH",
        uri: `/_apis/wit/workitems/${id}?api-version=${batchApiVersion}`,
        headers: { "Content-Type": "application/json-patch+json" },
        body: updates
          .filter((u) => u.id === id)
          .map(({ linkToId, type, comment }) => ({
            op: "add",
            path: "/relations/-",
            value: {
              rel: getLinkTypeRefName((type ?? "related") as LinkTypeName),
              url: `${orgUrl}/${projectId}/_apis/wit/workItems/${linkToId}`,
              attributes: { comment: comment ?? "" },
            },
          })),
      }));

      const result = await postBatch(connection, accessToken, userAgentProvider(), body);
      const subResponses = parseBatchValue(result);
      return textToolResult(formatBatchResultLines(`Link ${uniqueIds.length} work items`, subResponses));
    } catch (error) {
      return getErrorToolResult(error, "Failed to link work items batch.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: WORKITEM_TOOLS.link_azure_devops_work_items_batch,
    description: "Add relations between work items in one batch call. Supports parent/child/duplicate/related/successor/predecessor/tested by/tests.",
    inputSchema,
    handler,
  };

  return config;
}

function closeAndLinkWorkItemDuplicates(tokenProvider: () => Promise<AccessToken>, connectionProvider: () => Promise<WebApi>, userAgentProvider: () => string) {
  const inputSchema = {
    id: z.number().describe("Canonical work item ID. Duplicates will be linked to this one as 'Duplicate-Reverse'."),
    duplicateIds: z.array(z.number()).describe("Work item IDs to close and mark as duplicates."),
    project: z.string().describe("The name or ID of the Azure DevOps project."),
    state: z.string().optional().describe("State to apply to the duplicates. Defaults to 'Removed'."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ id, duplicateIds, project, state }) => {
    try {
      const effectiveState = state ?? "Removed";
      const connection = await connectionProvider();
      const projectId = await resolveProjectId(connection, project);
      const accessToken = await tokenProvider();

      const body = duplicateIds.map((duplicateId) => ({
        method: "PATCH",
        uri: `/_apis/wit/workitems/${duplicateId}?api-version=${batchApiVersion}`,
        headers: { "Content-Type": "application/json-patch+json" },
        body: [
          { op: "add", path: "/fields/System.State", value: effectiveState },
          {
            op: "add",
            path: "/relations/-",
            value: {
              rel: "System.LinkTypes.Duplicate-Reverse",
              url: `${connection.serverUrl}/${projectId}/_apis/wit/workItems/${id}`,
            },
          },
        ],
      }));

      const result = await postBatch(connection, accessToken, userAgentProvider(), body);
      const subResponses = parseBatchValue(result);
      return textToolResult(formatBatchResultLines(`Close ${duplicateIds.length} duplicates of #${id} (state: ${effectiveState})`, subResponses));
    } catch (error) {
      return getErrorToolResult(error, "Failed to close and link duplicates.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: WORKITEM_TOOLS.close_and_link_azure_devops_work_item_duplicates,
    description: "Mark a set of work items as duplicates of one canonical item, set their state, and add Duplicate-Reverse links — in one batch call.",
    inputSchema,
    handler,
  };

  return config;
}

export { WORKITEM_TOOLS };
