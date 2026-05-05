// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { AccessToken } from "@azure/identity";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { WebApi } from "azure-devops-node-api";
import {
  GitRef,
  GitVersionType,
  GitVersionDescriptor,
  GitPullRequest,
  GitRepository,
  IdentityRefWithVote,
  GitCommitRef,
  GitUserDate,
  GitChange,
  GitItem,
  GitBaseVersionDescriptor,
  GitTargetVersionDescriptor,
  GitCommitDiffs,
  PullRequestStatus,
  VersionControlChangeType,
  CommentThreadStatus,
  CommentThreadContext,
  Comment,
  GitPullRequestCommentThread,
  GitForkRef,
} from "azure-devops-node-api/interfaces/GitInterfaces.js";
import { TeamProjectReference, WebApiTagDefinition } from "azure-devops-node-api/interfaces/CoreInterfaces.js";
import { z, ZodRawShape } from "zod";
import { IdentityRef } from "azure-devops-node-api/interfaces/common/VSSInterfaces.js";
import { createTwoFilesPatch } from "diff";
import { formatPagination, formatRepositoryLines, stripRefHeads } from "../shared/format-utils.js";
import { getErrorToolResult, htmlToMarkdown, McpToolConfig, resolveProjectId, textToolResult, ToolHandler } from "../shared/tool-utils.js";
import { getEnumKeys } from "../utils.js";

const MAX_DIFF_INPUT_BYTES = 512 * 1024;
const MAX_DIFF_LINES = 2000;

interface ItemDiffResult {
  diff: string;
  truncated: boolean;
  totalLines: number;
}

export type IdentityResult = Pick<IdentityRef, "id" | "displayName" | "uniqueName">;

export type IdentityWithVoteResult = Pick<IdentityRefWithVote, "id" | "displayName" | "uniqueName" | "vote" | "hasDeclined" | "isRequired">;

export type ProjectResult = Pick<TeamProjectReference, "id" | "description" | "name">;

export type RepositoryResult = Pick<GitRepository, "id" | "defaultBranch" | "name" | "isFork"> & {
  project?: ProjectResult;
};

export type PullRequestResult = Pick<
  GitPullRequest,
  "closedDate" | "creationDate" | "description" | "pullRequestId" | "sourceRefName" | "status" | "targetRefName" | "title" | "workItemRefs" | "isDraft" | "commits"
> & {
  closedBy?: IdentityResult;
  createdBy?: IdentityResult;
  reviewers?: IdentityWithVoteResult[];
  changesFromAllCommits?: GitCommitDiffsResult;
};

export type GitUserDateResult = Pick<GitUserDate, "email" | "date">;

export type GitItemResult = Pick<GitItem, "gitObjectType" | "objectId" | "originalObjectId" | "isFolder" | "isSymLink" | "path">;

export type ProperChangeType = keyof typeof VersionControlChangeType;
export type GitChangeResult = Pick<GitChange, "changeId" | "originalPath"> & {
  changeType?: ProperChangeType;
  item?: GitItemResult;
};

export type GitCommitDiffsResult = Pick<GitCommitDiffs, "allChangesIncluded" | "baseCommit" | "commonCommit" | "targetCommit"> & {
  changes?: GitChangeResult[];
};

const toIdentityResult = (identityRef: IdentityRef): IdentityResult => {
  return {
    id: identityRef.id,
    displayName: identityRef.displayName,
    uniqueName: identityRef.uniqueName,
  };
};

const toIdentityWithVoteResult = (identityRef: IdentityRefWithVote): IdentityWithVoteResult => {
  return {
    id: identityRef.id,
    displayName: identityRef.displayName,
    uniqueName: identityRef.uniqueName,
    vote: identityRef.vote,
    hasDeclined: identityRef.hasDeclined,
    isRequired: identityRef.isRequired,
  };
};

const toProjectResult = (projectRef: TeamProjectReference): ProjectResult => {
  return {
    id: projectRef.id,
    description: projectRef.description,
    name: projectRef.name,
  };
};

const toRepositoryResult = (repo: GitRepository): RepositoryResult => {
  return {
    id: repo.id,
    defaultBranch: repo.defaultBranch,
    name: repo.name,
    isFork: repo.isFork,
    project: repo.project && toProjectResult(repo.project),
  };
};

const toPullRequestResult = (pr: GitPullRequest): PullRequestResult => {
  return {
    creationDate: pr.creationDate,
    closedDate: pr.closedDate,
    description: pr.description,
    pullRequestId: pr.pullRequestId,
    sourceRefName: pr.sourceRefName,
    targetRefName: pr.targetRefName,
    commits: pr.commits,
    status: pr.status,
    title: pr.title,
    workItemRefs: pr.workItemRefs,
    closedBy: pr.closedBy && toIdentityResult(pr.closedBy),
    createdBy: pr.createdBy && toIdentityResult(pr.createdBy),
    reviewers: pr.reviewers ? pr.reviewers.map((r) => toIdentityWithVoteResult(r)) : undefined,
  };
};

const toGitItemResult = (item: GitItem): GitItemResult => {
  return {
    gitObjectType: item.gitObjectType,
    objectId: item.objectId,
    originalObjectId: item.originalObjectId,
    isFolder: item.isFolder,
    isSymLink: item.isSymLink,
    path: item.path,
  };
};

const toGitChangeResult = (change: GitChange): GitChangeResult => {
  return {
    changeId: change.changeId,
    originalPath: change.originalPath,
    changeType: change.changeType ? (VersionControlChangeType[change.changeType] as ProperChangeType) : undefined,
    item: change.item && toGitItemResult(change.item),
  };
};

const toGitCommitDiffsResult = (commitDiffs: GitCommitDiffs): GitCommitDiffsResult => {
  return {
    changes: commitDiffs.changes ? commitDiffs.changes.map((c) => toGitChangeResult(c)) : undefined,
    baseCommit: commitDiffs.baseCommit,
    commonCommit: commitDiffs.commonCommit,
    targetCommit: commitDiffs.targetCommit,
    allChangesIncluded: commitDiffs.allChangesIncluded,
  };
};

const VOTE_LABELS: Record<number, string> = {
  10: "approved",
  5: "approved with suggestions",
  0: "no vote",
  [-5]: "waiting for author",
  [-10]: "rejected",
};

const formatIdentitySummary = (identity: IdentityResult): string => {
  return identity.uniqueName ? `${identity.displayName} <${identity.uniqueName}>` : (identity.displayName ?? "");
};

const formatReviewerLine = (reviewer: IdentityWithVoteResult): string => {
  const voteLabel = VOTE_LABELS[reviewer.vote ?? 0] ?? `vote ${reviewer.vote}`;
  const tags = [voteLabel];
  if (reviewer.isRequired) tags.push("required");
  if (reviewer.hasDeclined) tags.push("declined");
  return `- ${formatIdentitySummary(reviewer)}: ${tags.join(", ")}`;
};

const formatCommitLine = (commit: GitCommitRef): string => {
  const date = commit.author?.date ? new Date(commit.author.date).toISOString() : "";
  const subject = commit.comment?.split("\n", 1)[0] ?? "";
  return `- Commit ID: ${commit.commitId} (${date}): ${subject}`;
};

const formatChangeLine = (change: GitChangeResult): string => {
  const pathChanged = change.originalPath ? change.originalPath !== change.item?.path : false;
  const path = change.item?.path ?? change.originalPath ?? "";
  const objectId = change.item?.objectId;
  const itemType = change.item?.gitObjectType ?? "";
  const lines = [`- ${change.changeType} [${itemType}]: ${path}${pathChanged ? " (path changed)" : ""}`];
  if (objectId) {
    lines.push(`    - Object ID: ${objectId}`);
  }

  if (change.item?.originalObjectId) {
    lines.push(`    - Original Object ID: ${change.item?.originalObjectId}`);
  }

  return lines.join("\n");
};

const formatThreadContext = (context?: CommentThreadContext): string => {
  const filePath = context?.filePath;
  if (!filePath) {
    return "";
  }

  const pos = context?.rightFileStart ?? context?.leftFileStart;
  return pos?.line ? ` at ${filePath}:${pos.line}` : ` at ${filePath}`;
};

const formatCommentBlock = (comment: Comment): string[] => {
  const author = comment.author?.displayName ?? "(unknown author)";
  const date = comment.lastContentUpdatedDate ?? comment.publishedDate;
  const dateStr = date ? new Date(date).toISOString() : "";
  const header = `  - ${author} (${dateStr}):`;
  const body = htmlToMarkdown(comment.content)
    .split("\n")
    .map((l) => `      ${l}`);
  return [header, ...body];
};

function formatThreadLines(thread: GitPullRequestCommentThread): string[] {
  if (thread.id === undefined) {
    throw new Error("Thread is missing required field (id).");
  }

  const status = thread.status !== undefined ? CommentThreadStatus[thread.status] : "Unknown";
  const context = formatThreadContext(thread.threadContext);
  const lines: string[] = [`Thread #${thread.id} [${status}]${context}`];

  if (thread.publishedDate) {
    lines.push(`  Published: ${new Date(thread.publishedDate).toISOString()}`);
  }

  const comments = (thread.comments ?? []).filter((c) => !c.isDeleted);
  if (comments.length > 0) {
    lines.push(`  Comments (${comments.length}):`);
    comments.forEach((c) => lines.push(...formatCommentBlock(c)));
  }

  return lines;
}

interface FormatItemContentLinesOptions {
  itemPath?: string;
  commitId?: string;
  objectId?: string;
  startLine: number;
  limit: number;
}

function formatItemContentLines(content: string, options: FormatItemContentLinesOptions): string[] {
  const { itemPath, commitId, objectId, startLine, limit } = options;
  const allLines = content.split(/\r?\n/);
  if (allLines.length > 0 && allLines[allLines.length - 1] === "" && /\r?\n$/.test(content)) {
    allLines.pop();
  }

  const totalLines = allLines.length;
  const metadataParts: string[] = [];
  if (itemPath) metadataParts.push(`Path: ${itemPath}`);
  if (commitId) metadataParts.push(`Commit ID: ${commitId}`);
  if (objectId) metadataParts.push(`Object ID: ${objectId}`);
  const headerParts = [`--- METADATA ---`, `[${metadataParts.join(" | ")}]`, `[Total Lines: ${totalLines}]`];

  if (totalLines === 0) {
    return headerParts.concat(`File is empty.`);
  }

  if (startLine > totalLines) {
    return headerParts.concat(`startLine ${startLine} exceeds total (${totalLines} lines).`);
  }

  const startIdx = startLine - 1;
  const endIdx = Math.min(startIdx + limit, totalLines);
  const sliced = allLines.slice(startIdx, endIdx);

  headerParts.push("Lines are prefixed with [LNNN] markers. These markers are NOT part of the note content.");
  if (endIdx < totalLines) {
    headerParts.push(`More lines available, set startLine=${endIdx + 1} to continue`);
  }

  headerParts.push("", `--- CONTENT (lines ${startLine} - ${endIdx})`);

  const contentParts = sliced.map((line, index) => `[L${startIdx + index + 1}] ${line}`);
  return headerParts.concat(contentParts);
}

function formatPullRequestLines(pr: PullRequestResult): string[] {
  if (pr.pullRequestId === undefined || pr.status === undefined || !pr.title) {
    throw new Error("Pull request is missing required fields (id, status, or title).");
  }

  const lines: string[] = [];
  const draftSuffix = pr.isDraft ? " (draft)" : "";

  lines.push(`Pull Request #${pr.pullRequestId}: ${pr.title}`);
  lines.push(`Status: ${PullRequestStatus[pr.status]}${draftSuffix}`);
  if (pr.sourceRefName && pr.targetRefName) {
    lines.push(`Branches: ${stripRefHeads(pr.sourceRefName)} -> ${stripRefHeads(pr.targetRefName)}`);
  }
  if (pr.creationDate && pr.createdBy) {
    lines.push(`Created: ${new Date(pr.creationDate).toISOString()} by ${formatIdentitySummary(pr.createdBy)}`);
  }
  if (pr.closedDate && pr.closedBy) {
    lines.push(`Closed: ${new Date(pr.closedDate).toISOString()} by ${formatIdentitySummary(pr.closedBy)}`);
  }

  if (pr.description) {
    lines.push("", "Description:", ...pr.description.split("\n").map((l) => `  ${l}`));
  }

  const reviewers = pr.reviewers ?? [];
  if (reviewers.length > 0) {
    lines.push("", `Reviewers (${reviewers.length}):`, ...reviewers.map((r) => `  ${formatReviewerLine(r)}`));
  }

  const commits = pr.commits ?? [];
  if (commits.length > 0) {
    lines.push("", `Commits (${commits.length}):`, ...commits.map((c) => `  ${formatCommitLine(c)}`));
  }

  const workItemIds = (pr.workItemRefs ?? []).map((w) => w.id).filter((id): id is string => !!id);
  if (workItemIds.length > 0) {
    lines.push("", `Linked work items: ${workItemIds.join(", ")}`);
  }

  const diffs = pr.changesFromAllCommits;
  const changes = diffs?.changes ?? [];
  if (changes.length > 0) {
    const completeness = diffs?.allChangesIncluded === false ? " (truncated)" : "";
    const range = `, ${stripRefHeads(pr.sourceRefName ?? "")} -> ${stripRefHeads(pr.targetRefName ?? "")}`;
    lines.push("", `Changed Items (${changes.length}${completeness}${range}):`, ...changes.filter((change) => !change.item?.isFolder).map((c) => `  ${formatChangeLine(c)}`));
  }

  return lines;
}

export const REPO_TOOLS = {
  get_azure_devops_repositories: "get_azure_devops_repositories",
  get_azure_devops_pull_request_by_id: "get_azure_devops_pull_request_by_id",
  get_azure_devops_changes_by_commit: "get_azure_devops_changes_by_commit",
  get_azure_devops_item_content_by_commit: "get_azure_devops_item_content_by_commit",
  get_azure_devops_content_by_objectid: "get_azure_devops_content_by_objectid",
  create_azure_devops_pull_request_thread: "create_azure_devops_pull_request_thread",
  list_azure_devops_pull_request_threads: "list_azure_devops_pull_request_threads",
  list_azure_devops_pull_request_comments_by_thread: "list_azure_devops_pull_request_comments_by_thread",
  reply_to_azure_devops_pull_request_thread: "reply_to_azure_devops_pull_request_thread",
  resolve_azure_devops_pull_request_thread: "resolve_azure_devops_pull_request_thread",
  list_my_azure_devops_branches: "list_my_azure_devops_branches",
  get_azure_devops_branch_by_name: "get_azure_devops_branch_by_name",
  create_azure_devops_pull_request: "create_azure_devops_pull_request",
  get_azure_devops_item_diff_by_commits: "get_azure_devops_item_diff_by_commits",
  get_azure_devops_item_diff_by_object_ids: "get_azure_devops_item_diff_by_object_ids",
};

type ToolConfigType =
  | ReturnType<typeof getRepositories>
  | ReturnType<typeof getPullRequestById>
  | ReturnType<typeof getChangesByCommit>
  | ReturnType<typeof getItemContentByCommitTool>
  | ReturnType<typeof getContentByObjectIdTool>
  | ReturnType<typeof listPullRequestThreads>
  | ReturnType<typeof listPullRequestCommentsByThread>
  | ReturnType<typeof replyToComment>
  | ReturnType<typeof resolveComment>
  | ReturnType<typeof createPullRequestThread>
  | ReturnType<typeof createPullRequest>
  | ReturnType<typeof listMyBranches>
  | ReturnType<typeof getBranchByName>
  | ReturnType<typeof getItemDiffByCommitsTool>
  | ReturnType<typeof getItemDiffByObjectIdsTool>;

export function configureRepoTools(server: McpServer, tokenProvider: () => Promise<AccessToken>, connectionProvider: () => Promise<WebApi>): McpToolConfig<ZodRawShape>[] {
  const repoTools: ToolConfigType[] = [];
  repoTools.push(getRepositories(connectionProvider));
  repoTools.push(getPullRequestById(connectionProvider));
  repoTools.push(getChangesByCommit(connectionProvider));
  repoTools.push(getItemContentByCommitTool(connectionProvider));
  repoTools.push(getContentByObjectIdTool(connectionProvider));
  repoTools.push(listPullRequestThreads(connectionProvider));
  repoTools.push(listPullRequestCommentsByThread(connectionProvider));
  repoTools.push(replyToComment(connectionProvider));
  repoTools.push(resolveComment(connectionProvider));
  repoTools.push(createPullRequestThread(connectionProvider));
  repoTools.push(createPullRequest(connectionProvider));
  repoTools.push(listMyBranches(connectionProvider));
  repoTools.push(getBranchByName(connectionProvider));
  repoTools.push(getItemDiffByCommitsTool(connectionProvider));
  repoTools.push(getItemDiffByObjectIdsTool(connectionProvider));

  return repoTools as unknown as McpToolConfig<ZodRawShape>[];
}

function getRepositories(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    project: z.string().describe("The name or ID of the Azure DevOps project."),
    top: z.number().optional().describe("The maximum number of repositories to return."),
    skip: z.number().optional().describe("The number of repositories to skip."),
    repositoryNameOrId: z.string().optional().describe("Optional filter to search for repositories by name or ID."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ project, top, skip, repositoryNameOrId }) => {
    try {
      const effectiveTop = top ?? 100;
      const effectiveSkip = skip ?? 0;
      const reposResult = await getRepositoriesResult(connectionProvider, project, repositoryNameOrId);
      const sortedRepos = reposResult.sort((a, b) => a.name?.localeCompare(b.name ?? "") ?? 0);
      const paginatedRepositories = sortedRepos.slice(effectiveSkip, effectiveSkip + effectiveTop);
      const filterSuffix = repositoryNameOrId ? ` matching '${repositoryNameOrId}'` : "";

      if (paginatedRepositories.length === 0) {
        return textToolResult([`No repositories found in project '${project}'${filterSuffix}.`]);
      }

      const header = `Repositories in project '${project}'${filterSuffix}:`;
      const pagination = formatPagination(paginatedRepositories.length, effectiveTop, effectiveSkip, sortedRepos.length);
      return textToolResult([header, pagination, ...formatRepositoryLines(paginatedRepositories)]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to get repositories.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: REPO_TOOLS.get_azure_devops_repositories,
    description: "List Git repositories in an Azure DevOps project.",
    inputSchema,
    handler,
  };

  return config;
}

function getPullRequestById(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    repositoryId: z.string().describe("The ID of the repository where the pull request is located."),
    pullRequestId: z.number().describe("The ID of the pull request to retrieve."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ repositoryId, pullRequestId }) => {
    try {
      const prResult = await getPullRequestResult(connectionProvider, repositoryId, pullRequestId);
      return textToolResult(formatPullRequestLines(prResult));
    } catch (error) {
      return getErrorToolResult(error, "Failed to get pull request.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: REPO_TOOLS.get_azure_devops_pull_request_by_id,
    description: "Get a pull request by its ID.",
    inputSchema,
    handler,
  };

  return config;
}

function getChangesByCommit(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    repositoryId: z.string().describe("The ID of the repository where the commit is located."),
    commitId: z.string().describe("The ID of the commit to retrieve."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ repositoryId, commitId }) => {
    try {
      const connection = await connectionProvider();
      const gitApi = await connection.getGitApi();
      const commitChanges = await gitApi.getChanges(commitId, repositoryId);
      const changes = commitChanges?.changes?.map(toGitChangeResult) ?? [];

      if (changes.length === 0) {
        return textToolResult([`No changes found for commit ${commitId}.`]);
      }

      return textToolResult([`Changes in commit ${commitId} (${changes.length}):`, ...changes.map(formatChangeLine)]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to get changes.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: REPO_TOOLS.get_azure_devops_changes_by_commit,
    description: "Get a list of changes by a commit ID.",
    inputSchema,
    handler,
  };

  return config;
}

function getItemContentByCommitTool(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    repositoryId: z.string().describe("The ID of the repository where the item is located."),
    commitId: z.string().describe("The ID of the commit."),
    itemPath: z.string().describe("Item's path."),
    startLine: z.number().optional().describe("1-indexed line number to start reading from. Default: 1."),
    limit: z.number().optional().describe("Maximum number of lines to return starting from startLine. Default: 500"),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ repositoryId, commitId, itemPath, startLine, limit }) => {
    try {
      const content = await getItemContentByCommit(connectionProvider, repositoryId, commitId, itemPath);
      return textToolResult(formatItemContentLines(content, { itemPath, commitId, startLine: startLine || 1, limit: limit || 500 }));
    } catch (error) {
      return getErrorToolResult(error, `Failed to get content for ${itemPath} at commit ${commitId}.`);
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: REPO_TOOLS.get_azure_devops_item_content_by_commit,
    description: "Get content of an item at a specific commit. Paginated by line; use startLine/limit to page through large files.",
    inputSchema,
    handler,
  };

  return config;
}

function getContentByObjectIdTool(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    repositoryId: z.string().describe("The ID of the repository."),
    objectId: z.string().describe("The SHA-1 hash of the Git object (blob)."),
    startLine: z.number().optional().describe("1-indexed line number to start reading from. Default: 1."),
    limit: z.number().optional().describe("Maximum number of lines to return starting from startLine. Default: 500"),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ repositoryId, objectId, startLine, limit }) => {
    try {
      const content = await getContentByObjectId(connectionProvider, repositoryId, objectId);
      return textToolResult(formatItemContentLines(content, { objectId, startLine: startLine || 1, limit: limit || 500 }));
    } catch (error) {
      return getErrorToolResult(error, `Failed to get content for object ${objectId}.`);
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: REPO_TOOLS.get_azure_devops_content_by_objectid,
    description: "Get item content directly by Git object ID (SHA-1 hash) for item with blob gitObjectType. Paginated by line; use startLine/limit to page through large blobs.",
    inputSchema,
    handler,
  };

  return config;
}

function listPullRequestThreads(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    repositoryId: z.string().describe("The ID of the repository where the pull request is located."),
    pullRequestId: z.number().describe("The ID of the pull request for which to retrieve threads."),
    project: z.string().optional().describe("Project ID or project name (optional)"),
    iteration: z.number().optional().describe("The iteration ID for which to retrieve threads. Optional, defaults to the latest iteration."),
    baseIteration: z.number().optional().describe("The base iteration ID for which to retrieve threads. Optional, defaults to the latest base iteration."),
    top: z.number().optional().describe("The maximum number of threads to return."),
    skip: z.number().optional().describe("The number of threads to skip."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ repositoryId, pullRequestId, project, iteration, baseIteration, top, skip }) => {
    try {
      const effectiveTop = top ?? 100;
      const effectiveSkip = skip ?? 0;
      const connection = await connectionProvider();
      const gitApi = await connection.getGitApi();
      const threads = await gitApi.getThreads(repositoryId, pullRequestId, project, iteration, baseIteration);
      const sortedThreads = (threads ?? []).sort((a, b) => (a.id ?? 0) - (b.id ?? 0));
      const paginatedThreads = sortedThreads.slice(effectiveSkip, effectiveSkip + effectiveTop);

      if (paginatedThreads.length === 0) {
        return textToolResult([`No threads found in PR #${pullRequestId}.`]);
      }

      const header = `Threads in PR #${pullRequestId}:`;
      const pagination = formatPagination(paginatedThreads.length, effectiveTop, effectiveSkip, sortedThreads.length);
      const body = paginatedThreads.flatMap((t, i) => (i === 0 ? formatThreadLines(t) : ["", ...formatThreadLines(t)]));
      return textToolResult([header, pagination, "", ...body]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to get pull request threads.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: REPO_TOOLS.list_azure_devops_pull_request_threads,
    description: "Retrieve a list of comment threads for a pull request.",
    inputSchema,
    handler,
  };

  return config;
}

function listPullRequestCommentsByThread(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    repositoryId: z.string().describe("The ID of the repository where the pull request is located."),
    pullRequestId: z.number().describe("The ID of the pull request for which to retrieve thread comments."),
    threadId: z.number().describe("The ID of the thread for which to retrieve comments."),
    project: z.string().optional().describe("Project ID or project name (optional)"),
    top: z.number().optional().describe("The maximum number of comments to return."),
    skip: z.number().optional().describe("The number of comments to skip."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ repositoryId, pullRequestId, threadId, project, top, skip }) => {
    try {
      const effectiveTop = top ?? 100;
      const effectiveSkip = skip ?? 0;
      const connection = await connectionProvider();
      const gitApi = await connection.getGitApi();
      const comments = await gitApi.getComments(repositoryId, pullRequestId, threadId, project);
      const sortedComments = (comments ?? []).filter((c) => !c.isDeleted).sort((a, b) => (a.id ?? 0) - (b.id ?? 0));
      const paginatedComments = sortedComments.slice(effectiveSkip, effectiveSkip + effectiveTop);

      if (paginatedComments.length === 0) {
        return textToolResult([`No comments found in PR #${pullRequestId} thread #${threadId}.`]);
      }

      const header = `Comments in PR #${pullRequestId} thread #${threadId}:`;
      const pagination = formatPagination(paginatedComments.length, effectiveTop, effectiveSkip, sortedComments.length);
      const body = paginatedComments.flatMap((c) => formatCommentBlock(c));
      return textToolResult([header, pagination, "", ...body]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to get pull request thread comments.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: REPO_TOOLS.list_azure_devops_pull_request_comments_by_thread,
    description: "Retrieve a list of comments in a pull request thread.",
    inputSchema,
    handler,
  };

  return config;
}

function replyToComment(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    repositoryId: z.string().describe("The ID of the repository where the pull request is located."),
    pullRequestId: z.number().describe("The ID of the pull request where the comment thread exists."),
    threadId: z.number().describe("The ID of the thread to which the comment will be added."),
    content: z.string().describe("The content of the comment to be added."),
    project: z.string().optional().describe("Project ID or project name (optional)"),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ repositoryId, pullRequestId, threadId, content, project }) => {
    try {
      const connection = await connectionProvider();
      const projectId = project ? await resolveProjectId(connection, project) : undefined;
      const gitApi = await connection.getGitApi();
      const comment = await gitApi.createComment({ content }, repositoryId, pullRequestId, threadId, projectId);

      const header = `Reply added to PR #${pullRequestId} thread #${threadId} (comment id: ${comment.id}).`;
      return textToolResult([header, ...formatCommentBlock(comment)]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to reply to comment.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: REPO_TOOLS.reply_to_azure_devops_pull_request_thread,
    description: "Reply to a specific comment thread on a pull request.",
    inputSchema,
    handler,
  };

  return config;
}

function resolveComment(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    repositoryId: z.string().describe("The ID of the repository where the pull request is located."),
    pullRequestId: z.number().describe("The ID of the pull request where the comment thread exists."),
    threadId: z.number().describe("The ID of the thread to be resolved."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ repositoryId, pullRequestId, threadId }) => {
    try {
      const connection = await connectionProvider();
      const gitApi = await connection.getGitApi();
      const thread = await gitApi.updateThread({ status: CommentThreadStatus.Fixed }, repositoryId, pullRequestId, threadId);

      const status = thread.status !== undefined ? CommentThreadStatus[thread.status] : "Unknown";
      return textToolResult([`Thread #${threadId} in PR #${pullRequestId} resolved (status: ${status}).`]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to resolve comment thread.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: REPO_TOOLS.resolve_azure_devops_pull_request_thread,
    description: "Mark a pull request comment thread as resolved (status: Fixed).",
    inputSchema,
    handler,
  };

  return config;
}

function createPullRequestThread(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    repositoryId: z
      .string()
      .describe("The ID or name of the repository where the pull request is located. When using a repository name instead of a GUID, the project parameter must also be provided."),
    pullRequestId: z.coerce.number().min(1).describe("The ID of the pull request where the comment thread will be created."),
    content: z.string().describe("The content of the comment to be added."),
    project: z.string().optional().describe("Project ID or project name. Required when repositoryId is a repository name instead of a GUID."),
    filePath: z.string().optional().describe("The path of the file where the comment thread will be created."),
    status: z
      .enum(getEnumKeys(CommentThreadStatus) as [string, ...string[]])
      .optional()
      .default(CommentThreadStatus[CommentThreadStatus.Active])
      .describe("The status of the comment thread. Defaults to 'Active'."),
    rightFileStartLine: z.coerce.number().min(1).optional().describe("Position of first character of the thread's span in right file. 1-indexed line number."),
    rightFileStartOffset: z.number().optional().describe("Character offset (1-indexed) of first character. Must be set if rightFileStartLine is specified."),
    rightFileEndLine: z.number().optional().describe("Position of last character of the thread's span in right file. 1-indexed line number. Must be set if rightFileStartLine is specified."),
    rightFileEndOffset: z.number().optional().describe("Character offset (1-indexed) of last character. Must be set if rightFileEndLine is specified."),
  };

  const handler: ToolHandler<typeof inputSchema> = async ({
    repositoryId,
    pullRequestId,
    content,
    project,
    filePath,
    status,
    rightFileStartLine,
    rightFileStartOffset,
    rightFileEndLine,
    rightFileEndOffset,
  }) => {
    try {
      const normalizedFilePath = filePath && !filePath.startsWith("/") ? `/${filePath}` : filePath;
      const threadContext: CommentThreadContext = { filePath: normalizedFilePath };

      if (rightFileStartLine !== undefined) {
        threadContext.rightFileStart = { line: rightFileStartLine };
        if (rightFileStartOffset !== undefined) {
          if (rightFileStartOffset < 1) {
            return textToolResult(["rightFileStartOffset must be greater than or equal to 1."], true);
          }
          threadContext.rightFileStart.offset = rightFileStartOffset;
        }
      }

      if (rightFileEndLine !== undefined) {
        if (rightFileStartLine === undefined) {
          return textToolResult(["rightFileEndLine must only be specified if rightFileStartLine is also specified."], true);
        }
        if (rightFileEndLine < 1) {
          return textToolResult(["rightFileEndLine must be greater than or equal to 1."], true);
        }
        if (rightFileEndOffset === undefined) {
          return textToolResult(["rightFileEndOffset must be specified if rightFileEndLine is specified."], true);
        }
        if (rightFileEndOffset < 1) {
          return textToolResult(["rightFileEndOffset must be greater than or equal to 1."], true);
        }
        threadContext.rightFileEnd = { line: rightFileEndLine, offset: rightFileEndOffset };
      }

      if (rightFileEndOffset !== undefined && rightFileEndLine === undefined) {
        return textToolResult(["rightFileEndLine must be specified if rightFileEndOffset is specified."], true);
      }

      if (rightFileStartLine !== undefined && rightFileStartOffset !== undefined) {
        if (rightFileEndLine === undefined || rightFileEndOffset === undefined) {
          return textToolResult(["rightFileEndLine and rightFileEndOffset must both be specified when rightFileStartLine and rightFileStartOffset are both specified."], true);
        }
      }

      if (
        rightFileStartLine !== undefined &&
        rightFileEndLine !== undefined &&
        rightFileStartLine === rightFileEndLine &&
        rightFileEndOffset !== undefined &&
        rightFileStartOffset !== undefined &&
        rightFileEndOffset < rightFileStartOffset
      ) {
        return textToolResult(["rightFileEndOffset must be greater than or equal to rightFileStartOffset when both are on the same line."], true);
      }

      const connection = await connectionProvider();
      const projectId = project ? await resolveProjectId(connection, project) : undefined;
      const gitApi = await connection.getGitApi();
      const thread = await gitApi.createThread(
        { comments: [{ content }], threadContext, status: CommentThreadStatus[status as keyof typeof CommentThreadStatus] },
        repositoryId,
        pullRequestId,
        projectId
      );

      const header = `Thread #${thread.id} created in PR #${pullRequestId}.`;
      return textToolResult([header, "", ...formatThreadLines(thread)]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to create pull request thread.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: REPO_TOOLS.create_azure_devops_pull_request_thread,
    description: "Create a new comment thread on a pull request.",
    inputSchema,
    handler,
  };

  return config;
}

function createPullRequest(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    repositoryId: z.string().describe("The ID or name of the repository where the pull request will be created."),
    sourceRefName: z.string().describe("The source branch name for the pull request, e.g., 'refs/heads/feature-branch'."),
    targetRefName: z.string().describe("The target branch name for the pull request, e.g., 'refs/heads/main'."),
    title: z.string().describe("The title of the pull request."),
    description: z.string().max(4000).optional().describe("The description of the pull request. Must not be longer than 4000 characters."),
    isDraft: z.boolean().optional().describe("Indicates whether the pull request is a draft. Defaults to false."),
    workItems: z.string().optional().describe("Work item IDs to associate with the pull request, space-separated."),
    forkSourceRepositoryId: z.string().optional().describe("The ID of the fork repository that the pull request originates from."),
    labels: z.array(z.string()).optional().describe("Array of label names to add to the pull request after creation."),
  };

  const handler: ToolHandler<typeof inputSchema> = async ({ repositoryId, sourceRefName, targetRefName, title, description, isDraft, workItems, forkSourceRepositoryId, labels }) => {
    try {
      const connection = await connectionProvider();
      const gitApi = await connection.getGitApi();
      const workItemRefs = workItems ? workItems.split(" ").map((id) => ({ id: id.trim() })) : [];
      const forkSource: GitForkRef | undefined = forkSourceRepositoryId ? { repository: { id: forkSourceRepositoryId } } : undefined;
      const labelDefinitions: WebApiTagDefinition[] | undefined = labels ? labels.map((label) => ({ name: label })) : undefined;

      let pullRequest = await gitApi.createPullRequest(
        {
          sourceRefName,
          targetRefName,
          title,
          description,
          isDraft: isDraft ?? false,
          workItemRefs,
          forkSource,
          labels: labelDefinitions,
          supportsIterations: true,
        },
        repositoryId
      );

      if (!pullRequest) {
        const prs = await gitApi.getPullRequests(repositoryId, { sourceRefName, targetRefName, status: PullRequestStatus.Active }, undefined, undefined, 0, 1);
        if (prs && prs.length > 0) {
          pullRequest = prs[0];
        } else {
          return textToolResult(["Pull request created but API returned no data."]);
        }
      }

      return textToolResult(formatPullRequestLines(toPullRequestResult(pullRequest)));
    } catch (error) {
      return getErrorToolResult(error, "Failed to create pull request.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: REPO_TOOLS.create_azure_devops_pull_request,
    description: "Create a new pull request.",
    inputSchema,
    handler,
  };

  return config;
}

function listMyBranches(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    repositoryId: z.string().describe("The ID of the repository where the branches are located."),
    top: z.number().optional().describe("The maximum number of branches to return."),
    skip: z.number().optional().describe("The number of branches to skip."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ repositoryId, top, skip }) => {
    try {
      const effectiveTop = top ?? 100;
      const effectiveSkip = skip ?? 0;
      const connection = await connectionProvider();
      const gitApi = await connection.getGitApi();
      const refs = (await gitApi.getRefs(repositoryId, undefined, undefined, undefined, undefined, true)) ?? [];
      const branchNames = refs
        .filter((ref): ref is GitRef & { name: string } => !!ref.name && ref.name.startsWith("refs/heads/"))
        .map((ref) => stripRefHeads(ref.name))
        .sort((a, b) => b.localeCompare(a));
      const paginatedBranches = branchNames.slice(effectiveSkip, effectiveSkip + effectiveTop);

      if (paginatedBranches.length === 0) {
        return textToolResult([`No branches found for the authenticated user in repository ${repositoryId}.`]);
      }

      const header = `My branches in repository ${repositoryId}:`;
      const pagination = formatPagination(paginatedBranches.length, effectiveTop, effectiveSkip, branchNames.length);
      return textToolResult([header, pagination, ...paginatedBranches.map((b) => `- ${b}`)]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to list my branches.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: REPO_TOOLS.list_my_azure_devops_branches,
    description: "Retrieve a list of branches the authenticated user has in a given repository.",
    inputSchema,
    handler,
  };

  return config;
}

function getBranchByName(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    repositoryId: z.string().describe("The ID of the repository where the branch is located."),
    branchName: z.string().describe("The name of the branch to retrieve, e.g., 'main' or 'feature-branch'."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ repositoryId, branchName }) => {
    try {
      const connection = await connectionProvider();
      const gitApi = await connection.getGitApi();
      const refs = await gitApi.getRefs(repositoryId);
      const branch = refs?.find((ref) => ref.name === `refs/heads/${branchName}`);

      if (!branch) {
        return textToolResult([`Branch '${branchName}' not found in repository ${repositoryId}.`]);
      }

      const lines = [`Branch '${branchName}' in repository ${repositoryId}:`];
      if (branch.objectId) {
        lines.push(`- Object ID: ${branch.objectId}`);
      }
      if (branch.isLocked) {
        lines.push(`- Locked: yes`);
      }
      return textToolResult(lines);
    } catch (error) {
      return getErrorToolResult(error, "Failed to get branch.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: REPO_TOOLS.get_azure_devops_branch_by_name,
    description: "Get a branch by its name.",
    inputSchema,
    handler,
  };

  return config;
}

function buildItemDiff(oldContent: string, newContent: string, oldLabel: string, newLabel: string): ItemDiffResult {
  if (oldContent.length > MAX_DIFF_INPUT_BYTES || newContent.length > MAX_DIFF_INPUT_BYTES) {
    throw new Error(`File too large to diff (per-file limit: ${MAX_DIFF_INPUT_BYTES} bytes). old=${oldContent.length}, new=${newContent.length}.`);
  }

  const patch = createTwoFilesPatch(oldLabel, newLabel, oldContent, newContent, undefined, undefined, { context: 3 });
  const lines = patch.split("\n");
  if (lines.length <= MAX_DIFF_LINES) {
    return { diff: patch, truncated: false, totalLines: lines.length };
  }

  const truncated = lines.slice(0, MAX_DIFF_LINES).join("\n");
  return { diff: truncated, truncated: true, totalLines: lines.length };
}

function getItemDiffByCommitsTool(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    repositoryId: z.string().describe("Repository ID."),
    baseCommitId: z.string().describe("Commit ID of the base version."),
    targetCommitId: z.string().describe("Commit ID of the target version."),
    itemPath: z.string().describe("Path of the item, e.g. '/src/foo/bar.ts'."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ repositoryId, baseCommitId, targetCommitId, itemPath }) => {
    try {
      const [baseContent, targetContent] = await Promise.all([
        getItemContentByCommit(connectionProvider, repositoryId, baseCommitId, itemPath),
        getItemContentByCommit(connectionProvider, repositoryId, targetCommitId, itemPath),
      ]);

      const oldLabel = `${itemPath} @ ${baseCommitId.slice(0, 8)}`;
      const newLabel = `${itemPath} @ ${targetCommitId.slice(0, 8)}`;
      const result = buildItemDiff(baseContent, targetContent, oldLabel, newLabel);

      const headerAttrs = [`base: ${baseCommitId}`, `target: ${targetCommitId}`];
      if (result.truncated) {
        headerAttrs.push(`truncated: ${MAX_DIFF_LINES} of ${result.totalLines} lines`);
      }
      const header = `Diff for ${itemPath} (${headerAttrs.join(", ")}):`;
      return textToolResult([header, result.diff]);
    } catch (error) {
      return getErrorToolResult(error, `Failed to compute diff for ${itemPath}.`);
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: REPO_TOOLS.get_azure_devops_item_diff_by_commits,
    description: "Get a unified diff for one item between two commits. Truncates if either file exceeds size limits.",
    inputSchema,
    handler,
  };

  return config;
}

function getItemDiffByObjectIdsTool(connectionProvider: () => Promise<WebApi>) {
  const inputSchema = {
    repositoryId: z.string().describe("Repository ID."),
    baseObjectId: z.string().describe("SHA-1 of the base blob."),
    targetObjectId: z.string().describe("SHA-1 of the target blob."),
    itemPath: z.string().optional().describe("Optional item path, used only as a label in the diff header."),
  };
  const handler: ToolHandler<typeof inputSchema> = async ({ repositoryId, baseObjectId, targetObjectId, itemPath }) => {
    try {
      const [baseContent, targetContent] = await Promise.all([
        getContentByObjectId(connectionProvider, repositoryId, baseObjectId),
        getContentByObjectId(connectionProvider, repositoryId, targetObjectId),
      ]);

      const labelBase = itemPath ?? "(blob)";
      const oldLabel = `${labelBase} @ ${baseObjectId.slice(0, 8)}`;
      const newLabel = `${labelBase} @ ${targetObjectId.slice(0, 8)}`;
      const result = buildItemDiff(baseContent, targetContent, oldLabel, newLabel);

      const headerAttrs = [`baseObjectId: ${baseObjectId}`, `targetObjectId: ${targetObjectId}`];
      if (itemPath) {
        headerAttrs.unshift(`path: ${itemPath}`);
      }
      if (result.truncated) {
        headerAttrs.push(`truncated: ${MAX_DIFF_LINES} of ${result.totalLines} lines`);
      }
      const header = `Diff (${headerAttrs.join(", ")}):`;
      return textToolResult([header, result.diff]);
    } catch (error) {
      return getErrorToolResult(error, "Failed to compute diff between blobs.");
    }
  };

  const config: McpToolConfig<typeof inputSchema> = {
    name: REPO_TOOLS.get_azure_devops_item_diff_by_object_ids,
    description: "Get a unified diff between two blobs by SHA-1 object IDs. Pass itemPath only for labeling the diff header.",
    inputSchema,
    handler,
  };

  return config;
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

export async function getItemContentByCommit(connectionProvider: () => Promise<WebApi>, repositoryId: string, commitId: string, itemPath: string): Promise<string> {
  const connection = await connectionProvider();
  const gitApi = await connection.getGitApi();

  const itemVersion: GitVersionDescriptor = {
    version: commitId,
    versionType: GitVersionType["Commit"],
  };

  const stream = await gitApi.getItemContent(repositoryId, itemPath, undefined, undefined, undefined, true, undefined, false, itemVersion, true);
  return await streamToString(stream);
}

export async function getContentByObjectId(connectionProvider: () => Promise<WebApi>, repositoryId: string, objectId: string): Promise<string> {
  const connection = await connectionProvider();
  const gitApi = await connection.getGitApi();

  const stream = await gitApi.getBlobContent(repositoryId, objectId);
  return await streamToString(stream);
}

export async function getRepositoriesResult(connectionProvider: () => Promise<WebApi>, project: string, repositoryNameOrId?: string): Promise<RepositoryResult[]> {
  const connection = await connectionProvider();
  const gitApi = await connection.getGitApi();
  const repositories = await gitApi.getRepositories(project);

  const filter = repositoryNameOrId?.toLowerCase();
  const repos = filter ? repositories?.filter((repo) => repo.name?.toLowerCase().includes(filter) || repo.id === filter) : repositories;
  if (!repos || !repos.length) {
    throw new Error(`Repository ${repositoryNameOrId} not found in project ${project}`);
  }

  return repos.map((repo) => toRepositoryResult(repo));
}

export async function getPullRequestResult(connectionProvider: () => Promise<WebApi>, repositoryId: string, pullRequestId: number): Promise<PullRequestResult> {
  const connection = await connectionProvider();
  const gitApi = await connection.getGitApi();
  const pullRequest = await gitApi.getPullRequest(repositoryId, pullRequestId, undefined, undefined, undefined, undefined, true, true);
  if (!pullRequest) {
    throw new Error("Pull request not found.");
  }

  const prResult = toPullRequestResult(pullRequest);
  const commits = prResult.commits;
  if (commits && commits.length) {
    const latestCommitId = commits[0].commitId; // First commit is usually the latest
    const targetBranch = pullRequest.targetRefName ? stripRefHeads(pullRequest.targetRefName) : "main";
    const baseVersionDescriptor: GitBaseVersionDescriptor = {
      version: targetBranch,
      versionType: GitVersionType.Branch,
    };
    const targetVersionDescriptor: GitTargetVersionDescriptor = {
      version: latestCommitId,
      versionType: GitVersionType.Commit,
    };

    const diffs = await gitApi.getCommitDiffs(
      repositoryId,
      undefined,
      true, // diffCommonCommit
      100,
      undefined,
      baseVersionDescriptor,
      targetVersionDescriptor
    );
    prResult.changesFromAllCommits = diffs && toGitCommitDiffsResult(diffs);
  }
  return prResult;
}
