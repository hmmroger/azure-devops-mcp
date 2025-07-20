// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

import { AccessToken } from "@azure/identity";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { WebApi } from "azure-devops-node-api";
import {
  GitRef,
  PullRequestStatus,
  GitQueryCommitsCriteria,
  GitVersionType,
  GitVersionDescriptor,
  GitPullRequestQuery,
  GitPullRequestQueryInput,
  GitPullRequestQueryType,
  GitPullRequest,
  GitRepository,
  IdentityRefWithVote,
  GitCommitRef,
  GitUserDate,
  GitChange,
  GitItem,
  GitCommitChanges,
  GitBaseVersionDescriptor,
  GitTargetVersionDescriptor,
  GitCommitDiffs,
  VersionControlChangeType,
} from "azure-devops-node-api/interfaces/GitInterfaces.js";
import { TeamProjectReference } from "azure-devops-node-api/interfaces/CoreInterfaces.js";
import { z } from "zod";
import { getCurrentUserDetails } from "./auth.js";
import { IdentityRef } from "azure-devops-node-api/interfaces/common/VSSInterfaces.js";

export type IdentityResult = Pick<IdentityRef, "id" | "displayName" | "uniqueName">;

export type IdentityWithVoteResult = Pick<IdentityRefWithVote, "id" | "displayName" | "uniqueName" | "vote" | "hasDeclined" | "isRequired">;

export type ProjectResult = Pick<TeamProjectReference, "id" | "description" | "name">;

export type RepositoryResult = Pick<GitRepository, "id" | "defaultBranch" | "name" | "isFork"> & {
  project?: ProjectResult;
};

export type PullRequestResult = Pick<
  GitPullRequest,
  "closedDate" | "creationDate" | "description" | "pullRequestId" | "sourceRefName" | "status" | "targetRefName" | "title" | "workItemRefs" | "isDraft"
> & {
  commits?: GitCommitResult[];
  closedBy?: IdentityResult;
  createdBy?: IdentityResult;
  reviewers?: IdentityWithVoteResult[];
  pullRequestChanges?: GitCommitChangesResult;
};

export type GitUserDateResult = Pick<GitUserDate, "email" | "date">;

export type GitCommitResult = Pick<GitCommitRef, "commitId" | "comment" | "commentTruncated"> & {
  author?: GitUserDateResult;
  committer?: GitUserDateResult;
};

export type GitItemResult = Pick<GitItem, "gitObjectType" | "objectId" | "originalObjectId" | "isFolder" | "isSymLink" | "path">;

export type ProperChangeType = keyof typeof VersionControlChangeType;
export type GitChangeResult = Pick<GitChange, "changeId" | "originalPath"> & {
  changeType?: ProperChangeType;
  item?: GitItemResult;
};

export interface GitCommitChangesResult {
  changes?: GitChangeResult[];
}

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
    status: pr.status,
    title: pr.title,
    workItemRefs: pr.workItemRefs,
    closedBy: pr.closedBy && toIdentityResult(pr.closedBy),
    createdBy: pr.createdBy && toIdentityResult(pr.createdBy),
    reviewers: pr.reviewers ? pr.reviewers.map((r) => toIdentityWithVoteResult(r)) : undefined,
    commits: pr.commits ? pr.commits.map((c) => toGitCommitResult(c)) : undefined,
  };
};

const toGitUserDateResult = (user: GitUserDate): GitUserDateResult => {
  return {
    email: user.email,
    date: user.date,
  };
};

const toGitCommitResult = (commit: GitCommitRef): GitCommitResult => {
  return {
    commitId: commit.commitId,
    comment: commit.comment,
    commentTruncated: commit.commentTruncated,
    author: commit.author && toGitUserDateResult(commit.author),
    committer: commit.committer && toGitUserDateResult(commit.committer),
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

const toGitCommitChangesResult = (changes: GitCommitChanges): GitCommitChangesResult => {
  return {
    changes: changes.changes ? changes.changes.map((c) => toGitChangeResult(c)) : undefined,
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

const REPO_TOOLS = {
  get_azure_devops_repositories: "get_azure_devops_repositories",
  get_azure_devops_pull_request_by_id: "get_azure_devops_pull_request_by_id",
  get_azure_devops_changes_by_commit: "get_azure_devops_changes_by_commit",
  get_azure_devops_item_content_by_commit: "get_azure_devops_item_content_by_commit",
  get_azure_devops_content_by_objectid: "get_azure_devops_content_by_objectid",
  list_azure_devops_pull_request_threads: "list_azure_devops_pull_request_threads",
  list_azure_devops_pull_request_comments_by_thread: "list_azure_devops_pull_request_comments_by_thread",
  list_pull_requests_by_repo: "repo_list_pull_requests_by_repo",
  list_pull_requests_by_project: "repo_list_pull_requests_by_project",
  list_branches_by_repo: "repo_list_branches_by_repo",
  list_my_branches_by_repo: "repo_list_my_branches_by_repo",
  get_branch_by_name: "repo_get_branch_by_name",
  create_pull_request: "repo_create_pull_request",
  update_pull_request_status: "repo_update_pull_request_status",
  update_pull_request_reviewers: "repo_update_pull_request_reviewers",
  reply_to_comment: "repo_reply_to_comment",
  resolve_comment: "repo_resolve_comment",
  search_commits: "repo_search_commits",
  list_pull_requests_by_commits: "repo_list_pull_requests_by_commits",
};

function branchesFilterOutIrrelevantProperties(branches: GitRef[], top: number) {
  return branches
    ?.flatMap((branch) => (branch.name ? [branch.name] : []))
    ?.filter((branch) => branch.startsWith("refs/heads/"))
    .map((branch) => branch.replace("refs/heads/", ""))
    .sort((a, b) => b.localeCompare(a))
    .slice(0, top);
}

function pullRequestStatusStringToInt(status: string): number {
  switch (status) {
    case "abandoned":
      return PullRequestStatus.Abandoned.valueOf();
    case "active":
      return PullRequestStatus.Active.valueOf();
    case "all":
      return PullRequestStatus.All.valueOf();
    case "completed":
      return PullRequestStatus.Completed.valueOf();
    case "notSet":
      return PullRequestStatus.NotSet.valueOf();
    default:
      throw new Error(`Unknown pull request status: ${status}`);
  }
}

function configureRepoTools(server: McpServer, tokenProvider: () => Promise<AccessToken>, connectionProvider: () => Promise<WebApi>, disabledTools: Set<string>) {
  if (!disabledTools.has(REPO_TOOLS.get_azure_devops_repositories)) {
    server.tool(
      REPO_TOOLS.get_azure_devops_repositories,
      "[Azure DevOps] Get repositories for a given project with optional repository name filter.",
      {
        project: z.string().describe("The name or ID of the Azure DevOps project."),
        top: z.number().default(100).describe("The maximum number of repositories to return."),
        skip: z.number().default(0).describe("The number of repositories to skip. Defaults to 0."),
        repositoryNameOrId: z.string().optional().describe("Optional filter to search for repositories by name or ID."),
      },
      async ({ project, top, skip, repositoryNameOrId }) => {
        try {
          const reposResult = await getRepositoriesResult(connectionProvider, project, repositoryNameOrId);
          const paginatedRepositories = reposResult.sort((a, b) => a.name?.localeCompare(b.name ?? "") ?? 0).slice(skip, skip + top);
          return {
            content: [{ type: "text", text: JSON.stringify(paginatedRepositories) }],
          };
        } catch (error) {
          const message = (error as Error).message || "Failed to get repositories.";
          return {
            content: [{ type: "text", text: message }],
            isError: true,
          };
        }
      }
    );
  }

  if (!disabledTools.has(REPO_TOOLS.get_azure_devops_pull_request_by_id)) {
    server.tool(
      REPO_TOOLS.get_azure_devops_pull_request_by_id,
      "[Azure DevOps] Get a pull request by its ID.",
      {
        repositoryId: z.string().describe("The ID of the repository where the pull request is located."),
        pullRequestId: z.number().describe("The ID of the pull request to retrieve."),
      },
      async ({ repositoryId, pullRequestId }) => {
        try {
          const prRes = await getPullRequestResult(connectionProvider, repositoryId, pullRequestId);
          const text = JSON.stringify(prRes);
          return {
            content: [{ type: "text", text }],
          };
        } catch (error) {
          const message = (error as Error).message || "Failed to get pull request.";
          return {
            content: [{ type: "text", text: message }],
            isError: true,
          };
        }
      }
    );
  }

  if (!disabledTools.has(REPO_TOOLS.get_azure_devops_changes_by_commit)) {
    server.tool(
      REPO_TOOLS.get_azure_devops_changes_by_commit,
      "[Azure DevOps] Get a list of changes by a commit ID.",
      {
        repositoryId: z.string().describe("The ID of the repository where the pull request is located."),
        commitId: z.string().describe("The ID of the commit to retrieve."),
      },
      async ({ repositoryId, commitId }) => {
        const connection = await connectionProvider();
        const gitApi = await connection.getGitApi();
        const changes = await gitApi.getChanges(commitId, repositoryId);
        const text = changes ? JSON.stringify(toGitCommitChangesResult(changes)) : "No changes found.";
        return {
          content: [{ type: "text", text }],
        };
      }
    );
  }

  if (!disabledTools.has(REPO_TOOLS.get_azure_devops_item_content_by_commit)) {
    server.tool(
      REPO_TOOLS.get_azure_devops_item_content_by_commit,
      "[Azure DevOps] Get content of an item by commit ID.",
      {
        repositoryId: z.string().describe("The ID of the repository where the pull request is located."),
        commitId: z.string().describe("The ID of the commit."),
        itemPath: z.string().describe("Item's path."),
      },
      async ({ repositoryId, commitId, itemPath }) => {
        try {
          const content = await getItemContentByCommit(connectionProvider, repositoryId, commitId, itemPath);
          return {
            content: [{ type: "text", text: JSON.stringify(content) }],
          };
        } catch (error) {
          return {
            content: [{ type: "text", text: `Error getting content for item ${itemPath} at commit ${commitId}: ${error instanceof Error ? error.message : String(error)}` }],
            isError: true,
          };
        }
      }
    );
  }

  if (!disabledTools.has(REPO_TOOLS.get_azure_devops_content_by_objectid)) {
    server.tool(
      REPO_TOOLS.get_azure_devops_content_by_objectid,
      "[Azure DevOps] Get item content directly by Git object ID (SHA-1 hash) for item with blob gitObjectType.",
      {
        repositoryId: z.string().describe("The ID of the repository."),
        objectId: z.string().describe("The SHA-1 hash of the Git object (blob)."),
      },
      async ({ repositoryId, objectId }) => {
        try {
          const content = await getContentByObjectId(connectionProvider, repositoryId, objectId);
          return {
            content: [{ type: "text", text: content }],
          };
        } catch (error) {
          return {
            content: [{ type: "text", text: `Error getting content for object ${objectId}: ${error instanceof Error ? error.message : String(error)}` }],
            isError: true,
          };
        }
      }
    );
  }

  if (!disabledTools.has(REPO_TOOLS.list_azure_devops_pull_request_threads)) {
    server.tool(
      REPO_TOOLS.list_azure_devops_pull_request_threads,
      "[Azure DevOps] Retrieve a list of comment threads for a pull request.",
      {
        repositoryId: z.string().describe("The ID of the repository where the pull request is located."),
        pullRequestId: z.number().describe("The ID of the pull request for which to retrieve threads."),
        project: z.string().optional().describe("Project ID or project name (optional)"),
        iteration: z.number().optional().describe("The iteration ID for which to retrieve threads. Optional, defaults to the latest iteration."),
        baseIteration: z.number().optional().describe("The base iteration ID for which to retrieve threads. Optional, defaults to the latest base iteration."),
        top: z.number().default(100).describe("The maximum number of threads to return."),
        skip: z.number().default(0).describe("The number of threads to skip."),
      },
      async ({ repositoryId, pullRequestId, project, iteration, baseIteration, top, skip }) => {
        const connection = await connectionProvider();
        const gitApi = await connection.getGitApi();

        const threads = await gitApi.getThreads(repositoryId, pullRequestId, project, iteration, baseIteration);

        const paginatedThreads = threads?.sort((a, b) => (a.id ?? 0) - (b.id ?? 0)).slice(skip, skip + top);

        return {
          content: [{ type: "text", text: JSON.stringify(paginatedThreads, null, 2) }],
        };
      }
    );
  }

  if (!disabledTools.has(REPO_TOOLS.list_azure_devops_pull_request_comments_by_thread)) {
    server.tool(
      REPO_TOOLS.list_azure_devops_pull_request_comments_by_thread,
      "[Azure DevOps] Retrieve a list of comments in a pull request thread.",
      {
        repositoryId: z.string().describe("The ID of the repository where the pull request is located."),
        pullRequestId: z.number().describe("The ID of the pull request for which to retrieve thread comments."),
        threadId: z.number().describe("The ID of the thread for which to retrieve comments."),
        project: z.string().optional().describe("Project ID or project name (optional)"),
        top: z.number().default(100).describe("The maximum number of comments to return."),
        skip: z.number().default(0).describe("The number of comments to skip."),
      },
      async ({ repositoryId, pullRequestId, threadId, project, top, skip }) => {
        const connection = await connectionProvider();
        const gitApi = await connection.getGitApi();

        // Get thread comments - GitApi uses getComments for retrieving comments from a specific thread
        const comments = await gitApi.getComments(repositoryId, pullRequestId, threadId, project);

        const paginatedComments = comments?.sort((a, b) => (a.id ?? 0) - (b.id ?? 0)).slice(skip, skip + top);

        return {
          content: [{ type: "text", text: JSON.stringify(paginatedComments, null, 2) }],
        };
      }
    );
  }

  if (!disabledTools.has(REPO_TOOLS.create_pull_request)) {
    server.tool(
      REPO_TOOLS.create_pull_request,
      "Create a new pull request.",
      {
        repositoryId: z.string().describe("The ID of the repository where the pull request will be created."),
        sourceRefName: z.string().describe("The source branch name for the pull request, e.g., 'refs/heads/feature-branch'."),
        targetRefName: z.string().describe("The target branch name for the pull request, e.g., 'refs/heads/main'."),
        title: z.string().describe("The title of the pull request."),
        description: z.string().optional().describe("The description of the pull request. Optional."),
        isDraft: z.boolean().optional().default(false).describe("Indicates whether the pull request is a draft. Defaults to false."),
        workItems: z.string().optional().describe("Work item IDs to associate with the pull request, space-separated."),
      },
      async ({ repositoryId, sourceRefName, targetRefName, title, description, isDraft, workItems }) => {
        const connection = await connectionProvider();
        const gitApi = await connection.getGitApi();
        const workItemRefs = workItems ? workItems.split(" ").map((id) => ({ id: id.trim() })) : [];

        const pullRequest = await gitApi.createPullRequest(
          {
            sourceRefName,
            targetRefName,
            title,
            description,
            isDraft,
            workItemRefs: workItemRefs,
          },
          repositoryId
        );

        return {
          content: [{ type: "text", text: JSON.stringify(pullRequest, null, 2) }],
        };
      }
    );
  }

  if (!disabledTools.has(REPO_TOOLS.update_pull_request_status)) {
    server.tool(
      REPO_TOOLS.update_pull_request_status,
      "Update status of an existing pull request to active or abandoned.",
      {
        repositoryId: z.string().describe("The ID of the repository where the pull request exists."),
        pullRequestId: z.number().describe("The ID of the pull request to be published."),
        status: z.enum(["active", "abandoned"]).describe("The new status of the pull request. Can be 'active' or 'abandoned'."),
      },
      async ({ repositoryId, pullRequestId, status }) => {
        const connection = await connectionProvider();
        const gitApi = await connection.getGitApi();
        const statusValue = status === "active" ? 3 : 2;

        const updatedPullRequest = await gitApi.updatePullRequest({ status: statusValue }, repositoryId, pullRequestId);

        return {
          content: [{ type: "text", text: JSON.stringify(updatedPullRequest, null, 2) }],
        };
      }
    );
  }

  if (!disabledTools.has(REPO_TOOLS.update_pull_request_reviewers)) {
    server.tool(
      REPO_TOOLS.update_pull_request_reviewers,
      "Add or remove reviewers for an existing pull request.",
      {
        repositoryId: z.string().describe("The ID of the repository where the pull request exists."),
        pullRequestId: z.number().describe("The ID of the pull request to update."),
        reviewerIds: z.array(z.string()).describe("List of reviewer ids to add or remove from the pull request."),
        action: z.enum(["add", "remove"]).describe("Action to perform on the reviewers. Can be 'add' or 'remove'."),
      },
      async ({ repositoryId, pullRequestId, reviewerIds, action }) => {
        const connection = await connectionProvider();
        const gitApi = await connection.getGitApi();

        let updatedPullRequest;
        if (action === "add") {
          updatedPullRequest = await gitApi.createPullRequestReviewers(
            reviewerIds.map((id) => ({ id: id })),
            repositoryId,
            pullRequestId
          );

          return {
            content: [{ type: "text", text: JSON.stringify(updatedPullRequest, null, 2) }],
          };
        } else {
          for (const reviewerId of reviewerIds) {
            await gitApi.deletePullRequestReviewer(repositoryId, pullRequestId, reviewerId);
          }

          return {
            content: [{ type: "text", text: `Reviewers with IDs ${reviewerIds.join(", ")} removed from pull request ${pullRequestId}.` }],
          };
        }
      }
    );
  }

  if (!disabledTools.has(REPO_TOOLS.list_pull_requests_by_repo)) {
    server.tool(
      REPO_TOOLS.list_pull_requests_by_repo,
      "Retrieve a list of pull requests for a given repository.",
      {
        repositoryId: z.string().describe("The ID of the repository where the pull requests are located."),
        top: z.number().default(100).describe("The maximum number of pull requests to return."),
        skip: z.number().default(0).describe("The number of pull requests to skip."),
        created_by_me: z.boolean().default(false).describe("Filter pull requests created by the current user."),
        i_am_reviewer: z.boolean().default(false).describe("Filter pull requests where the current user is a reviewer."),
        status: z.enum(["abandoned", "active", "all", "completed", "notSet"]).default("active").describe("Filter pull requests by status. Defaults to 'active'."),
      },
      async ({ repositoryId, top, skip, created_by_me, i_am_reviewer, status }) => {
        const connection = await connectionProvider();
        const gitApi = await connection.getGitApi();

        // Build the search criteria
        const searchCriteria: {
          status: number;
          repositoryId: string;
          creatorId?: string;
          reviewerId?: string;
        } = {
          status: pullRequestStatusStringToInt(status),
          repositoryId: repositoryId,
        };

        if (created_by_me || i_am_reviewer) {
          const data = await getCurrentUserDetails(tokenProvider, connectionProvider);
          const userId = data.authenticatedUser.id;
          if (created_by_me) {
            searchCriteria.creatorId = userId;
          }
          if (i_am_reviewer) {
            searchCriteria.reviewerId = userId;
          }
        }

        const pullRequests = await gitApi.getPullRequests(
          repositoryId,
          searchCriteria,
          undefined, // project
          undefined, // maxCommentLength
          skip,
          top
        );

        const prResult = pullRequests ? pullRequests.map((pr) => toPullRequestResult(pr)) : [];

        return {
          content: [{ type: "text", text: JSON.stringify(prResult) }],
        };
      }
    );
  }

  if (!disabledTools.has(REPO_TOOLS.list_pull_requests_by_project)) {
    server.tool(
      REPO_TOOLS.list_pull_requests_by_project,
      "Retrieve a list of pull requests for a given project Id or Name.",
      {
        project: z.string().describe("The name or ID of the Azure DevOps project."),
        top: z.number().default(100).describe("The maximum number of pull requests to return."),
        skip: z.number().default(0).describe("The number of pull requests to skip."),
        created_by_me: z.boolean().default(false).describe("Filter pull requests created by the current user."),
        i_am_reviewer: z.boolean().default(false).describe("Filter pull requests where the current user is a reviewer."),
        status: z.enum(["abandoned", "active", "all", "completed", "notSet"]).default("active").describe("Filter pull requests by status. Defaults to 'active'."),
      },
      async ({ project, top, skip, created_by_me, i_am_reviewer, status }) => {
        const connection = await connectionProvider();
        const gitApi = await connection.getGitApi();

        // Build the search criteria
        const gitPullRequestSearchCriteria: {
          status: number;
          creatorId?: string;
          reviewerId?: string;
        } = {
          status: pullRequestStatusStringToInt(status),
        };

        if (created_by_me || i_am_reviewer) {
          const data = await getCurrentUserDetails(tokenProvider, connectionProvider);
          const userId = data.authenticatedUser.id;
          if (created_by_me) {
            gitPullRequestSearchCriteria.creatorId = userId;
          }
          if (i_am_reviewer) {
            gitPullRequestSearchCriteria.reviewerId = userId;
          }
        }

        const pullRequests = await gitApi.getPullRequestsByProject(
          project,
          gitPullRequestSearchCriteria,
          undefined, // maxCommentLength
          skip,
          top
        );

        const prResult = pullRequests ? pullRequests.map((pr) => toPullRequestResult(pr)) : [];

        return {
          content: [{ type: "text", text: JSON.stringify(prResult) }],
        };
      }
    );
  }

  if (!disabledTools.has(REPO_TOOLS.list_branches_by_repo)) {
    server.tool(
      REPO_TOOLS.list_branches_by_repo,
      "Retrieve a list of branches for a given repository.",
      {
        repositoryId: z.string().describe("The ID of the repository where the branches are located."),
        top: z.number().default(100).describe("The maximum number of branches to return. Defaults to 100."),
      },
      async ({ repositoryId, top }) => {
        const connection = await connectionProvider();
        const gitApi = await connection.getGitApi();
        const branches = await gitApi.getRefs(repositoryId, undefined);

        const filteredBranches = branchesFilterOutIrrelevantProperties(branches, top);

        return {
          content: [{ type: "text", text: JSON.stringify(filteredBranches, null, 2) }],
        };
      }
    );
  }

  if (!disabledTools.has(REPO_TOOLS.list_my_branches_by_repo)) {
    server.tool(
      REPO_TOOLS.list_my_branches_by_repo,
      "Retrieve a list of my branches for a given repository Id.",
      {
        repositoryId: z.string().describe("The ID of the repository where the branches are located."),
        top: z.number().default(100).describe("The maximum number of branches to return."),
      },
      async ({ repositoryId, top }) => {
        const connection = await connectionProvider();
        const gitApi = await connection.getGitApi();
        const branches = await gitApi.getRefs(repositoryId, undefined, undefined, undefined, undefined, true);

        const filteredBranches = branchesFilterOutIrrelevantProperties(branches, top);

        return {
          content: [{ type: "text", text: JSON.stringify(filteredBranches, null, 2) }],
        };
      }
    );
  }

  if (!disabledTools.has(REPO_TOOLS.get_branch_by_name)) {
    server.tool(
      REPO_TOOLS.get_branch_by_name,
      "Get a branch by its name.",
      {
        repositoryId: z.string().describe("The ID of the repository where the branch is located."),
        branchName: z.string().describe("The name of the branch to retrieve, e.g., 'main' or 'feature-branch'."),
      },
      async ({ repositoryId, branchName }) => {
        const connection = await connectionProvider();
        const gitApi = await connection.getGitApi();
        const branches = await gitApi.getRefs(repositoryId);
        const branch = branches?.find((branch) => branch.name === `refs/heads/${branchName}`);
        if (!branch) {
          return {
            content: [
              {
                type: "text",
                text: `Branch ${branchName} not found in repository ${repositoryId}`,
              },
            ],
          };
        }
        return {
          content: [{ type: "text", text: JSON.stringify(branch, null, 2) }],
        };
      }
    );
  }

  if (!disabledTools.has(REPO_TOOLS.reply_to_comment)) {
    server.tool(
      REPO_TOOLS.reply_to_comment,
      "Replies to a specific comment on a pull request.",
      {
        repositoryId: z.string().describe("The ID of the repository where the pull request is located."),
        pullRequestId: z.number().describe("The ID of the pull request where the comment thread exists."),
        threadId: z.number().describe("The ID of the thread to which the comment will be added."),
        content: z.string().describe("The content of the comment to be added."),
        project: z.string().optional().describe("Project ID or project name (optional)"),
      },
      async ({ repositoryId, pullRequestId, threadId, content, project }) => {
        const connection = await connectionProvider();
        const gitApi = await connection.getGitApi();
        const comment = await gitApi.createComment({ content }, repositoryId, pullRequestId, threadId, project);

        return {
          content: [{ type: "text", text: JSON.stringify(comment, null, 2) }],
        };
      }
    );
  }

  if (!disabledTools.has(REPO_TOOLS.resolve_comment)) {
    server.tool(
      REPO_TOOLS.resolve_comment,
      "Resolves a specific comment thread on a pull request.",
      {
        repositoryId: z.string().describe("The ID of the repository where the pull request is located."),
        pullRequestId: z.number().describe("The ID of the pull request where the comment thread exists."),
        threadId: z.number().describe("The ID of the thread to be resolved."),
      },
      async ({ repositoryId, pullRequestId, threadId }) => {
        const connection = await connectionProvider();
        const gitApi = await connection.getGitApi();
        const thread = await gitApi.updateThread(
          { status: 2 }, // 2 corresponds to "Resolved" status
          repositoryId,
          pullRequestId,
          threadId
        );

        return {
          content: [{ type: "text", text: JSON.stringify(thread, null, 2) }],
        };
      }
    );
  }

  if (!disabledTools.has(REPO_TOOLS.search_commits)) {
    const gitVersionTypeStrings = Object.values(GitVersionType).filter((value): value is string => typeof value === "string");

    server.tool(
      REPO_TOOLS.search_commits,
      "Searches for commits in a repository",
      {
        project: z.string().describe("Project name or ID"),
        repository: z.string().describe("Repository name or ID"),
        fromCommit: z.string().optional().describe("Starting commit ID"),
        toCommit: z.string().optional().describe("Ending commit ID"),
        version: z.string().optional().describe("The name of the branch, tag or commit to filter commits by"),
        versionType: z
          .enum(gitVersionTypeStrings as [string, ...string[]])
          .optional()
          .default(GitVersionType[GitVersionType.Branch])
          .describe("The meaning of the version parameter, e.g., branch, tag or commit"),
        skip: z.number().optional().default(0).describe("Number of commits to skip"),
        top: z.number().optional().default(10).describe("Maximum number of commits to return"),
        includeLinks: z.boolean().optional().default(false).describe("Include commit links"),
        includeWorkItems: z.boolean().optional().default(false).describe("Include associated work items"),
      },
      async ({ project, repository, fromCommit, toCommit, version, versionType, skip, top, includeLinks, includeWorkItems }) => {
        try {
          const connection = await connectionProvider();
          const gitApi = await connection.getGitApi();

          const searchCriteria: GitQueryCommitsCriteria = {
            fromCommitId: fromCommit,
            toCommitId: toCommit,
            includeLinks: includeLinks,
            includeWorkItems: includeWorkItems,
          };

          if (version) {
            const itemVersion: GitVersionDescriptor = {
              version: version,
              versionType: GitVersionType[versionType as keyof typeof GitVersionType],
            };
            searchCriteria.itemVersion = itemVersion;
          }

          const commits = await gitApi.getCommits(
            repository,
            searchCriteria,
            project,
            skip, // skip
            top
          );

          return {
            content: [{ type: "text", text: JSON.stringify(commits, null, 2) }],
          };
        } catch (error) {
          return {
            content: [
              {
                type: "text",
                text: `Error searching commits: ${error instanceof Error ? error.message : String(error)}`,
              },
            ],
            isError: true,
          };
        }
      }
    );
  }

  if (!disabledTools.has(REPO_TOOLS.list_pull_requests_by_commits)) {
    const pullRequestQueryTypesStrings = Object.values(GitPullRequestQueryType).filter((value): value is string => typeof value === "string");

    server.tool(
      REPO_TOOLS.list_pull_requests_by_commits,
      "Lists pull requests by commit IDs to find which pull requests contain specific commits",
      {
        project: z.string().describe("Project name or ID"),
        repository: z.string().describe("Repository name or ID"),
        commits: z.array(z.string()).describe("Array of commit IDs to query for"),
        queryType: z
          .enum(pullRequestQueryTypesStrings as [string, ...string[]])
          .optional()
          .default(GitPullRequestQueryType[GitPullRequestQueryType.LastMergeCommit])
          .describe("Type of query to perform"),
      },
      async ({ project, repository, commits, queryType }) => {
        try {
          const connection = await connectionProvider();
          const gitApi = await connection.getGitApi();

          const query: GitPullRequestQuery = {
            queries: [
              {
                items: commits,
                type: GitPullRequestQueryType[queryType as keyof typeof GitPullRequestQueryType],
              } as GitPullRequestQueryInput,
            ],
          };

          const queryResult = await gitApi.getPullRequestQuery(query, repository, project);

          return {
            content: [{ type: "text", text: JSON.stringify(queryResult, null, 2) }],
          };
        } catch (error) {
          return {
            content: [
              {
                type: "text",
                text: `Error querying pull requests by commits: ${error instanceof Error ? error.message : String(error)}`,
              },
            ],
            isError: true,
          };
        }
      }
    );
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
  const pullRequest = await gitApi.getPullRequest(repositoryId, pullRequestId);
  if (!pullRequest) {
    throw new Error("Pull request not found.");
  }

  const prResult = toPullRequestResult(pullRequest);
  const commits = await gitApi.getPullRequestCommits(repositoryId, pullRequestId);
  prResult.commits = commits ? commits.map((commit) => toGitCommitResult(commit)) : [];

  if (commits && commits.length) {
    const latestCommitId = commits[0].commitId; // First commit is usually the latest
    const targetBranch = pullRequest.targetRefName?.replace("refs/heads/", "") || "main";
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
      false, // diffCommonCommit
      100,
      undefined,
      baseVersionDescriptor,
      targetVersionDescriptor
    );
    prResult.pullRequestChanges = diffs && toGitCommitDiffsResult(diffs);
  }
  return prResult;
}

export { REPO_TOOLS, configureRepoTools };
