/* eslint-disable header/header */
import type { IdentityBase } from "azure-devops-node-api/interfaces/IdentitiesInterfaces.js";
import type { TeamProjectReference, WebApiTeam } from "azure-devops-node-api/interfaces/CoreInterfaces.js";
import type { GitRepository } from "azure-devops-node-api/interfaces/GitInterfaces.js";

type RepositoryLike = Pick<GitRepository, "id" | "name" | "defaultBranch" | "isFork">;

function getIdentityDescriptor(identity: IdentityBase): string | undefined {
  // Runtime shape is often a plain string even though the type declares IdentityDescriptor.
  const descriptor = identity.descriptor as unknown;
  if (typeof descriptor === "string") return descriptor;
  if (descriptor && typeof descriptor === "object" && "identifier" in descriptor) {
    const identifier = (descriptor as { identifier?: string }).identifier;
    return typeof identifier === "string" ? identifier : undefined;
  }
  return undefined;
}

export function formatProject(project: TeamProjectReference): string {
  const name = project.name ?? "(unnamed)";
  const id = project.id ?? "unknown-id";
  const projectNameWithMetadata = `${name} (ID: ${id}, state: ${project.state})`;
  return project.description ? `${projectNameWithMetadata} - ${project.description}` : projectNameWithMetadata;
}

export function formatProjectLines(projects: TeamProjectReference[]): string[] {
  return projects.map((p) => `- ${formatProject(p)}`);
}

export function formatTeam(team: WebApiTeam): string {
  const name = team.name ?? "(unnamed)";
  const id = team.id ?? "unknown-id";
  const teamNameWithMetadata = `${name} (ID: ${id})`;
  return team.description ? `${teamNameWithMetadata} - ${team.description}` : teamNameWithMetadata;
}

export function formatTeamLines(teams: WebApiTeam[]): string[] {
  return teams.map((t) => `- ${formatTeam(t)}`);
}

export function formatIdentity(identity: IdentityBase): string {
  const name = identity.customDisplayName || identity.providerDisplayName || "(no display name)";
  const id = identity.id ?? "unknown-id";
  const attrs: string[] = [`ID: ${id}`];
  const descriptor = getIdentityDescriptor(identity);
  if (descriptor) {
    attrs.push(`descriptor: ${descriptor}`);
  }

  return `${name} (${attrs.join(", ")})`;
}

export function formatIdentityLines(identities: IdentityBase[]): string[] {
  return identities.map((i) => `- ${formatIdentity(i)}`);
}

export function stripRefHeads(ref: string): string {
  return ref.replace(/^refs\/heads\//, "");
}

export function formatRepository(repo: RepositoryLike): string {
  const name = repo.name ?? "(unnamed)";
  const id = repo.id ?? "unknown-id";
  const defaultBranch = repo.defaultBranch ? stripRefHeads(repo.defaultBranch) : "(no default)";
  return `${name} (ID: ${id}, default branch: ${defaultBranch})`;
}

export function formatRepositoryLines(repos: RepositoryLike[]): string[] {
  return repos.map((r) => `- ${formatRepository(r)}`);
}

/**
 * Builds a self-contained pagination line meant to live on its own line in tool output, e.g.
 * "[Total: 150 | Range: 1-25 | Top: 25 | Skip: 0 | More: yes (set skip=25)]".
 * Callers are expected to invoke this only when their tool actually paginates, so `top`/`skip`
 * are required. `total` stays optional: when the caller server-paginates and can't know the
 * grand total, `Total` is omitted and `More` falls back to a full-page heuristic.
 */
export function formatPagination(shown: number, top: number, skip: number, total?: number): string {
  const range = shown > 0 ? `${skip + 1}-${skip + shown}` : "0";
  const parts: string[] = [];

  if (total !== undefined) {
    parts.push(`Total: ${total}`);
  }
  parts.push(`Range: ${range}`, `Top: ${top}`, `Skip: ${skip}`);

  let more: string;
  if (total !== undefined) {
    more = total - (skip + shown) > 0 ? `yes (set skip=${skip + shown})` : "no";
  } else {
    more = shown >= top ? `maybe (set skip=${skip + shown} to check)` : "no";
  }
  parts.push(`More: ${more}`);

  return `[${parts.join(" | ")}]`;
}
