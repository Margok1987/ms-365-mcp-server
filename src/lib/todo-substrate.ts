export const EXO_TASKS_SCOPE = 'https://outlook.office.com/Tasks.ReadWrite';
export const TODO_MOVE_FEATURE_ENV = 'MS365_MCP_ENABLE_TODO_MOVE';
export const SUBSTRATE_TODO_BASE = 'https://substrate.office.com/todob2/api/v1';

const EXO_AUDIENCES = new Set([
  'https://outlook.office.com',
  '00000002-0000-0ff1-ce00-000000000000',
]);

export interface DelegatedPrincipal {
  oid: string;
  tid: string;
  aud: string;
  scopes: string[];
}

function decodeJwtClaims(token: string): Record<string, unknown> {
  const parts = token.split('.');
  if (parts.length < 2 || !parts[1]) {
    throw new Error('Access token is not a JWT with a readable delegated-principal payload.');
  }
  try {
    const parsed = JSON.parse(Buffer.from(parts[1], 'base64url').toString('utf8'));
    if (!parsed || typeof parsed !== 'object') throw new Error('claims are not an object');
    return parsed as Record<string, unknown>;
  } catch (error) {
    throw new Error(`Access token claims could not be decoded: ${(error as Error).message}`);
  }
}

export function delegatedPrincipalFromAccessToken(token: string): DelegatedPrincipal {
  const claims = decodeJwtClaims(token);
  const oid = typeof claims.oid === 'string' ? claims.oid : '';
  const tid = typeof claims.tid === 'string' ? claims.tid : '';
  const aud = typeof claims.aud === 'string' ? claims.aud : '';
  const scopes =
    typeof claims.scp === 'string'
      ? claims.scp
          .split(/\s+/)
          .map((scope) => scope.trim())
          .filter(Boolean)
      : [];

  if (!oid || !tid) {
    throw new Error('Access token does not expose delegated oid/tid claims.');
  }
  return { oid, tid, aud, scopes };
}

export function assertTodoMoveTokenPair(
  graphAccessToken: string,
  exchangeAccessToken: string
): { anchorMailbox: string; oid: string; tid: string } {
  const graph = delegatedPrincipalFromAccessToken(graphAccessToken);
  const exchange = delegatedPrincipalFromAccessToken(exchangeAccessToken);

  if (graph.oid !== exchange.oid || graph.tid !== exchange.tid) {
    throw new Error(
      'Graph and Exchange Online tokens belong to different delegated principals; refusing To Do move.'
    );
  }

  if (!EXO_AUDIENCES.has(exchange.aud.replace(/\/$/, ''))) {
    throw new Error(
      `Exchange Online token has unexpected audience ${JSON.stringify(exchange.aud)}; refusing To Do move.`
    );
  }

  if (!exchange.scopes.some((scope) => scope.toLowerCase() === 'tasks.readwrite')) {
    throw new Error('Exchange Online token lacks delegated Tasks.ReadWrite; refusing To Do move.');
  }

  return {
    anchorMailbox: `OID:${exchange.oid}@${exchange.tid}`,
    oid: exchange.oid,
    tid: exchange.tid,
  };
}

export interface TodoMoveResult {
  previousTaskId: string;
  taskId: string;
  sourceListId: string;
  destinationListId: string;
  parentFolderId: string;
}

export interface MoveTodoTaskInput {
  accessToken: string;
  sourceListId: string;
  taskId: string;
  destinationListId: string;
  anchorMailbox?: string;
  fetchImpl?: typeof fetch;
}

export class TodoMoveOutcomeUnknownError extends Error {
  readonly code = 'TODO_MOVE_OUTCOME_UNKNOWN';
  readonly unknownOutcome = true;

  constructor(message: string) {
    super(message);
    this.name = 'TodoMoveOutcomeUnknownError';
  }
}

export class TodoMoveRejectedError extends Error {
  readonly code = 'TODO_MOVE_REJECTED';

  constructor(
    message: string,
    readonly status: number
  ) {
    super(message);
    this.name = 'TodoMoveRejectedError';
  }
}

export function todoMoveEnabled(env: NodeJS.ProcessEnv = process.env): boolean {
  const raw = env[TODO_MOVE_FEATURE_ENV]?.trim().toLowerCase();
  return raw === '1' || raw === 'true';
}

function assertOpaqueId(value: string, name: string): void {
  if (value.trim() === '') {
    throw new Error(`${name} must be a non-empty Microsoft To Do native id.`);
  }
}

function substrateTaskUrl(destinationListId: string, taskId: string): string {
  return `${SUBSTRATE_TODO_BASE}/taskfolders/${encodeURIComponent(destinationListId)}/tasks/${encodeURIComponent(taskId)}`;
}

function responseSnippet(raw: string): string {
  const compact = raw.replace(/\s+/g, ' ').trim();
  return compact.length > 512 ? `${compact.slice(0, 512)}...` : compact;
}

/**
 * Re-parent exactly one Microsoft To Do task through the first-party To Do
 * Substrate endpoint.
 *
 * Deliberately no retry and no Graph copy/delete fallback. A transport failure,
 * 5xx, or a 2xx response that does not confirm the new parent is
 * TODO_MOVE_OUTCOME_UNKNOWN. The caller must reconcile source and destination
 * state before any later dispatch.
 */
export async function moveTodoTaskViaSubstrate(input: MoveTodoTaskInput): Promise<TodoMoveResult> {
  assertOpaqueId(input.sourceListId, 'sourceListId');
  assertOpaqueId(input.taskId, 'taskId');
  assertOpaqueId(input.destinationListId, 'destinationListId');
  if (input.sourceListId === input.destinationListId) {
    throw new Error('sourceListId and destinationListId must differ.');
  }
  if (input.accessToken.trim() === '') {
    throw new Error('Exchange Online access token is required.');
  }

  const fetchImpl = input.fetchImpl ?? fetch;
  const headers: Record<string, string> = {
    authorization: `Bearer ${input.accessToken}`,
    accept: 'application/json',
    'content-type': 'application/json',
  };
  if (input.anchorMailbox?.trim()) {
    headers['x-anchormailbox'] = input.anchorMailbox.trim();
  }

  let response: Response;
  try {
    response = await fetchImpl(substrateTaskUrl(input.destinationListId, input.taskId), {
      method: 'PATCH',
      headers,
      body: JSON.stringify({ ParentFolderId: input.destinationListId }),
    });
  } catch (error) {
    throw new TodoMoveOutcomeUnknownError(
      `Substrate move transport failed after dispatch could have started: ${(error as Error).message}`
    );
  }

  const raw = await response.text().catch(() => '');
  if (!response.ok) {
    const detail = responseSnippet(raw);
    const suffix = detail ? `: ${detail}` : '';
    if (response.status >= 500) {
      throw new TodoMoveOutcomeUnknownError(
        `Substrate move returned HTTP ${response.status}; reconcile native state before retry${suffix}`
      );
    }
    throw new TodoMoveRejectedError(
      `Substrate move rejected with HTTP ${response.status}${suffix}`,
      response.status
    );
  }

  let body: unknown;
  try {
    body = raw ? JSON.parse(raw) : null;
  } catch {
    throw new TodoMoveOutcomeUnknownError(
      'Substrate move returned success with a non-JSON body; reconcile native state before retry.'
    );
  }

  const record = body as { Id?: unknown; ParentFolderId?: unknown } | null;
  const newTaskId = typeof record?.Id === 'string' ? record.Id : '';
  const parentFolderId = typeof record?.ParentFolderId === 'string' ? record.ParentFolderId : '';

  if (!newTaskId || parentFolderId !== input.destinationListId) {
    throw new TodoMoveOutcomeUnknownError(
      'Substrate move response did not confirm both the new task identity and destination parent; reconcile native state before retry.'
    );
  }

  return {
    previousTaskId: input.taskId,
    taskId: newTaskId,
    sourceListId: input.sourceListId,
    destinationListId: input.destinationListId,
    parentFolderId,
  };
}
