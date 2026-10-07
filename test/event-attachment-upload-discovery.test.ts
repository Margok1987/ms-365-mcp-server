import { describe, expect, it, vi } from 'vitest';
import { registerDiscoveryTools } from '../src/graph-tools.js';

type Handler = (args: any) => Promise<any>;

function discoveryHarness(allowedScopes?: string) {
  const legacyHandlers = new Map<string, Handler>();
  const registeredHandlers = new Map<string, Handler>();
  const registeredConfigs = new Map<string, any>();

  const server = {
    tool: vi.fn((name: string, ...args: any[]) => {
      legacyHandlers.set(name, args[args.length - 1] as Handler);
    }),
    registerTool: vi.fn((name: string, config: any, handler: Handler) => {
      registeredConfigs.set(name, config);
      registeredHandlers.set(name, handler);
    }),
    server: {
      _requestHandlers: new Map(),
      setRequestHandler: vi.fn(),
    },
  };

  registerDiscoveryTools(
    server as any,
    {} as any,
    false,
    true,
    undefined,
    false,
    [],
    undefined,
    allowedScopes,
    true
  );

  return { legacyHandlers, registeredHandlers, registeredConfigs };
}

function textJson(result: any) {
  return JSON.parse(result.content.find((item: any) => item.type === 'text').text);
}

describe('large event attachment discovery contract', () => {
  it('discovers the semantic uploader in the calendar category', async () => {
    const { legacyHandlers } = discoveryHarness('Calendars.ReadWrite');
    const search = legacyHandlers.get('search-tools')!;
    const result = await search({
      query: 'upload large attachment to calendar event',
      category: 'calendar',
      limit: 20,
    });
    expect(textJson(result).tools.some((tool: any) => tool.name === 'upload-large-event-attachment')).toBe(
      true
    );
  });

  it('publishes immutable-id and host-file schema requirements', async () => {
    const { legacyHandlers } = discoveryHarness('Calendars.ReadWrite');
    const schemaHandler = legacyHandlers.get('get-tool-schema')!;
    const result = await schemaHandler({ tool_name: 'upload-large-event-attachment' });
    const body = textJson(result);
    const serialized = JSON.stringify(body);
    expect(serialized).toContain('restImmutableEntryId');
    expect(serialized).toContain('eventId');
    expect(serialized).toContain('file');
    expect(serialized).toContain('confirm');
  });

  it('requires the top-level ChatGPT file parameter', async () => {
    const { registeredHandlers } = discoveryHarness('Calendars.ReadWrite');
    const execute = registeredHandlers.get('execute-tool')!;
    const result = await execute({
      tool_name: 'upload-large-event-attachment',
      parameters: {
        eventId: 'immutable-event',
        idKind: 'restImmutableEntryId',
        confirm: true,
      },
    });
    expect(result.isError).toBe(true);
    expect(textJson(result).error).toContain('requires a ChatGPT file parameter');
  });

  it('rejects a manually nested parameters.file value', async () => {
    const { registeredHandlers } = discoveryHarness('Calendars.ReadWrite');
    const execute = registeredHandlers.get('execute-tool')!;
    const result = await execute({
      tool_name: 'upload-large-event-attachment',
      parameters: {
        eventId: 'immutable-event',
        idKind: 'restImmutableEntryId',
        confirm: true,
        file: {
          download_url: 'https://files.example.test/file',
          file_id: 'forged',
        },
      },
    });
    expect(result.isError).toBe(true);
    expect(textJson(result).error).toContain('top-level execute-tool.file');
  });

  it('enforces confirm=true before staging a host file', async () => {
    const { registeredHandlers } = discoveryHarness('Calendars.ReadWrite');
    const execute = registeredHandlers.get('execute-tool')!;
    const result = await execute({
      tool_name: 'upload-large-event-attachment',
      parameters: {
        eventId: 'immutable-event',
        idKind: 'restImmutableEntryId',
        confirm: false,
      },
      file: {
        download_url: 'https://files.example.test/file',
        file_id: 'file_123',
        file_name: 'calendar-qa.bin',
      },
    });
    expect(result.isError).toBe(true);
    expect(textJson(result).error).toContain('confirm=true');
  });

  it('is removed from discovery when Calendars.ReadWrite is outside allowed scopes', async () => {
    const { legacyHandlers } = discoveryHarness('Mail.Read');
    const search = legacyHandlers.get('search-tools')!;
    const result = await search({
      query: 'upload large attachment to calendar event',
      category: 'calendar',
      limit: 20,
    });
    expect(textJson(result).tools.some((tool: any) => tool.name === 'upload-large-event-attachment')).toBe(
      false
    );
  });
});
