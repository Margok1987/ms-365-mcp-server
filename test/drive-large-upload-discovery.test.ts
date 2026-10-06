import { describe, expect, it, vi } from 'vitest';
import { registerDiscoveryTools } from '../src/graph-tools.js';

type Handler = (args: any) => Promise<any>;

function discoveryHarness() {
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
    undefined,
    true
  );

  return { server, legacyHandlers, registeredHandlers, registeredConfigs };
}

function textJson(result: any) {
  return JSON.parse(result.content.find((item: any) => item.type === 'text').text);
}

describe('large file upload discovery contract', () => {
  it('keeps the public discovery surface at three tools and exposes execute-tool file metadata', () => {
    const { server, registeredConfigs } = discoveryHarness();
    expect(server.tool).toHaveBeenCalledTimes(2);
    expect(server.registerTool).toHaveBeenCalledTimes(1);

    const execute = registeredConfigs.get('execute-tool');
    expect(execute).toBeTruthy();
    expect(execute._meta).toEqual({ 'openai/fileParams': ['file'] });
    expect(execute.annotations).toMatchObject({
      readOnlyHint: false,
      destructiveHint: true,
      openWorldHint: true,
    });
  });

  it('discovers upload-drive-file inside the files category', async () => {
    const { legacyHandlers } = discoveryHarness();
    const search = legacyHandlers.get('search-tools')!;
    const result = await search({ query: 'large file upload', category: 'files', limit: 20 });
    const body = textJson(result);
    expect(body.tools.some((tool: any) => tool.name === 'upload-drive-file')).toBe(true);
  });

  it('requires the top-level ChatGPT file parameter', async () => {
    const { registeredHandlers } = discoveryHarness();
    const execute = registeredHandlers.get('execute-tool')!;
    const result = await execute({
      tool_name: 'upload-drive-file',
      parameters: { driveId: 'drive', parentItemId: 'parent', confirm: true },
    });
    expect(result.isError).toBe(true);
    expect(textJson(result).error).toContain('requires a ChatGPT file parameter');
  });

  it('rejects a manually nested parameters.file value', async () => {
    const { registeredHandlers } = discoveryHarness();
    const execute = registeredHandlers.get('execute-tool')!;
    const result = await execute({
      tool_name: 'upload-drive-file',
      parameters: {
        driveId: 'drive',
        parentItemId: 'parent',
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

  it('still enforces confirm=true after file injection', async () => {
    const { registeredHandlers } = discoveryHarness();
    const execute = registeredHandlers.get('execute-tool')!;
    const result = await execute({
      tool_name: 'upload-drive-file',
      parameters: { driveId: 'drive', parentItemId: 'parent', confirm: false },
      file: {
        download_url: 'https://files.example.test/file',
        file_id: 'file_123',
        file_name: 'report.pdf',
      },
    });
    expect(result.isError).toBe(true);
    expect(textJson(result).error).toContain('confirm=true');
  });
});
