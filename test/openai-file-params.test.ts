import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { describe, expect, it, vi } from 'vitest';
import { zodToJsonSchema } from 'zod-to-json-schema';
import type { GraphClient } from '../src/graph-client.js';
import { registerDiscoveryTools, registerGraphTools, UTILITY_TOOLS } from '../src/graph-tools.js';
import { describeUtilityToolSchema } from '../src/lib/tool-schema.js';

const fileObject = {
  download_url: 'https://1.1.1.1/file',
  file_id: 'file-1',
  mime_type: 'application/pdf',
  file_name: 'report.pdf',
};

describe('OpenAI MCP file parameter contract', () => {
  it('does not add a separate attachment utility', () => {
    expect(UTILITY_TOOLS.some((entry) => entry.name === 'upload-attachment')).toBe(false);
    expect(UTILITY_TOOLS.some((entry) => entry.name === 'get-upload-url')).toBe(true);
  });

  it('registers get-upload-url with openai/fileParams metadata in direct mode', () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const registerTool = vi.spyOn(server, 'registerTool').mockImplementation((() => {}) as never);
    vi.spyOn(server, 'tool').mockImplementation((() => {}) as never);

    registerGraphTools(server, {} as GraphClient, { orgMode: true });

    const call = registerTool.mock.calls.find((entry) => entry[0] === 'get-upload-url');
    expect(call).toBeDefined();
    const config = call?.[1] as {
      inputSchema: { safeParse: (value: unknown) => { success: boolean } };
      _meta?: Record<string, unknown>;
    };
    expect(config._meta).toEqual({ 'openai/fileParams': ['file'] });
    expect(
      config.inputSchema.safeParse({
        target: '/me/messages/AAA/attachments',
        name: 'report.pdf',
        file: fileObject,
      }).success
    ).toBe(true);
  });

  it('reports the host file as File in discovery schema with the required OpenAI fields', () => {
    const utility = UTILITY_TOOLS.find((entry) => entry.name === 'get-upload-url');
    expect(utility).toBeDefined();
    const schema = describeUtilityToolSchema(utility!, {
      graphClient: {} as GraphClient,
      multiAccount: false,
      accountNames: [],
    });
    expect(schema.fileParams).toEqual(['file']);
    const file = schema.parameters.find((parameter) => parameter.name === 'file');
    expect(file?.in).toBe('File');
    expect(file?.required).toBe(false);
    expect(file?.schema).toMatchObject({
      type: 'object',
      required: ['download_url', 'file_id'],
      additionalProperties: false,
      properties: {
        download_url: { type: 'string' },
        file_id: { type: 'string' },
        mime_type: { type: 'string' },
        file_name: { type: 'string' },
      },
    });
  });

  it('exposes one top-level file parameter on discovery execute-tool', () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const registerTool = vi.spyOn(server, 'registerTool').mockImplementation((() => {}) as never);
    vi.spyOn(server, 'tool').mockImplementation((() => {}) as never);

    registerDiscoveryTools(server, {} as GraphClient, { orgMode: true });

    const call = registerTool.mock.calls.find((entry) => entry[0] === 'execute-tool');
    expect(call).toBeDefined();
    const config = call?.[1] as {
      inputSchema: Parameters<typeof zodToJsonSchema>[0];
      _meta?: Record<string, unknown>;
    };
    expect(config._meta).toEqual({ 'openai/fileParams': ['file'] });
    const json = zodToJsonSchema(config.inputSchema, {
      target: 'jsonSchema7',
      $refStrategy: 'none',
    }) as Record<string, unknown>;
    expect(json).toMatchObject({
      type: 'object',
      properties: {
        file: {
          type: 'object',
          required: ['download_url', 'file_id'],
          additionalProperties: false,
        },
      },
    });
  });

  it('fails closed when a file is supplied to a non-file tool or nested in parameters', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const registerTool = vi.spyOn(server, 'registerTool').mockImplementation((() => {}) as never);
    vi.spyOn(server, 'tool').mockImplementation((() => {}) as never);

    registerDiscoveryTools(server, {} as GraphClient, { orgMode: true });

    const call = registerTool.mock.calls.find((entry) => entry[0] === 'execute-tool');
    const handler = call?.[2] as (args: Record<string, unknown>) => Promise<{
      content: Array<{ text: string }>;
      isError?: boolean;
    }>;

    const graphResult = await handler({
      tool_name: 'list-mail-messages',
      file: fileObject,
    });
    expect(graphResult.isError).toBe(true);
    expect(JSON.parse(graphResult.content[0].text).error).toBe('tool_does_not_accept_file');

    const nestedResult = await handler({
      tool_name: 'get-upload-url',
      parameters: { file: fileObject },
    });
    expect(nestedResult.isError).toBe(true);
    expect(JSON.parse(nestedResult.content[0].text).error).toBe('file_must_be_top_level');
  });
});
