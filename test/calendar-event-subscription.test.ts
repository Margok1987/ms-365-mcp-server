import { describe, expect, it, vi } from 'vitest';
import { buildScopesFromEndpoints } from '../src/auth.js';
import { UTILITY_TOOLS } from '../src/graph-tools.js';
import { TOOL_CATEGORIES } from '../src/tool-categories.js';

function jsonText(result: any) {
  return JSON.parse(result.content.find((item: any) => item.type === 'text').text);
}

const utility = UTILITY_TOOLS.find((item) => item.name === 'create-calendar-event-subscription')!;

describe('Calendar immutable event subscription utility', () => {
  it('is present in the Calendar preset and derives Calendars.Read', () => {
    expect(
      new RegExp(TOOL_CATEGORIES.calendar.pattern.source, 'i').test(
        'create-calendar-event-subscription'
      )
    ).toBe(true);
    expect(
      buildScopesFromEndpoints(false, '^create-calendar-event-subscription$', false)
    ).toContain('Calendars.Read');
  });

  it('requires confirmation before calling Graph', async () => {
    const graphClient = { makeRequest: vi.fn() };
    const result = await utility.execute(
      {
        notificationUrl: 'https://qa.example.test/notify',
        lifecycleNotificationUrl: 'https://qa.example.test/lifecycle',
        expirationDateTime: '2099-01-01T00:00:00Z',
        clientState: 'secret-state',
        confirm: false,
      },
      { graphClient: graphClient as any, multiAccount: false, accountNames: [] }
    );
    expect(result.isError).toBe(true);
    expect(graphClient.makeRequest).not.toHaveBeenCalled();
  });

  it('refuses non-HTTPS callback URLs', async () => {
    const graphClient = { makeRequest: vi.fn() };
    const result = await utility.execute(
      {
        notificationUrl: 'http://qa.example.test/notify',
        lifecycleNotificationUrl: 'https://qa.example.test/lifecycle',
        expirationDateTime: '2099-01-01T00:00:00Z',
        clientState: 'secret-state',
        confirm: true,
      },
      { graphClient: graphClient as any, multiAccount: false, accountNames: [] }
    );
    expect(result.isError).toBe(true);
    expect(graphClient.makeRequest).not.toHaveBeenCalled();
  });

  it('creates only /me/events subscriptions with immutable ids and never echoes clientState', async () => {
    const graphClient = {
      makeRequest: vi.fn(async () => ({
        id: 'sub-1',
        resource: '/me/events',
        changeType: 'created,updated,deleted',
        expirationDateTime: '2099-01-01T00:00:00Z',
        notificationUrl: 'https://qa.example.test/notify',
        lifecycleNotificationUrl: 'https://qa.example.test/lifecycle',
        clientState: 'secret-state',
      })),
    };

    const result = await utility.execute(
      {
        notificationUrl: 'https://qa.example.test/notify',
        lifecycleNotificationUrl: 'https://qa.example.test/lifecycle',
        expirationDateTime: '2099-01-01T00:00:00Z',
        clientState: 'secret-state',
        confirm: true,
      },
      { graphClient: graphClient as any, multiAccount: false, accountNames: [] }
    );

    expect(result.isError).not.toBe(true);
    expect(graphClient.makeRequest).toHaveBeenCalledTimes(1);
    const [endpoint, options] = graphClient.makeRequest.mock.calls[0];
    expect(endpoint).toBe('/subscriptions');
    expect(options.headers.Prefer).toBe('IdType="ImmutableId"');
    const body = JSON.parse(options.body);
    expect(body).toMatchObject({
      resource: '/me/events',
      changeType: 'created,updated,deleted',
      clientState: 'secret-state',
      latestSupportedTlsVersion: 'v1_2',
    });

    const receipt = jsonText(result);
    expect(receipt.immutableEventIdsRequested).toBe(true);
    expect(receipt.clientState).toBeUndefined();
    expect(JSON.stringify(receipt)).not.toContain('secret-state');
  });
});
