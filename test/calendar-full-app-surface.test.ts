import assert from 'node:assert/strict';
import test from 'node:test';
import endpoints from '../src/endpoints.json' with { type: 'json' };

type Endpoint = {
  pathPattern: string;
  method: string;
  toolName: string;
  presets?: string[];
  scopes?: string[];
  workScopes?: string[];
  readOnly?: boolean;
};

const entries = endpoints as Endpoint[];
const byName = new Map(entries.map((entry) => [entry.toolName, entry]));

test('calendar full-app provider additions have unique tool names', () => {
  assert.equal(new Set(entries.map((entry) => entry.toolName)).size, entries.length);
});

test('calendar full-app provider additions are present', () => {
  const expected = [
  "get-primary-calendar",
  "get-calendar",
  "permanently-delete-calendar",
  "translate-exchange-ids",
  "permanently-delete-calendar-event",
  "list-calendar-groups",
  "create-calendar-group",
  "get-calendar-group",
  "update-calendar-group",
  "delete-calendar-group",
  "list-calendar-reminders",
  "list-event-attachments",
  "get-event-attachment",
  "add-event-attachment",
  "delete-event-attachment",
  "create-event-attachment-upload-session",
  "list-rooms",
  "list-room-lists",
  "list-workspaces",
  "get-my-calendar-permission",
  "list-shared-calendars",
  "get-shared-calendar",
  "create-shared-calendar-event",
  "update-shared-calendar-event",
  "cancel-shared-calendar-event",
  "accept-shared-calendar-event",
  "tentatively-accept-shared-calendar-event",
  "decline-shared-calendar-event",
  "get-outlook-category",
  "update-outlook-category",
  "delete-outlook-category",
  "get-work-hours-and-locations",
  "update-work-hours-and-locations",
  "list-work-plan-recurrences",
  "create-work-plan-recurrence",
  "update-work-plan-recurrence",
  "delete-work-plan-recurrence",
  "get-work-plan-occurrences-view",
  "create-time-off-work-plan-occurrence",
  "update-time-off-work-plan-occurrence",
  "delete-time-off-work-plan-occurrence",
  "set-current-work-location",
  "list-event-open-extensions",
  "get-event-open-extension",
  "create-event-open-extension",
  "update-event-open-extension",
  "delete-event-open-extension"
];
  for (const name of expected) {
    assert.ok(byName.has(name), `missing endpoint tool: ${name}`);
  }
});

test('modern Work Hours & Locations preserves PUT semantics for full replacements', () => {
  assert.equal(byName.get('update-work-plan-recurrence')?.method, 'put');
  assert.equal(byName.get('update-time-off-work-plan-occurrence')?.method, 'put');
});

test('translateExchangeIds remains read-only despite POST transport', () => {
  assert.equal(byName.get('translate-exchange-ids')?.method, 'post');
  assert.equal(byName.get('translate-exchange-ids')?.readOnly, true);
});

test('external-effect and cross-family routes are visible in the calendar preset', () => {
  for (const name of [
    'list-shared-calendar-events',
    'get-shared-calendar-view',
    'get-group-calendar-view',
    'list-group-events',
    'get-mailbox-settings',
    'update-mailbox-settings',
    'create-subscription',
    'get-subscription',
    'update-subscription',
    'delete-subscription',
    'reauthorize-subscription',
  ]) {
    assert.ok(byName.get(name)?.presets?.includes('calendar'), `${name} missing calendar preset`);
  }
});
