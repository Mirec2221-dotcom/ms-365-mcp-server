import { beforeEach, describe, expect, it, vi } from 'vitest';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { registerGraphTools } from '../src/graph-tools.js';
import GraphClient from '../src/graph-client.js';

vi.mock('../src/logger.js', () => ({
  default: {
    info: vi.fn(),
    error: vi.fn(),
  },
}));

vi.mock('../src/generated/client.js', () => ({
  api: {
    endpoints: [
      {
        alias: 'list-mail-messages',
        method: 'GET',
        path: '/me/messages',
        description: 'List mail messages',
      },
      { alias: 'send-mail', method: 'POST', path: '/me/sendMail', description: 'Send mail' },
      {
        alias: 'list-calendar-events',
        method: 'GET',
        path: '/me/events',
        description: 'List calendar events',
      },
      {
        alias: 'list-excel-worksheets',
        method: 'GET',
        path: '/workbook/worksheets',
        description: 'List Excel worksheets',
      },
      { alias: 'get-current-user', method: 'GET', path: '/me', description: 'Get current user' },
    ],
  },
}));

describe('Tool Filtering', () => {
  let server: McpServer;
  let graphClient: GraphClient;
  let toolSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    server = new McpServer({ name: 'test', version: '1.0.0' });
    graphClient = {} as GraphClient;
    toolSpy = vi.spyOn(server, 'tool').mockImplementation(() => {});
  });

  it('should register all tools when no filter is provided', async () => {
    await registerGraphTools(server, graphClient, false);

    // 5 endpoints + 10 base tools (2 category + 1 code + 7 skill) = 15 total
    expect(toolSpy).toHaveBeenCalledTimes(15);
    expect(toolSpy).toHaveBeenCalledWith(
      'list-mail-messages',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
    expect(toolSpy).toHaveBeenCalledWith(
      'send-mail',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
    expect(toolSpy).toHaveBeenCalledWith(
      'list-calendar-events',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
    expect(toolSpy).toHaveBeenCalledWith(
      'list-excel-worksheets',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
    expect(toolSpy).toHaveBeenCalledWith(
      'get-current-user',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
  });

  it('should filter tools by regex pattern - mail only', async () => {
    await registerGraphTools(server, graphClient, false, 'mail');

    // 2 mail endpoints + 10 base tools = 12 total
    expect(toolSpy).toHaveBeenCalledTimes(12);
    expect(toolSpy).toHaveBeenCalledWith(
      'list-mail-messages',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
    expect(toolSpy).toHaveBeenCalledWith(
      'send-mail',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
  });

  it('should filter tools by regex pattern - calendar or excel', async () => {
    await registerGraphTools(server, graphClient, false, 'calendar|excel');

    // 2 endpoints (calendar + excel) + 10 base tools = 12 total
    expect(toolSpy).toHaveBeenCalledTimes(12);
    expect(toolSpy).toHaveBeenCalledWith(
      'list-calendar-events',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
    expect(toolSpy).toHaveBeenCalledWith(
      'list-excel-worksheets',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
  });

  it('should handle invalid regex patterns gracefully', async () => {
    await registerGraphTools(server, graphClient, false, '[invalid regex');

    // Falls back to all: 5 endpoints + 10 base tools = 15 total
    expect(toolSpy).toHaveBeenCalledTimes(15);
  });

  it('should combine read-only and filtering correctly', async () => {
    await registerGraphTools(server, graphClient, true, 'mail');

    // 1 mail GET endpoint + 10 base tools = 11 total
    expect(toolSpy).toHaveBeenCalledTimes(11);
    expect(toolSpy).toHaveBeenCalledWith(
      'list-mail-messages',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
  });

  it('should register no tools when pattern matches nothing', async () => {
    await registerGraphTools(server, graphClient, false, 'nonexistent');

    // 0 endpoints match, but 10 base tools still registered
    expect(toolSpy).toHaveBeenCalledTimes(10);
  });
});
