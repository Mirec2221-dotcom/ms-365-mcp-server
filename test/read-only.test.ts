import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { parseArgs } from '../src/cli.js';
import { registerGraphTools } from '../src/graph-tools.js';
import type { GraphClient } from '../src/graph-client.js';

vi.mock('../src/cli.js', () => {
  const parseArgsMock = vi.fn();
  return {
    parseArgs: parseArgsMock,
  };
});

vi.mock('../src/generated/client.js', () => {
  return {
    api: {
      endpoints: [
        {
          alias: 'list-mail-messages',
          method: 'get',
          path: '/me/messages',
          parameters: [],
        },
        {
          alias: 'send-mail',
          method: 'post',
          path: '/me/sendMail',
          parameters: [],
        },
        {
          alias: 'delete-mail-message',
          method: 'delete',
          path: '/me/messages/{message-id}',
          parameters: [],
        },
      ],
    },
  };
});

vi.mock('../src/logger.js', () => {
  return {
    default: {
      info: vi.fn(),
      error: vi.fn(),
    },
  };
});

describe('Read-Only Mode', () => {
  let mockServer: { tool: ReturnType<typeof vi.fn> };

  beforeEach(() => {
    vi.clearAllMocks();

    delete process.env.READ_ONLY;

    mockServer = {
      tool: vi.fn(),
    };
  });

  afterEach(() => {
    vi.resetAllMocks();
  });

  it('should respect --read-only flag from CLI', async () => {
    vi.mocked(parseArgs).mockReturnValue({ readOnly: true } as ReturnType<typeof parseArgs>);

    const options = parseArgs();
    expect(options.readOnly).toBe(true);

    await registerGraphTools(mockServer, {} as GraphClient, options.readOnly);

    // Now includes: 1 GET endpoint + 3 meta tools (categories, code execution) + 7 skill tools = 11 total
    expect(mockServer.tool).toHaveBeenCalledTimes(11);

    const toolCalls = mockServer.tool.mock.calls.map((call: unknown[]) => call[0]);
    expect(toolCalls).toContain('list-mail-messages');
    expect(toolCalls).not.toContain('send-mail');
    expect(toolCalls).not.toContain('delete-mail-message');

    // Verify meta tools are registered
    expect(toolCalls).toContain('list-m365-categories');
    expect(toolCalls).toContain('list-category-tools');
    expect(toolCalls).toContain('execute-m365-code');

    // Verify skill tools are registered
    expect(toolCalls).toContain('create-m365-skill');
    expect(toolCalls).toContain('list-m365-skills');
    expect(toolCalls).toContain('execute-m365-skill');
  });

  it('should register all endpoints when not in read-only mode', async () => {
    vi.mocked(parseArgs).mockReturnValue({ readOnly: false } as ReturnType<typeof parseArgs>);

    const options = parseArgs();
    expect(options.readOnly).toBe(false);

    await registerGraphTools(mockServer, {} as GraphClient, options.readOnly);

    // Now includes: 3 endpoints + 3 meta tools + 7 skill tools = 13 total
    expect(mockServer.tool).toHaveBeenCalledTimes(13);

    const toolCalls = mockServer.tool.mock.calls.map((call: unknown[]) => call[0]);
    expect(toolCalls).toContain('list-mail-messages');
    expect(toolCalls).toContain('send-mail');
    expect(toolCalls).toContain('delete-mail-message');
  });
});
