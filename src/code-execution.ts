import { Script, createContext } from 'vm';
import GraphClient from './graph-client.js';
import logger from './logger.js';

interface ExecutionOptions {
  timeout?: number; // milliseconds
  maxMemory?: number; // bytes (not enforced by vm, but can monitor)
}

interface M365Client {
  mail: {
    list: (options?: Record<string, unknown>) => Promise<unknown>;
    get: (messageId: string) => Promise<unknown>;
    send: (message: Record<string, unknown>) => Promise<unknown>;
    delete: (messageId: string) => Promise<unknown>;
  };
  calendar: {
    list: (options?: Record<string, unknown>) => Promise<unknown>;
    get: (eventId: string) => Promise<unknown>;
    create: (event: Record<string, unknown>) => Promise<unknown>;
    update: (eventId: string, event: Record<string, unknown>) => Promise<unknown>;
    delete: (eventId: string) => Promise<unknown>;
  };
  teams: {
    list: () => Promise<unknown>;
    getChannels: (teamId: string) => Promise<unknown>;
    getMessages: (teamId: string, channelId: string) => Promise<unknown>;
  };
  files: {
    list: (driveId?: string) => Promise<unknown>;
    get: (driveId: string, itemId: string) => Promise<unknown>;
    upload: (driveId: string, itemId: string, content: string) => Promise<unknown>;
  };
  sharepoint: {
    searchSites: (query: string) => Promise<unknown>;
    getSite: (siteId: string) => Promise<unknown>;
    getLists: (siteId: string) => Promise<unknown>;
    getListItems: (siteId: string, listId: string) => Promise<unknown>;
  };
  planner: {
    listTasks: () => Promise<unknown>;
    getTask: (taskId: string) => Promise<unknown>;
    createTask: (task: Record<string, unknown>) => Promise<unknown>;
    updateTask: (taskId: string, task: Record<string, unknown>, etag?: string) => Promise<unknown>;
  };
  todo: {
    listLists: () => Promise<unknown>;
    listTasks: (listId: string) => Promise<unknown>;
    createTask: (listId: string, task: Record<string, unknown>) => Promise<unknown>;
    updateTask: (listId: string, taskId: string, task: Record<string, unknown>) => Promise<unknown>;
  };
}

/**
 * Creates a sandboxed M365 client for code execution
 */
function createM365Client(graphClient: GraphClient): M365Client {
  return {
    mail: {
      async list(options = {}) {
        const params = new URLSearchParams();
        if (options.filter) params.append('$filter', String(options.filter));
        if (options.select) params.append('$select', String(options.select));
        if (options.top) params.append('$top', String(options.top));
        if (options.skip) params.append('$skip', String(options.skip));
        if (options.orderby) params.append('$orderby', String(options.orderby));

        const path = `/me/messages${params.toString() ? '?' + params.toString() : ''}`;
        const response = await graphClient.graphRequest(path, { method: 'GET' });
        return JSON.parse(response.content[0].text);
      },

      async get(messageId: string) {
        const response = await graphClient.graphRequest(`/me/messages/${messageId}`, {
          method: 'GET',
        });
        return JSON.parse(response.content[0].text);
      },

      async send(message: Record<string, unknown>) {
        const response = await graphClient.graphRequest('/me/sendMail', {
          method: 'POST',
          body: JSON.stringify({ message, saveToSentItems: true }),
        });
        return response.content[0].text ? JSON.parse(response.content[0].text) : { success: true };
      },

      async delete(messageId: string) {
        await graphClient.graphRequest(`/me/messages/${messageId}`, { method: 'DELETE' });
        return { success: true };
      },
    },

    calendar: {
      async list(options = {}) {
        const params = new URLSearchParams();
        if (options.filter) params.append('$filter', String(options.filter));
        if (options.select) params.append('$select', String(options.select));
        if (options.top) params.append('$top', String(options.top));

        const path = `/me/events${params.toString() ? '?' + params.toString() : ''}`;
        const response = await graphClient.graphRequest(path, { method: 'GET' });
        return JSON.parse(response.content[0].text);
      },

      async get(eventId: string) {
        const response = await graphClient.graphRequest(`/me/events/${eventId}`, {
          method: 'GET',
        });
        return JSON.parse(response.content[0].text);
      },

      async create(event: Record<string, unknown>) {
        const response = await graphClient.graphRequest('/me/events', {
          method: 'POST',
          body: JSON.stringify(event),
        });
        return JSON.parse(response.content[0].text);
      },

      async update(eventId: string, event: Record<string, unknown>) {
        const response = await graphClient.graphRequest(`/me/events/${eventId}`, {
          method: 'PATCH',
          body: JSON.stringify(event),
        });
        return JSON.parse(response.content[0].text);
      },

      async delete(eventId: string) {
        await graphClient.graphRequest(`/me/events/${eventId}`, { method: 'DELETE' });
        return { success: true };
      },
    },

    teams: {
      async list() {
        const response = await graphClient.graphRequest('/me/joinedTeams', { method: 'GET' });
        return JSON.parse(response.content[0].text);
      },

      async getChannels(teamId: string) {
        const response = await graphClient.graphRequest(`/teams/${teamId}/channels`, {
          method: 'GET',
        });
        return JSON.parse(response.content[0].text);
      },

      async getMessages(teamId: string, channelId: string) {
        const response = await graphClient.graphRequest(
          `/teams/${teamId}/channels/${channelId}/messages`,
          { method: 'GET' }
        );
        return JSON.parse(response.content[0].text);
      },
    },

    files: {
      async list(driveId = 'me/drive') {
        const response = await graphClient.graphRequest(`/${driveId}/root/children`, {
          method: 'GET',
        });
        return JSON.parse(response.content[0].text);
      },

      async get(driveId: string, itemId: string) {
        const response = await graphClient.graphRequest(
          `/drives/${driveId}/items/${itemId}/content`,
          { method: 'GET' }
        );
        return response.content[0].text;
      },

      async upload(driveId: string, itemId: string, content: string) {
        const response = await graphClient.graphRequest(
          `/drives/${driveId}/items/${itemId}/content`,
          {
            method: 'PUT',
            body: content,
          }
        );
        return response.content[0].text ? JSON.parse(response.content[0].text) : { success: true };
      },
    },

    sharepoint: {
      async searchSites(query: string) {
        const params = new URLSearchParams({ search: query });
        const response = await graphClient.graphRequest(`/sites?${params.toString()}`, {
          method: 'GET',
        });
        return JSON.parse(response.content[0].text);
      },

      async getSite(siteId: string) {
        const response = await graphClient.graphRequest(`/sites/${siteId}`, { method: 'GET' });
        return JSON.parse(response.content[0].text);
      },

      async getLists(siteId: string) {
        const response = await graphClient.graphRequest(`/sites/${siteId}/lists`, {
          method: 'GET',
        });
        return JSON.parse(response.content[0].text);
      },

      async getListItems(siteId: string, listId: string) {
        const response = await graphClient.graphRequest(`/sites/${siteId}/lists/${listId}/items`, {
          method: 'GET',
        });
        return JSON.parse(response.content[0].text);
      },
    },

    planner: {
      async listTasks() {
        const response = await graphClient.graphRequest('/me/planner/tasks', { method: 'GET' });
        return JSON.parse(response.content[0].text);
      },

      async getTask(taskId: string) {
        const response = await graphClient.graphRequest(`/planner/tasks/${taskId}`, {
          method: 'GET',
          includeHeaders: true,
        });
        return JSON.parse(response.content[0].text);
      },

      async createTask(task: Record<string, unknown>) {
        const response = await graphClient.graphRequest('/planner/tasks', {
          method: 'POST',
          body: JSON.stringify(task),
        });
        return JSON.parse(response.content[0].text);
      },

      async updateTask(taskId: string, task: Record<string, unknown>, etag?: string) {
        const headers: Record<string, string> = {};
        if (etag) headers['If-Match'] = etag;

        const response = await graphClient.graphRequest(`/planner/tasks/${taskId}`, {
          method: 'PATCH',
          headers,
          body: JSON.stringify(task),
        });
        return response.content[0].text ? JSON.parse(response.content[0].text) : { success: true };
      },
    },

    todo: {
      async listLists() {
        const response = await graphClient.graphRequest('/me/todo/lists', { method: 'GET' });
        return JSON.parse(response.content[0].text);
      },

      async listTasks(listId: string) {
        const response = await graphClient.graphRequest(`/me/todo/lists/${listId}/tasks`, {
          method: 'GET',
        });
        return JSON.parse(response.content[0].text);
      },

      async createTask(listId: string, task: Record<string, unknown>) {
        const response = await graphClient.graphRequest(`/me/todo/lists/${listId}/tasks`, {
          method: 'POST',
          body: JSON.stringify(task),
        });
        return JSON.parse(response.content[0].text);
      },

      async updateTask(listId: string, taskId: string, task: Record<string, unknown>) {
        const response = await graphClient.graphRequest(
          `/me/todo/lists/${listId}/tasks/${taskId}`,
          {
            method: 'PATCH',
            body: JSON.stringify(task),
          }
        );
        return JSON.parse(response.content[0].text);
      },
    },
  };
}

/**
 * Executes user-provided JavaScript code in a sandboxed environment with M365 client access
 */
export async function executeM365Code(
  code: string,
  graphClient: GraphClient,
  options: ExecutionOptions = {}
): Promise<unknown> {
  const timeout = options.timeout || 30000; // 30 seconds default
  const startTime = Date.now();

  logger.info(`Executing code in sandbox with ${timeout}ms timeout`);

  try {
    // Create M365 client
    const m365 = createM365Client(graphClient);

    // Create sandbox context with limited globals
    const sandbox = {
      m365,
      console: {
        log: (...args: unknown[]) => logger.info('Sandbox console.log:', ...args),
        error: (...args: unknown[]) => logger.error('Sandbox console.error:', ...args),
        warn: (...args: unknown[]) => logger.warn('Sandbox console.warn:', ...args),
      },
      setTimeout: undefined, // Disable setTimeout
      setInterval: undefined, // Disable setInterval
      setImmediate: undefined, // Disable setImmediate
      process: undefined, // Disable process access
      require: undefined, // Disable require
      __dirname: undefined,
      __filename: undefined,
      global: undefined,
      Promise, // Allow Promises
      Array,
      Object,
      String,
      Number,
      Boolean,
      Date,
      Math,
      JSON,
      Set,
      Map,
    };

    const context = createContext(sandbox);

    // Wrap code in async function to support await
    const wrappedCode = `
      (async function() {
        ${code}
      })()
    `;

    const script = new Script(wrappedCode, {
      filename: 'user-code.js',
      timeout,
    });

    // Execute script and return result
    const resultPromise = script.runInContext(context, { timeout }) as Promise<unknown>;

    // Wait for the result with timeout
    const result = await Promise.race([
      resultPromise,
      new Promise((_, reject) => setTimeout(() => reject(new Error('Execution timeout')), timeout)),
    ]);

    const executionTime = Date.now() - startTime;
    logger.info(`Code execution completed in ${executionTime}ms`);

    return result;
  } catch (error) {
    const executionTime = Date.now() - startTime;
    logger.error(`Code execution failed after ${executionTime}ms:`, error);
    throw error;
  }
}
