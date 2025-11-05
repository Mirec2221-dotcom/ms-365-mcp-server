import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import logger from './logger.js';
import GraphClient from './graph-client.js';
import { api } from './generated/client.js';
import { z } from 'zod';
import { readFileSync } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { executeM365Code } from './code-execution.js';
import { registerSkillTools } from './skill-tools.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

interface EndpointConfig {
  pathPattern: string;
  method: string;
  toolName: string;
  scopes?: string[];
  workScopes?: string[];
  returnDownloadUrl?: boolean;
  category?: string;
}

const endpointsData = JSON.parse(
  readFileSync(path.join(__dirname, 'endpoints.json'), 'utf8')
) as EndpointConfig[];

type TextContent = {
  type: 'text';
  text: string;
  [key: string]: unknown;
};

type ImageContent = {
  type: 'image';
  data: string;
  mimeType: string;
  [key: string]: unknown;
};

type AudioContent = {
  type: 'audio';
  data: string;
  mimeType: string;
  [key: string]: unknown;
};

type ResourceTextContent = {
  type: 'resource';
  resource: {
    text: string;
    uri: string;
    mimeType?: string;
    [key: string]: unknown;
  };
  [key: string]: unknown;
};

type ResourceBlobContent = {
  type: 'resource';
  resource: {
    blob: string;
    uri: string;
    mimeType?: string;
    [key: string]: unknown;
  };
  [key: string]: unknown;
};

type ResourceContent = ResourceTextContent | ResourceBlobContent;

type ContentItem = TextContent | ImageContent | AudioContent | ResourceContent;

interface CallToolResult {
  content: ContentItem[];
  _meta?: Record<string, unknown>;
  isError?: boolean;

  [key: string]: unknown;
}

export async function registerGraphTools(
  server: McpServer,
  graphClient: GraphClient,
  readOnly: boolean = false,
  enabledToolsPattern?: string,
  orgMode: boolean = false
): Promise<void> {
  let enabledToolsRegex: RegExp | undefined;
  if (enabledToolsPattern) {
    try {
      enabledToolsRegex = new RegExp(enabledToolsPattern, 'i');
      logger.info(`Tool filtering enabled with pattern: ${enabledToolsPattern}`);
    } catch {
      logger.error(`Invalid tool filter regex pattern: ${enabledToolsPattern}. Ignoring filter.`);
    }
  }

  for (const tool of api.endpoints) {
    const endpointConfig = endpointsData.find((e) => e.toolName === tool.alias);
    if (!orgMode && endpointConfig && !endpointConfig.scopes && endpointConfig.workScopes) {
      logger.info(`Skipping work account tool ${tool.alias} - not in org mode`);
      continue;
    }

    if (readOnly && tool.method.toUpperCase() !== 'GET') {
      logger.info(`Skipping write operation ${tool.alias} in read-only mode`);
      continue;
    }

    if (enabledToolsRegex && !enabledToolsRegex.test(tool.alias)) {
      logger.info(`Skipping tool ${tool.alias} - doesn't match filter pattern`);
      continue;
    }

    const paramSchema: Record<string, unknown> = {};
    if (tool.parameters && tool.parameters.length > 0) {
      for (const param of tool.parameters) {
        paramSchema[param.name] = param.schema || z.any();
      }
    }

    if (tool.method.toUpperCase() === 'GET' && tool.path.includes('/')) {
      paramSchema['fetchAllPages'] = z
        .boolean()
        .describe('Automatically fetch all pages of results')
        .optional();
    }

    // Add includeHeaders parameter for all tools to capture ETags and other headers
    paramSchema['includeHeaders'] = z
      .boolean()
      .describe('Include response headers (including ETag) in the response metadata')
      .optional();

    // Add excludeResponse parameter to only return success/failure indication
    paramSchema['excludeResponse'] = z
      .boolean()
      .describe('Exclude the full response body and only return success or failure indication')
      .optional();

    // Add If-Match header support for Planner API endpoints (required for PATCH/DELETE)
    if (
      tool.path.includes('/planner/') &&
      ['PATCH', 'DELETE'].includes(tool.method.toUpperCase())
    ) {
      paramSchema['If-Match'] = z
        .string()
        .describe(
          'ETag value from the task object (required for Planner updates). Get this by calling the corresponding GET endpoint with includeHeaders=true and using the _etag value from _meta.'
        )
        .optional();
    }

    // Add preferTextContent parameter for mail/message endpoints (HTML to text conversion)
    if (
      tool.method.toUpperCase() === 'GET' &&
      (tool.path.includes('/messages') || tool.path.includes('/mailFolders'))
    ) {
      paramSchema['preferTextContent'] = z
        .boolean()
        .describe(
          'Request email body in plain text format instead of HTML. Uses Microsoft Graph Prefer header for server-side conversion, with automatic fallback to client-side HTML-to-text conversion for better LLM consumption. Reduces token usage by removing HTML tags.'
        )
        .optional();
    }

    server.tool(
      tool.alias,
      tool.description || `Execute ${tool.method.toUpperCase()} request to ${tool.path}`,
      paramSchema,
      {
        title: tool.alias,
        readOnlyHint: tool.method.toUpperCase() === 'GET',
      },
      async (params) => {
        logger.info(`Tool ${tool.alias} called with params: ${JSON.stringify(params)}`);
        try {
          logger.info(`params: ${JSON.stringify(params)}`);

          const parameterDefinitions = tool.parameters || [];

          let path = tool.path;
          const queryParams: Record<string, string> = {};
          const headers: Record<string, string> = {};
          let body: unknown = null;

          for (let [paramName, paramValue] of Object.entries(params)) {
            // Skip pagination control parameter - it's not part of the Microsoft Graph API - I think 🤷
            if (paramName === 'fetchAllPages') {
              continue;
            }

            // Skip headers control parameter - it's not part of the Microsoft Graph API
            if (paramName === 'includeHeaders') {
              continue;
            }

            // Skip excludeResponse control parameter - it's not part of the Microsoft Graph API
            if (paramName === 'excludeResponse') {
              continue;
            }

            // Handle preferTextContent parameter - add Prefer header for text content
            if (paramName === 'preferTextContent') {
              if (paramValue === true) {
                headers['Prefer'] = 'outlook.body-content-type="text"';
                logger.info('Added Prefer header for text content');
              }
              continue;
            }

            // Handle If-Match header for Planner API
            if (paramName === 'If-Match') {
              headers['If-Match'] = paramValue as string;
              continue;
            }

            // Ok, so, MCP clients (such as claude code) doesn't support $ in parameter names,
            // and others might not support __, so we strip them in hack.ts and restore them here
            const odataParams = [
              'filter',
              'select',
              'expand',
              'orderby',
              'skip',
              'top',
              'count',
              'search',
              'format',
            ];
            const fixedParamName = odataParams.includes(paramName.toLowerCase())
              ? `$${paramName.toLowerCase()}`
              : paramName;
            const paramDef = parameterDefinitions.find((p) => p.name === paramName);

            if (paramDef) {
              switch (paramDef.type) {
                case 'Path':
                  path = path
                    .replace(`{${paramName}}`, encodeURIComponent(paramValue as string))
                    .replace(`:${paramName}`, encodeURIComponent(paramValue as string));
                  break;

                case 'Query':
                  queryParams[fixedParamName] = `${paramValue}`;
                  break;

                case 'Body':
                  if (paramDef.schema) {
                    const parseResult = paramDef.schema.safeParse(paramValue);
                    if (!parseResult.success) {
                      const wrapped = { [paramName]: paramValue };
                      const wrappedResult = paramDef.schema.safeParse(wrapped);
                      if (wrappedResult.success) {
                        logger.info(
                          `Auto-corrected parameter '${paramName}': AI passed nested field directly, wrapped it as {${paramName}: ...}`
                        );
                        body = wrapped;
                      } else {
                        body = paramValue;
                      }
                    } else {
                      body = paramValue;
                    }
                  } else {
                    body = paramValue;
                  }
                  break;

                case 'Header':
                  headers[fixedParamName] = `${paramValue}`;
                  break;
              }
            } else if (paramName === 'body') {
              body = paramValue;
              logger.info(`Set body param: ${JSON.stringify(body)}`);
            }
          }

          if (Object.keys(queryParams).length > 0) {
            const queryString = Object.entries(queryParams)
              .map(([key, value]) => `${encodeURIComponent(key)}=${encodeURIComponent(value)}`)
              .join('&');
            path = `${path}${path.includes('?') ? '&' : '?'}${queryString}`;
          }

          const options: {
            method: string;
            headers: Record<string, string>;
            body?: string;
            rawResponse?: boolean;
            includeHeaders?: boolean;
            excludeResponse?: boolean;
          } = {
            method: tool.method.toUpperCase(),
            headers,
          };

          if (options.method !== 'GET' && body) {
            options.body = typeof body === 'string' ? body : JSON.stringify(body);
          }

          const isProbablyMediaContent =
            tool.errors?.some((error) => error.description === 'Retrieved media content') ||
            path.endsWith('/content');

          if (endpointConfig?.returnDownloadUrl && path.endsWith('/content')) {
            path = path.replace(/\/content$/, '');
            logger.info(
              `Auto-returning download URL for ${tool.alias} (returnDownloadUrl=true in endpoints.json)`
            );
          } else if (isProbablyMediaContent) {
            options.rawResponse = true;
          }

          // Set includeHeaders if requested
          if (params.includeHeaders === true) {
            options.includeHeaders = true;
          }

          // Set excludeResponse if requested
          if (params.excludeResponse === true) {
            options.excludeResponse = true;
          }

          logger.info(`Making graph request to ${path} with options: ${JSON.stringify(options)}`);
          let response = await graphClient.graphRequest(path, options);

          const fetchAllPages = params.fetchAllPages === true;
          if (fetchAllPages && response && response.content && response.content.length > 0) {
            try {
              let combinedResponse = JSON.parse(response.content[0].text);
              let allItems = combinedResponse.value || [];
              let nextLink = combinedResponse['@odata.nextLink'];
              let pageCount = 1;

              while (nextLink) {
                logger.info(`Fetching page ${pageCount + 1} from: ${nextLink}`);

                const url = new URL(nextLink);
                const nextPath = url.pathname.replace('/v1.0', '');
                const nextOptions = { ...options };

                const nextQueryParams: Record<string, string> = {};
                for (const [key, value] of url.searchParams.entries()) {
                  nextQueryParams[key] = value;
                }
                nextOptions.queryParams = nextQueryParams;

                const nextResponse = await graphClient.graphRequest(nextPath, nextOptions);
                if (nextResponse && nextResponse.content && nextResponse.content.length > 0) {
                  const nextJsonResponse = JSON.parse(nextResponse.content[0].text);
                  if (nextJsonResponse.value && Array.isArray(nextJsonResponse.value)) {
                    allItems = allItems.concat(nextJsonResponse.value);
                  }
                  nextLink = nextJsonResponse['@odata.nextLink'];
                  pageCount++;

                  if (pageCount > 100) {
                    logger.warn(`Reached maximum page limit (100) for pagination`);
                    break;
                  }
                } else {
                  break;
                }
              }

              combinedResponse.value = allItems;
              if (combinedResponse['@odata.count']) {
                combinedResponse['@odata.count'] = allItems.length;
              }
              delete combinedResponse['@odata.nextLink'];

              response.content[0].text = JSON.stringify(combinedResponse);

              logger.info(
                `Pagination complete: collected ${allItems.length} items across ${pageCount} pages`
              );
            } catch (e) {
              logger.error(`Error during pagination: ${e}`);
            }
          }

          if (response && response.content && response.content.length > 0) {
            const responseText = response.content[0].text;
            const responseSize = responseText.length;
            logger.info(`Response size: ${responseSize} characters`);

            try {
              const jsonResponse = JSON.parse(responseText);
              if (jsonResponse.value && Array.isArray(jsonResponse.value)) {
                logger.info(`Response contains ${jsonResponse.value.length} items`);
                if (jsonResponse.value.length > 0 && jsonResponse.value[0].body) {
                  logger.info(
                    `First item has body field with size: ${JSON.stringify(jsonResponse.value[0].body).length} characters`
                  );
                }
              }
              if (jsonResponse['@odata.nextLink']) {
                logger.info(`Response has pagination nextLink: ${jsonResponse['@odata.nextLink']}`);
              }
              const preview = responseText.substring(0, 500);
              logger.info(`Response preview: ${preview}${responseText.length > 500 ? '...' : ''}`);
            } catch {
              const preview = responseText.substring(0, 500);
              logger.info(
                `Response preview (non-JSON): ${preview}${responseText.length > 500 ? '...' : ''}`
              );
            }
          }

          // Convert McpResponse to CallToolResult with the correct structure
          const content: ContentItem[] = response.content.map((item) => {
            // GraphClient only returns text content items, so create proper TextContent items
            const textContent: TextContent = {
              type: 'text',
              text: item.text,
            };
            return textContent;
          });

          const result: CallToolResult = {
            content,
            _meta: response._meta,
            isError: response.isError,
          };

          return result;
        } catch (error) {
          logger.error(`Error in tool ${tool.alias}: ${(error as Error).message}`);
          const errorContent: TextContent = {
            type: 'text',
            text: JSON.stringify({
              error: `Error in tool ${tool.alias}: ${(error as Error).message}`,
            }),
          };

          return {
            content: [errorContent],
            isError: true,
          };
        }
      }
    );
  }

  // Register meta-tools for progressive tool discovery
  server.tool(
    'list-m365-categories',
    'Get a list of all available Microsoft 365 tool categories with counts. Use this to discover which categories of tools are available before loading specific tools.',
    {},
    {
      title: 'list-m365-categories',
      readOnlyHint: true,
    },
    async () => {
      const categoryCounts: Record<string, number> = {};
      const categoryDescriptions: Record<string, string> = {
        mail: 'Email and message operations (Outlook)',
        calendar: 'Calendar and event management',
        contacts: 'Contact management',
        teams: 'Microsoft Teams operations',
        chats: 'Teams chat operations',
        files: 'OneDrive and file operations',
        sharepoint: 'SharePoint sites and lists',
        excel: 'Excel workbook operations',
        planner: 'Microsoft Planner task management',
        todo: 'Microsoft To Do task management',
        onenote: 'OneNote notebook operations',
        search: 'Search operations across Microsoft 365',
        users: 'User information and management',
        other: 'Miscellaneous operations',
      };

      for (const endpoint of endpointsData) {
        const category = endpoint.category || 'other';
        categoryCounts[category] = (categoryCounts[category] || 0) + 1;
      }

      const categories = Object.entries(categoryCounts)
        .map(([name, count]) => ({
          name,
          description: categoryDescriptions[name] || 'Other operations',
          toolCount: count,
        }))
        .sort((a, b) => b.toolCount - a.toolCount);

      const result = {
        totalCategories: categories.length,
        totalTools: endpointsData.length,
        categories,
      };

      const content: TextContent = {
        type: 'text',
        text: JSON.stringify(result, null, 2),
      };

      return {
        content: [content],
      };
    }
  );

  server.tool(
    'list-category-tools',
    'Get a list of all tools in a specific Microsoft 365 category. Use this after discovering available categories to see what operations are available in each category.',
    {
      category: z.string().describe('The category name (e.g., "mail", "calendar", "teams")'),
    },
    {
      title: 'list-category-tools',
      readOnlyHint: true,
    },
    async (params) => {
      const { category } = params as { category: string };
      const categoryTools = endpointsData
        .filter((e) => e.category === category)
        .map((e) => {
          const apiTool = api.endpoints.find((t) => t.alias === e.toolName);
          return {
            name: e.toolName,
            description: apiTool?.description || `${e.method.toUpperCase()} ${e.pathPattern}`,
            method: e.method.toUpperCase(),
            readOnly: e.method.toUpperCase() === 'GET',
          };
        });

      if (categoryTools.length === 0) {
        const content: TextContent = {
          type: 'text',
          text: JSON.stringify({
            error: `Category "${category}" not found. Use list-m365-categories to see available categories.`,
          }),
        };
        return {
          content: [content],
          isError: true,
        };
      }

      const result = {
        category,
        toolCount: categoryTools.length,
        tools: categoryTools,
      };

      const content: TextContent = {
        type: 'text',
        text: JSON.stringify(result, null, 2),
      };

      return {
        content: [content],
      };
    }
  );

  // Register code execution tool for advanced data filtering and processing
  server.tool(
    'execute-m365-code',
    'Execute JavaScript code in a sandboxed environment with access to Microsoft 365 APIs. Use this for advanced data filtering, aggregation, and multi-step operations. The code has access to an `m365` object with methods like m365.mail.list(), m365.calendar.list(), etc. This significantly reduces token usage by processing data locally before returning results.',
    {
      code: z
        .string()
        .describe(
          'JavaScript code to execute. The code should return a value. Example: const messages = await m365.mail.list({ filter: "isRead eq false" }); return messages.filter(m => m.importance === "high").map(m => ({ from: m.from.emailAddress.address, subject: m.subject }));'
        ),
      timeout: z
        .number()
        .optional()
        .describe('Execution timeout in milliseconds (default: 30000, max: 60000)'),
    },
    {
      title: 'execute-m365-code',
      readOnlyHint: false,
    },
    async (params) => {
      const { code, timeout = 30000 } = params as { code: string; timeout?: number };

      // Validate timeout
      const maxTimeout = 60000; // 60 seconds max
      const actualTimeout = Math.min(timeout, maxTimeout);

      try {
        logger.info(`Executing M365 code (timeout: ${actualTimeout}ms)`);

        const result = await executeM365Code(code, graphClient, { timeout: actualTimeout });

        const content: TextContent = {
          type: 'text',
          text: JSON.stringify(
            {
              success: true,
              result,
              executedAt: new Date().toISOString(),
            },
            null,
            2
          ),
        };

        return {
          content: [content],
        };
      } catch (error) {
        logger.error(`Code execution error: ${(error as Error).message}`);

        const content: TextContent = {
          type: 'text',
          text: JSON.stringify({
            success: false,
            error: (error as Error).message,
            stack: (error as Error).stack,
          }),
        };

        return {
          content: [content],
          isError: true,
        };
      }
    }
  );

  // Register skill management tools
  await registerSkillTools(server, graphClient);
}
