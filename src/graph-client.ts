import logger from './logger.js';
import AuthManager from './auth.js';
import { refreshAccessToken } from './lib/microsoft-auth.js';
import { convert } from 'html-to-text';

interface GraphRequestOptions {
  headers?: Record<string, string>;
  method?: string;
  body?: string;
  rawResponse?: boolean;
  includeHeaders?: boolean;
  excludeResponse?: boolean;
  accessToken?: string;
  refreshToken?: string;

  [key: string]: unknown;
}

interface ContentItem {
  type: 'text';
  text: string;

  [key: string]: unknown;
}

interface McpResponse {
  content: ContentItem[];
  _meta?: Record<string, unknown>;
  isError?: boolean;

  [key: string]: unknown;
}

class GraphClient {
  private authManager: AuthManager;
  private accessToken: string | null = null;
  private refreshToken: string | null = null;

  constructor(authManager: AuthManager) {
    this.authManager = authManager;
  }

  setOAuthTokens(accessToken: string, refreshToken?: string): void {
    this.accessToken = accessToken;
    this.refreshToken = refreshToken || null;
  }

  async makeRequest(endpoint: string, options: GraphRequestOptions = {}): Promise<unknown> {
    // Use OAuth tokens if available, otherwise fall back to authManager
    let accessToken =
      options.accessToken || this.accessToken || (await this.authManager.getToken());
    let refreshToken = options.refreshToken || this.refreshToken;

    if (!accessToken) {
      throw new Error('No access token available');
    }

    try {
      let response = await this.performRequest(endpoint, accessToken, options);

      if (response.status === 401 && refreshToken) {
        // Token expired, try to refresh
        await this.refreshAccessToken(refreshToken);

        // Update token for retry
        accessToken = this.accessToken || accessToken;
        if (!accessToken) {
          throw new Error('Failed to refresh access token');
        }

        // Retry the request with new token
        response = await this.performRequest(endpoint, accessToken, options);
      }

      if (response.status === 403) {
        const errorText = await response.text();
        if (errorText.includes('scope') || errorText.includes('permission')) {
          throw new Error(
            `Microsoft Graph API scope error: ${response.status} ${response.statusText} - ${errorText}. This tool requires organization mode. Please restart with --org-mode flag.`
          );
        }
        throw new Error(
          `Microsoft Graph API error: ${response.status} ${response.statusText} - ${errorText}`
        );
      }

      if (!response.ok) {
        throw new Error(
          `Microsoft Graph API error: ${response.status} ${response.statusText} - ${await response.text()}`
        );
      }

      const text = await response.text();
      let result: any;

      if (text === '') {
        result = { message: 'OK!' };
      } else {
        try {
          result = JSON.parse(text);
        } catch {
          result = { message: 'OK!', rawResponse: text };
        }
      }

      // If includeHeaders is requested, add response headers to the result
      if (options.includeHeaders) {
        const etag = response.headers.get('ETag') || response.headers.get('etag');

        // Simple approach: just add ETag to the result if it's an object
        if (result && typeof result === 'object' && !Array.isArray(result)) {
          return {
            ...result,
            _etag: etag || 'no-etag-found',
          };
        }
      }

      return result;
    } catch (error) {
      logger.error('Microsoft Graph API request failed:', error);
      throw error;
    }
  }

  private async refreshAccessToken(refreshToken: string): Promise<void> {
    const tenantId = process.env.MS365_MCP_TENANT_ID || 'common';
    const clientId = process.env.MS365_MCP_CLIENT_ID || '084a3e9f-a9f4-43f7-89f9-d229cf97853e';
    const clientSecret = process.env.MS365_MCP_CLIENT_SECRET;

    if (!clientSecret) {
      throw new Error('MS365_MCP_CLIENT_SECRET not configured');
    }

    const response = await refreshAccessToken(refreshToken, clientId, clientSecret, tenantId);
    this.accessToken = response.access_token;
    if (response.refresh_token) {
      this.refreshToken = response.refresh_token;
    }
  }

  private async performRequest(
    endpoint: string,
    accessToken: string,
    options: GraphRequestOptions
  ): Promise<Response> {
    const url = `https://graph.microsoft.com/v1.0${endpoint}`;

    const headers: Record<string, string> = {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      ...options.headers,
    };

    return fetch(url, {
      method: options.method || 'GET',
      headers,
      body: options.body,
    });
  }

  async graphRequest(endpoint: string, options: GraphRequestOptions = {}): Promise<McpResponse> {
    try {
      logger.info(`Calling ${endpoint} with options: ${JSON.stringify(options)}`);

      // Use new OAuth-aware request method
      const result = await this.makeRequest(endpoint, options);

      return this.formatJsonResponse(result, options.rawResponse, options.excludeResponse);
    } catch (error) {
      logger.error(`Error in Graph API request: ${error}`);
      return {
        content: [{ type: 'text', text: JSON.stringify({ error: (error as Error).message }) }],
        isError: true,
      };
    }
  }

  /**
   * Converts HTML content to plain text optimized for LLM consumption
   * @param html HTML string to convert
   * @returns Plain text string
   */
  private convertHtmlToText(html: string): string {
    try {
      return convert(html, {
        wordwrap: 80,
        preserveNewlines: true,
        selectors: [
          // Links: show only text, hide href
          { selector: 'a', options: { ignoreHref: true } },
          // Images: skip completely
          { selector: 'img', format: 'skip' },
          // Tables: preserve structure
          { selector: 'table', options: { uppercaseHeadings: false } },
          // Skip style and script tags
          { selector: 'style', format: 'skip' },
          { selector: 'script', format: 'skip' },
        ],
      });
    } catch (error) {
      logger.warn(`Failed to convert HTML to text: ${error}`);
      // Fallback: return original HTML if conversion fails
      return html;
    }
  }

  /**
   * Processes message body objects and converts HTML to text if needed
   * @param obj Any object that might contain message body
   */
  private processMessageBodies(obj: any): void {
    if (!obj || typeof obj !== 'object') {
      return;
    }

    // Handle arrays (like list of messages)
    if (Array.isArray(obj)) {
      obj.forEach((item) => this.processMessageBodies(item));
      return;
    }

    // Check if this object has a body property with contentType
    if (obj.body && typeof obj.body === 'object') {
      if (obj.body.contentType === 'html' && obj.body.content) {
        logger.info('Converting HTML email body to text');
        obj.body.content = this.convertHtmlToText(obj.body.content);
        obj.body.contentType = 'text';
      }
    }

    // Check for uniqueBody (used in message threads)
    if (obj.uniqueBody && typeof obj.uniqueBody === 'object') {
      if (obj.uniqueBody.contentType === 'html' && obj.uniqueBody.content) {
        logger.info('Converting HTML uniqueBody to text');
        obj.uniqueBody.content = this.convertHtmlToText(obj.uniqueBody.content);
        obj.uniqueBody.contentType = 'text';
      }
    }

    // Recursively process nested objects
    Object.keys(obj).forEach((key) => {
      if (typeof obj[key] === 'object') {
        this.processMessageBodies(obj[key]);
      }
    });
  }

  formatJsonResponse(data: unknown, rawResponse = false, excludeResponse = false): McpResponse {
    // If excludeResponse is true, only return success indication
    if (excludeResponse) {
      return {
        content: [{ type: 'text', text: JSON.stringify({ success: true }) }],
      };
    }

    // Handle the case where data includes headers metadata
    if (data && typeof data === 'object' && '_headers' in data) {
      const responseData = data as {
        data: unknown;
        _headers: Record<string, string>;
        _etag?: string;
      };

      const meta: Record<string, unknown> = {};
      if (responseData._etag) {
        meta.etag = responseData._etag;
      }
      if (responseData._headers) {
        meta.headers = responseData._headers;
      }

      if (rawResponse) {
        return {
          content: [{ type: 'text', text: JSON.stringify(responseData.data) }],
          _meta: meta,
        };
      }

      if (responseData.data === null || responseData.data === undefined) {
        return {
          content: [{ type: 'text', text: JSON.stringify({ success: true }) }],
          _meta: meta,
        };
      }

      // Process message bodies to convert HTML to text
      this.processMessageBodies(responseData.data);

      // Remove OData properties
      const removeODataProps = (obj: Record<string, unknown>): void => {
        if (typeof obj === 'object' && obj !== null) {
          Object.keys(obj).forEach((key) => {
            if (key.startsWith('@odata.')) {
              delete obj[key];
            } else if (typeof obj[key] === 'object') {
              removeODataProps(obj[key] as Record<string, unknown>);
            }
          });
        }
      };

      removeODataProps(responseData.data as Record<string, unknown>);

      return {
        content: [{ type: 'text', text: JSON.stringify(responseData.data, null, 2) }],
        _meta: meta,
      };
    }

    // Original handling for backward compatibility
    if (rawResponse) {
      return {
        content: [{ type: 'text', text: JSON.stringify(data) }],
      };
    }

    if (data === null || data === undefined) {
      return {
        content: [{ type: 'text', text: JSON.stringify({ success: true }) }],
      };
    }

    // Process message bodies to convert HTML to text
    this.processMessageBodies(data);

    // Remove OData properties
    const removeODataProps = (obj: Record<string, unknown>): void => {
      if (typeof obj === 'object' && obj !== null) {
        Object.keys(obj).forEach((key) => {
          if (key.startsWith('@odata.')) {
            delete obj[key];
          } else if (typeof obj[key] === 'object') {
            removeODataProps(obj[key] as Record<string, unknown>);
          }
        });
      }
    };

    removeODataProps(data as Record<string, unknown>);

    return {
      content: [{ type: 'text', text: JSON.stringify(data, null, 2) }],
    };
  }
}

export default GraphClient;
