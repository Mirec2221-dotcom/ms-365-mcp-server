import { z } from 'zod';
import { TextContent } from '@modelcontextprotocol/sdk/types.js';
import { SkillStorage } from './skills-storage.js';
import { executeM365Code } from './code-execution.js';
import { loadBuiltinSkills } from './builtin-skills.js';
import GraphClient from './graph-client.js';
import logger from './logger.js';

/**
 * Register skill management tools with MCP server
 */
export async function registerSkillTools(server: any, graphClient: GraphClient): Promise<void> {
  const storage = new SkillStorage();
  await storage.init();

  // Load built-in skills
  await loadBuiltinSkills(storage);

  // 1. CREATE SKILL
  server.tool(
    'create-m365-skill',
    'Create a reusable skill (JavaScript function) for M365 operations. Skills can be saved and reused across sessions for common tasks like filtering emails, analyzing calendar data, etc.',
    {
      name: z
        .string()
        .describe(
          'Skill name in camelCase (e.g., getUnreadUrgentEmails, summarizeTodaysMeetings). Must be unique.'
        ),
      description: z
        .string()
        .describe('Clear description of what the skill does and when to use it'),
      category: z
        .enum(['mail', 'calendar', 'teams', 'files', 'sharepoint', 'planner', 'todo', 'general'])
        .describe('Primary M365 category this skill operates on'),
      code: z
        .string()
        .describe(
          'JavaScript code as async function body. Has access to m365 client and params object. Example: const messages = await m365.mail.list({filter: "isRead eq false"}); return messages.value;'
        ),
      parameters: z
        .record(
          z.object({
            type: z.enum(['string', 'number', 'boolean', 'object', 'array']),
            description: z.string(),
            required: z.boolean(),
            default: z.any().optional(),
          })
        )
        .optional()
        .describe('Optional parameter definitions if skill accepts inputs'),
      tags: z
        .array(z.string())
        .optional()
        .describe('Tags for categorization and search (e.g., urgent, daily, reports)'),
      isPublic: z
        .boolean()
        .optional()
        .describe('Whether skill can be shared with other users (default: false)'),
    },
    {
      title: 'create-m365-skill',
      readOnlyHint: false,
    },
    async (params) => {
      try {
        const { name, description, category, code, parameters, tags, isPublic } = params as {
          name: string;
          description: string;
          category: string;
          code: string;
          parameters?: Record<string, any>;
          tags?: string[];
          isPublic?: boolean;
        };

        // Validate code for security
        const validation = storage.validateCode(code);
        if (!validation.valid) {
          const content: TextContent = {
            type: 'text',
            text: JSON.stringify(
              {
                success: false,
                error: 'Code validation failed',
                details: validation.errors,
              },
              null,
              2
            ),
          };
          return { content: [content], isError: true };
        }

        // Check if name already exists
        const existing = await storage.getByName(name);
        if (existing) {
          const content: TextContent = {
            type: 'text',
            text: JSON.stringify(
              {
                success: false,
                error: `Skill with name '${name}' already exists. Use update-m365-skill to modify it.`,
              },
              null,
              2
            ),
          };
          return { content: [content], isError: true };
        }

        const skill = await storage.save({
          name,
          description,
          category: category as any,
          code,
          parameters,
          tags,
          isPublic: isPublic || false,
        });

        logger.info(`Skill created: ${skill.name} (${skill.id})`);

        const content: TextContent = {
          type: 'text',
          text: JSON.stringify(
            {
              success: true,
              message: `Skill '${skill.name}' created successfully`,
              skillId: skill.id,
              skill: {
                id: skill.id,
                name: skill.name,
                description: skill.description,
                category: skill.category,
                tags: skill.tags,
                createdAt: skill.createdAt,
              },
            },
            null,
            2
          ),
        };

        return { content: [content] };
      } catch (error) {
        logger.error('Error creating skill:', error);
        const content: TextContent = {
          type: 'text',
          text: JSON.stringify(
            {
              success: false,
              error: (error as Error).message,
            },
            null,
            2
          ),
        };
        return { content: [content], isError: true };
      }
    }
  );

  // 2. LIST SKILLS
  server.tool(
    'list-m365-skills',
    'List all saved M365 skills with optional filtering. Shows skill metadata including usage statistics.',
    {
      category: z
        .string()
        .optional()
        .describe(
          'Filter by category (mail, calendar, teams, files, sharepoint, planner, todo, general)'
        ),
      tags: z
        .array(z.string())
        .optional()
        .describe('Filter by tags (shows skills with ANY of these tags)'),
      isPublic: z.boolean().optional().describe('Filter by public/private skills'),
      isBuiltin: z.boolean().optional().describe('Filter by built-in skills'),
    },
    {
      title: 'list-m365-skills',
      readOnlyHint: true,
    },
    async (params) => {
      try {
        const skills = await storage.list(params as any);

        const content: TextContent = {
          type: 'text',
          text: JSON.stringify(
            {
              count: skills.length,
              skills: skills.map((s) => ({
                id: s.id,
                name: s.name,
                description: s.description,
                category: s.category,
                usageCount: s.usageCount,
                tags: s.tags,
                isPublic: s.isPublic,
                isBuiltin: s.isBuiltin,
                createdAt: s.createdAt,
                updatedAt: s.updatedAt,
              })),
            },
            null,
            2
          ),
        };

        return { content: [content] };
      } catch (error) {
        logger.error('Error listing skills:', error);
        const content: TextContent = {
          type: 'text',
          text: JSON.stringify({ success: false, error: (error as Error).message }, null, 2),
        };
        return { content: [content], isError: true };
      }
    }
  );

  // 3. GET SKILL
  server.tool(
    'get-m365-skill',
    'Get details of a specific skill including its full code. Can retrieve by ID or name.',
    {
      skillId: z
        .string()
        .describe(
          'Skill ID (UUID format) or skill name. Will search by ID first, then by name if not found.'
        ),
    },
    {
      title: 'get-m365-skill',
      readOnlyHint: true,
    },
    async (params) => {
      try {
        const { skillId } = params as { skillId: string };

        // Try by ID first
        let skill = await storage.get(skillId);

        // If not found, try by name
        if (!skill) {
          skill = await storage.getByName(skillId);
        }

        if (!skill) {
          const content: TextContent = {
            type: 'text',
            text: JSON.stringify(
              {
                success: false,
                error: `Skill not found: ${skillId}`,
              },
              null,
              2
            ),
          };
          return { content: [content], isError: true };
        }

        const content: TextContent = {
          type: 'text',
          text: JSON.stringify(
            {
              success: true,
              skill,
            },
            null,
            2
          ),
        };

        return { content: [content] };
      } catch (error) {
        logger.error('Error getting skill:', error);
        const content: TextContent = {
          type: 'text',
          text: JSON.stringify({ success: false, error: (error as Error).message }, null, 2),
        };
        return { content: [content], isError: true };
      }
    }
  );

  // 4. EXECUTE SKILL
  server.tool(
    'execute-m365-skill',
    'Execute a saved skill with optional parameters. Automatically tracks usage statistics.',
    {
      skillId: z.string().describe('Skill ID or name to execute'),
      params: z
        .record(z.any())
        .optional()
        .describe('Parameters to pass to the skill (if skill accepts parameters)'),
      timeout: z
        .number()
        .optional()
        .describe('Execution timeout in milliseconds (default: 30000, max: 60000)'),
    },
    {
      title: 'execute-m365-skill',
      readOnlyHint: false,
    },
    async (params) => {
      try {
        const {
          skillId,
          params: skillParams,
          timeout,
        } = params as {
          skillId: string;
          params?: Record<string, any>;
          timeout?: number;
        };

        // Get skill
        let skill = await storage.get(skillId);
        if (!skill) {
          skill = await storage.getByName(skillId);
        }

        if (!skill) {
          const content: TextContent = {
            type: 'text',
            text: JSON.stringify({ success: false, error: `Skill not found: ${skillId}` }, null, 2),
          };
          return { content: [content], isError: true };
        }

        logger.info(`Executing skill: ${skill.name} (${skill.id})`);

        // Wrap code with parameter injection
        const wrappedCode = `
          const params = ${JSON.stringify(skillParams || {})};
          ${skill.code}
        `;

        const startTime = Date.now();
        const result = await executeM365Code(wrappedCode, graphClient, {
          timeout: timeout || 30000,
        });
        const executionTime = Date.now() - startTime;

        // Increment usage counter
        await storage.incrementUsage(skill.id);

        logger.info(`Skill executed successfully: ${skill.name} (${executionTime}ms)`);

        const content: TextContent = {
          type: 'text',
          text: JSON.stringify(
            {
              success: true,
              skillName: skill.name,
              skillId: skill.id,
              usageCount: skill.usageCount + 1,
              executionTime,
              result,
            },
            null,
            2
          ),
        };

        return { content: [content] };
      } catch (error) {
        logger.error('Error executing skill:', error);
        const content: TextContent = {
          type: 'text',
          text: JSON.stringify(
            {
              success: false,
              error: (error as Error).message,
              stack: (error as Error).stack,
            },
            null,
            2
          ),
        };
        return { content: [content], isError: true };
      }
    }
  );

  // 5. UPDATE SKILL
  server.tool(
    'update-m365-skill',
    'Update an existing skill. Can modify any field except ID and usage statistics.',
    {
      skillId: z.string().describe('Skill ID to update'),
      updates: z
        .object({
          name: z.string().optional(),
          description: z.string().optional(),
          code: z.string().optional(),
          category: z
            .enum([
              'mail',
              'calendar',
              'teams',
              'files',
              'sharepoint',
              'planner',
              'todo',
              'general',
            ])
            .optional(),
          tags: z.array(z.string()).optional(),
          parameters: z.record(z.any()).optional(),
          isPublic: z.boolean().optional(),
        })
        .describe('Fields to update'),
    },
    {
      title: 'update-m365-skill',
      readOnlyHint: false,
    },
    async (params) => {
      try {
        const { skillId, updates } = params as { skillId: string; updates: Record<string, any> };

        const skill = await storage.get(skillId);
        if (!skill) {
          const content: TextContent = {
            type: 'text',
            text: JSON.stringify({ success: false, error: `Skill not found: ${skillId}` }, null, 2),
          };
          return { content: [content], isError: true };
        }

        // Validate code if updating
        if (updates.code) {
          const validation = storage.validateCode(updates.code);
          if (!validation.valid) {
            const content: TextContent = {
              type: 'text',
              text: JSON.stringify(
                {
                  success: false,
                  error: 'Code validation failed',
                  details: validation.errors,
                },
                null,
                2
              ),
            };
            return { content: [content], isError: true };
          }
        }

        // Apply updates
        Object.assign(skill, updates);
        const updated = await storage.save(skill);

        logger.info(`Skill updated: ${updated.name} (${updated.id})`);

        const content: TextContent = {
          type: 'text',
          text: JSON.stringify(
            {
              success: true,
              message: `Skill '${updated.name}' updated successfully`,
              skill: {
                id: updated.id,
                name: updated.name,
                description: updated.description,
                category: updated.category,
                updatedAt: updated.updatedAt,
              },
            },
            null,
            2
          ),
        };

        return { content: [content] };
      } catch (error) {
        logger.error('Error updating skill:', error);
        const content: TextContent = {
          type: 'text',
          text: JSON.stringify({ success: false, error: (error as Error).message }, null, 2),
        };
        return { content: [content], isError: true };
      }
    }
  );

  // 6. DELETE SKILL
  server.tool(
    'delete-m365-skill',
    'Delete a saved skill. This action cannot be undone. Built-in skills cannot be deleted.',
    {
      skillId: z.string().describe('Skill ID to delete'),
    },
    {
      title: 'delete-m365-skill',
      readOnlyHint: false,
    },
    async (params) => {
      try {
        const { skillId } = params as { skillId: string };

        // Check if skill is builtin
        const skill = await storage.get(skillId);
        if (skill?.isBuiltin) {
          const content: TextContent = {
            type: 'text',
            text: JSON.stringify(
              {
                success: false,
                error: 'Cannot delete built-in skills',
              },
              null,
              2
            ),
          };
          return { content: [content], isError: true };
        }

        const deleted = await storage.delete(skillId);

        if (!deleted) {
          const content: TextContent = {
            type: 'text',
            text: JSON.stringify({ success: false, error: `Skill not found: ${skillId}` }, null, 2),
          };
          return { content: [content], isError: true };
        }

        logger.info(`Skill deleted: ${skillId}`);

        const content: TextContent = {
          type: 'text',
          text: JSON.stringify(
            {
              success: true,
              message: `Skill deleted successfully`,
              skillId,
            },
            null,
            2
          ),
        };

        return { content: [content] };
      } catch (error) {
        logger.error('Error deleting skill:', error);
        const content: TextContent = {
          type: 'text',
          text: JSON.stringify({ success: false, error: (error as Error).message }, null, 2),
        };
        return { content: [content], isError: true };
      }
    }
  );

  // 7. SEARCH SKILLS
  server.tool(
    'search-m365-skills',
    'Search skills by query string. Searches across name, description, and tags.',
    {
      query: z.string().describe('Search query (searches name, description, tags)'),
    },
    {
      title: 'search-m365-skills',
      readOnlyHint: true,
    },
    async (params) => {
      try {
        const { query } = params as { query: string };
        const skills = await storage.search(query);

        const content: TextContent = {
          type: 'text',
          text: JSON.stringify(
            {
              query,
              count: skills.length,
              skills: skills.map((s) => ({
                id: s.id,
                name: s.name,
                description: s.description,
                category: s.category,
                usageCount: s.usageCount,
                tags: s.tags,
              })),
            },
            null,
            2
          ),
        };

        return { content: [content] };
      } catch (error) {
        logger.error('Error searching skills:', error);
        const content: TextContent = {
          type: 'text',
          text: JSON.stringify({ success: false, error: (error as Error).message }, null, 2),
        };
        return { content: [content], isError: true };
      }
    }
  );

  logger.info('Skill management tools registered (7 tools)');
}
