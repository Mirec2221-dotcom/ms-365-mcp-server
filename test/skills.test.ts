import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { randomUUID } from 'crypto';
import fs from 'fs/promises';
import path from 'path';
import { SkillStorage } from '../src/skills-storage.js';
import { loadBuiltinSkills } from '../src/builtin-skills.js';
import type { M365Skill, SkillFilters } from '../src/types/skill.js';

// Mock logger to avoid console noise during tests
vi.mock('../src/logger.js', () => ({
  default: {
    info: vi.fn(),
    warn: vi.fn(),
    error: vi.fn(),
    debug: vi.fn(),
  },
}));

describe('Skill Persistence', () => {
  const testDataDir = './test-data/skills';
  let storage: SkillStorage;

  beforeEach(async () => {
    // Create fresh storage instance for each test
    storage = new SkillStorage(testDataDir);
    await storage.init();
  });

  afterEach(async () => {
    // Clean up test data directory
    try {
      await fs.rm(testDataDir, { recursive: true, force: true });
    } catch (error) {
      // Directory might not exist, ignore
    }
    vi.clearAllMocks();
  });

  describe('SkillStorage', () => {
    describe('Initialization', () => {
      it('should create storage directory if it does not exist', async () => {
        await storage.init();
        const stats = await fs.stat(testDataDir);
        expect(stats.isDirectory()).toBe(true);
      });

      it('should not fail if storage directory already exists', async () => {
        await storage.init();
        await storage.init(); // Call again
        const stats = await fs.stat(testDataDir);
        expect(stats.isDirectory()).toBe(true);
      });
    });

    describe('Save Operations', () => {
      it('should create a new skill with generated ID', async () => {
        const skillData: Partial<M365Skill> = {
          name: 'testSkill',
          description: 'A test skill',
          category: 'general',
          code: 'return "test";',
        };

        const saved = await storage.save(skillData);

        expect(saved.id).toBeDefined();
        expect(saved.name).toBe('testSkill');
        expect(saved.createdAt).toBeDefined();
        expect(saved.updatedAt).toBeDefined();
        expect(saved.usageCount).toBe(0);
      });

      it('should update an existing skill', async () => {
        const skill = await storage.save({
          name: 'testSkill',
          description: 'Original description',
          category: 'general',
          code: 'return "test";',
        });

        const originalUpdatedAt = skill.updatedAt;

        // Wait a bit to ensure timestamp difference
        await new Promise((resolve) => setTimeout(resolve, 10));

        skill.description = 'Updated description';
        const updated = await storage.save(skill);

        expect(updated.id).toBe(skill.id);
        expect(updated.description).toBe('Updated description');
        expect(updated.updatedAt).not.toBe(originalUpdatedAt);
        expect(updated.createdAt).toBe(skill.createdAt);
      });

      it('should persist skill to file system', async () => {
        const skill = await storage.save({
          name: 'persistTest',
          description: 'Test persistence',
          category: 'mail',
          code: 'return "persisted";',
        });

        const filePath = path.join(testDataDir, `${skill.id}.json`);
        const fileContent = await fs.readFile(filePath, 'utf-8');
        const parsed = JSON.parse(fileContent);

        expect(parsed.id).toBe(skill.id);
        expect(parsed.name).toBe('persistTest');
      });
    });

    describe('Get Operations', () => {
      it('should retrieve skill by ID', async () => {
        const saved = await storage.save({
          name: 'getByIdTest',
          description: 'Test get by ID',
          category: 'calendar',
          code: 'return "found";',
        });

        const retrieved = await storage.get(saved.id);

        expect(retrieved).not.toBeNull();
        expect(retrieved?.id).toBe(saved.id);
        expect(retrieved?.name).toBe('getByIdTest');
      });

      it('should return null for non-existent ID', async () => {
        const nonExistentId = randomUUID();
        const retrieved = await storage.get(nonExistentId);

        expect(retrieved).toBeNull();
      });

      it('should retrieve skill by name', async () => {
        await storage.save({
          name: 'uniqueSkillName',
          description: 'Test get by name',
          category: 'teams',
          code: 'return "found by name";',
        });

        const retrieved = await storage.getByName('uniqueSkillName');

        expect(retrieved).not.toBeNull();
        expect(retrieved?.name).toBe('uniqueSkillName');
      });

      it('should return null for non-existent name', async () => {
        const retrieved = await storage.getByName('nonExistentSkill');
        expect(retrieved).toBeNull();
      });
    });

    describe('List Operations', () => {
      beforeEach(async () => {
        // Create test skills
        await storage.save({
          name: 'mailSkill1',
          description: 'Mail skill 1',
          category: 'mail',
          code: 'return "mail1";',
          tags: ['email', 'urgent'],
          isPublic: true,
        });

        await storage.save({
          name: 'mailSkill2',
          description: 'Mail skill 2',
          category: 'mail',
          code: 'return "mail2";',
          tags: ['email'],
          isPublic: false,
        });

        await storage.save({
          name: 'calendarSkill',
          description: 'Calendar skill',
          category: 'calendar',
          code: 'return "calendar";',
          tags: ['meetings'],
          isPublic: true,
        });
      });

      it('should list all skills', async () => {
        const skills = await storage.list();
        expect(skills.length).toBe(3);
      });

      it('should filter skills by category', async () => {
        const filters: SkillFilters = { category: 'mail' };
        const skills = await storage.list(filters);

        expect(skills.length).toBe(2);
        expect(skills.every((s) => s.category === 'mail')).toBe(true);
      });

      it('should filter skills by tags', async () => {
        const filters: SkillFilters = { tags: ['urgent'] };
        const skills = await storage.list(filters);

        expect(skills.length).toBe(1);
        expect(skills[0].name).toBe('mailSkill1');
      });

      it('should filter skills by public status', async () => {
        const filters: SkillFilters = { isPublic: true };
        const skills = await storage.list(filters);

        expect(skills.length).toBe(2);
        expect(skills.every((s) => s.isPublic === true)).toBe(true);
      });

      it('should filter skills by builtin status', async () => {
        await storage.save({
          name: 'builtinSkill',
          description: 'Built-in skill',
          category: 'general',
          code: 'return "builtin";',
          isBuiltin: true,
        });

        const filters: SkillFilters = { isBuiltin: true };
        const skills = await storage.list(filters);

        expect(skills.length).toBe(1);
        expect(skills[0].isBuiltin).toBe(true);
      });

      it('should combine multiple filters', async () => {
        const filters: SkillFilters = {
          category: 'mail',
          isPublic: true,
        };
        const skills = await storage.list(filters);

        expect(skills.length).toBe(1);
        expect(skills[0].name).toBe('mailSkill1');
      });

      it('should sort skills by usage count descending', async () => {
        const skill1 = await storage.save({
          name: 'lowUsage',
          description: 'Low usage',
          category: 'general',
          code: 'return 1;',
        });

        const skill2 = await storage.save({
          name: 'highUsage',
          description: 'High usage',
          category: 'general',
          code: 'return 2;',
        });

        // Increment usage
        await storage.incrementUsage(skill2.id);
        await storage.incrementUsage(skill2.id);
        await storage.incrementUsage(skill2.id);
        await storage.incrementUsage(skill1.id);

        const skills = await storage.list();
        expect(skills[0].name).toBe('highUsage');
        expect(skills[0].usageCount).toBe(3);
      });
    });

    describe('Delete Operations', () => {
      it('should delete skill by ID', async () => {
        const skill = await storage.save({
          name: 'deleteMe',
          description: 'To be deleted',
          category: 'general',
          code: 'return "delete";',
        });

        const deleted = await storage.delete(skill.id);
        expect(deleted).toBe(true);

        const retrieved = await storage.get(skill.id);
        expect(retrieved).toBeNull();
      });

      it('should return false when deleting non-existent skill', async () => {
        const nonExistentId = randomUUID();
        const deleted = await storage.delete(nonExistentId);
        expect(deleted).toBe(false);
      });

      it('should remove file from file system', async () => {
        const skill = await storage.save({
          name: 'removeFile',
          description: 'File removal test',
          category: 'general',
          code: 'return "remove";',
        });

        const filePath = path.join(testDataDir, `${skill.id}.json`);
        await storage.delete(skill.id);

        await expect(fs.access(filePath)).rejects.toThrow();
      });
    });

    describe('Usage Tracking', () => {
      it('should increment usage count', async () => {
        const skill = await storage.save({
          name: 'usageTest',
          description: 'Usage tracking test',
          category: 'general',
          code: 'return "usage";',
        });

        expect(skill.usageCount).toBe(0);

        await storage.incrementUsage(skill.id);
        const after1 = await storage.get(skill.id);
        expect(after1?.usageCount).toBe(1);

        await storage.incrementUsage(skill.id);
        const after2 = await storage.get(skill.id);
        expect(after2?.usageCount).toBe(2);
      });

      it('should handle incrementing non-existent skill', async () => {
        const nonExistentId = randomUUID();
        // Should not throw
        await expect(storage.incrementUsage(nonExistentId)).resolves.not.toThrow();
      });
    });

    describe('Search Operations', () => {
      beforeEach(async () => {
        await storage.save({
          name: 'emailProcessor',
          description: 'Process urgent emails quickly',
          category: 'mail',
          code: 'return "process";',
          tags: ['email', 'urgent', 'automation'],
        });

        await storage.save({
          name: 'meetingAnalyzer',
          description: 'Analyze meeting patterns',
          category: 'calendar',
          code: 'return "analyze";',
          tags: ['meetings', 'analysis'],
        });

        await storage.save({
          name: 'taskManager',
          description: 'Manage daily tasks efficiently',
          category: 'todo',
          code: 'return "manage";',
          tags: ['tasks', 'productivity'],
        });
      });

      it('should search by name', async () => {
        const results = await storage.search('email');
        expect(results.length).toBe(1);
        expect(results[0].name).toBe('emailProcessor');
      });

      it('should search by description', async () => {
        const results = await storage.search('analyze');
        expect(results.length).toBe(1);
        expect(results[0].name).toBe('meetingAnalyzer');
      });

      it('should search by tags', async () => {
        const results = await storage.search('urgent');
        expect(results.length).toBe(1);
        expect(results[0].name).toBe('emailProcessor');
      });

      it('should perform case-insensitive search', async () => {
        const results = await storage.search('EMAIL');
        expect(results.length).toBe(1);
        expect(results[0].name).toBe('emailProcessor');
      });

      it('should return multiple matches', async () => {
        const results = await storage.search('tasks');
        expect(results.length).toBeGreaterThanOrEqual(1);
      });

      it('should return empty array for no matches', async () => {
        const results = await storage.search('nonexistentquery12345');
        expect(results.length).toBe(0);
      });
    });

    describe('Code Validation', () => {
      it('should reject code with eval', () => {
        const code = 'eval("malicious code")';
        const result = storage.validateCode(code);
        expect(result.valid).toBe(false);
        expect(result.errors.some((e) => e.includes('eval'))).toBe(true);
      });

      it('should reject code with Function constructor', () => {
        const code = 'new Function("return 1")';
        const result = storage.validateCode(code);
        expect(result.valid).toBe(false);
        expect(result.errors.some((e) => e.includes('Function'))).toBe(true);
      });

      it('should reject code with require', () => {
        const code = 'const fs = require("fs")';
        const result = storage.validateCode(code);
        expect(result.valid).toBe(false);
        expect(result.errors.some((e) => e.includes('require'))).toBe(true);
      });

      it('should reject code with import', () => {
        const code = 'import("child_process")';
        const result = storage.validateCode(code);
        expect(result.valid).toBe(false);
        expect(result.errors.some((e) => e.includes('import'))).toBe(true);
      });

      it('should reject code with process.exit', () => {
        const code = 'process.exit(1)';
        const result = storage.validateCode(code);
        expect(result.valid).toBe(false);
        expect(result.errors.some((e) => e.includes('process.exit'))).toBe(true);
      });

      it('should reject code with child_process', () => {
        const code = 'child_process.exec("rm -rf /")';
        const result = storage.validateCode(code);
        expect(result.valid).toBe(false);
        expect(result.errors.some((e) => e.includes('child_process'))).toBe(true);
      });

      it('should accept safe code', () => {
        const code = `
          const messages = await m365.mail.list({ top: 10 });
          return messages.value.filter(m => m.isRead === false);
        `;
        const result = storage.validateCode(code);
        expect(result.valid).toBe(true);
        expect(result.errors.length).toBe(0);
      });

      it('should accept code with await and m365 client', () => {
        const code = `
          const today = new Date();
          const events = await m365.calendar.list({ filter: "start ge '" + today.toISOString() + "'" });
          return events.value.length;
        `;
        const result = storage.validateCode(code);
        expect(result.valid).toBe(true);
      });
    });
  });

  describe('Built-in Skills', () => {
    it('should load built-in skills on initialization', async () => {
      const loaded = await loadBuiltinSkills(storage);
      expect(loaded).toBeGreaterThan(0);

      const skills = await storage.list({ isBuiltin: true });
      expect(skills.length).toBe(loaded);
    });

    it('should not duplicate built-in skills on multiple loads', async () => {
      const firstLoad = await loadBuiltinSkills(storage);
      const secondLoad = await loadBuiltinSkills(storage);

      expect(secondLoad).toBe(0); // No new skills loaded
      expect(firstLoad).toBeGreaterThan(0);
    });

    it('should mark built-in skills with isBuiltin flag', async () => {
      await loadBuiltinSkills(storage);

      const skill = await storage.getByName('summarizeTodaysEmails');
      expect(skill).not.toBeNull();
      expect(skill?.isBuiltin).toBe(true);
    });

    it('should load all expected built-in skills', async () => {
      await loadBuiltinSkills(storage);

      const expectedSkills = [
        'summarizeTodaysEmails',
        'getUnreadUrgentEmails',
        'analyzeTodaysMeetings',
        'getOverdueTasks',
        'getTodaysTodoTasks',
        'dailyProductivitySummary',
      ];

      for (const skillName of expectedSkills) {
        const skill = await storage.getByName(skillName);
        expect(skill).not.toBeNull();
        expect(skill?.name).toBe(skillName);
      }
    });

    it('should categorize built-in skills correctly', async () => {
      await loadBuiltinSkills(storage);

      const mailSkills = await storage.list({ category: 'mail', isBuiltin: true });
      expect(mailSkills.length).toBeGreaterThan(0);

      const calendarSkills = await storage.list({ category: 'calendar', isBuiltin: true });
      expect(calendarSkills.length).toBeGreaterThan(0);

      const plannerSkills = await storage.list({ category: 'planner', isBuiltin: true });
      expect(plannerSkills.length).toBeGreaterThan(0);

      const todoSkills = await storage.list({ category: 'todo', isBuiltin: true });
      expect(todoSkills.length).toBeGreaterThan(0);

      const generalSkills = await storage.list({ category: 'general', isBuiltin: true });
      expect(generalSkills.length).toBeGreaterThan(0);
    });

    it('should tag built-in skills appropriately', async () => {
      await loadBuiltinSkills(storage);

      const urgentSkills = await storage.search('urgent');
      expect(urgentSkills.some((s) => s.isBuiltin)).toBe(true);

      const dailySkills = await storage.search('daily');
      expect(dailySkills.some((s) => s.isBuiltin)).toBe(true);
    });
  });

  describe('Skill Parameters', () => {
    it('should save and retrieve skill with parameters', async () => {
      const skill = await storage.save({
        name: 'parameterizedSkill',
        description: 'Skill with parameters',
        category: 'general',
        code: 'return params.value * 2;',
        parameters: {
          value: {
            type: 'number',
            description: 'Value to double',
            required: true,
          },
          multiplier: {
            type: 'number',
            description: 'Optional multiplier',
            required: false,
            default: 2,
          },
        },
      });

      expect(skill.parameters).toBeDefined();
      expect(skill.parameters?.value.type).toBe('number');
      expect(skill.parameters?.value.required).toBe(true);
      expect(skill.parameters?.multiplier.default).toBe(2);
    });

    it('should support various parameter types', async () => {
      const skill = await storage.save({
        name: 'multiTypeSkill',
        description: 'Multiple parameter types',
        category: 'general',
        code: 'return { string: params.text, number: params.count, boolean: params.flag };',
        parameters: {
          text: { type: 'string', description: 'Text param', required: true },
          count: { type: 'number', description: 'Number param', required: true },
          flag: { type: 'boolean', description: 'Boolean param', required: false, default: false },
          data: { type: 'object', description: 'Object param', required: false },
          items: { type: 'array', description: 'Array param', required: false },
        },
      });

      expect(skill.parameters?.text.type).toBe('string');
      expect(skill.parameters?.count.type).toBe('number');
      expect(skill.parameters?.flag.type).toBe('boolean');
      expect(skill.parameters?.data.type).toBe('object');
      expect(skill.parameters?.items.type).toBe('array');
    });
  });

  describe('Skill Metadata', () => {
    it('should track creation and update timestamps', async () => {
      const skill = await storage.save({
        name: 'timestampTest',
        description: 'Timestamp tracking',
        category: 'general',
        code: 'return Date.now();',
      });

      expect(skill.createdAt).toBeDefined();
      expect(skill.updatedAt).toBeDefined();
      expect(new Date(skill.createdAt).getTime()).toBeLessThanOrEqual(
        new Date(skill.updatedAt).getTime()
      );
    });

    it('should support tags for categorization', async () => {
      const skill = await storage.save({
        name: 'taggedSkill',
        description: 'Skill with tags',
        category: 'mail',
        code: 'return "tagged";',
        tags: ['email', 'urgent', 'automation', 'productivity'],
      });

      expect(skill.tags).toBeDefined();
      expect(skill.tags?.length).toBe(4);
      expect(skill.tags).toContain('urgent');
    });

    it('should support public/private visibility', async () => {
      const publicSkill = await storage.save({
        name: 'publicSkill',
        description: 'Public skill',
        category: 'general',
        code: 'return "public";',
        isPublic: true,
      });

      const privateSkill = await storage.save({
        name: 'privateSkill',
        description: 'Private skill',
        category: 'general',
        code: 'return "private";',
        isPublic: false,
      });

      expect(publicSkill.isPublic).toBe(true);
      expect(privateSkill.isPublic).toBe(false);
    });

    it('should support author metadata', async () => {
      const skill = await storage.save({
        name: 'authoredSkill',
        description: 'Skill with author',
        category: 'general',
        code: 'return "authored";',
        author: 'john.doe@company.com',
      });

      expect(skill.author).toBe('john.doe@company.com');
    });

    it('should support return type documentation', async () => {
      const skill = await storage.save({
        name: 'typedSkill',
        description: 'Skill with return type',
        category: 'general',
        code: 'return { count: 10, items: [] };',
        returnType: '{ count: number, items: Array }',
      });

      expect(skill.returnType).toBeDefined();
    });
  });

  describe('Edge Cases', () => {
    it('should handle skills with no tags', async () => {
      const skill = await storage.save({
        name: 'noTags',
        description: 'No tags',
        category: 'general',
        code: 'return 1;',
      });

      const searched = await storage.search('noTags');
      expect(searched.length).toBe(1);
    });

    it('should handle skills with special characters in name', async () => {
      const skill = await storage.save({
        name: 'skill-with-dashes_and_underscores',
        description: 'Special characters',
        category: 'general',
        code: 'return "special";',
      });

      const retrieved = await storage.getByName('skill-with-dashes_and_underscores');
      expect(retrieved).not.toBeNull();
    });

    it('should handle empty skill list', async () => {
      const skills = await storage.list();
      expect(Array.isArray(skills)).toBe(true);
    });

    it('should handle empty search results', async () => {
      const results = await storage.search('definitelynotfound');
      expect(Array.isArray(results)).toBe(true);
      expect(results.length).toBe(0);
    });

    it('should handle very long skill code', async () => {
      // Create actually long code (not just code that generates long output)
      const longCode = `
        // This is a very long skill with lots of comments and code
        const data = await m365.mail.list({ top: 100 });
        ${'// '.repeat(500)}
        const filtered = data.value.filter(m => m.isRead === false);
        ${'// '.repeat(500)}
        return filtered;
      `;
      const skill = await storage.save({
        name: 'longCode',
        description: 'Very long code',
        category: 'general',
        code: longCode,
      });

      expect(skill.code.length).toBeGreaterThan(1000);
    });

    it('should handle skill with all optional fields', async () => {
      const minimalSkill = await storage.save({
        name: 'minimal',
        description: 'Minimal skill',
        category: 'general',
        code: 'return 42;',
      });

      expect(minimalSkill.id).toBeDefined();
      expect(minimalSkill.createdAt).toBeDefined();
      expect(minimalSkill.updatedAt).toBeDefined();
      expect(minimalSkill.usageCount).toBe(0);
    });
  });
});
