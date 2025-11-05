import fs from 'fs/promises';
import path from 'path';
import { randomUUID } from 'crypto';
import { M365Skill, SkillFilters } from './types/skill.js';
import logger from './logger.js';

/**
 * SkillStorage - File-based storage for M365 Skills
 */
export class SkillStorage {
  private skillsDir: string;
  private initialized = false;

  constructor(dataDir = './data/skills') {
    this.skillsDir = dataDir;
  }

  /**
   * Initialize storage directory
   */
  async init(): Promise<void> {
    if (this.initialized) return;

    try {
      await fs.mkdir(this.skillsDir, { recursive: true });
      this.initialized = true;
      logger.info(`Skill storage initialized at: ${this.skillsDir}`);
    } catch (error) {
      logger.error('Failed to initialize skill storage:', error);
      throw error;
    }
  }

  /**
   * Save a skill (create or update)
   */
  async save(skill: Partial<M365Skill>): Promise<M365Skill> {
    await this.init();

    // Generate ID if new skill
    if (!skill.id) {
      skill.id = randomUUID();
      skill.createdAt = new Date().toISOString();
      skill.usageCount = 0;
      // Explicitly set isBuiltin to false for custom skills if not already set
      if (skill.isBuiltin === undefined) {
        skill.isBuiltin = false;
      }
    }

    // Update timestamp
    skill.updatedAt = new Date().toISOString();

    // Validate required fields
    if (!skill.name || !skill.description || !skill.code || !skill.category) {
      throw new Error('Missing required skill fields: name, description, code, category');
    }

    const fullSkill = skill as M365Skill;
    const filePath = path.join(this.skillsDir, `${fullSkill.id}.json`);

    try {
      await fs.writeFile(filePath, JSON.stringify(fullSkill, null, 2), 'utf-8');
      logger.info(`Skill saved: ${fullSkill.name} (${fullSkill.id})`);
      return fullSkill;
    } catch (error) {
      logger.error(`Failed to save skill ${fullSkill.id}:`, error);
      throw error;
    }
  }

  /**
   * Get a skill by ID
   */
  async get(id: string): Promise<M365Skill | null> {
    await this.init();

    try {
      const filePath = path.join(this.skillsDir, `${id}.json`);
      const content = await fs.readFile(filePath, 'utf-8');
      const skill = JSON.parse(content) as M365Skill;
      return skill;
    } catch (error) {
      if (error && typeof error === 'object' && 'code' in error && error.code === 'ENOENT') {
        return null;
      }
      logger.error(`Failed to get skill ${id}:`, error);
      throw error;
    }
  }

  /**
   * Get a skill by name
   */
  async getByName(name: string): Promise<M365Skill | null> {
    const skills = await this.list();
    return skills.find((s) => s.name === name) || null;
  }

  /**
   * List all skills with optional filtering
   */
  async list(filters?: SkillFilters): Promise<M365Skill[]> {
    await this.init();

    try {
      const files = await fs.readdir(this.skillsDir);
      const skills: M365Skill[] = [];

      for (const file of files) {
        if (!file.endsWith('.json')) continue;

        try {
          const content = await fs.readFile(path.join(this.skillsDir, file), 'utf-8');
          const skill = JSON.parse(content) as M365Skill;

          // Apply filters
          if (filters?.category && skill.category !== filters.category) continue;
          if (filters?.author && skill.author !== filters.author) continue;
          if (filters?.isPublic !== undefined && skill.isPublic !== filters.isPublic) continue;
          // Treat undefined isBuiltin as false for custom skills
          if (filters?.isBuiltin !== undefined) {
            const skillIsBuiltin = skill.isBuiltin === true;
            if (skillIsBuiltin !== filters.isBuiltin) continue;
          }
          if (filters?.tags && !filters.tags.some((tag) => skill.tags?.includes(tag))) continue;

          skills.push(skill);
        } catch (error) {
          logger.warn(`Failed to parse skill file ${file}:`, error);
        }
      }

      // Sort by usage count (most used first)
      return skills.sort((a, b) => b.usageCount - a.usageCount);
    } catch (error) {
      logger.error('Failed to list skills:', error);
      throw error;
    }
  }

  /**
   * Delete a skill
   */
  async delete(id: string): Promise<boolean> {
    await this.init();

    try {
      const filePath = path.join(this.skillsDir, `${id}.json`);
      await fs.unlink(filePath);
      logger.info(`Skill deleted: ${id}`);
      return true;
    } catch (error) {
      if (error && typeof error === 'object' && 'code' in error && error.code === 'ENOENT') {
        return false;
      }
      logger.error(`Failed to delete skill ${id}:`, error);
      throw error;
    }
  }

  /**
   * Increment usage counter
   */
  async incrementUsage(id: string): Promise<void> {
    const skill = await this.get(id);
    if (skill) {
      skill.usageCount++;
      await this.save(skill);
    }
  }

  /**
   * Search skills by query (searches name, description, tags)
   */
  async search(query: string): Promise<M365Skill[]> {
    const skills = await this.list();
    const lowerQuery = query.toLowerCase();

    return skills.filter(
      (skill) =>
        skill.name.toLowerCase().includes(lowerQuery) ||
        skill.description.toLowerCase().includes(lowerQuery) ||
        skill.tags?.some((tag) => tag.toLowerCase().includes(lowerQuery))
    );
  }

  /**
   * Validate skill code for security
   */
  validateCode(code: string): { valid: boolean; errors: string[] } {
    const errors: string[] = [];
    const forbidden = [
      'eval(',
      'Function(',
      'require(',
      'import(',
      '__dirname',
      '__filename',
      'process.exit',
      'child_process',
    ];

    for (const pattern of forbidden) {
      if (code.includes(pattern)) {
        errors.push(`Forbidden pattern detected: ${pattern}`);
      }
    }

    return {
      valid: errors.length === 0,
      errors,
    };
  }
}

// Export singleton instance
export const skillStorage = new SkillStorage();
