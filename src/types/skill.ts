/**
 * Skill Parameter Definition
 */
export interface SkillParameter {
  type: 'string' | 'number' | 'boolean' | 'object' | 'array';
  description: string;
  required: boolean;
  default?: unknown;
}

/**
 * M365 Skill Definition
 */
export interface M365Skill {
  id: string;
  name: string;
  description: string;
  category: 'mail' | 'calendar' | 'teams' | 'files' | 'sharepoint' | 'planner' | 'todo' | 'general';
  code: string;
  parameters?: Record<string, SkillParameter>;
  returnType?: string;
  author?: string;
  tags?: string[];
  createdAt: string;
  updatedAt: string;
  usageCount: number;
  isPublic: boolean;
  isBuiltin?: boolean;
}

/**
 * Skill Filters
 */
export interface SkillFilters {
  category?: string;
  tags?: string[];
  author?: string;
  isPublic?: boolean;
  isBuiltin?: boolean;
}

/**
 * Skill Execution Result
 */
export interface SkillExecutionResult {
  success: boolean;
  skillName: string;
  skillId: string;
  usageCount: number;
  executionTime: number;
  result?: unknown;
  error?: string;
}
