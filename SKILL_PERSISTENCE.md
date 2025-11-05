# Skill Persistence - Reusable M365 Automation

## Overview

Skill Persistence allows you to save, manage, and reuse JavaScript functions (called "skills") that perform common M365 operations. Instead of writing the same filtering/aggregation logic repeatedly, you can create a skill once and execute it many times.

**Key Benefits:**

- ✅ **Reusability** - Write once, use many times
- ✅ **Consistency** - Same logic across different sessions
- ✅ **Discoverability** - Browse and search available skills
- ✅ **Versioning** - Update skills without changing calling code
- ✅ **Analytics** - Track usage statistics for popular skills
- ✅ **Sharing** - Optional public skills for team collaboration

## Architecture

```
┌────────────────────────────────────────────────────────────┐
│                    MCP Client (Claude)                      │
│  - create-m365-skill                                        │
│  - list-m365-skills                                         │
│  - execute-m365-skill                                       │
│  - get-m365-skill                                           │
│  - update-m365-skill                                        │
│  - delete-m365-skill                                        │
│  - search-m365-skills                                       │
└────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌────────────────────────────────────────────────────────────┐
│                  Skill Management Layer                     │
│  ├─ skill-tools.ts (MCP tool registration)                 │
│  ├─ skills-storage.ts (File I/O operations)                │
│  └─ builtin-skills.ts (Pre-installed skills)               │
└────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌────────────────────────────────────────────────────────────┐
│                   Code Execution Sandbox                    │
│  ├─ code-execution.ts (VM isolation)                       │
│  └─ M365 Client SDK (mail, calendar, teams, etc.)          │
└────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌────────────────────────────────────────────────────────────┐
│                File System Storage (JSON)                   │
│  data/skills/                                               │
│  ├─ {skill-id-1}.json                                       │
│  ├─ {skill-id-2}.json                                       │
│  └─ ...                                                     │
└────────────────────────────────────────────────────────────┘
```

## Data Model

### Skill Structure

```typescript
interface M365Skill {
  id: string; // UUID
  name: string; // camelCase identifier
  description: string; // What the skill does
  category: string; // mail | calendar | teams | files | etc.
  code: string; // JavaScript function body
  parameters?: {
    // Optional input parameters
    [key: string]: {
      type: 'string' | 'number' | 'boolean' | 'object' | 'array';
      description: string;
      required: boolean;
      default?: any;
    };
  };
  returnType?: string; // Description of return value
  author?: string; // Creator user ID
  tags?: string[]; // Categorization tags
  createdAt: string; // ISO timestamp
  updatedAt: string; // ISO timestamp
  usageCount: number; // Execution counter
  isPublic: boolean; // Shareable flag
  isBuiltin?: boolean; // Pre-installed skill
}
```

### Storage Format

Skills are stored as individual JSON files in `data/skills/`:

```
data/skills/
├─ 3f7a9c8e-4b1d-4e2a-9f6c-1234567890ab.json
├─ 8a2d5e7f-9c3b-4e1a-8f5c-9876543210cd.json
└─ builtin/
   └─ .gitkeep
```

## Available Tools

### 1. create-m365-skill

Create a new reusable skill.

**Parameters:**

- `name` (string, required) - Skill name in camelCase (e.g., `getUnreadUrgentEmails`)
- `description` (string, required) - Clear description of functionality
- `category` (enum, required) - `mail` | `calendar` | `teams` | `files` | `sharepoint` | `planner` | `todo` | `general`
- `code` (string, required) - JavaScript function body
- `parameters` (object, optional) - Parameter definitions
- `tags` (string[], optional) - Tags for search/categorization
- `isPublic` (boolean, optional) - Share with other users (default: false)

**Example:**

```javascript
{
  "name": "getUnreadUrgentEmails",
  "description": "Get all unread high-priority emails",
  "category": "mail",
  "code": `
    const messages = await m365.mail.list({
      filter: "isRead eq false and importance eq 'high'",
      select: 'from,subject,receivedDateTime'
    });

    return messages.value.map(m => ({
      from: m.from.emailAddress.address,
      subject: m.subject,
      received: m.receivedDateTime
    }));
  `,
  "tags": ["email", "urgent", "filter"]
}
```

**Response:**

```json
{
  "success": true,
  "message": "Skill 'getUnreadUrgentEmails' created successfully",
  "skillId": "3f7a9c8e-4b1d-4e2a-9f6c-1234567890ab",
  "skill": {
    "id": "3f7a9c8e-4b1d-4e2a-9f6c-1234567890ab",
    "name": "getUnreadUrgentEmails",
    "description": "Get all unread high-priority emails",
    "category": "mail",
    "tags": ["email", "urgent", "filter"],
    "createdAt": "2025-11-05T10:00:00.000Z"
  }
}
```

### 2. list-m365-skills

List all saved skills with optional filtering.

**Parameters:**

- `category` (string, optional) - Filter by category
- `tags` (string[], optional) - Filter by tags (OR logic)
- `isPublic` (boolean, optional) - Filter public/private
- `isBuiltin` (boolean, optional) - Filter built-in skills

**Example:**

```javascript
{
  "category": "mail",
  "tags": ["urgent"]
}
```

**Response:**

```json
{
  "count": 2,
  "skills": [
    {
      "id": "3f7a9c8e-4b1d-4e2a-9f6c-1234567890ab",
      "name": "getUnreadUrgentEmails",
      "description": "Get all unread high-priority emails",
      "category": "mail",
      "usageCount": 42,
      "tags": ["email", "urgent", "filter"],
      "isPublic": true,
      "isBuiltin": false,
      "createdAt": "2025-11-05T10:00:00.000Z",
      "updatedAt": "2025-11-05T10:00:00.000Z"
    }
  ]
}
```

### 3. get-m365-skill

Get full details of a specific skill including code.

**Parameters:**

- `skillId` (string, required) - Skill ID or name

**Example:**

```javascript
{
  "skillId": "getUnreadUrgentEmails"  // or UUID
}
```

**Response:**

```json
{
  "success": true,
  "skill": {
    "id": "3f7a9c8e-4b1d-4e2a-9f6c-1234567890ab",
    "name": "getUnreadUrgentEmails",
    "description": "Get all unread high-priority emails",
    "category": "mail",
    "code": "const messages = await m365.mail.list(...); ...",
    "usageCount": 42,
    "tags": ["email", "urgent", "filter"],
    "isPublic": true,
    "createdAt": "2025-11-05T10:00:00.000Z",
    "updatedAt": "2025-11-05T10:00:00.000Z"
  }
}
```

### 4. execute-m365-skill

Execute a saved skill with optional parameters.

**Parameters:**

- `skillId` (string, required) - Skill ID or name
- `params` (object, optional) - Parameters to pass to skill
- `timeout` (number, optional) - Execution timeout in ms (default: 30000, max: 60000)

**Example:**

```javascript
{
  "skillId": "getUnreadUrgentEmails"
}
```

**Response:**

```json
{
  "success": true,
  "skillName": "getUnreadUrgentEmails",
  "skillId": "3f7a9c8e-4b1d-4e2a-9f6c-1234567890ab",
  "usageCount": 43,
  "executionTime": 1250,
  "result": [
    {
      "from": "boss@company.com",
      "subject": "Urgent: Q4 Report Needed",
      "received": "2025-11-05T09:30:00Z"
    }
  ]
}
```

### 5. update-m365-skill

Update an existing skill.

**Parameters:**

- `skillId` (string, required) - Skill ID to update
- `updates` (object, required) - Fields to update

**Example:**

```javascript
{
  "skillId": "3f7a9c8e-4b1d-4e2a-9f6c-1234567890ab",
  "updates": {
    "description": "Updated description",
    "tags": ["email", "urgent", "priority", "inbox"]
  }
}
```

### 6. delete-m365-skill

Delete a saved skill (cannot delete built-in skills).

**Parameters:**

- `skillId` (string, required) - Skill ID

**Example:**

```javascript
{
  "skillId": "3f7a9c8e-4b1d-4e2a-9f6c-1234567890ab"
}
```

### 7. search-m365-skills

Search skills by query string (searches name, description, tags).

**Parameters:**

- `query` (string, required) - Search query

**Example:**

```javascript
{
  "query": "urgent email"
}
```

## Built-in Skills

Six pre-installed skills are available immediately:

### 1. **summarizeTodaysEmails**

- **Category**: mail
- **Description**: Comprehensive summary of today's emails
- **Returns**: Total count, unread count, urgent emails, top 5 senders

### 2. **getUnreadUrgentEmails**

- **Category**: mail
- **Description**: All unread high-priority emails
- **Returns**: Array of urgent unread emails with sender/subject/preview

### 3. **analyzeTodaysMeetings**

- **Category**: calendar
- **Description**: Analyze today's meetings
- **Returns**: Meeting count, duration, attendee stats, online/offline split

### 4. **getOverdueTasks**

- **Category**: planner
- **Description**: All overdue Planner tasks across plans
- **Returns**: Array of overdue tasks with days overdue

### 5. **getTodaysTodoTasks**

- **Category**: todo
- **Description**: All To Do tasks due today
- **Returns**: Array of today's tasks sorted by importance

### 6. **dailyProductivitySummary**

- **Category**: general
- **Description**: Combined summary of emails, meetings, tasks
- **Returns**: High-level productivity overview for the day

## Usage Examples

### Example 1: Create and Execute Simple Skill

```javascript
// 1. Create skill
await createSkill({
  name: 'countUnreadEmails',
  description: 'Count unread emails',
  category: 'mail',
  code: `
    const messages = await m365.mail.list({
      filter: 'isRead eq false',
      select: 'id'
    });
    return { count: messages.value.length };
  `,
  tags: ['email', 'count'],
});

// 2. Execute skill
await executeSkill({ skillId: 'countUnreadEmails' });
// → { success: true, result: { count: 15 }, executionTime: 450 }
```

### Example 2: Skill with Parameters

```javascript
// 1. Create skill with parameters
await createSkill({
  name: 'getEmailsByDate',
  description: 'Get emails from specific date range',
  category: 'mail',
  code: `
    const { startDate, endDate } = params;
    const messages = await m365.mail.list({
      filter: \`receivedDateTime ge \${startDate} and receivedDateTime lt \${endDate}\`,
      select: 'from,subject,receivedDateTime'
    });
    return messages.value;
  `,
  parameters: {
    startDate: { type: 'string', description: 'ISO date', required: true },
    endDate: { type: 'string', description: 'ISO date', required: true },
  },
  tags: ['email', 'date', 'filter'],
});

// 2. Execute with parameters
await executeSkill({
  skillId: 'getEmailsByDate',
  params: {
    startDate: '2025-11-01T00:00:00Z',
    endDate: '2025-11-05T23:59:59Z',
  },
});
```

### Example 3: Complex Multi-Service Skill

```javascript
await createSkill({
  name: 'weeklyProductivityReport',
  description: 'Generate weekly productivity report',
  category: 'general',
  code: `
    const weekAgo = new Date();
    weekAgo.setDate(weekAgo.getDate() - 7);
    const weekAgoISO = weekAgo.toISOString();
    const nowISO = new Date().toISOString();

    // Get emails
    const emails = await m365.mail.list({
      filter: \`receivedDateTime ge \${weekAgoISO}\`,
      select: 'isRead,importance',
      top: 500
    });

    // Get meetings
    const meetings = await m365.calendar.list({
      filter: \`start/dateTime ge '\${weekAgoISO}'\`,
      select: 'start,end'
    });

    // Get completed tasks
    const lists = await m365.todo.listTaskLists();
    let completedTasks = 0;
    for (const list of lists.value) {
      const tasks = await m365.todo.listTasks({ listId: list.id });
      completedTasks += tasks.value.filter(t =>
        t.status === 'completed' &&
        t.completedDateTime?.dateTime >= weekAgoISO
      ).length;
    }

    // Calculate meeting time
    let meetingMinutes = 0;
    for (const event of meetings.value) {
      const start = new Date(event.start.dateTime);
      const end = new Date(event.end.dateTime);
      meetingMinutes += (end - start) / (1000 * 60);
    }

    return {
      period: { start: weekAgoISO, end: nowISO },
      emails: {
        total: emails.value.length,
        unread: emails.value.filter(e => !e.isRead).length,
        urgent: emails.value.filter(e => e.importance === 'high').length
      },
      meetings: {
        count: meetings.value.length,
        totalHours: Math.round(meetingMinutes / 60 * 10) / 10
      },
      tasksCompleted: completedTasks,
      summary: \`Processed \${emails.value.length} emails, attended \${meetings.value.length} meetings (\${Math.round(meetingMinutes / 60)}h), completed \${completedTasks} tasks\`
    };
  `,
  tags: ['report', 'weekly', 'productivity', 'summary'],
});
```

### Example 4: Using Built-in Skills

```javascript
// List all built-in skills
await listSkills({ isBuiltin: true });

// Execute built-in skill
await executeSkill({ skillId: 'summarizeTodaysEmails' });
// → { total: 47, unread: 12, urgent: 3, topSenders: [...] }

await executeSkill({ skillId: 'analyzeTodaysMeetings' });
// → { totalMeetings: 5, totalDuration: 240, onlineMeetings: 3, ... }
```

## Security

### Code Validation

All skill code is validated before storage:

**Forbidden Patterns:**

- `eval(`
- `Function(`
- `require(`
- `import(`
- `__dirname`
- `__filename`
- `process.exit`
- `child_process`

Skills with forbidden patterns are rejected during creation/update.

### Sandbox Isolation

Skills execute in the same sandboxed environment as `execute-m365-code`:

- ✅ No file system access
- ✅ No network access (except M365 API via client)
- ✅ No process manipulation
- ✅ Timeout protection (30s default, 60s max)
- ✅ Context isolation using Node.js VM module

### Built-in Skill Protection

Built-in skills cannot be:

- ❌ Deleted
- ✅ Can be viewed
- ✅ Can be executed

## Best Practices

### 1. Naming Conventions

- Use **camelCase** for skill names
- Start with verb: `get`, `list`, `analyze`, `summarize`, `create`
- Be specific: `getUnreadUrgentEmails` not `getEmails`

### 2. Code Structure

```javascript
// ✅ GOOD - Clear, focused, documented
const messages = await m365.mail.list({
  filter: "isRead eq false and importance eq 'high'",
  select: 'from,subject,receivedDateTime', // Only needed fields
});

// Map to minimal output
return messages.value.map((m) => ({
  from: m.from.emailAddress.address,
  subject: m.subject,
  received: m.receivedDateTime,
}));

// ❌ BAD - No filtering, returns everything
const messages = await m365.mail.list({});
return messages;
```

### 3. Error Handling

```javascript
// Skills should handle errors gracefully
try {
  const messages = await m365.mail.list({ filter: 'invalid' });
  return messages;
} catch (error) {
  return { error: error.message, timestamp: new Date().toISOString() };
}
```

### 4. Token Optimization

```javascript
// ✅ Filter and aggregate BEFORE returning
const messages = await m365.mail.list({ top: 1000 });
const summary = {
  total: messages.value.length,
  bySender: /* aggregation */
};
return summary; // Small output

// ❌ Return all data unprocessed
return messages.value; // Huge token cost
```

### 5. Tags and Discovery

```javascript
// ✅ GOOD - Descriptive tags
tags: ['email', 'urgent', 'unread', 'priority', 'inbox'];

// ❌ BAD - Generic tags
tags: ['mail', 'test'];
```

## Performance

### Usage Statistics

Skills track automatic metrics:

- **usageCount**: Incremented on each execution
- **executionTime**: Milliseconds taken for last run
- Sorted by usage when listing (most used first)

### Token Reduction

Skills provide token reduction through:

1. **Reuse** - Code not sent repeatedly
2. **Aggregation** - Filter/process before returning
3. **Caching** - Frequently used logic pre-saved

**Example Token Savings:**

```
Without skill:
- Code execution: 1500 tokens
- Full email data: 35,000 tokens
- Total: 36,500 tokens

With skill:
- Skill execution: 50 tokens
- Aggregated result: 500 tokens
- Total: 550 tokens

Savings: 98.5% per execution after skill creation
```

## Troubleshooting

### Skill Not Found

```
Error: Skill not found: mySkillName
```

**Solution**: Check skill name spelling or list all skills to verify

### Code Validation Failed

```
Error: Code validation failed
Details: ["Forbidden pattern detected: require("]
```

**Solution**: Remove forbidden patterns from code

### Execution Timeout

```
Error: Script execution timed out after 30000ms
```

**Solution**:

- Optimize skill code (reduce API calls)
- Increase timeout: `{ timeout: 60000 }`
- Break into smaller skills

### Cannot Delete Built-in Skill

```
Error: Cannot delete built-in skills
```

**Solution**: Built-in skills are protected. Create custom version instead.

## API Reference

See [CODE_EXECUTION.md](CODE_EXECUTION.md) for details on:

- M365 Client SDK methods (`m365.mail.*`, `m365.calendar.*`, etc.)
- Sandbox environment
- Security restrictions
- Available globals

## Migration from Code Execution

### Before (Direct Code Execution)

```javascript
(await execute) -
  m365 -
  code({
    code: `
    const messages = await m365.mail.list({ filter: "isRead eq false" });
    return messages.value.length;
  `,
  });
```

### After (Using Skills)

```javascript
// One-time: Create skill
(await create) -
  m365 -
  skill({
    name: 'countUnreadEmails',
    category: 'mail',
    code: `
    const messages = await m365.mail.list({ filter: "isRead eq false" });
    return messages.value.length;
  `,
  });

// Every time: Execute by name
(await execute) - m365 - skill({ skillId: 'countUnreadEmails' });
```

**Benefits:**

- ✅ Shorter MCP tool calls
- ✅ Reusable across sessions
- ✅ Discoverable by name
- ✅ Version control
- ✅ Usage analytics

---

**Last Updated:** November 5, 2025
**Repository:** https://github.com/softeria/ms-365-mcp-server
**License:** MIT
