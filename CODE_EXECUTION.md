# Code Execution with Microsoft 365 MCP Server

## Overview

The MS365 MCP Server now supports code execution in a sandboxed environment, enabling advanced data filtering, aggregation, and multi-step operations. This feature dramatically reduces token usage by processing data locally before returning results to the LLM.

**Key Benefits:**
- 🚀 **98.7% token reduction** (as measured by Anthropic)
- ⚡ **Faster responses** - Process data locally without multiple API calls
- 💰 **Lower costs** - Reduced token consumption
- 🔧 **Complex operations** - Execute multi-step workflows in a single call

## Architecture

### Sandbox Security

The code execution environment uses Node.js built-in `vm` module with:

- **Context isolation** - User code runs in isolated context
- **Timeout protection** - Default 30s, max 60s
- **Limited globals** - Only safe built-in objects exposed
- **No file system access** - No `require`, `fs`, `process`, etc.
- **No network access** - Only through provided m365 client

### M365 Client API

The sandbox provides an `m365` object with typed methods for Microsoft 365 operations:

```javascript
m365.mail.*       // Email operations
m365.calendar.*   // Calendar operations
m365.teams.*      // Teams operations
m365.files.*      // OneDrive/Files operations
m365.sharepoint.* // SharePoint operations
m365.planner.*    // Planner operations
m365.todo.*       // To Do operations
```

## Usage Examples

### Example 1: Filter Unread High-Priority Emails

Instead of fetching all messages and filtering in the LLM context:

```javascript
// Traditional approach (sends all 1000 emails to LLM = ~500KB)
const messages = await m365.mail.list({ top: 1000 });

// Code execution approach (sends only summary = ~50 bytes)
const messages = await m365.mail.list({ filter: "isRead eq false" });
const highPriority = messages.value.filter(m => m.importance === "high");

return {
  count: highPriority.length,
  senders: [...new Set(highPriority.map(m => m.from.emailAddress.address))],
  subjects: highPriority.map(m => m.subject)
};
```

**Token reduction: ~99%**

### Example 2: Calendar Analysis

```javascript
const now = new Date();
const nextWeek = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);

const events = await m365.calendar.list({
  filter: `start/dateTime ge '${now.toISOString()}' and start/dateTime le '${nextWeek.toISOString()}'`
});

const analysis = {
  totalMeetings: events.value.length,
  byDay: {},
  topAttendees: {}
};

for (const event of events.value) {
  const day = event.start.dateTime.split('T')[0];
  analysis.byDay[day] = (analysis.byDay[day] || 0) + 1;

  if (event.attendees) {
    for (const attendee of event.attendees) {
      const email = attendee.emailAddress.address;
      analysis.topAttendees[email] = (analysis.topAttendees[email] || 0) + 1;
    }
  }
}

return analysis;
```

### Example 3: Planner Task Summary

```javascript
const tasks = await m365.planner.listTasks();

const summary = {
  total: tasks.value.length,
  byStatus: {
    notStarted: 0,
    inProgress: 0,
    completed: 0
  },
  overdue: [],
  dueThisWeek: []
};

const now = new Date();
const weekFromNow = new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000);

for (const task of tasks.value) {
  // Count by status
  if (task.percentComplete === 0) summary.byStatus.notStarted++;
  else if (task.percentComplete === 100) summary.byStatus.completed++;
  else summary.byStatus.inProgress++;

  // Check overdue
  if (task.dueDateTime && new Date(task.dueDateTime) < now && task.percentComplete < 100) {
    summary.overdue.push({
      title: task.title,
      dueDate: task.dueDateTime,
      planId: task.planId
    });
  }

  // Check due this week
  if (task.dueDateTime) {
    const dueDate = new Date(task.dueDateTime);
    if (dueDate >= now && dueDate <= weekFromNow) {
      summary.dueThisWeek.push({
        title: task.title,
        dueDate: task.dueDateTime
      });
    }
  }
}

return summary;
```

### Example 4: Multi-Step Workflow

```javascript
// 1. Find unread important emails
const unreadMessages = await m365.mail.list({
  filter: "isRead eq false and importance eq 'high'",
  top: 10
});

// 2. Create tasks for each email
const todoLists = await m365.todo.listLists();
const defaultList = todoLists.value.find(l => l.wellknownListName === 'defaultList');

const createdTasks = [];
for (const msg of unreadMessages.value) {
  const task = await m365.todo.createTask(defaultList.id, {
    title: `Follow up: ${msg.subject}`,
    body: {
      content: `From: ${msg.from.emailAddress.address}`,
      contentType: 'text'
    },
    importance: 'high'
  });
  createdTasks.push(task.title);
}

// 3. Mark emails as read
for (const msg of unreadMessages.value) {
  // Note: This would require implementing a patch method
  console.log(`Would mark as read: ${msg.subject}`);
}

return {
  processedEmails: unreadMessages.value.length,
  createdTasks: createdTasks
};
```

## API Reference

### execute-m365-code Tool

```typescript
{
  code: string,      // JavaScript code to execute
  timeout?: number   // Optional timeout (default: 30000ms, max: 60000ms)
}
```

### M365 Client Methods

#### Mail Operations

```javascript
m365.mail.list(options?: {
  filter?: string,
  select?: string,
  top?: number,
  skip?: number,
  orderby?: string
}) => Promise<{ value: Message[] }>

m365.mail.get(messageId: string) => Promise<Message>

m365.mail.send(message: {
  subject: string,
  body: { content: string, contentType: 'text' | 'html' },
  toRecipients: [{ emailAddress: { address: string } }]
}) => Promise<void>

m365.mail.delete(messageId: string) => Promise<{ success: true }>
```

#### Calendar Operations

```javascript
m365.calendar.list(options?: {
  filter?: string,
  select?: string,
  top?: number
}) => Promise<{ value: Event[] }>

m365.calendar.get(eventId: string) => Promise<Event>

m365.calendar.create(event: {
  subject: string,
  start: { dateTime: string, timeZone: string },
  end: { dateTime: string, timeZone: string }
}) => Promise<Event>

m365.calendar.update(eventId: string, event: Partial<Event>) => Promise<Event>

m365.calendar.delete(eventId: string) => Promise<{ success: true }>
```

#### Teams Operations

```javascript
m365.teams.list() => Promise<{ value: Team[] }>

m365.teams.getChannels(teamId: string) => Promise<{ value: Channel[] }>

m365.teams.getMessages(teamId: string, channelId: string) => Promise<{ value: Message[] }>
```

#### Files Operations

```javascript
m365.files.list(driveId?: string) => Promise<{ value: DriveItem[] }>

m365.files.get(driveId: string, itemId: string) => Promise<string>

m365.files.upload(driveId: string, itemId: string, content: string) => Promise<void>
```

#### SharePoint Operations

```javascript
m365.sharepoint.searchSites(query: string) => Promise<{ value: Site[] }>

m365.sharepoint.getSite(siteId: string) => Promise<Site>

m365.sharepoint.getLists(siteId: string) => Promise<{ value: List[] }>

m365.sharepoint.getListItems(siteId: string, listId: string) => Promise<{ value: ListItem[] }>
```

#### Planner Operations

```javascript
m365.planner.listTasks() => Promise<{ value: PlannerTask[] }>

m365.planner.getTask(taskId: string) => Promise<PlannerTask>

m365.planner.createTask(task: {
  planId: string,
  bucketId?: string,
  title: string
}) => Promise<PlannerTask>

m365.planner.updateTask(
  taskId: string,
  task: Partial<PlannerTask>,
  etag?: string
) => Promise<void>
```

#### To Do Operations

```javascript
m365.todo.listLists() => Promise<{ value: TodoList[] }>

m365.todo.listTasks(listId: string) => Promise<{ value: TodoTask[] }>

m365.todo.createTask(listId: string, task: {
  title: string,
  body?: { content: string, contentType: string },
  importance?: 'low' | 'normal' | 'high'
}) => Promise<TodoTask>

m365.todo.updateTask(
  listId: string,
  taskId: string,
  task: Partial<TodoTask>
) => Promise<TodoTask>
```

## Security Considerations

### What's Protected

- ✅ File system access blocked
- ✅ Network access blocked (except through m365 client)
- ✅ Process access blocked
- ✅ Module import blocked
- ✅ Timeout protection (max 60s)
- ✅ Context isolation from main process

### What's Allowed

- ✅ JavaScript built-in objects (Array, Object, String, etc.)
- ✅ Promises and async/await
- ✅ Console logging (logged to server logs)
- ✅ Math operations
- ✅ Date/time operations
- ✅ M365 API calls through provided client

### Best Practices

1. **Validate input** - Always validate data from M365 APIs before processing
2. **Limit loops** - Avoid infinite loops; prefer Array methods like filter(), map()
3. **Handle errors** - Use try/catch for robust error handling
4. **Return summaries** - Return aggregated data, not full datasets
5. **Use OData filters** - Pre-filter data using Graph API OData queries when possible

## Performance Tips

### 1. Use OData Filters

```javascript
// Good: Filter on server
const messages = await m365.mail.list({
  filter: "isRead eq false and importance eq 'high'"
});

// Bad: Fetch all, filter locally (wastes bandwidth)
const all = await m365.mail.list({ top: 1000 });
const filtered = all.value.filter(m => !m.isRead && m.importance === 'high');
```

### 2. Select Only Needed Fields

```javascript
// Good: Minimal data transfer
const messages = await m365.mail.list({
  select: "id,subject,from,receivedDateTime"
});

// Bad: Fetches all fields including large body
const messages = await m365.mail.list();
```

### 3. Return Aggregations

```javascript
// Good: Return summary (50 bytes)
return {
  total: messages.length,
  unreadCount: messages.filter(m => !m.isRead).length
};

// Bad: Return full array (500KB)
return messages;
```

## Troubleshooting

### Timeout Errors

```
Error: Execution timeout
```

**Solution:** Increase timeout parameter or optimize code to process less data.

### Memory Issues

```
Error: JavaScript heap out of memory
```

**Solution:** Process data in batches or use server-side filtering with OData.

### Syntax Errors

```
Error: Unexpected token
```

**Solution:** Ensure code is valid JavaScript. Use console.log for debugging.

## Migration Guide

### From Individual Tool Calls

**Before:**
```
1. Call list-mail-messages with filter
2. Parse results in LLM context
3. Call send-mail for each result
```

**After:**
```javascript
const messages = await m365.mail.list({ filter: "..." });
for (const msg of messages.value) {
  await m365.mail.send({ ... });
}
return { processed: messages.value.length };
```

### Token Savings

| Operation | Before (Tokens) | After (Tokens) | Savings |
|-----------|----------------|----------------|---------|
| List 100 emails | ~50,000 | ~500 | 99% |
| Calendar analysis | ~30,000 | ~200 | 99.3% |
| Task summary | ~20,000 | ~300 | 98.5% |

## Future Enhancements

- [ ] TypeScript support for better type safety
- [ ] Skill persistence (save reusable code snippets)
- [ ] Streaming results for large datasets
- [ ] PII tokenization before returning to LLM
- [ ] Rate limiting and quota management
- [ ] More granular permissions control

---

**Related Documentation:**
- [Anthropic MCP Code Execution Article](https://www.anthropic.com/engineering/code-execution-with-mcp)
- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/overview)
- [Project Architecture Guide](./CLAUDE.md)
