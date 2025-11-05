import { M365Skill } from './types/skill.js';
import { SkillStorage } from './skills-storage.js';
import logger from './logger.js';

/**
 * Built-in Skills Library
 * These skills are pre-installed and available immediately
 */
const BUILTIN_SKILLS: Partial<M365Skill>[] = [
  // 1. MAIL: Summarize today's emails
  {
    name: 'summarizeTodaysEmails',
    description:
      "Get a comprehensive summary of today's emails including count, unread count, urgent emails, and top senders",
    category: 'mail',
    code: `
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      const todayISO = today.toISOString();

      const tomorrow = new Date(today);
      tomorrow.setDate(tomorrow.getDate() + 1);
      const tomorrowISO = tomorrow.toISOString();

      const messages = await m365.mail.list({
        filter: \`receivedDateTime ge \${todayISO} and receivedDateTime lt \${tomorrowISO}\`,
        select: 'from,subject,importance,isRead,hasAttachments',
        top: 100
      });

      const bySender = {};
      let unreadCount = 0;
      let urgentCount = 0;

      for (const msg of messages.value) {
        const sender = msg.from.emailAddress.address;
        if (!bySender[sender]) {
          bySender[sender] = { count: 0, unread: 0, urgent: 0 };
        }
        bySender[sender].count++;
        if (!msg.isRead) {
          bySender[sender].unread++;
          unreadCount++;
        }
        if (msg.importance === 'high') {
          bySender[sender].urgent++;
          urgentCount++;
        }
      }

      const topSenders = Object.entries(bySender)
        .sort((a, b) => b[1].count - a[1].count)
        .slice(0, 5)
        .map(([email, stats]) => ({ email, ...stats }));

      return {
        date: today.toISOString().split('T')[0],
        total: messages.value.length,
        unread: unreadCount,
        urgent: urgentCount,
        topSenders
      };
    `,
    tags: ['email', 'summary', 'daily', 'report'],
    isPublic: true,
    isBuiltin: true,
  },

  // 2. MAIL: Get unread urgent emails
  {
    name: 'getUnreadUrgentEmails',
    description: 'Get all unread high-priority emails with sender, subject, and received time',
    category: 'mail',
    code: `
      const messages = await m365.mail.list({
        filter: "isRead eq false and importance eq 'high'",
        select: 'from,subject,receivedDateTime,bodyPreview',
        orderby: ['receivedDateTime desc'],
        top: 50
      });

      return messages.value.map(m => ({
        from: m.from.emailAddress.address,
        subject: m.subject,
        preview: m.bodyPreview?.substring(0, 100),
        received: m.receivedDateTime
      }));
    `,
    tags: ['email', 'urgent', 'unread', 'filter'],
    isPublic: true,
    isBuiltin: true,
  },

  // 3. CALENDAR: Analyze today's meetings
  {
    name: 'analyzeTodaysMeetings',
    description:
      "Analyze today's calendar meetings including duration, attendee count, and meeting types",
    category: 'calendar',
    code: `
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      const todayISO = today.toISOString();

      const tomorrow = new Date(today);
      tomorrow.setDate(tomorrow.getDate() + 1);
      const tomorrowISO = tomorrow.toISOString();

      const events = await m365.calendar.list({
        filter: \`start/dateTime ge '\${todayISO}' and start/dateTime lt '\${tomorrowISO}'\`,
        select: 'subject,start,end,attendees,isOnlineMeeting,organizer'
      });

      let totalDuration = 0;
      let onlineMeetings = 0;
      let totalAttendees = 0;
      const meetings = [];

      for (const event of events.value) {
        const start = new Date(event.start.dateTime);
        const end = new Date(event.end.dateTime);
        const duration = (end - start) / (1000 * 60); // minutes

        totalDuration += duration;
        if (event.isOnlineMeeting) onlineMeetings++;
        totalAttendees += (event.attendees?.length || 0);

        meetings.push({
          subject: event.subject,
          start: event.start.dateTime,
          duration: Math.round(duration),
          attendeeCount: event.attendees?.length || 0,
          isOnline: event.isOnlineMeeting
        });
      }

      return {
        date: today.toISOString().split('T')[0],
        totalMeetings: events.value.length,
        totalDuration: Math.round(totalDuration),
        averageDuration: events.value.length > 0 ? Math.round(totalDuration / events.value.length) : 0,
        onlineMeetings,
        totalAttendees,
        meetings: meetings.sort((a, b) => a.start.localeCompare(b.start))
      };
    `,
    tags: ['calendar', 'meetings', 'analysis', 'daily'],
    isPublic: true,
    isBuiltin: true,
  },

  // 4. PLANNER: Get overdue tasks
  {
    name: 'getOverdueTasks',
    description: 'Get all overdue Planner tasks across all plans',
    category: 'planner',
    code: `
      const plans = await m365.planner.listUserPlans();
      const overdueTasks = [];
      const today = new Date().toISOString();

      for (const plan of plans.value) {
        const tasks = await m365.planner.listPlanTasks({ planId: plan.id });

        for (const task of tasks.value) {
          if (task.dueDateTime && task.dueDateTime < today && task.percentComplete < 100) {
            const daysOverdue = Math.floor(
              (new Date() - new Date(task.dueDateTime)) / (1000 * 60 * 60 * 24)
            );

            overdueTasks.push({
              planName: plan.title,
              taskTitle: task.title,
              dueDate: task.dueDateTime.split('T')[0],
              daysOverdue,
              priority: task.priority,
              percentComplete: task.percentComplete
            });
          }
        }
      }

      return overdueTasks.sort((a, b) => b.daysOverdue - a.daysOverdue);
    `,
    tags: ['planner', 'tasks', 'overdue', 'report'],
    isPublic: true,
    isBuiltin: true,
  },

  // 5. TODO: Get today's tasks
  {
    name: 'getTodaysTodoTasks',
    description: 'Get all To Do tasks due today across all lists',
    category: 'todo',
    code: `
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      const todayISO = today.toISOString().split('T')[0];

      const lists = await m365.todo.listTaskLists();
      const todaysTasks = [];

      for (const list of lists.value) {
        const tasks = await m365.todo.listTasks({ listId: list.id });

        for (const task of tasks.value) {
          if (task.dueDateTime?.dateTime) {
            const dueDate = task.dueDateTime.dateTime.split('T')[0];
            if (dueDate === todayISO && task.status !== 'completed') {
              todaysTasks.push({
                listName: list.displayName,
                taskTitle: task.title,
                importance: task.importance,
                isReminderOn: task.isReminderOn,
                status: task.status
              });
            }
          }
        }
      }

      return {
        date: todayISO,
        count: todaysTasks.length,
        tasks: todaysTasks.sort((a, b) => {
          const importanceOrder = { high: 0, normal: 1, low: 2 };
          return importanceOrder[a.importance] - importanceOrder[b.importance];
        })
      };
    `,
    tags: ['todo', 'tasks', 'daily', 'due'],
    isPublic: true,
    isBuiltin: true,
  },

  // 6. GENERAL: Daily productivity summary
  {
    name: 'dailyProductivitySummary',
    description: 'Combined summary of emails, meetings, and tasks for today',
    category: 'general',
    code: `
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      const todayISO = today.toISOString();
      const tomorrow = new Date(today);
      tomorrow.setDate(tomorrow.getDate() + 1);
      const tomorrowISO = tomorrow.toISOString();

      // Get emails
      const emails = await m365.mail.list({
        filter: \`receivedDateTime ge \${todayISO} and receivedDateTime lt \${tomorrowISO}\`,
        select: 'isRead,importance',
        top: 100
      });

      // Get meetings
      const meetings = await m365.calendar.list({
        filter: \`start/dateTime ge '\${todayISO}' and start/dateTime lt '\${tomorrowISO}'\`,
        select: 'start,end'
      });

      // Calculate meeting duration
      let meetingMinutes = 0;
      for (const event of meetings.value) {
        const start = new Date(event.start.dateTime);
        const end = new Date(event.end.dateTime);
        meetingMinutes += (end - start) / (1000 * 60);
      }

      const emailStats = {
        total: emails.value.length,
        unread: emails.value.filter(e => !e.isRead).length,
        urgent: emails.value.filter(e => e.importance === 'high').length
      };

      return {
        date: today.toISOString().split('T')[0],
        emails: emailStats,
        meetings: {
          count: meetings.value.length,
          totalMinutes: Math.round(meetingMinutes)
        },
        summary: \`\${emailStats.total} emails (\${emailStats.unread} unread, \${emailStats.urgent} urgent), \${meetings.value.length} meetings (\${Math.round(meetingMinutes / 60)}h \${Math.round(meetingMinutes % 60)}m)\`
      };
    `,
    tags: ['productivity', 'summary', 'daily', 'report', 'overview'],
    isPublic: true,
    isBuiltin: true,
  },
];

/**
 * Load built-in skills into storage
 * Only creates skills that don't already exist
 */
export async function loadBuiltinSkills(storage: SkillStorage): Promise<number> {
  let loaded = 0;

  for (const skillDef of BUILTIN_SKILLS) {
    try {
      // Check if skill already exists
      if (skillDef.name) {
        const existing = await storage.getByName(skillDef.name);
        if (existing) {
          logger.debug(`Built-in skill '${skillDef.name}' already exists, skipping`);
          continue;
        }
      }

      // Save the skill
      await storage.save(skillDef);
      loaded++;
      logger.info(`Loaded built-in skill: ${skillDef.name}`);
    } catch (error) {
      logger.error(`Failed to load built-in skill '${skillDef.name}':`, error);
    }
  }

  if (loaded > 0) {
    logger.info(`Loaded ${loaded} built-in skills`);
  }

  return loaded;
}

export { BUILTIN_SKILLS };
