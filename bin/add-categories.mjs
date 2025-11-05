#!/usr/bin/env node

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, '..');
const endpointsFile = path.join(rootDir, 'src', 'endpoints.json');

console.log('📂 Adding categories to endpoints...');

const endpoints = JSON.parse(fs.readFileSync(endpointsFile, 'utf8'));

// Define category mapping based on tool name patterns
function determineCategory(toolName, pathPattern) {
  // Mail & Messages
  if (toolName.includes('mail') || toolName.includes('message') || toolName.includes('attachment')) {
    return 'mail';
  }

  // Calendar & Events
  if (toolName.includes('calendar') || toolName.includes('event') || toolName.includes('meeting-time')) {
    return 'calendar';
  }

  // Contacts
  if (toolName.includes('contact')) {
    return 'contacts';
  }

  // Teams
  if (toolName.includes('team') && !toolName.includes('channel-message')) {
    return 'teams';
  }

  // Chats
  if (toolName.includes('chat')) {
    return 'chats';
  }

  // Files & Drives (OneDrive/SharePoint files)
  if (toolName.includes('onedrive') || toolName.includes('drive') || toolName.includes('folder') && !toolName.includes('mail')) {
    return 'files';
  }

  // SharePoint Sites & Lists
  if (toolName.includes('sharepoint-site') || toolName.includes('list')) {
    return 'sharepoint';
  }

  // Excel
  if (toolName.includes('excel')) {
    return 'excel';
  }

  // Planner
  if (toolName.includes('planner')) {
    return 'planner';
  }

  // Todo
  if (toolName.includes('todo')) {
    return 'todo';
  }

  // OneNote
  if (toolName.includes('onenote')) {
    return 'onenote';
  }

  // Search
  if (toolName.includes('search')) {
    return 'search';
  }

  // Users
  if (toolName.includes('user') || toolName === 'get-current-user') {
    return 'users';
  }

  return 'other';
}

// Add category to each endpoint
let categoryCounts = {};
for (const endpoint of endpoints) {
  const category = determineCategory(endpoint.toolName, endpoint.pathPattern);
  endpoint.category = category;
  categoryCounts[category] = (categoryCounts[category] || 0) + 1;
}

// Sort endpoints by category for better organization
endpoints.sort((a, b) => {
  if (a.category !== b.category) {
    return a.category.localeCompare(b.category);
  }
  return a.toolName.localeCompare(b.toolName);
});

// Write back to file
fs.writeFileSync(endpointsFile, JSON.stringify(endpoints, null, 2) + '\n');

console.log('✅ Categories added successfully!\n');
console.log('📊 Category distribution:');
Object.entries(categoryCounts)
  .sort((a, b) => b[1] - a[1])
  .forEach(([category, count]) => {
    console.log(`   ${category.padEnd(15)} ${count} tools`);
  });
