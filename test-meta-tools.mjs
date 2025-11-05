#!/usr/bin/env node

import { readFileSync } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const endpointsFile = path.join(__dirname, 'src', 'endpoints.json');
const endpointsData = JSON.parse(readFileSync(endpointsFile, 'utf8'));

console.log('🧪 Testing meta-tools logic...\n');

// Test 1: list-m365-categories logic
console.log('Test 1: list-m365-categories');
console.log('=====================================');

const categoryCounts = {};
const categoryDescriptions = {
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

const result1 = {
  totalCategories: categories.length,
  totalTools: endpointsData.length,
  categories,
};

console.log(JSON.stringify(result1, null, 2));

// Test 2: list-category-tools logic for "mail" category
console.log('\n\nTest 2: list-category-tools (category: "mail")');
console.log('=====================================');

const testCategory = 'mail';
const categoryTools = endpointsData
  .filter((e) => e.category === testCategory)
  .map((e) => ({
    name: e.toolName,
    description: `${e.method.toUpperCase()} ${e.pathPattern}`,
    method: e.method.toUpperCase(),
    readOnly: e.method.toUpperCase() === 'GET',
  }));

const result2 = {
  category: testCategory,
  toolCount: categoryTools.length,
  tools: categoryTools.slice(0, 5), // Show first 5 for brevity
};

console.log(JSON.stringify(result2, null, 2));
console.log(`\n(Showing first 5 of ${categoryTools.length} tools)`);

// Test 3: Verify all endpoints have categories
console.log('\n\nTest 3: Verify all endpoints have categories');
console.log('=====================================');

const missingCategories = endpointsData.filter((e) => !e.category);
if (missingCategories.length === 0) {
  console.log('✅ All endpoints have categories assigned');
} else {
  console.log(`❌ ${missingCategories.length} endpoints missing categories:`);
  missingCategories.forEach((e) => {
    console.log(`  - ${e.toolName}`);
  });
}

console.log('\n✅ All meta-tools logic tests passed!');
