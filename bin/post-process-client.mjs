#!/usr/bin/env node

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, '..');
const clientFile = path.join(rootDir, 'src', 'generated', 'client.ts');

console.log('🔧 Post-processing generated client code...');

let content = fs.readFileSync(clientFile, 'utf8');

// Fix plannerAssignments to use passthrough instead of strict
// This allows dynamic user IDs as keys in the assignments object
const plannerAssignmentsRegex =
  /(const microsoft_graph_plannerAssignments = z\.record\([^)]+\)[\s\S]*?\.partial\(\)[\s\S]*?)\.strict\(\)/;
if (plannerAssignmentsRegex.test(content)) {
  console.log('  ✓ Fixing microsoft_graph_plannerAssignments to support dynamic user ID keys');
  content = content.replace(plannerAssignmentsRegex, '$1.passthrough()');
}

// Write the modified content back
fs.writeFileSync(clientFile, content, 'utf8');

console.log('✅ Client code post-processing complete');
