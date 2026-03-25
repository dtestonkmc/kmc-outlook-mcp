#!/usr/bin/env node
/**
 * One-time setup: creates the managed folder structure in your Outlook mailbox.
 * Run: node setup-folders.js --config /path/to/retention-config.json
 * Safe to re-run — existing folders are skipped (409 treated as success).
 */
require('dotenv').config();
const fs = require('fs');
const path = require('path');
const { getAccessToken } = require('./auth/token-manager');
const { callGraphAPI } = require('./utils/graph-api');

// Well-known folder names that Graph API exposes as top-level aliases
const WELL_KNOWN_FOLDERS = {
  'inbox': 'inbox',
  'drafts': 'drafts',
  'sentitems': 'sentitems',
  'sent items': 'sentitems',
  'deleteditems': 'deleteditems',
  'deleted items': 'deleteditems',
  'junkemail': 'junkemail',
  'junk email': 'junkemail',
  'archive': 'archive',
};

async function resolveWellKnownFolder(token, folderName) {
  const alias = WELL_KNOWN_FOLDERS[folderName.toLowerCase()];
  if (!alias) return null;

  try {
    const result = await callGraphAPI(token, 'GET', `me/mailFolders/${alias}`);
    console.log(`  Resolved well-known folder: ${folderName} (${result.id})`);
    return result.id;
  } catch {
    return null;
  }
}

async function findExistingFolder(token, parentFolderId, folderName) {
  const endpoint = parentFolderId
    ? `me/mailFolders/${parentFolderId}/childFolders`
    : 'me/mailFolders';
  const res = await callGraphAPI(token, 'GET', endpoint, null, { $filter: `displayName eq '${folderName}'` });
  return res.value[0]?.id || null;
}

async function createFolder(token, parentFolderId, folderName) {
  try {
    const endpoint = parentFolderId
      ? `me/mailFolders/${parentFolderId}/childFolders`
      : 'me/mailFolders';

    const result = await callGraphAPI(token, 'POST', endpoint, { displayName: folderName });
    console.log(`  Created: ${folderName}`);
    return result.id;
  } catch (err) {
    if (err.message && err.message.includes('409')) {
      console.log(`  Exists:  ${folderName}`);
      return await findExistingFolder(token, parentFolderId, folderName);
    }
    throw err;
  }
}

async function main() {
  const configArg = process.argv.indexOf('--config');
  const configPath = configArg >= 0 ? process.argv[configArg + 1]
    : path.join(__dirname, 'retention-config.json');

  if (!fs.existsSync(configPath)) {
    console.error(`Config file not found: ${configPath}`);
    console.error('Copy retention-config.example.json to retention-config.json and edit it first.');
    process.exit(1);
  }

  const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
  const token = await getAccessToken();
  if (!token) {
    console.error('No valid token. Run "npm run auth-server" first.');
    process.exit(1);
  }

  // Build unique folder paths from config
  const folderPaths = [...new Set(config.map(entry => entry.folder))];

  console.log(`\nCreating ${folderPaths.length} folder paths...\n`);

  const folderIdCache = {};

  for (const folderPath of folderPaths) {
    // e.g. "Inbox/Notifications/NinjaOne" -> ["Inbox", "Notifications", "NinjaOne"]
    const parts = folderPath.split('/');
    let parentId = null;
    let builtPath = '';

    for (const part of parts) {
      builtPath = builtPath ? `${builtPath}/${part}` : part;
      if (!folderIdCache[builtPath]) {
        // First segment might be a well-known folder (Inbox, Drafts, etc.)
        if (!parentId) {
          const wellKnownId = await resolveWellKnownFolder(token, part);
          if (wellKnownId) {
            folderIdCache[builtPath] = wellKnownId;
            parentId = wellKnownId;
            continue;
          }
        }
        folderIdCache[builtPath] = await createFolder(token, parentId, part);
      }
      parentId = folderIdCache[builtPath];
    }
  }

  console.log('\nFolder setup complete.');
}

main().catch(err => {
  console.error('Setup failed:', err.message);
  process.exit(1);
});
