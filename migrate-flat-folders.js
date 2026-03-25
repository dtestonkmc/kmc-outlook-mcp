#!/usr/bin/env node
/**
 * One-time migration: moves emails from old flat folders to new nested structure.
 * Run: node migrate-flat-folders.js
 * Safe to re-run — skips folders that are empty or don't exist.
 */
require('dotenv').config({ path: require('path').join(__dirname, '.env') });
const { getAccessToken } = require('./auth/token-manager');
const { callGraphAPI } = require('./utils/graph-api');

// Map of old flat folder name -> new nested path
const MIGRATIONS = {
  'Barracuda': 'Inbox/Security/Barracuda',
  'CDW': 'Inbox/Vendors/CDW',
  'Navisite': 'Inbox/Vendors/Navisite',
  // Empty folders to delete (no emails to move)
  'Alexandria': null,
  'DMARC': null,
  'Eaton': null,
};

// Well-known folder aliases
const WELL_KNOWN = { 'inbox': 'inbox' };

async function resolveWellKnown(token, name) {
  const alias = WELL_KNOWN[name.toLowerCase()];
  if (!alias) return null;
  const r = await callGraphAPI(token, 'GET', `me/mailFolders/${alias}`, null, { $select: 'id' });
  return r.id;
}

async function findChildFolder(token, parentId, name) {
  const r = await callGraphAPI(token, 'GET', `me/mailFolders/${parentId}/childFolders`, null, {
    $filter: `displayName eq '${name}'`, $select: 'id,displayName,totalItemCount'
  });
  return r.value?.[0] || null;
}

async function resolvePath(token, path) {
  const parts = path.split('/');
  let currentId = null;
  for (const part of parts) {
    if (!currentId) {
      currentId = await resolveWellKnown(token, part);
      if (!currentId) {
        console.error(`  Could not resolve "${part}"`);
        return null;
      }
    } else {
      const child = await findChildFolder(token, currentId, part);
      if (!child) {
        console.error(`  Could not find child "${part}"`);
        return null;
      }
      currentId = child.id;
    }
  }
  return currentId;
}

async function moveAllEmails(token, sourceFolderId, destFolderId, folderName) {
  let moved = 0;
  let hasMore = true;

  while (hasMore) {
    const msgs = await callGraphAPI(token, 'GET', `me/mailFolders/${sourceFolderId}/messages`, null, {
      $select: 'id', $top: '50'
    });

    const emails = msgs.value || [];
    if (emails.length === 0) {
      hasMore = false;
      break;
    }

    for (const email of emails) {
      try {
        await callGraphAPI(token, 'POST', `me/messages/${email.id}/move`, { destinationId: destFolderId });
        moved++;
      } catch (err) {
        console.error(`  Failed to move message ${email.id}: ${err.message}`);
      }
    }
    console.log(`  Moved ${moved} emails so far from ${folderName}...`);
  }

  return moved;
}

async function deleteFolder(token, folderId, name) {
  try {
    await callGraphAPI(token, 'DELETE', `me/mailFolders/${folderId}`);
    console.log(`  Deleted empty folder: ${name}`);
  } catch (err) {
    console.error(`  Failed to delete folder ${name}: ${err.message}`);
  }
}

async function main() {
  const token = await getAccessToken();
  if (!token) {
    console.error('No valid token. Run "npm run auth-server" first.');
    process.exit(1);
  }

  const inboxId = await resolveWellKnown(token, 'inbox');
  if (!inboxId) {
    console.error('Could not resolve Inbox');
    process.exit(1);
  }

  // First, ensure Inbox/Vendors/Navisite exists
  console.log('\nEnsuring Inbox/Vendors/Navisite exists...');
  const vendorsId = await resolvePath(token, 'Inbox/Vendors');
  if (vendorsId) {
    const existing = await findChildFolder(token, vendorsId, 'Navisite');
    if (!existing) {
      await callGraphAPI(token, 'POST', `me/mailFolders/${vendorsId}/childFolders`, { displayName: 'Navisite' });
      console.log('  Created Inbox/Vendors/Navisite');
    } else {
      console.log('  Already exists');
    }
  } else {
    console.error('  Could not find Inbox/Vendors — run setup-folders.js first');
    process.exit(1);
  }

  console.log('\nStarting migration...\n');

  for (const [flatName, nestedPath] of Object.entries(MIGRATIONS)) {
    console.log(`Processing: ${flatName}`);

    // Find the flat folder under Inbox
    const flatFolder = await findChildFolder(token, inboxId, flatName);
    if (!flatFolder) {
      console.log(`  Skipped — folder not found under Inbox`);
      continue;
    }

    if (nestedPath === null) {
      // Just delete empty folders
      if (flatFolder.totalItemCount === 0) {
        await deleteFolder(token, flatFolder.id, flatName);
      } else {
        console.log(`  WARNING: ${flatName} has ${flatFolder.totalItemCount} items — skipping delete`);
      }
      continue;
    }

    // Resolve destination
    const destId = await resolvePath(token, nestedPath);
    if (!destId) {
      console.error(`  Could not resolve destination "${nestedPath}" — skipping`);
      continue;
    }

    // Check if source and destination are the same folder
    if (flatFolder.id === destId) {
      console.log(`  Skipped — source and destination are the same folder`);
      continue;
    }

    if (flatFolder.totalItemCount === 0) {
      console.log(`  Empty — deleting flat folder`);
      await deleteFolder(token, flatFolder.id, flatName);
      continue;
    }

    // Move emails
    console.log(`  Moving ${flatFolder.totalItemCount} emails to ${nestedPath}...`);
    const moved = await moveAllEmails(token, flatFolder.id, destId, flatName);
    console.log(`  Done — moved ${moved} emails`);

    // Delete the now-empty flat folder
    await deleteFolder(token, flatFolder.id, flatName);
  }

  // Clean up other empty oddball folders
  for (const name of ['Inbox', 'Junk E-Mail']) {
    const folder = await findChildFolder(token, inboxId, name);
    if (folder && folder.totalItemCount === 0) {
      console.log(`Cleaning up empty folder: ${name}`);
      await deleteFolder(token, folder.id, name);
    }
  }

  console.log('\nMigration complete.');
}

main().catch(err => {
  console.error('Migration failed:', err.message);
  process.exit(1);
});
