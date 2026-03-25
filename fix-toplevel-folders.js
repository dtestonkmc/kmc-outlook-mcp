#!/usr/bin/env node
/**
 * Fix: moves emails from top-level duplicate folders (Notifications, Security,
 * Vendors, Invoices) into the correct Inbox subfolders, then deletes the
 * top-level duplicates.
 */
require('dotenv').config({ path: require('path').join(__dirname, '.env') });
const { getAccessToken } = require('./auth/token-manager');
const { callGraphAPI } = require('./utils/graph-api');

async function resolveWellKnown(token, name) {
  const aliases = { 'inbox': 'inbox', 'deleteditems': 'deleteditems' };
  const alias = aliases[name.toLowerCase()];
  if (!alias) return null;
  const r = await callGraphAPI(token, 'GET', `me/mailFolders/${alias}`, null, { $select: 'id' });
  return r.id;
}

async function findChildFolder(token, parentId, name) {
  const r = await callGraphAPI(token, 'GET', `me/mailFolders/${parentId}/childFolders`, null, {
    $filter: `displayName eq '${name}'`, $select: 'id,displayName,totalItemCount,childFolderCount'
  });
  return r.value?.[0] || null;
}

async function findTopLevelFolder(token, name) {
  const r = await callGraphAPI(token, 'GET', 'me/mailFolders', null, {
    $filter: `displayName eq '${name}'`, $select: 'id,displayName,totalItemCount,childFolderCount,parentFolderId'
  });
  return r.value?.[0] || null;
}

async function moveAllMessages(token, srcFolderId, destFolderId, label) {
  let moved = 0;
  while (true) {
    const msgs = await callGraphAPI(token, 'GET', `me/mailFolders/${srcFolderId}/messages`, null, {
      $select: 'id', $top: '50'
    });
    if (!msgs.value || msgs.value.length === 0) break;
    for (const msg of msgs.value) {
      try {
        await callGraphAPI(token, 'POST', `me/messages/${msg.id}/move`, { destinationId: destFolderId });
        moved++;
      } catch (err) {
        console.error(`  Failed to move message: ${err.message}`);
      }
    }
    console.log(`  ${label}: moved ${moved} so far...`);
  }
  return moved;
}

async function safeDeleteFolder(token, folderId, name) {
  const info = await callGraphAPI(token, 'GET', `me/mailFolders/${folderId}`, null, {
    $select: 'totalItemCount,childFolderCount'
  });
  if (info.totalItemCount > 0) {
    console.log(`  SKIP delete "${name}" — still has ${info.totalItemCount} items`);
    return false;
  }
  if (info.childFolderCount > 0) {
    console.log(`  SKIP delete "${name}" — still has ${info.childFolderCount} child folders`);
    return false;
  }
  await callGraphAPI(token, 'DELETE', `me/mailFolders/${folderId}`);
  console.log(`  Deleted: ${name}`);
  return true;
}

async function main() {
  const token = await getAccessToken();
  if (!token) { console.error('No token'); process.exit(1); }

  const inboxId = await resolveWellKnown(token, 'inbox');

  // Top-level category folders that should only exist under Inbox
  const categories = ['Notifications', 'Security', 'Vendors', 'Invoices'];

  for (const category of categories) {
    console.log(`\n=== ${category} ===`);

    // Find top-level folder (the wrong one)
    const topLevel = await findTopLevelFolder(token, category);
    if (!topLevel) {
      console.log('  No top-level folder found — OK');
      continue;
    }

    // Check if it's actually under Inbox (not a true top-level)
    if (topLevel.parentFolderId === inboxId) {
      console.log('  This IS the Inbox subfolder, not a duplicate — skipping');
      continue;
    }

    // Find the correct Inbox subfolder
    const inboxSub = await findChildFolder(token, inboxId, category);
    if (!inboxSub) {
      console.log(`  WARNING: Inbox/${category} does not exist — creating it`);
      await callGraphAPI(token, 'POST', `me/mailFolders/${inboxId}/childFolders`, { displayName: category });
    }
    const inboxSubFolder = await findChildFolder(token, inboxId, category);

    // Get child folders of the top-level duplicate
    const topChildren = await callGraphAPI(token, 'GET', `me/mailFolders/${topLevel.id}/childFolders`, null, {
      $top: '100', $select: 'id,displayName,totalItemCount,childFolderCount'
    });

    // Move messages from each top-level child to the corresponding Inbox child
    for (const topChild of (topChildren.value || [])) {
      console.log(`  Processing ${category}/${topChild.displayName} (${topChild.totalItemCount} items)`);

      // Find or create matching subfolder under Inbox/Category
      let inboxChild = await findChildFolder(token, inboxSubFolder.id, topChild.displayName);
      if (!inboxChild) {
        console.log(`    Creating Inbox/${category}/${topChild.displayName}`);
        await callGraphAPI(token, 'POST', `me/mailFolders/${inboxSubFolder.id}/childFolders`, { displayName: topChild.displayName });
        inboxChild = await findChildFolder(token, inboxSubFolder.id, topChild.displayName);
      }

      if (topChild.totalItemCount > 0) {
        const moved = await moveAllMessages(token, topChild.id, inboxChild.id, `${category}/${topChild.displayName}`);
        console.log(`    Moved ${moved} emails to Inbox/${category}/${topChild.displayName}`);
      }

      // Delete the now-empty top-level child
      await safeDeleteFolder(token, topChild.id, `${category}/${topChild.displayName}`);
    }

    // Move any messages directly in the top-level category folder
    if (topLevel.totalItemCount > 0) {
      const moved = await moveAllMessages(token, topLevel.id, inboxSubFolder.id, category);
      console.log(`  Moved ${moved} loose emails to Inbox/${category}`);
    }

    // Delete the top-level category folder
    await safeDeleteFolder(token, topLevel.id, category);
  }

  console.log('\nDone. Top-level duplicates cleaned up.');
}

main().catch(err => {
  console.error('Failed:', err.message);
  process.exit(1);
});
