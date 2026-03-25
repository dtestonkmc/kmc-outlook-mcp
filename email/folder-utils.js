/**
 * Email folder utilities
 */
const { callGraphAPI } = require('../utils/graph-api');

/**
 * Cache of folder information to reduce API calls
 * Format: { userId: { folderName: { id, path } } }
 */
const folderCache = {};

/**
 * Well-known folder names and their Graph API aliases
 */
const WELL_KNOWN_ALIASES = {
  'inbox': 'inbox',
  'drafts': 'drafts',
  'sent': 'sentitems',
  'sent items': 'sentitems',
  'deleted': 'deleteditems',
  'deleted items': 'deleteditems',
  'junk': 'junkemail',
  'junk email': 'junkemail',
  'archive': 'archive'
};

/**
 * Well-known folder names and their endpoints (for message listing)
 */
const WELL_KNOWN_FOLDERS = {
  'inbox': 'me/mailFolders/inbox/messages',
  'drafts': 'me/mailFolders/drafts/messages',
  'sent': 'me/mailFolders/sentItems/messages',
  'deleted': 'me/mailFolders/deletedItems/messages',
  'junk': 'me/mailFolders/junkemail/messages',
  'archive': 'me/mailFolders/archive/messages'
};

/**
 * Resolve a folder name to its endpoint path
 * @param {string} accessToken - Access token
 * @param {string} folderName - Folder name or path to resolve
 * @returns {Promise<string>} - Resolved endpoint path
 */
async function resolveFolderPath(accessToken, folderName) {

  // Default to inbox if no folder specified
  if (!folderName) {
    return WELL_KNOWN_FOLDERS['inbox'];
  }

  // Check if it's a well-known folder (case-insensitive)
  const lowerFolderName = folderName.toLowerCase();
  if (WELL_KNOWN_FOLDERS[lowerFolderName]) {
    console.error(`Using well-known folder path for "${folderName}"`);
    return WELL_KNOWN_FOLDERS[lowerFolderName];
  }

  try {
    // Try to find the folder by name (supports nested paths)
    const folderId = await getFolderIdByName(accessToken, folderName);
    if (folderId) {
      const path = `me/mailFolders/${folderId}/messages`;
      console.error(`Resolved folder "${folderName}" to path: ${path}`);
      return path;
    }

    // If not found, fall back to inbox
    console.error(`Couldn't find folder "${folderName}", falling back to inbox`);
    return WELL_KNOWN_FOLDERS['inbox'];
  } catch (error) {
    console.error(`Error resolving folder "${folderName}": ${error.message}`);
    return WELL_KNOWN_FOLDERS['inbox'];
  }
}

/**
 * Resolve a well-known folder name to its Graph API folder ID
 * @param {string} accessToken - Access token
 * @param {string} name - Folder name (e.g. "Inbox")
 * @returns {Promise<string|null>} - Folder ID or null
 */
async function resolveWellKnownFolderId(accessToken, name) {
  const alias = WELL_KNOWN_ALIASES[name.toLowerCase()];
  if (!alias) return null;
  try {
    const result = await callGraphAPI(accessToken, 'GET', `me/mailFolders/${alias}`, null, { $select: 'id' });
    return result.id || null;
  } catch {
    return null;
  }
}

/**
 * Find a child folder by name under a given parent folder ID
 * @param {string} accessToken - Access token
 * @param {string} parentFolderId - Parent folder ID
 * @param {string} childName - Child folder display name
 * @returns {Promise<string|null>} - Child folder ID or null
 */
async function findChildFolder(accessToken, parentFolderId, childName) {
  try {
    const response = await callGraphAPI(
      accessToken,
      'GET',
      `me/mailFolders/${parentFolderId}/childFolders`,
      null,
      { $filter: `displayName eq '${childName}'`, $select: 'id,displayName' }
    );
    if (response.value && response.value.length > 0) {
      return response.value[0].id;
    }
    // Case-insensitive fallback
    const allChildren = await callGraphAPI(
      accessToken,
      'GET',
      `me/mailFolders/${parentFolderId}/childFolders`,
      null,
      { $top: '100', $select: 'id,displayName' }
    );
    if (allChildren.value) {
      const match = allChildren.value.find(
        f => f.displayName.toLowerCase() === childName.toLowerCase()
      );
      if (match) return match.id;
    }
    return null;
  } catch (error) {
    console.error(`Error finding child folder "${childName}": ${error.message}`);
    return null;
  }
}

/**
 * Get the ID of a mail folder by its name or path
 * Supports nested paths like "Inbox/Notifications/UniFi"
 * @param {string} accessToken - Access token
 * @param {string} folderName - Name or slash-separated path of the folder
 * @returns {Promise<string|null>} - Folder ID or null if not found
 */
async function getFolderIdByName(accessToken, folderName) {
  try {
    console.error(`Looking for folder "${folderName}"`);

    // If it contains slashes, walk the path
    if (folderName.includes('/')) {
      const parts = folderName.split('/');
      let currentId = null;

      for (const part of parts) {
        if (!currentId) {
          // First segment — try well-known alias first
          currentId = await resolveWellKnownFolderId(accessToken, part);
          if (currentId) {
            console.error(`Resolved well-known folder "${part}" -> ${currentId}`);
            continue;
          }
          // Fall back to top-level search
          currentId = await findTopLevelFolder(accessToken, part);
          if (!currentId) {
            console.error(`Top-level folder "${part}" not found`);
            return null;
          }
        } else {
          // Subsequent segments — search children
          const childId = await findChildFolder(accessToken, currentId, part);
          if (!childId) {
            console.error(`Child folder "${part}" not found under parent`);
            return null;
          }
          currentId = childId;
        }
      }

      console.error(`Resolved path "${folderName}" -> ${currentId}`);
      return currentId;
    }

    // Simple name — search top-level
    return await findTopLevelFolder(accessToken, folderName);
  } catch (error) {
    console.error(`Error finding folder "${folderName}": ${error.message}`);
    return null;
  }
}

/**
 * Find a top-level folder by display name
 * @param {string} accessToken - Access token
 * @param {string} name - Folder display name
 * @returns {Promise<string|null>} - Folder ID or null
 */
async function findTopLevelFolder(accessToken, name) {
  // Try exact match filter
  const response = await callGraphAPI(
    accessToken,
    'GET',
    'me/mailFolders',
    null,
    { $filter: `displayName eq '${name}'` }
  );

  if (response.value && response.value.length > 0) {
    console.error(`Found folder "${name}" with ID: ${response.value[0].id}`);
    return response.value[0].id;
  }

  // Case-insensitive fallback
  console.error(`No exact match for "${name}", trying case-insensitive search`);
  const allFoldersResponse = await callGraphAPI(
    accessToken,
    'GET',
    'me/mailFolders',
    null,
    { $top: '100' }
  );

  if (allFoldersResponse.value) {
    const match = allFoldersResponse.value.find(
      f => f.displayName.toLowerCase() === name.toLowerCase()
    );
    if (match) {
      console.error(`Found case-insensitive match for "${name}" with ID: ${match.id}`);
      return match.id;
    }
  }

  console.error(`No folder found matching "${name}"`);
  return null;
}

/**
 * Get all mail folders (recursively)
 * @param {string} accessToken - Access token
 * @returns {Promise<Array>} - Array of folder objects
 */
async function getAllFolders(accessToken) {
  try {
    const selectFields = 'id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount';

    // Get top-level folders
    const response = await callGraphAPI(
      accessToken,
      'GET',
      'me/mailFolders',
      null,
      { $top: '100', $select: selectFields }
    );

    if (!response.value) {
      return [];
    }

    const allFolders = [...response.value];

    // Recursively get children
    async function getChildren(parentFolders) {
      const withChildren = parentFolders.filter(f => f.childFolderCount > 0);
      if (withChildren.length === 0) return;

      const childPromises = withChildren.map(async (folder) => {
        try {
          const childResponse = await callGraphAPI(
            accessToken,
            'GET',
            `me/mailFolders/${folder.id}/childFolders`,
            null,
            { $select: selectFields }
          );
          return childResponse.value || [];
        } catch (error) {
          console.error(`Error getting child folders for "${folder.displayName}": ${error.message}`);
          return [];
        }
      });

      const childArrays = await Promise.all(childPromises);
      const children = childArrays.flat();
      allFolders.push(...children);

      // Recurse into grandchildren
      await getChildren(children);
    }

    await getChildren(response.value);
    return allFolders;
  } catch (error) {
    console.error(`Error getting all folders: ${error.message}`);
    return [];
  }
}

module.exports = {
  WELL_KNOWN_FOLDERS,
  resolveFolderPath,
  getFolderIdByName,
  getAllFolders
};
