/**
 * Move email to trash functionality
 */
const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

/**
 * Move email to Deleted Items handler
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleMoveToTrash(args) {
  const { messageId } = args;

  if (!messageId) {
    return {
      content: [{
        type: 'text',
        text: 'Error: messageId is required'
      }]
    };
  }

  try {
    const accessToken = await ensureAuthenticated();
    await callGraphAPI(
      accessToken,
      'POST',
      `me/messages/${messageId}/move`,
      { destinationId: 'deleteditems' }
    );
    return {
      content: [{
        type: 'text',
        text: `Message ${messageId} moved to Deleted Items`
      }]
    };
  } catch (error) {
    return {
      content: [{
        type: 'text',
        text: `Error moving message to trash: ${error.message}`
      }]
    };
  }
}

const moveToTrashTool = {
  name: 'move-to-trash',
  description: 'Move an email message to the Deleted Items folder',
  inputSchema: {
    type: 'object',
    properties: {
      messageId: {
        type: 'string',
        description: 'The ID of the message to move to Deleted Items'
      }
    },
    required: ['messageId']
  },
  handler: handleMoveToTrash
};

module.exports = { handleMoveToTrash, moveToTrashTool };
