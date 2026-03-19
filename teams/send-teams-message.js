const { callGraphAPI } = require('../utils/graph-api');
const { ensureAuthenticated } = require('../auth');

async function handleSendTeamsMessage(args) {
  const { message } = args;

  if (!message) {
    return { success: false, error: 'message is required' };
  }

  const chatId = process.env.TEAMS_SELF_CHAT_ID;
  if (!chatId) {
    return { success: false, error: 'TEAMS_SELF_CHAT_ID is not set in environment. Run get-self-chat-id.js to find your chat ID.' };
  }

  try {
    const accessToken = await ensureAuthenticated();
    await callGraphAPI(accessToken, 'POST', `/me/chats/${chatId}/messages`, {
      body: { content: message, contentType: 'text' }
    });
    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

const sendTeamsMessageTool = {
  name: 'send-teams-message',
  description: 'Send a message to the configured Teams self-chat',
  inputSchema: {
    type: 'object',
    properties: {
      message: {
        type: 'string',
        description: 'The text content to send to Teams self-chat'
      }
    },
    required: ['message']
  },
  handler: handleSendTeamsMessage
};

module.exports = { handleSendTeamsMessage, sendTeamsMessageTool };
