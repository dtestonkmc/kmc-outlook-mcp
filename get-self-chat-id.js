#!/usr/bin/env node
/**
 * One-time setup: retrieves your Teams self-chat ID from Microsoft Graph.
 * Run: node get-self-chat-id.js
 * Copy the printed ID to TEAMS_SELF_CHAT_ID in your .env file.
 */
require('dotenv').config();
const { getAccessToken } = require('./auth/token-manager');
const { callGraphAPI } = require('./utils/graph-api');

async function main() {
  const token = await getAccessToken();
  if (!token) {
    console.error('No valid token. Run "npm run auth-server" first to authenticate.');
    process.exit(1);
  }

  try {
    // Get current user info first
    const me = await callGraphAPI(token, 'GET', 'me', null, { $select: 'id' });
    const myId = me.id;

    // List oneOnOne chats and find the one where both members are you
    const response = await callGraphAPI(token, 'GET', 'me/chats', null, {
      chatType: 'oneOnOne',
      $expand: 'members',
      $top: '50'
    });
    const chats = response.value || [];

    const selfChat = chats.find(chat => {
      const members = chat.members || [];
      return members.length === 1 ||
        (members.length === 2 && members.every(m => (m.userId || '').toLowerCase() === myId.toLowerCase()));
    });

    if (!selfChat) {
      console.error('No self-chat found. Open Microsoft Teams and send yourself a message first.');
      process.exit(1);
    }

    console.log('\nYour Teams self-chat ID:');
    console.log(selfChat.id);
    console.log('\nAdd this to your .env file:');
    console.log(`TEAMS_SELF_CHAT_ID=${selfChat.id}`);
  } catch (err) {
    console.error('Error fetching self-chat:', err.message);
    process.exit(1);
  }
}

main();
