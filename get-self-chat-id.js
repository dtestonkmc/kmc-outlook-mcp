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
    const response = await callGraphAPI(token, 'GET', "/me/chats?$filter=chatType eq 'selfChat'");
    const chats = response.value || [];

    if (chats.length === 0) {
      console.error('No self-chat found. Open Microsoft Teams and send yourself a message first.');
      process.exit(1);
    }

    const selfChat = chats[0];
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
