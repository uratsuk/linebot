require('dotenv').config();
const express = require('express');
const axios = require('axios');
const line = require('@line/bot-sdk');

const app = express();
app.use(express.json());

const config = {
  channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN,
  channelSecret: process.env.LINE_CHANNEL_SECRET,
};

const client = new line.Client(config);

app.post('/webhook', line.middleware(config), async (req, res) => {
  res.sendStatus(200); // 👈 先にLINEに「OK」と返す！

  const events = req.body.events;
  console.log('📩 Webhook received:', JSON.stringify(events, null, 2));

  for (const event of events) {
    if (event.type === 'message' && event.message.type === 'text') {
      const userMessage = event.message.text;
      console.log('💬 User message:', userMessage);

      try {
        const gptResponse = await axios.post(
          'https://api.openai.com/v1/chat/completions',
          {
            model: 'gpt-3.5-turbo',
            messages: [{ role: 'user', content: userMessage }],
            temperature: 0,
            max_tokens: 200,
          },
          {
            headers: {
              Authorization: `Bearer ${process.env.OPENAI_API_KEY}`,
              'Content-Type': 'application/json',
            },
          }
        );

        const replyText = gptResponse.data.choices[0].message.content;
        console.log('🤖 GPT reply:', replyText);

        await client.replyMessage(event.replyToken, {
          type: 'text',
          text: replyText,
        });

      } catch (error) {
        console.error('❌ ChatGPT error:', error.response?.data || error.message);
        await client.replyMessage(event.replyToken, {
          type: 'text',
          text: 'ごめんなさい、今ちょっと考え中です…💦',
        });
      }
    }
  }
});

app.listen(3000, () => {
  console.log('🤖 ChatGPT Bot is running on http://localhost:3000');
});
