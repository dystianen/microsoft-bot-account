const dotenv = require('dotenv');
dotenv.config();

module.exports = {
  concurrencyLimit: 2,
  microsoftUrl:
    'https://www.microsoft.com/en-gb/microsoft-365/enterprise/office-365-plans-and-pricing#plans',
  adsPower: {
    baseUrl: process.env.ADSPOWER_BASE_URL,
    apiKey: process.env.ADSPOWER_API_KEY,
    groupId: process.env.ADSPOWER_GROUP_ID,
  },
  proxy: {
    host: process.env.PROXY_HOST,
    port: process.env.PROXY_PORT,
    username: process.env.PROXY_USERNAME,
    password: process.env.PROXY_PASSWORD,
    type: 'socks5',
  },
  maxAccountsPerPayment: 3,
  telegram: {
    token: process.env.TELEGRAM_BOT_TOKEN,
    logChatId: process.env.TELEGRAM_LOG_CHAT_ID,
  },
  headless: process.env.HEADLESS === 'true',
  hardTimeout: 1.5 * 60 * 1000, // 90 seconds
  paymentTimeout: 5 * 60 * 1000, // 300 seconds
};
