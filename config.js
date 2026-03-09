const dotenv = require("dotenv");
dotenv.config();

module.exports = {
  concurrencyLimit: 10,
  microsoftUrl:
    "https://www.microsoft.com/en-us/microsoft-365/business/microsoft-365-business-basic",
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
    type: "socks5",
  },
};
