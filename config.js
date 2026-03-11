const dotenv = require("dotenv");
dotenv.config();

module.exports = {
  concurrencyLimit: 2, // Set to 1 if you want to run one account at a time
  // microsoftUrl:
  //   "https://signup.microsoft.com/get-started/signup?products=91dcd8b1-3b1b-444d-9cdb-0bc0da3eb40d&mproducts=CFQ7TTC0LH18:0002&fmproducts=CFQ7TTC0LH18:0002&renewalterm=P1Y&renewalbillingterm=P1Y&culture=en-us&country=us&ali=1", // US
  microsoftUrl:
    "https://signup.microsoft.com/get-started/signup?products=2ef55228-927f-4abc-ac51-befd4cdcd850&mproducts=CFQ7TTC0LF8S:0009&fmproducts=CFQ7TTC0LF8S:0009&renewalterm=P1Y&renewalbillingterm=P1Y&culture=ph-ph&country=ph&ali=1", // PH
  // microsoftUrl: "https://signup.microsoft.com/get-started/signup?products=2ef55228-927f-4abc-ac51-befd4cdcd850&mproducts=CFQ7TTC0LF8S:0009&fmproducts=CFQ7TTC0LF8S:0009&renewalterm=P1Y&renewalbillingterm=P1Y&culture=id-id&country=id&ali=1",
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
  maxAccountsPerPayment: 3, // Maximum number of Microsoft accounts per payment method
  telegram: {
    token: process.env.TELEGRAM_BOT_TOKEN,
  },
  headless: process.env.HEADLESS === "true", // Run browser in headless mode
};
