const dotenv = require("dotenv");
dotenv.config();

module.exports = {
  microsoftUrl:
    "https://www.microsoft.com/en-us/microsoft-365/business/microsoft-365-business-basic",
  adsPower: {
    baseUrl: process.env.ADSPOWER_BASE_URL || "http://local.adspower.net:50325",
    apiKey:
      process.env.ADSPOWER_API_KEY ||
      "a74c1696169a859321288618723ac1340051adfbdc2511a1",
    groupId: process.env.ADSPOWER_GROUP_ID || "0",
  },
  proxy: {
    host: process.env.PROXY_HOST || "gw.dataimpulse.com",
    port: process.env.PROXY_PORT || "824",
    username:
      process.env.PROXY_USERNAME || "2028ce76fd49cc77d009__cr.fi_rotate",
    password: process.env.PROXY_PASSWORD || "2daec922ca463aaa",
    type: "socks5",
  },
  microsoftAccount: {
    email: "elowen.finch@desmoinesiowa.us",
    firstName: "Elowen",
    lastName: "Finch",
    companyName: "Des Moines Agricultural Cooperative",
    companySize: "1 person",
    phone: "+15155556789",
    jobTitle: "Corn & Soybean Commodity Trader",
    address: "400 Grand Avenue",
    city: "Des Moines",
    state: "Iowa",
    postalCode: "50309",
    country: "United States",
    password: "AgriElowen6789@DesMoines",
  },
  payment: {
    cardNumber: "5198939816602718",
    cvv: "213",
    expMonth: "03",
    expYear: "30",
  },
};
