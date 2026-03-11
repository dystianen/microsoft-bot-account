const adsPowerHelper = require("./adspower_helper");
const MicrosoftBot = require("./microsoft_bot");
const config = require("./config");
const { SuccessAccount } = require("./models");

let _saveQueue = Promise.resolve();

async function saveToDB(result, telegram_id) {
  if (result.status !== "SUCCESS") return;
  try {
    const successAcc = new SuccessAccount({
      email: result.email || "unknown",
      password: result.domainPassword || result.password || "unknown",
      domainEmail: result.domainEmail,
      domainPassword: result.domainPassword,
      telegram_id: telegram_id.toString(),
    });
    await successAcc.save();
    console.log(`[DB] Account saved for user ${telegram_id}: ${result.domainEmail}`);
  } catch (err) {
    console.error(`[DB] Error saving to DB: ${err.message}`);
  }
}

// File-based history removed in favor of MongoDB

async function processSingleAccount(accountConfig, index, total) {
  const profileName = `MS-Account-${Date.now()}-${index}`;

  console.log(
    `\n--- Starting Account ${index + 1} of ${total}: ${accountConfig.microsoftAccount.email} ---`,
  );

  let currentProfileId = null;
  let bot = null;
  let result = null;
  let executionResult = null;

  try {
    // 1. Create AdsPower profile
    console.log(`[Account ${index + 1}] Creating AdsPower profile...`);
    const proxyOverride = (accountConfig.proxyUsername && accountConfig.proxyPassword) ? {
      username: accountConfig.proxyUsername,
      password: accountConfig.proxyPassword
    } : null;
    
    currentProfileId = await adsPowerHelper.createProfile(profileName, proxyOverride);
    console.log(`[Account ${index + 1}] Created profile: ${currentProfileId}`);

    // 2. Start browser
    console.log(`[Account ${index + 1}] Starting browser...`);
    const headlessOverride = accountConfig.headless !== undefined ? accountConfig.headless : null;
    const { wsUrl } = await adsPowerHelper.startBrowser(currentProfileId, headlessOverride);
    console.log(`[Account ${index + 1}] Browser started. WS URL: ${wsUrl}`);

    // 3. Run Microsoft automation
    bot = new MicrosoftBot(wsUrl, accountConfig);
    result = await bot.run();

    if (result && result.success) {
      console.log(
        `[Account ${index + 1}] Automation finished successfully. Domain: ${result.domainEmail} Password: ${accountConfig.microsoftAccount.password}`,
      );
      executionResult = {
        status: "SUCCESS",
        domainEmail: result.domainEmail,
        domainPassword: accountConfig.microsoftAccount.password,
        log: "Completed successfully",
      };
    } else {
      console.error(
        `[Account ${index + 1}] Automation failed: ${result?.error || "Unknown error"}`,
      );
      executionResult = {
        status: "FAILED",
        domainEmail: "",
        domainPassword: "",
        log: result?.error || "Unknown automation error",
      };
    }
  } catch (err) {
    console.error(`\n[ERROR Account ${index + 1}] failed:`, err.message);
    executionResult = {
      status: "ERROR",
      domainEmail: "",
      domainPassword: "",
      log: err.message,
    };
  } finally {
    console.log(`[Account ${index + 1}] Starting cleanup...`);

    // 1. Close the browser instance through Playwright first
    if (bot) {
      try {
        await bot.cleanup();
      } catch (e) {
        console.error(`[Account ${index + 1}] Bot cleanup error:`, e.message);
      }
    }

    // 2. Stop & Delete profil AdsPower
    if (currentProfileId) {
      try {
        await adsPowerHelper.stopBrowser(currentProfileId);
        console.log(`[Account ${index + 1}] Browser stopped.`);
      } catch (e) {
        console.warn(
          `[Account ${index + 1}] stopBrowser warning (proceeding anyway):`,
          e.message,
        );
      }

      await new Promise((resolve) => setTimeout(resolve, 3000)); // 3s cukup, 10s terlalu lama

      try {
        await adsPowerHelper.deleteProfile(currentProfileId);
        console.log(`[Account ${index + 1}] AdsPower profile deleted.`);
      } catch (e) {
        console.error(`[Account ${index + 1}] deleteProfile error:`, e.message);
      }
    }

    if (!executionResult) {
      executionResult = {
        status: "FAILED",
        domainEmail: "",
        domainPassword: "",
        log: "Incomplete execution",
      };
    }

    // Save to global history if success
    if (executionResult.status === "SUCCESS") {
      executionResult.email = accountConfig.microsoftAccount.email;
      await saveToDB(executionResult, accountConfig.telegram_id);
    }
  }

  return executionResult;
}

module.exports = {
  processSingleAccount,
};
