const adsPowerHelper = require('./utils/adspower_helper');
const MicrosoftBot = require('./bots/microsoft_bot');
const { SuccessAccount } = require('./db/models');
const remoteLogger = require('./utils/logger');

async function saveToDB(result, telegram_id) {
  if (result.status !== 'SUCCESS') return;
  try {
    const successAcc = new SuccessAccount({
      email: result.email || 'unknown',
      password: result.domainPassword || result.password || 'unknown',
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

async function processSingleAccount(accountConfig, index, total, onPaymentSaved) {
  console.log(
    `\n--- Starting Account ${index + 1} of ${total}: ${accountConfig.microsoftAccount.email} ---`
  );

  let currentProfileId = null;
  let bot = null;
  let result = null;
  let executionResult = null;

  try {
    console.log(`[Account ${index + 1}] Creating AdsPower profile...`);
    const proxyOverride =
      accountConfig.proxyUsername && accountConfig.proxyPassword
        ? {
            username: accountConfig.proxyUsername,
            password: accountConfig.proxyPassword,
          }
        : null;

    currentProfileId = await adsPowerHelper.createProfile('', proxyOverride);
    console.log(`[Account ${index + 1}] Created profile: ${currentProfileId}`);

    console.log(`[Account ${index + 1}] Starting browser...`);
    const headlessOverride = accountConfig.headless !== undefined ? accountConfig.headless : null;
    const { wsUrl } = await adsPowerHelper.startBrowser(currentProfileId, headlessOverride);
    console.log(`[Account ${index + 1}] Browser started. WS URL: ${wsUrl}`);

    // ── Wrap onPaymentSaved dengan logging agar kita tahu kapan dipanggil ──
    const wrappedOnPaymentSaved = async () => {
      console.log(`[Account ${index + 1}] onPaymentSaved triggered — decrementing VCC...`);
      try {
        await onPaymentSaved();
        console.log(`[Account ${index + 1}] onPaymentSaved completed.`);
      } catch (err) {
        console.error(`[Account ${index + 1}] onPaymentSaved threw:`, err.message);
      }
    };
    bot = new MicrosoftBot(wsUrl, accountConfig, wrappedOnPaymentSaved);
    result = await bot.run();

    if (result && result.success) {
      console.log(
        `[Account ${index + 1}] Automation finished successfully. Domain: ${result.domainEmail}`
      );
      executionResult = {
        status: 'SUCCESS',
        domainEmail: result.domainEmail,
        domainPassword: accountConfig.microsoftAccount.password,
        log: 'Completed successfully',
      };

      const successDetails = `📧 <b>Login:</b> <code>${executionResult.domainEmail}</code>\n🔑 <b>Pass:</b> <code>${executionResult.domainPassword}</code>`;
      await remoteLogger.logSuccess(
        accountConfig.microsoftAccount.email,
        'Berhasil membuat akun Microsoft!',
        successDetails
      );
    } else {
      console.error(
        `[Account ${index + 1}] Automation failed: ${result?.error || 'Unknown error'}`
      );
      executionResult = {
        status: 'FAILED',
        domainEmail: '',
        domainPassword: '',
        log: result?.error || 'Unknown automation error',
      };
      await remoteLogger.logError(
        accountConfig.microsoftAccount.email,
        'Proses otomatisasi gagal',
        executionResult.log
      );
    }
  } catch (err) {
    console.error(`\n[ERROR Account ${index + 1}] failed:`, err.message);
    executionResult = {
      status: 'FAILED',
      domainEmail: '',
      domainPassword: '',
      log: err.message,
    };
    await remoteLogger.logError(
      accountConfig.microsoftAccount.email,
      'Terjadi kesalahan fatal',
      err.message
    );
  } finally {
    if (executionResult && executionResult.status !== 'SUCCESS') {
      console.log(`[Account ${index + 1}] Waiting 5s before cleanup...`);
      await new Promise((resolve) => setTimeout(resolve, 5000));
    }

    console.log(`[Account ${index + 1}] Starting cleanup...`);

    if (bot) {
      try {
        await bot.cleanup();
      } catch (e) {
        console.error(`[Account ${index + 1}] Bot cleanup error:`, e.message);
      }
    }

    if (currentProfileId) {
      try {
        await adsPowerHelper.stopBrowser(currentProfileId);
        console.log(`[Account ${index + 1}] Browser stopped.`);
      } catch (e) {
        console.warn(`[Account ${index + 1}] stopBrowser warning:`, e.message);
      }

      await new Promise((resolve) => setTimeout(resolve, 3000));

      try {
        await adsPowerHelper.deleteProfile(currentProfileId);
        console.log(`[Account ${index + 1}] AdsPower profile deleted.`);
      } catch (e) {
        console.error(`[Account ${index + 1}] deleteProfile error:`, e.message);
      }
    }

    if (!executionResult) {
      executionResult = {
        status: 'FAILED',
        domainEmail: '',
        domainPassword: '',
        log: 'Incomplete execution',
      };
    }

    if (executionResult.status === 'SUCCESS') {
      executionResult.email = accountConfig.microsoftAccount.email;
      await saveToDB(executionResult, accountConfig.telegram_id);
    }
  }

  return executionResult;
}

module.exports = {
  processSingleAccount,
};
