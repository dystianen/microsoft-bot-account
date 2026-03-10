const adsPowerHelper = require("./adspower_helper");
const MicrosoftBot = require("./microsoft_bot");
const config = require("./config");
const fs = require("fs");
const XLSX = require("xlsx");

const EXCEL_FILE = "./accounts_result.xlsx";

function generateExcelReport(accounts, results) {
  const data = accounts.map((acc, index) => {
    const res = results[index] || { status: "FAILED", domainEmail: "", domainPassword: "", log: "Incomplete execution" };
    
    return {
      "Email": acc.microsoftAccount.email,
      "Password": acc.microsoftAccount.password,
      "First Name": acc.microsoftAccount.firstName,
      "Last Name": acc.microsoftAccount.lastName,
      "Company Name": acc.microsoftAccount.companyName,
      "Company Size": acc.microsoftAccount.companySize || "1 person",
      "Phone": acc.microsoftAccount.phone,
      "Job Title": acc.microsoftAccount.jobTitle,
      "Address": acc.microsoftAccount.address,
      "City": acc.microsoftAccount.city,
      "State": acc.microsoftAccount.state,
      "Postal Code": acc.microsoftAccount.postalCode,
      "Country": acc.microsoftAccount.country || "United States",
      "Card Number": acc.payment.cardNumber,
      "CVV": acc.payment.cvv,
      "Exp Month": acc.payment.expMonth,
      "Exp Year": acc.payment.expYear,
      "Status": res.status,
      "Domain Email": res.domainEmail || "",
      "Domain Password": res.domainPassword || "",
      "Log": res.log || ""
    };
  });

  const worksheet = XLSX.utils.json_to_sheet(data);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Results");

  // Adjust column widths
  worksheet["!cols"] = [
    { wch: 40 }, { wch: 25 }, { wch: 15 }, { wch: 15 }, { wch: 30 },
    { wch: 15 }, { wch: 15 }, { wch: 25 }, { wch: 30 }, { wch: 15 },
    { wch: 15 }, { wch: 12 }, { wch: 15 }, { wch: 20 }, { wch: 8 },
    { wch: 10 }, { wch: 10 }, { wch: 15 }, { wch: 45 }, { wch: 25 }, 
    { wch: 50 },
  ];

  XLSX.writeFile(workbook, EXCEL_FILE);
  console.log(`\n[Excel] Saved results to ${EXCEL_FILE}.`);
}

function loadAccounts() {
  const paymentsFile = './payments.json';
  const accountsFile = './microsoft_accounts.json';

  if (!fs.existsSync(paymentsFile) || !fs.existsSync(accountsFile)) {
    throw new Error('Required files (payments.json or microsoft_accounts.json) are missing.');
  }

  const payments = JSON.parse(fs.readFileSync(paymentsFile, 'utf8'));
  const microsoftAccounts = JSON.parse(fs.readFileSync(accountsFile, 'utf8'));
  const maxPerPayment = config.maxAccountsPerPayment || 5;

  console.log(`[Load] Found ${payments.length} payment methods and ${microsoftAccounts.length} Microsoft accounts.`);
  console.log(`[Load] Max accounts per payment: ${maxPerPayment}`);

  const pairedAccounts = [];

  microsoftAccounts.forEach((msAcc, index) => {
    // Calculate which payment to use
    const paymentIndex = Math.floor(index / maxPerPayment);
    
    if (paymentIndex >= payments.length) {
      console.warn(`[Warning] No payment method available for account ${index + 1} (${msAcc.email}). Skipping.`);
      return;
    }

    pairedAccounts.push({
      microsoftAccount: msAcc,
      payment: payments[paymentIndex]
    });
  });

  console.log(`[Load] Paired ${pairedAccounts.length} accounts for processing.`);
  return pairedAccounts;
}

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
    currentProfileId = await adsPowerHelper.createProfile(profileName);
    console.log(`[Account ${index + 1}] Created profile: ${currentProfileId}`);

    // 2. Start browser
    console.log(`[Account ${index + 1}] Starting browser...`);
    const { wsUrl } = await adsPowerHelper.startBrowser(currentProfileId);
    console.log(`[Account ${index + 1}] Browser started. WS URL: ${wsUrl}`);

    // 3. Run Microsoft automation
    bot = new MicrosoftBot(wsUrl, accountConfig);
    result = await bot.run();

    if (result && result.success) {
      console.log(
        `[Account ${index + 1}] Automation finished successfully. Domain: ${result.domainEmail} Password: ${accountConfig.microsoftAccount.password}`,
      );
      executionResult = { status: "SUCCESS", domainEmail: result.domainEmail, domainPassword: accountConfig.microsoftAccount.password, log: "Completed successfully" };
    } else {
      console.error(
        `[Account ${index + 1}] Automation failed: ${result?.error || "Unknown error"}`,
      );
      executionResult = { status: "FAILED", domainEmail: "", domainPassword: "", log: result?.error || "Unknown automation error" };
    }
  } catch (err) {
    console.error(`\n[ERROR Account ${index + 1}] failed:`, err.message);
    executionResult = { status: "ERROR", domainEmail: "", domainPassword: "", log: err.message };
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
        // Wait a bit before deleting to ensure API process is ready
        await new Promise((resolve) => setTimeout(resolve, 5000));
        await adsPowerHelper.deleteProfile(currentProfileId);
        console.log(`[Account ${index + 1}] AdsPower profile cleaned up.`);
      } catch (cleanupError) {
        console.error(
          `[Account ${index + 1}] AdsPower cleanup error:`,
          cleanupError.message,
        );
      }
    }

    if (!executionResult) {
      executionResult = { status: "FAILED", domainEmail: "", domainPassword: "", log: "Incomplete execution" };
    }
  }
  
  return executionResult;
}

async function main() {
  try {
    // Load paired accounts
    const accounts = loadAccounts();

    const concurrencyLimit = config.concurrencyLimit || 3;
    console.log(
      `Loaded ${accounts.length} paired accounts. Concurrency limit: ${concurrencyLimit}`,
    );

    const executing = new Set();
    const tasks = [];
    const results = new Array(accounts.length);

    for (let i = 0; i < accounts.length; i++) {
      const accountConfig = accounts[i];

      const promise = processSingleAccount(
        accountConfig,
        i,
        accounts.length,
      ).then((res) => {
        results[i] = res; // Store result
        executing.delete(promise);
      });

      tasks.push(promise);
      executing.add(promise);

      if (executing.size >= concurrencyLimit) {
        await Promise.race(executing);
      }

      // Optional: Add a small staggered startup delay (e.g., 2-5 seconds)
      // to avoid triggering anti-bot by opening many browsers exactly at the same time
      if (i < accounts.length - 1) {
        const staggerDelay = 2000;
        await new Promise((resolve) => setTimeout(resolve, staggerDelay));
      }
    }

    await Promise.all(tasks);
    console.log("\nAll accounts processing attempts finished!");

    // Generate Excel report at the very end
    generateExcelReport(accounts, results);
  } catch (error) {
    console.error("Fatal execution error:", error.message);
    process.exit(1);
  }
}

module.exports = {
  processSingleAccount,
  generateExcelReport,
  loadAccounts
};

if (require.main === module) {
  main();
}
