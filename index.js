const adsPowerHelper = require('./adspower_helper');
const MicrosoftBot = require('./microsoft_bot');

const fs = require('fs');

async function main() {
  try {
    // Read accounts from JSON file
    const accountsData = fs.readFileSync('./accounts.json', 'utf8');
    const accounts = JSON.parse(accountsData);

    console.log(`Loaded ${accounts.length} accounts for processing.`);

    for (let i = 0; i < accounts.length; i++) {
      const accountConfig = accounts[i];
      console.log(`\n--- Processing Account ${i + 1} of ${accounts.length}: ${accountConfig.microsoftAccount.email} ---`);
      
      const profileName = `MS-Account-${Date.now()}`;
      
      // 1. Create AdsPower profile
      console.log('Creating AdsPower profile...');
      const profileId = await adsPowerHelper.createProfile("");
      console.log(`Created profile: ${profileId}`);

      // 2. Start browser
      console.log('Starting browser...');
      const { wsUrl } = await adsPowerHelper.startBrowser(profileId);
      console.log(`Browser started. WS URL: ${wsUrl}`);

      // 3. Run Microsoft automation
      try {
        const bot = new MicrosoftBot(wsUrl, accountConfig);
        await bot.run();
        console.log(`Automation finished successfully for account ${i + 1}.`);
      } catch (err) {
        console.error(`\n[ERROR] Account ${i + 1} failed:`, err.message);
        console.log("Cleaning up profile due to failure...");
        
        // Stop browser & Delete profile on failure
        await adsPowerHelper.stopBrowser(profileId);
        
        // Kasih jeda sebentar agar browser benar-benar tertutup sebelum didelete
        await new Promise(resolve => setTimeout(resolve, 2000));
        await adsPowerHelper.deleteProfile(profileId);
      }
      
      // Add delay between accounts (e.g., 10-20 seconds) to avoid spamming
      if (i < accounts.length - 1) {
        const delay = Math.floor(Math.random() * 10000) + 10000;
        console.log(`Waiting for ${delay / 1000} seconds before next account...`);
        await new Promise(resolve => setTimeout(resolve, delay));
      }
    }

    console.log('\nAll accounts processed!');
  } catch (error) {
    console.error('Fatal execution error:', error.message);
    process.exit(1);
  }
}

main();
