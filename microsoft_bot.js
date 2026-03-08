const { chromium } = require("playwright-core");
const config = require("./config");

class MicrosoftBot {
  constructor(wsUrl) {
    this.wsUrl = wsUrl;
  }

  async humanDelay(min = 500, max = 1500) {
    const delay = Math.floor(Math.random() * (max - min + 1) + min);
    await this.page.waitForTimeout(delay);
  }

  async randomMouseMove() {
    const { width, height } = this.page.viewportSize() || {
      width: 1280,
      height: 720,
    };
    const x = Math.floor(Math.random() * width);
    const y = Math.floor(Math.random() * height);
    await this.page.mouse.move(x, y, { steps: 10 });
  }

  async waitForPage(selector) {
    if (selector) {
      // Tunggu elemen spesifik muncul = page sudah siap
      await this.page.waitForSelector(selector, {
        state: "attached",
        timeout: 120000,
      });
    } else {
      await this.page.waitForLoadState("domcontentloaded", { timeout: 150000 });
    }
  }

  async connect() {
    console.log("[STEP 1] Connecting to browser");

    this.browser = await chromium.connectOverCDP(this.wsUrl);

    const contexts = this.browser.contexts();
    this.context =
      contexts.length > 0 ? contexts[0] : await this.browser.newContext();

    const pages = this.context.pages();
    this.page = pages.length > 0 ? pages[0] : await this.context.newPage();
    this.profileId = this.wsUrl.split("/").pop(); // Extract profile ID from WS URL if possible
  }

  async openMicrosoftPage() {
    console.log("[STEP 2] Opening Microsoft page");

    await this.page.goto(config.microsoftUrl, {
      waitUntil: "domcontentloaded",
      timeout: 100000,
    });
  }

  async clickTryButton() {
    console.log("[STEP 3] Clicking Try button");

    await this.waitForPage("#action-oc5f9e");

    // Tangkap new page sebelum click
    const [newPage] = await Promise.all([
      this.context.waitForEvent("page"),
      this.page.evaluate(() => {
        document.querySelector("#action-oc5f9e").click();
      }),
    ]);

    // Switch this.page ke tab baru
    await newPage.waitForLoadState("domcontentloaded");
    this.page = newPage;

    console.log("[STEP 3] Switched to new page:", this.page.url());
  }

  async clickBuildCartNextButton() {
    console.log("[STEP 4] Clicking Next button");

    await this.waitForPage('[data-bi-id="BuildCartNext"]');

    await this.page.evaluate(() => {
      document.querySelector('[data-bi-id="BuildCartNext"]').click();
    });
  }

  async fillEmail() {
    console.log("[STEP 5] Filling email");

    // Tunggu sampai input email muncul
    await this.waitForPage('[data-bi-id="Email"]');

    const emailInput = this.page.locator('[data-bi-id="Email"]');
    await emailInput.pressSequentially(config.microsoftAccount.email, {
      delay: Math.floor(Math.random() * 50) + 70,
    });
    await this.humanDelay(1000, 2000);
  }

  async clickCollectEmailNextButton() {
    console.log("[STEP 6] Clicking CollectEmail Next button");

    await this.waitForPage('[data-bi-id="CollectEmailNext"]');

    await this.page.evaluate(() => {
      document.querySelector('[data-bi-id="CollectEmailNext"]').click();
    });

    // Tunggu verifikasi manual selesai, baru lanjut ke setup account
    console.log("[INFO] Waiting for email verification to complete...");
    await this.waitForPage('[data-bi-id="ConfirmEmailSetupAccount"]');
    console.log("[INFO] Verification complete, setup account button detected");
  }

  async clickConfirmEmailSetupAccountButton() {
    console.log("[STEP 7] Clicking Setup Account button");

    await this.page.evaluate(() => {
      document.querySelector('[data-bi-id="ConfirmEmailSetupAccount"]').click();
    });
  }

  async fillBasicInfo() {
    console.log("[STEP 8] Filling basic info");

    await this.waitForPage('[data-testid="firstNameField"]');

    // Fill semua text fields secara human-like
    const fields = [
      {
        selector: '[data-testid="firstNameField"]',
        value: config.microsoftAccount.firstName,
      },
      {
        selector: '[data-testid="lastNameField"]',
        value: config.microsoftAccount.lastName,
      },
      {
        selector: '[data-testid="companyNameField"]',
        value: config.microsoftAccount.companyName,
      },
      {
        selector: '[data-testid="phoneNumberField"]',
        value: config.microsoftAccount.phone,
      },
      {
        selector: '[data-testid="jobTitle"]',
        value: config.microsoftAccount.jobTitle,
      },
    ];

    for (const field of fields) {
      const locator = this.page.locator(field.selector);
      await locator.click();
      await locator.pressSequentially(field.value, {
        delay: Math.floor(Math.random() * 40) + 60,
      });
      await this.humanDelay(400, 800);
    }

    await this.page.locator("#address_line1").click();
    await this.page
      .locator("#address_line1")
      .pressSequentially(config.microsoftAccount.address, {
        delay: Math.floor(Math.random() * 30) + 50,
      });
    await this.humanDelay(500, 1000);

    await this.page.locator("#city").click();
    await this.page
      .locator("#city")
      .pressSequentially(config.microsoftAccount.city, {
        delay: Math.floor(Math.random() * 30) + 50,
      });
    await this.humanDelay(500, 1000);

    await this.page.locator("#postal_code").click();
    await this.page
      .locator("#postal_code")
      .pressSequentially(config.microsoftAccount.postalCode, {
        delay: Math.floor(Math.random() * 30) + 50,
      });
    await this.humanDelay(800, 1500);

    // Pilih company size (random)
    await this.selectDropdownByText(
      '[data-testid="companySizeDropdown"]',
      config.microsoftAccount.companySize,
    );
    await this.humanDelay(600, 1200);

    // Pilih state Alabama (sesuai config)
    await this.selectDropdownByText("#input_region", "Alabama");
    await this.humanDelay(600, 1200);

    // Pilih No untuk website
    await this.selectDropdownByText('[data-testid="websiteDropdown"]', "No");
    await this.humanDelay(600, 1200);

    // Check partner checkbox
    await this.page.evaluate(() => {
      const checkbox = document.querySelector("#partner-checkbox");
      if (checkbox && !checkbox.checked) checkbox.click();
    });

    await this.humanDelay(1000, 2000);

    // Click Next
    await this.page.evaluate(() => {
      document.querySelector('[data-bi-id="SignupNext"]').click();
    });
  }

  async selectRandomDropdown(selector) {
    // Open dropdown
    await this.page.evaluate((sel) => {
      document.querySelector(sel).click();
    }, selector);

    // Tunggu options muncul
    await this.page.waitForSelector(".ms-Dropdown-items", {
      state: "attached",
      timeout: 10000,
    });

    // Pilih random option (skip index 0 karena biasanya placeholder)
    await this.page.evaluate((sel) => {
      const dropdown = document.querySelector(sel);
      const options =
        dropdown
          .closest(".ms-Dropdown-container")
          ?.querySelectorAll(".ms-Dropdown-item") ||
        document.querySelectorAll(".ms-Dropdown-items .ms-Dropdown-item");

      const validOptions = Array.from(options).filter(
        (o) => !o.classList.contains("is-disabled"),
      );
      const randomIndex = Math.floor(Math.random() * validOptions.length);
      validOptions[randomIndex]?.click();
    }, selector);
  }

  async selectDropdownByText(selector, text) {
    // Open dropdown
    await this.page.evaluate((sel) => {
      document.querySelector(sel).click();
    }, selector);

    await this.page.waitForSelector(".ms-Dropdown-items", {
      state: "attached",
      timeout: 10000,
    });

    // Pilih option by text (partial match supported for month/year formats)
    await this.page.evaluate((text) => {
      const options = document.querySelectorAll(
        ".ms-Dropdown-items .ms-Dropdown-item",
      );
      const target = Array.from(options).find((o) => {
        const itemText = o.textContent.trim().toLowerCase();
        const search = text.toString().toLowerCase();
        return itemText === search || itemText.startsWith(search);
      });
      if (target) target.click();
    }, text);
  }

  async waitForManualSteps() {
    console.log(
      "[INFO] Waiting for manual verification (captcha / phone / payment)",
    );

    await this.page.waitForTimeout(100000);
  }

  async clickUseThisAddressButton() {
    console.log("[STEP 10] Checking for address confirmation button...");

    try {
      const selector =
        "#pidlddc-button-addressUseButton, #pidlddc-button-userEnteredButton";

      const btn = this.page.locator(selector);
      await btn.waitFor({ state: "visible", timeout: 15000 });

      await this.randomMouseMove();
      await btn.click();

      console.log("[STEP 10] Address confirmation button clicked");
      await this.humanDelay(1000, 2000);
    } catch {
      console.log(
        "[STEP 10] Address confirmation button not found, skipping...",
      );
    }
  }

  async waitForDomainSuggestion() {
    console.log("[INFO] Waiting for domain suggestion to appear...");

    await this.page.waitForFunction(
      () => {
        const el = document.querySelector(
          'input.ms-TextField-field[maxlength="27"]',
        );
        return el && el.value && el.value.length > 5;
      },
      { timeout: 60000 },
    );

    console.log("[INFO] Domain suggestion detected");
  }

  async fillPassword() {
    console.log("[STEP 11] Filling password");

    await this.waitForPage('[data-testid="pwdField"]');

    await this.waitForDomainSuggestion();

    await this.page.waitForTimeout(1000);

    await this.page.locator('[data-testid="pwdField"]').click();
    await this.page
      .locator('[data-testid="pwdField"]')
      .pressSequentially(config.microsoftAccount.password, {
        delay: Math.floor(Math.random() * 40) + 60,
      });

    await this.humanDelay(500, 1000);

    await this.page.locator('[data-testid="cPwdField"]').click();
    await this.page
      .locator('[data-testid="cPwdField"]')
      .pressSequentially(config.microsoftAccount.password, {
        delay: Math.floor(Math.random() * 40) + 60,
      });

    await this.humanDelay(1000, 2000);

    const nextBtn = this.page.locator('[data-bi-id="AutoDomainNext"]');
    await nextBtn.waitFor({ state: "visible" });
    await this.randomMouseMove();
    await nextBtn.click();
  }

  async handleOptionalSignIn() {
    console.log("[STEP 11.5] Checking for optional Sign In prompt...");

    try {
      // Tunggu sebentar untuk lihat apakah button Sign In muncul
      const signInBtn = this.page.locator('button:has-text("Sign In")');
      const isVisible = await signInBtn
        .isVisible({ timeout: 10000 })
        .catch(() => false);

      if (isVisible) {
        console.log("Sign In button detected, clicking...");
        await this.randomMouseMove();
        await signInBtn.click();

        await this.humanDelay(2000, 4000);

        // Setelah click Sign In, biasanya ada prompt "Stay signed in?"
        const staySignedInBtn = this.page.locator("#idSIButton9");
        if (
          await staySignedInBtn.isVisible({ timeout: 15000 }).catch(() => false)
        ) {
          console.log("Stay signed in? prompt detected, clicking Yes...");
          await this.randomMouseMove();
          await staySignedInBtn.click();
        }
      } else {
        console.log("No Sign In button detected, proceeding...");
      }
    } catch (e) {
      console.log("Optional Sign In handler skipped or errored:", e.message);
    }
  }

  async goToPaymentPage() {
    console.log("[STEP 12] Waiting until payment page appears");

    await this.page
      .locator("text=Add payment method")
      .waitFor({ timeout: 100000 });

    console.log("Payment page detected");
  }

  async fillPaymentDetails() {
    console.log("[STEP 13] Filling VCC payment details");

    // Tunggu input card number muncul
    await this.page.waitForSelector("#accountToken", { timeout: 60000 });

    // Fill Card Number
    console.log("Typing card number...");
    await this.page.locator("#accountToken").click();
    await this.page
      .locator("#accountToken")
      .pressSequentially(config.payment.cardNumber, {
        delay: Math.floor(Math.random() * 30) + 50,
      });
    await this.humanDelay(800, 1500);

    // Fill CVV
    console.log("Typing CVV...");
    await this.page.locator("#cvvToken").click();
    await this.page.locator("#cvvToken").pressSequentially(config.payment.cvv, {
      delay: Math.floor(Math.random() * 40) + 70,
    });
    await this.humanDelay(1000, 2000);

    // Select Expiry Month
    console.log("Selecting expiry month:", config.payment.expMonth);
    await this.selectDropdownByText(
      "#input_expiryMonth",
      config.payment.expMonth,
    );
    await this.humanDelay(800, 1500);

    // Select Expiry Year
    console.log("Selecting expiry year:", config.payment.expYear);
    await this.selectDropdownByText(
      "#input_expiryYear",
      config.payment.expYear,
    );
    await this.humanDelay(1000, 2000);

    console.log("VCC details filled");
  }

  async clickSavePaymentButton() {
    console.log("[STEP 14] Clicking Save progress button");

    const saveBtn = this.page.locator('[data-bi-id="SignupSave"]');
    await saveBtn.waitFor({ state: "visible" });
    await this.randomMouseMove();
    await saveBtn.click();
  }

  // async clickStartTrialButton() {
  //   console.log("[STEP 15] Finalizing - Waiting for trial agreement checkbox...");

  //   // Tunggu sampai checkbox muncul
  //   try {
  //     const checkboxLocator = this.page.locator(".ms-Checkbox");
  //     await checkboxLocator.waitFor({ state: "visible", timeout: 120000 });

  //     const input = checkboxLocator.locator('input[type="checkbox"]');
  //     const isChecked = await input.isChecked();

  //     if (!isChecked) {
  //       console.log("Checkbox not checked, clicking label...");
  //       await this.randomMouseMove();

  //       // Klik label atau container lebih reliable untuk UI Microsoft
  //       await checkboxLocator.locator("label").click({ force: true });

  //       // Verifikasi apakah sudah ter-check
  //       await this.page.waitForTimeout(1000);
  //       if (!(await input.isChecked())) {
  //           console.log("Still not checked, trying direct input click...");
  //           await input.click({ force: true });
  //       }

  //       console.log("Agreement checkbox checked");
  //     } else {
  //       console.log("Agreement checkbox is already checked");
  //     }
  //   } catch (e) {
  //     console.warn("Agreement checkbox not found or timeout (120s):", e.message);
  //   }
  //   await this.humanDelay(1500, 3000);

  //   console.log("Waiting for Start Trial button...");

  //   const startTrialBtn = this.page.getByRole("button", { name: /Start/i });

  //   await startTrialBtn.waitFor({
  //     state: "visible",
  //     timeout: 60000,
  //   });

  //   // Tunggu sampai button tidak disabled
  //   await this.page.waitForFunction(
  //     () => {
  //       const btn = Array.from(document.querySelectorAll("button")).find((b) =>
  //         /start/i.test(b.innerText),
  //       );
  //       return (
  //         btn && !btn.disabled && btn.getAttribute("aria-disabled") !== "true"
  //       );
  //     },
  //     { timeout: 20000 },
  //   );

  //   await this.randomMouseMove();

  //   await startTrialBtn.click();

  //   console.log("Start trial button clicked");
  // }
  async clickStartTrialButton() {
    console.log("[SAVE] Waiting for checklist checkbox...");

    // Wait for checkbox to appear by finding its label text
    try {
      // Find the container based on the agreement text to avoid relying on dynamic IDs/classes
      const checkboxContainer = this.page
        .locator(".ms-Checkbox")
        .filter({
          hasText: /authorize recurring payments|by checking the box/i,
        });
      const checkboxInput = checkboxContainer.locator('input[type="checkbox"]');

      await checkboxInput.waitFor({ state: "attached", timeout: 120000 });

      const isChecked = await checkboxInput.getAttribute("aria-checked");

      if (isChecked !== "true") {
        console.log("Checkbox not checked, attempting to check...");
        await this.randomMouseMove();

        // Try clicking the label inside the matched container
        const label = checkboxContainer.locator("label");
        const labelExists = await label.count();

        if (labelExists > 0) {
          await label.click({ force: true });
        } else {
          await checkboxInput.click({ force: true });
        }

        await this.page.waitForTimeout(1000);

        // Verify checked via aria-checked (Microsoft Fabric uses this, not native checked)
        const rechecked = await checkboxInput.getAttribute("aria-checked");
        if (rechecked !== "true") {
          console.log("Still not checked, trying JS dispatch...");
          await checkboxInput.evaluate((el) => {
            if (el) {
              el.click();
              el.dispatchEvent(new Event("change", { bubbles: true }));
            }
          });
        }

        console.log("Checklist checkbox checked");
      } else {
        console.log("Checkbox already checked");
      }
    } catch (e) {
      console.warn("Checkbox not found or error:", e.message);
    }

    await this.humanDelay(1000, 2000);

    console.log("Waiting for Save/Start Trial button to become enabled...");

    // Target the specific disabled button and wait for it to become enabled
    const saveBtn = this.page
      .locator("button.ms-Button--primary")
      .filter({ hasText: /start trial|save/i });

    await saveBtn.waitFor({ state: "visible", timeout: 30000 });

    // Poll until aria-disabled is removed
    await this.page.waitForFunction(
      () => {
        const btn = document.querySelector("button.ms-Button--primary");
        return (
          btn &&
          !btn.disabled &&
          btn.getAttribute("aria-disabled") !== "true" &&
          !btn.classList.contains("is-disabled")
        );
      },
      { timeout: 30000 },
    );

    await this.randomMouseMove();
    await this.humanDelay(500, 1000);

    await saveBtn.click({ force: true });
    console.log("Save/Start Trial button clicked");
  }

  async pauseForManualPayment() {
    console.log(
      "Please enter payment details manually. Automation will wait...",
    );

    await this.page.waitForTimeout(180000);
  }

  async checkForError() {
    const hasError = await this.page.evaluate(() => {
      const text = document.body.innerText;
      return (
        text.includes("Something went wrong") ||
        text.includes("Something happened")
      );
    });

    if (hasError) {
      console.log(
        "[ERROR] Error page detected, closing browser and deleting profile...",
      );
      await this.cleanup();
      return true;
    }

    return false;
  }

  async cleanup() {
    try {
      await this.browser.close();
    } catch (e) {
      console.error("Error closing browser:", e);
    }

    // Note: Profile deletion usually requires AdsPower API, not just fs.rmSync
    // if the profile is managed by the app.
    if (config.profilePath && fs.existsSync(config.profilePath)) {
      try {
        fs.rmSync(config.profilePath, { recursive: true, force: true });
        console.log(
          "[CLEANUP] Local profile folder deleted:",
          config.profilePath,
        );
      } catch (e) {
        console.warn("[CLEANUP] Could not delete profile folder:", e.message);
      }
    }
  }

  async run() {
    try {
      await this.connect();
      await this.humanDelay(1000, 3000);
      await this.openMicrosoftPage();
      if (await this.checkForError()) return;
      await this.humanDelay(2000, 5000);

      await this.clickTryButton();
      if (await this.checkForError()) return;
      await this.humanDelay(2000, 4000);

      await this.clickBuildCartNextButton();
      if (await this.checkForError()) return;
      await this.humanDelay(1500, 3000);

      await this.fillEmail();
      if (await this.checkForError()) return;
      await this.humanDelay(1000, 2500);

      await this.clickCollectEmailNextButton();
      if (await this.checkForError()) return;
      await this.humanDelay(2000, 4000);

      await this.clickConfirmEmailSetupAccountButton();
      if (await this.checkForError()) return;
      await this.humanDelay(2000, 4000);

      await this.fillBasicInfo();
      if (await this.checkForError()) return;
      await this.humanDelay(1500, 3500);

      await this.clickUseThisAddressButton();
      if (await this.checkForError()) return;
      await this.humanDelay(1500, 3000);

      await this.fillPassword();
      if (await this.checkForError()) return;
      await this.humanDelay(2000, 5000);

      // Cek apakah minta Sign In manual setelah isi password
      await this.handleOptionalSignIn();
      await this.humanDelay(2000, 5000);

      //   await this.waitForManualSteps();
      await this.goToPaymentPage();
      await this.humanDelay(2000, 4000);

      await this.fillPaymentDetails();
      await this.humanDelay(2000, 5000);

      await this.clickSavePaymentButton();

      // Tunggu page loading/proses simpan (bisa lama)
      console.log("[INFO] Waiting for page transition after Save...");
      try {
        await this.page.waitForSelector(
          '#pidlddc-button-userEnteredButton, input[type="checkbox"]',
          {
            state: "visible",
            timeout: 120000,
          },
        );
      } catch (e) {
        console.log(
          "[INFO] Timeout waiting for next step elements, checking manually...",
        );
      }

      // Selalu cek konfirmasi alamat (jika ada)
      await this.clickUseThisAddressButton();
      await this.humanDelay(1500, 3000);

      // Terakhir klik Start Trial
      await this.clickStartTrialButton();
      await this.humanDelay(5000, 10000);

      await this.pauseForManualPayment();

      console.log("Automation completed safely");
    } catch (error) {
      console.error("Automation error:", error);
      await this.cleanup();
    }
  }
}

module.exports = MicrosoftBot;
