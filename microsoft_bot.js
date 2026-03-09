const { chromium } = require("playwright-core");
const fs = require("fs");
const config = require("./config");

class MicrosoftBot {
  constructor(wsUrl, accountConfig) {
    this.wsUrl = wsUrl;
    this.browser = null;
    this.context = null;
    this.page = null;
    this.accountConfig = accountConfig; // Store the specific account configuration
  }

  async humanDelay(min = 100, max = 400) {
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

  getGenericLocator(keyword, elementType = "input") {
    // Cari elemen berdasarkan substring keyword yang case-insensitive di berbagai attribute
    return this.page
      .locator(
        `${elementType}[id*="${keyword}" i], ${elementType}[data-testid*="${keyword}" i], ${elementType}[data-bi-id*="${keyword}" i], ${elementType}[name*="${keyword}" i], ${elementType}[aria-label*="${keyword}" i]`,
      )
      .first();
  }

  getGenericButton(keyword) {
    return this.page
      .locator(
        `button[id*="${keyword}" i], button[data-testid*="${keyword}" i], button[data-bi-id*="${keyword}" i], a[data-bi-id*="${keyword}" i], button:has-text("${keyword}"), a:has-text("${keyword}")`,
      )
      .first();
  }

  async waitForPage(selector) {
    if (selector) {
      // Tunggu elemen spesifik muncul = page sudah siap
      await this.page.waitForSelector(selector, {
        state: "attached",
        timeout: 150000,
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

  // async clickTryButton() {
  //   console.log("[STEP 3] Clicking Try button");

  //   await this.waitForPage("#action-oc5f9e");

  //   // Tangkap new page sebelum click
  //   const [newPage] = await Promise.all([
  //     this.context.waitForEvent("page"),
  //     this.page.evaluate(() => {
  //       document.querySelector("#action-oc5f9e").click();
  //     }),
  //   ]);

  //   // Switch this.page ke tab baru
  //   await newPage.waitForLoadState("domcontentloaded");
  //   this.page = newPage;

  //   console.log("[STEP 3] Switched to new page:", this.page.url());
  // }

  async clickBuildCartNextButton() {
    console.log("[STEP 4] Clicking Next button");

    const nextBtn = this.getGenericButton("Next");
    await nextBtn.waitFor({ state: "visible", timeout: 60000 });
    await this.randomMouseMove();
    await nextBtn.click();
  }

  async fillEmail() {
    const email = this.accountConfig.microsoftAccount.email;

    console.log("[STEP 5] Filling email:", email);

    // Gunakan locator generik untuk email
    const emailInput = this.getGenericLocator("email");

    await emailInput.waitFor({ state: "visible", timeout: 30000 });

    await this.randomMouseMove();

    await emailInput.click();

    await emailInput.pressSequentially(email, {
      delay: Math.floor(Math.random() * 30) + 50,
    });

    await this.humanDelay(100, 300);
  }

  async clickCollectEmailNextButton() {
    console.log("[STEP 6] Clicking CollectEmail Next button");

    const nextBtn = this.getGenericButton("Next");
    await nextBtn.waitFor({ state: "visible", timeout: 60000 });
    await this.randomMouseMove();
    await nextBtn.click();

    // Tunggu verifikasi manual selesai, baru lanjut ke setup account
    console.log("[INFO] Waiting for email verification to complete...");

    // Tunggu button Setup muncul
    const setupBtn = this.getGenericButton("Setup");
    await setupBtn.waitFor({ state: "visible", timeout: 150000 });
    console.log("[INFO] Verification complete, setup account button detected");
  }

  async clickConfirmEmailSetupAccountButton() {
    console.log("[STEP 7] Clicking Setup Account button");

    const setupBtn = this.getGenericButton("Setup");
    await this.randomMouseMove();
    await setupBtn.click();
  }

  async fillBasicInfo() {
    console.log("[STEP 8] Filling basic info");

    // Tunggu field first name / nama awal muncul
    await this.getGenericLocator("first").waitFor({
      state: "visible",
      timeout: 30000,
    });

    // Fill semua text fields secara human-like
    const fields = [
      {
        locator: this.getGenericLocator("first"),
        value: this.accountConfig.microsoftAccount.firstName,
      },
      {
        locator: this.getGenericLocator("last"),
        value: this.accountConfig.microsoftAccount.lastName,
      },
      {
        locator: this.getGenericLocator("company"),
        value: this.accountConfig.microsoftAccount.companyName,
      },
      {
        locator: this.getGenericLocator("phone"),
        value: this.accountConfig.microsoftAccount.phone,
      },
      {
        locator: this.getGenericLocator("job"),
        value: this.accountConfig.microsoftAccount.jobTitle,
      },
    ];

    for (const field of fields) {
      await field.locator.click();
      await field.locator.pressSequentially(field.value, {
        delay: Math.floor(Math.random() * 40) + 60,
      });
      await this.humanDelay(400, 800);
    }

    const addressLocator = this.getGenericLocator("address");
    await addressLocator.click();
    await addressLocator.pressSequentially(
      this.accountConfig.microsoftAccount.address,
      {
        delay: Math.floor(Math.random() * 30) + 50,
      },
    );
    await this.humanDelay(100, 300);

    // Input City
    const cityLocator = this.getGenericLocator("city");
    await cityLocator.click();
    for (const char of this.accountConfig.microsoftAccount.city) {
      await this.page.keyboard.type(char, { delay: Math.random() * 50 + 50 });
    }
    await this.humanDelay(100, 300);

    // Postal code (optional)
    const postalLocator = this.page
      .locator(
        'input[id*="postal" i], input[id*="zip" i], input[data-testid*="postal" i], input[data-testid*="zip" i]',
      )
      .first();

    if (
      this.accountConfig.microsoftAccount.postalCode &&
      (await postalLocator.count()) > 0
    ) {
      try {
        await postalLocator.click();

        await postalLocator.pressSequentially(
          this.accountConfig.microsoftAccount.postalCode,
          {
            delay: Math.floor(Math.random() * 30) + 50,
          },
        );

        console.log("Postal code filled");
        await this.humanDelay(150, 400);
      } catch (err) {
        console.log("Postal code field found but could not fill, skipping...");
      }
    } else {
      console.log("Postal code not provided or field not found, skipping...");
    }

    // Pilih company size (random)
    await this.selectDropdownByText(
      'div[role="combobox"][id*="size" i], div[role="combobox"][data-testid*="size" i], select[id*="size" i]',
      this.accountConfig.microsoftAccount.companySize,
    );
    await this.humanDelay(600, 1200);

    // Pilih state Alabama (sesuai config)
    const regionInput = this.page
      .locator('input[id*="region" i], input[id*="state" i]')
      .first();

    if ((await regionInput.count()) > 0) {
      await regionInput.click();

      await regionInput.pressSequentially(
        this.accountConfig.microsoftAccount.state || "Alabama",
        { delay: Math.floor(Math.random() * 30) + 50 },
      );

      console.log("Region filled as text input");
    } else {
      // fallback kalau ternyata dropdown
      await this.selectDropdownByText(
        'div[role="combobox"][id*="region" i], div[role="combobox"][id*="state" i], select[id*="region" i]',
        this.accountConfig.microsoftAccount.state || "Alabama",
      );
    }

    await this.humanDelay(600, 1200);

    // Pilih No untuk website
    await this.selectDropdownByText(
      'div[role="combobox"][id*="website" i], div[role="combobox"][data-testid*="website" i], select[id*="website" i]',
      "No",
    );
    await this.humanDelay(600, 1200);

    // Check partner checkbox
    try {
      console.log("Checking partner checkbox...");

      // Selector utama (paling stabil)
      let partnerCheckbox = this.page.locator("#partner-checkbox");

      // fallback jika id berubah
      if ((await partnerCheckbox.count()) === 0) {
        partnerCheckbox = this.page.locator(
          'input[type="checkbox"][aria-label*="share my information" i]',
        );
      }

      // fallback kedua
      if ((await partnerCheckbox.count()) === 0) {
        partnerCheckbox = this.page
          .locator('input[type="checkbox"]')
          .filter({ hasText: /partner|privacy/i })
          .first();
      }

      if ((await partnerCheckbox.count()) > 0) {
        await partnerCheckbox.waitFor({
          state: "visible",
          timeout: 10000,
        });

        const isChecked = await partnerCheckbox.isChecked();

        if (!isChecked) {
          await this.randomMouseMove();

          await partnerCheckbox.check({
            force: true,
          });

          console.log("Partner checkbox checked");
        } else {
          console.log("Partner checkbox already checked");
        }
      } else {
        console.log("Partner checkbox not found, skipping...");
      }
    } catch (err) {
      console.log("Checkbox error:", err.message);
    }

    await this.humanDelay(200, 500);

    // Click Next
    const nextBtn = this.getGenericButton("Next");
    await this.randomMouseMove();
    await nextBtn.click();
  }

  async selectRandomDropdown(selector) {
    // Open dropdown
    const dropdown = this.page.locator(selector).first();
    await dropdown.waitFor({ state: "visible", timeout: 15000 });
    await this.randomMouseMove();
    await dropdown.click({ force: true });

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
    const dropdown = this.page.locator(selector).first();
    await dropdown.waitFor({ state: "visible", timeout: 15000 });
    await this.randomMouseMove();
    await dropdown.click({ force: true });

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
      const btn = this.getGenericButton("Use this address");
      await btn.waitFor({ state: "visible", timeout: 15000 });

      await this.randomMouseMove();
      await btn.click();

      console.log("[STEP 10.5] Address confirmation button clicked");
      await this.humanDelay(200, 500);
    } catch {
      console.log(
        "[STEP 10.5] Address confirmation button not found, skipping...",
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

    // Menggunakan regex untuk membedakan password utama dan retype password
    const passwordLocator = this.page
      .locator(
        'input[type="password"]:not([id*="retype" i]):not([id*="confirm" i]):not([data-testid*="cpwd" i])',
      )
      .first();
    const confirmPasswordLocator = this.page
      .locator('input[type="password"]')
      .nth(1);

    await passwordLocator.waitFor({ state: "visible", timeout: 30000 });

    await this.randomMouseMove();
    await passwordLocator.click({ force: true }).catch(() => {});
    await passwordLocator.pressSequentially(
      this.accountConfig.microsoftAccount.password,
      {
        delay: Math.floor(Math.random() * 100) + 100, // Ketik pelan
      },
    );

    await this.humanDelay(100, 300);

    await confirmPasswordLocator.click({ force: true }).catch(() => {});
    await confirmPasswordLocator.pressSequentially(
      this.accountConfig.microsoftAccount.password,
      {
        delay: Math.floor(Math.random() * 40) + 60,
      },
    );

    await this.humanDelay(200, 500);

    const nextBtn = this.getGenericButton("Next");
    await nextBtn.waitFor({ state: "visible", timeout: 30000 });
    await this.randomMouseMove();
    await nextBtn.click();
  }

  async handleOptionalSignIn() {
    console.log("[STEP 11.5] Checking for optional Sign In prompt...");

    try {
      // Tunggu page load dasar
      await this.page.waitForLoadState("domcontentloaded");

      const signInBtn = this.getGenericButton("Sign In");

      // Retry hingga 3 kali untuk tombol Sign In
      let signInDetected = false;
      for (let attempt = 1; attempt <= 3; attempt++) {
        signInDetected = await signInBtn
          .waitFor({ state: "visible", timeout: 20000 })
          .then(() => true)
          .catch(() => false);

        if (signInDetected) break;

        console.log(`Sign In button not detected, retry ${attempt}...`);
        await this.humanDelay(200, 500); // delay sebelum retry
      }

      if (!signInDetected) {
        console.log("No Sign In button detected after retries, skipping...");
        return;
      }

      console.log("Sign In button detected, clicking...");
      await this.randomMouseMove();
      await signInBtn.click();
      await this.humanDelay(400, 800);

      // Tangani "Stay signed in?" prompt jika muncul
      const staySignedInBtn = this.page
        .locator(
          'button[id*="idSIButton" i], input[id*="idSIButton" i], button[type="submit"], input[type="submit"]',
        )
        .first();

      const staySignedInVisible = await staySignedInBtn
        .waitFor({ state: "visible", timeout: 10000 })
        .then(() => true)
        .catch(() => false);

      if (staySignedInVisible) {
        console.log("Stay signed in prompt detected, clicking...");
        await this.randomMouseMove();
        await staySignedInBtn.click();
        await this.humanDelay(200, 500);
      } else {
        console.log("No 'Stay signed in?' prompt detected, proceeding...");
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

    // Tunggu input card number muncul dengan locator lebih fleksibel
    const cardLocator = this.page
      .locator(
        'input[id*="accounttoken" i], input[id*="card" i], input[data-testid*="card" i]',
      )
      .first();
    await cardLocator.waitFor({ state: "visible", timeout: 60000 });

    // Fill Card Number
    console.log("Typing card number...");
    await cardLocator.click();
    await cardLocator.pressSequentially(this.accountConfig.payment.cardNumber, {
      delay: Math.floor(Math.random() * 30) + 50,
    });
    await this.humanDelay(150, 400);

    // Fill CVV
    console.log("Typing CVV...");
    const cvvLocator = this.page
      .locator(
        'input[id*="cvv" i], input[data-testid*="cvv" i], input[name*="cvv" i]',
      )
      .first();
    await cvvLocator.click();
    await cvvLocator.pressSequentially(this.accountConfig.payment.cvv, {
      delay: Math.floor(Math.random() * 30) + 50,
    });
    await this.humanDelay(150, 400);

    // Select Expiry Month
    let expMonth = this.accountConfig.payment.expMonth.toString();
    if (expMonth.length === 1) {
      expMonth = "0" + expMonth;
    }
    console.log("Selecting expiry month:", expMonth);
    // Untuk dropdown, cari custom selector atau fallback
    const expMonthLocatorString =
      'div[role="combobox"][id*="month" i], div[role="combobox"][data-testid*="month" i], select[id*="month" i]';
    await this.selectDropdownByText(expMonthLocatorString, expMonth);
    await this.humanDelay(150, 400);

    // Select Expiry Year
    console.log("Selecting expiry year:", this.accountConfig.payment.expYear);
    const expYearLocatorString =
      'div[role="combobox"][id*="year" i], div[role="combobox"][data-testid*="year" i], select[id*="year" i]';
    await this.selectDropdownByText(
      expYearLocatorString,
      this.accountConfig.payment.expYear,
    );
    await this.humanDelay(200, 500);

    console.log("VCC details filled");
  }

  async clickSavePaymentButton() {
    console.log("[STEP 14] Clicking Save progress button");

    const saveBtn = this.getGenericButton("Save");

    await saveBtn.waitFor({ state: "visible", timeout: 60000 });

    await this.randomMouseMove();
    await saveBtn.click();

    console.log("[INFO] Waiting for payment response...");

    // Tunggu sedikit agar DOM update
    await this.page.waitForTimeout(3000);

    const useAddressBtn = this.page
      .locator('button:has-text("Use this address")')
      .first();

    const exists = await useAddressBtn.isVisible().catch(() => false);

    if (exists) {
      console.log("[INFO] Address confirmation detected");

      await this.randomMouseMove();
      await useAddressBtn.click();

      console.log("[INFO] Use this address clicked");
    } else {
      console.log("[INFO] No address confirmation needed, skipping...");
    }
  }

  async clickStartTrialButton() {
    console.log("[SAVE] Checking if checklist checkbox exists...");

    try {
      const checkboxContainer = this.page.locator(".ms-Checkbox").filter({
        hasText: /authorize recurring payments|by checking the box/i,
      });

      const checkboxInput = checkboxContainer.locator('input[type="checkbox"]');

      const exists = (await checkboxInput.count()) > 0;

      if (!exists) {
        console.log("Checklist checkbox not found, skipping...");
      } else {
        console.log("Checklist checkbox found");

        await checkboxInput.waitFor({ state: "visible", timeout: 10000 });

        const isChecked = await checkboxInput.getAttribute("aria-checked");

        if (isChecked !== "true") {
          console.log("Checkbox not checked, checking...");

          await this.randomMouseMove();

          const label = checkboxContainer.locator("label");

          if ((await label.count()) > 0) {
            await label.click({ force: true });
          } else {
            await checkboxInput.click({ force: true });
          }

          await this.page.waitForTimeout(500);

          const rechecked = await checkboxInput.getAttribute("aria-checked");

          if (rechecked !== "true") {
            await checkboxInput.evaluate((el) => {
              el.click();
              el.dispatchEvent(new Event("change", { bubbles: true }));
            });
          }

          console.log("Checklist checkbox checked");
        } else {
          console.log("Checkbox already checked");
        }
      }
    } catch (e) {
      console.log("Checkbox handling skipped:", e.message);
    }

    await this.humanDelay(200, 500);

    console.log("Waiting for Save/Start Trial button to become enabled...");

    // Target the specific disabled button and wait for it to become enabled
    const saveBtn = this.page
      .locator('button[class*="primary" i], button[data-bi-id*="save" i]')
      .filter({ hasText: /start trial|save/i })
      .first();

    await saveBtn.waitFor({ state: "visible", timeout: 30000 });

    // Poll until aria-disabled is removed
    await this.page.waitForFunction(
      () => {
        const btn = Array.from(document.querySelectorAll("button")).find((b) =>
          /start trial|save/i.test(b.textContent),
        );
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
    await this.humanDelay(100, 300);

    await saveBtn.click({ force: true });
    console.log("Save/Start Trial button clicked");
  }

  async clickPostTrialNextButton() {
    console.log("[STEP 15] Clicking Next button after Start Trial");

    const nextBtn = this.getGenericButton("Next");
    await nextBtn.waitFor({ state: "visible", timeout: 120000 });
    await this.randomMouseMove();
    await this.humanDelay(300, 600);
    await nextBtn.click();

    console.log(
      "[STEP 15] Next button clicked, waiting for confirmation page...",
    );
    await this.waitForPage();
  }

  async extractDomainEmail() {
    console.log("[STEP 16] Extracting domain email from confirmation page...");

    await this.humanDelay(2000, 4000);

    const emailLocator = this.page.locator("#displayName");

    const found = await emailLocator
      .waitFor({ state: "visible", timeout: 30000 })
      .then(() => true)
      .catch(() => false);

    if (!found) {
      console.warn("[STEP 16] displayName not found");
      return { domainEmail: "", domainPassword: "" };
    }

    const rawText = (await emailLocator.textContent())?.trim() || "";

    const emailMatch = rawText.match(/[\w.+-]+@[\w.-]+\.onmicrosoft\.com/i);
    const domainEmail = emailMatch ? emailMatch[0] : rawText;

    // ambil sebelum .onmicrosoft.com
    const domainPassword = domainEmail.replace(/\.onmicrosoft\.com$/i, "");

    console.log("[STEP 16] Domain email:", domainEmail);
    console.log("[STEP 16] Extracted password:", domainPassword);

    return { domainEmail, domainPassword };
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
    let currentStep = "Initializing";
    try {
      currentStep = "Connecting to browser";
      await this.connect();
      await this.humanDelay(1000, 3000);

      currentStep = "Opening Microsoft page";
      await this.openMicrosoftPage();
      if (await this.checkForError())
        throw new Error("Microsoft error page detected during initial load");
      await this.humanDelay(400, 800);

      // currentStep = "Clicking Try button";
      // await this.clickTryButton();
      // if (await this.checkForError())
      //   throw new Error("Microsoft error page detected after Try button");
      // await this.humanDelay(400, 800);

      currentStep = "Building cart";
      await this.clickBuildCartNextButton();
      if (await this.checkForError())
        throw new Error("Microsoft error page detected during building cart");
      await this.humanDelay(300, 600);

      currentStep = "Filling email";
      await this.fillEmail();
      if (await this.checkForError())
        throw new Error("Microsoft error page detected after filling email");
      await this.humanDelay(1000, 2500);

      currentStep = "Confirming email email";
      await this.clickCollectEmailNextButton();
      if (await this.checkForError())
        throw new Error("Microsoft error page detected after confirming email");
      await this.humanDelay(400, 800);

      currentStep = "Setup account button";
      await this.clickConfirmEmailSetupAccountButton();
      if (await this.checkForError())
        throw new Error("Microsoft error page detected after setup button");
      await this.humanDelay(400, 800);

      currentStep = "Filling basic info";
      await this.fillBasicInfo();
      if (await this.checkForError())
        throw new Error("Microsoft error page detected after basic info");
      await this.humanDelay(1500, 3500);

      currentStep = "Confirming address (Stage 1)";
      await this.clickUseThisAddressButton();
      if (await this.checkForError())
        throw new Error(
          "Microsoft error page detected after address confirmation",
        );
      await this.humanDelay(300, 600);

      currentStep = "Filling password";
      await this.fillPassword();
      if (await this.checkForError())
        throw new Error("Microsoft error page detected after filling password");
      await this.humanDelay(400, 800);

      currentStep = "Handling manual sign in (if any)";
      await this.handleOptionalSignIn();
      await this.humanDelay(400, 800);

      currentStep = "Going to payment page";
      await this.goToPaymentPage();
      await this.humanDelay(400, 800);

      currentStep = "Filling VCC payment details";
      await this.fillPaymentDetails();
      await this.humanDelay(400, 800);

      currentStep = "Saving payment";
      await this.clickSavePaymentButton();

      currentStep = "Confirming address (Stage 2)";
      await this.clickUseThisAddressButton();
      await this.humanDelay(300, 600);

      currentStep = "Clicking Start Trial";
      await this.clickStartTrialButton();
      await this.humanDelay(800, 1500);

      currentStep = "Finishing trial setup";
      await this.clickPostTrialNextButton();
      await this.humanDelay(800, 1500);

      currentStep = "Extracting domain email";
      const { domainEmail, domainPassword } = await this.extractDomainEmail();

      console.log("Automation completed safely");
      return { success: true, domainEmail, domainPassword };
    } catch (error) {
      console.error(`Automation error at step [${currentStep}]:`, error);
      return {
        success: false,
        domainEmail: "",
        domainPassword: "",
        error: `Step: ${currentStep} - Error: ${error.message}`,
      };
    }
  }
}

module.exports = MicrosoftBot;
