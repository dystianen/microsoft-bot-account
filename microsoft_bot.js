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

  async clickButtonWithPossibleNames(names, options = {}) {
    const {
      timeout = 60000,
      waitVisible = true,
      clickOptions = { force: true },
    } = options;
    let targetBtn = null;
    let foundName = null;

    for (const name of names) {
      try {
        const btn = this.getGenericButton(name);
        if (waitVisible) {
          await btn.waitFor({ state: "visible", timeout: 3000 });
        } else if (await btn.count()) {
          targetBtn = btn;
          foundName = name;
          break;
        }

        if (await btn.isVisible()) {
          targetBtn = btn;
          foundName = name;
          break;
        }
      } catch {}
    }

    if (!targetBtn) {
      // Fallback search using filter if no specific button found
      const regex = new RegExp(names.join("|"), "i");
      targetBtn = this.page
        .locator("button, a")
        .filter({ hasText: regex })
        .first();

      if (!(await targetBtn.count())) {
        throw new Error(`Button not found. Tried: ${names.join(", ")}`);
      }
    }

    console.log(`[INFO] Clicking button: ${foundName || "regex match"}`);
    await this.randomMouseMove();
    await this.humanDelay(200, 500);

    try {
      await targetBtn.click({ timeout: 5000, ...clickOptions });
    } catch {
      console.log("Playwright click blocked, fallback to JS click...");
      await targetBtn.evaluate((el) => el.click());
    }

    return foundName;
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

    this.browser = await Promise.race([
      chromium.connectOverCDP(this.wsUrl),
      new Promise((_, reject) =>
        setTimeout(
          () => reject(new Error("CDP connection timeout after 30s")),
          30000,
        ),
      ),
    ]);

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

  async clickBuildCartNextButton() {
    console.log("[STEP 4] Clicking Next button");
    await this.clickButtonWithPossibleNames([
      "Next",
      "Selanjutnya",
      "Continue",
    ]);
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
    console.log("[STEP 6] Clicking Next button for email");
    await this.clickButtonWithPossibleNames(["Next", "Selanjutnya"]);

    console.log("[INFO] Waiting for email verification...");
    const setupNames = ["Setup", "Atur", "Set up"];

    const start = Date.now();
    const interval = setInterval(() => {
      console.log(
        `[INFO] Still waiting... ${Math.round((Date.now() - start) / 1000)}s`,
      );
    }, 15000);

    try {
      // Find which setup name works
      let setupBtn = null;
      for (const name of setupNames) {
        const btn = this.getGenericButton(name);
        if (await btn.count()) {
          setupBtn = btn;
          break;
        }
      }

      if (!setupBtn) {
        setupBtn = this.getGenericButton(setupNames[0]);
      }

      await this.waitWithCheck(setupBtn, 150000);
      this._setupBtnReady = true;
    } finally {
      clearInterval(interval);
    }
  }

  async clickConfirmEmailSetupAccountButton() {
    console.log("[STEP 7] Clicking Setup Account button");
    const setupNames = ["Setup Account", "Setup", "Atur Akun", "Mulai"];

    await this.clickButtonWithPossibleNames(setupNames);
    this._setupBtnReady = false; // reset flag
  }

  async fillBasicInfo() {
    console.log("[STEP 8] Filling basic info");

    // Tunggu field first name / nama awal muncul
    await this.waitWithCheck(this.getGenericLocator("first"), 30000);

    const fieldDefs = [
      {
        keyword: "first",
        value: this.accountConfig.microsoftAccount.firstName,
      },
      { keyword: "last", value: this.accountConfig.microsoftAccount.lastName },
      {
        keyword: "company",
        value: this.accountConfig.microsoftAccount.companyName,
      },
      { keyword: "phone", value: this.accountConfig.microsoftAccount.phone },
      { keyword: "job", value: this.accountConfig.microsoftAccount.jobTitle },
    ];

    for (const { keyword, value } of fieldDefs) {
      const locator = this.getGenericLocator(keyword);
      await locator.waitFor({ state: "visible", timeout: 15000 });
      await locator.click();
      await locator.pressSequentially(value, {
        delay: Math.floor(Math.random() * 40) + 60,
      });
      await this.humanDelay(400, 800);
    }

    const addressLocator = this.getGenericLocator("address");
    await addressLocator.waitFor({ state: "visible", timeout: 15000 });
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
    await cityLocator.waitFor({ state: "visible", timeout: 15000 });
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

    const regionIsInput = await regionInput
      .waitFor({ state: "visible", timeout: 8000 })
      .then(() => true)
      .catch(() => false);

    if (regionIsInput) {
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
    const possibleTexts = ["No", "Tidak", "Tidak ada"];

    for (const text of possibleTexts) {
      try {
        await this.selectDropdownByText(
          'div[role="combobox"][id*="website" i], div[role="combobox"][data-testid*="website" i], select[id*="website" i]',
          text,
        );
        console.log(`Website dropdown selected: ${text}`);
        break;
      } catch {}
    }
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
    await this.clickButtonWithPossibleNames([
      "Next",
      "Selanjutnya",
      "Continue",
    ]);
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
    await this.page.evaluate((textToSelect) => {
      const options = document.querySelectorAll(
        ".ms-Dropdown-items .ms-Dropdown-item",
      );
      const target = Array.from(options).find((o) => {
        if (!o || !o.textContent) return false;
        const itemText = o.textContent.trim().toLowerCase();
        const search = (textToSelect || "").toString().toLowerCase();
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

    const possibleNames = [
      "Use this address",
      "Use address",
      "Gunakan alamat ini",
    ];

    const clickedName = await this.clickButtonWithPossibleNames(possibleNames);

    // Jika bahasa Indonesia → pilih radio pertama (ini khusus logic address)
    if (clickedName === "Gunakan alamat ini") {
      console.log(
        "[STEP 10.1] Indonesian detected, ensure radio is selected if needed",
      );
      // Logic radio bisa ditambahkan jika clickButtonWithPossibleNames tidak cukup
      // Tapi biasanya button klik sudah cukup jika radio auto-selected atau tidak wajib
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

    // Extract domain email prefix from input field (e.g. PortlandDesignStudio156)
    try {
      const domainInput = this.page
        .locator('input.ms-TextField-field[maxlength="27"]')
        .first();
      await domainInput.waitFor({ state: "visible", timeout: 15000 });

      // Wait for suggestion to populate if it's empty
      await this.page
        .waitForFunction(
          (el) => el && el.value && el.value.length > 3,
          await domainInput.elementHandle(),
          { timeout: 15000 },
        )
        .catch(() => {});

      const prefix = await domainInput.inputValue();
      if (prefix) {
        this.extractedDomainEmail = `${prefix}.onmicrosoft.com`;
        this.extractedDomainPassword =
          this.accountConfig.microsoftAccount.password;
        console.log(
          `[INFO] Extracted Domain Email: ${this.extractedDomainEmail}`,
        );
      }
    } catch (e) {
      console.log(
        "[WARN] Could not extract domain prefix in fillPassword step:",
        e.message,
      );
    }

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

    const nextBtnNames = ["Next", "Selanjutnya", "Finish", "Selesai"];
    await this.clickButtonWithPossibleNames(nextBtnNames);
  }

  async handleOptionalSignIn() {
    console.log("[STEP 11.5] Checking for optional Sign In prompt...");

    try {
      await this.page.waitForLoadState("domcontentloaded");

      const signInBtn = this.getGenericButton("Sign In");

      let signInDetected = false;
      for (let attempt = 1; attempt <= 3; attempt++) {
        // Check for error page before each attempt
        if (await this.checkForError()) {
          throw new Error(
            "MICROSOFT_ERROR_PAGE: Terdeteksi saat pengecekan Sign In.",
          );
        }

        signInDetected = await signInBtn
          .waitFor({ state: "visible", timeout: 10000 })
          .then(() => true)
          .catch(() => false);

        if (signInDetected) break;

        console.log(`Sign In button not detected, retry ${attempt}...`);
        await this.humanDelay(200, 500);
      }

      if (!signInDetected) {
        console.log("No Sign In button detected after retries, skipping...");
        return;
      }

      console.log("Sign In button detected, clicking...");

      await this.randomMouseMove();

      // Handle popup window
      const [popup] = await Promise.all([
        this.page.waitForEvent("popup").catch(() => null),
        signInBtn.click(),
      ]);

      if (!popup) {
        console.log("No popup detected after Sign In click.");
        return;
      }

      console.log("Popup window detected");

      await popup.waitForLoadState("domcontentloaded");

      const yesBtn = popup.locator(
        'button:has-text("Yes"), input[value="Yes"], #idSIButton9',
      );

      const yesVisible = await yesBtn
        .waitFor({ state: "visible", timeout: 15000 })
        .then(() => true)
        .catch(() => false);

      if (yesVisible) {
        console.log("Stay signed in prompt detected, clicking Yes...");
        await this.randomMouseMove();
        await yesBtn.click();
      } else {
        console.log("No 'Stay signed in?' prompt detected in popup.");
      }

      await popup.waitForLoadState("networkidle").catch(() => {});

      console.log("Sign In popup handled successfully");
    } catch (e) {
      if (e.message.includes("MICROSOFT_ERROR_PAGE")) throw e;
      console.log("Optional Sign In handler skipped or errored:", e.message);
    }
  }

  async goToPaymentPage() {
    console.log("[STEP 12] Waiting until payment page appears");

    await this.page
      .locator("text=Add payment method")
      .waitFor({ timeout: 100000 })
      .catch(async () => {
        if (await this.checkForError())
          throw new Error(
            "MICROSOFT_ERROR_PAGE: Terdeteksi saat menunggu halaman pembayaran.",
          );
        throw new Error("Timeout waiting for payment page");
      });

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
    const saveNames = ["Save", "Simpan", "Next", "Selanjutnya"];
    await this.clickButtonWithPossibleNames(saveNames, { timeout: 60000 });

    console.log("[INFO] Waiting for payment response...");

    const TIMEOUT = 45000;
    let resolved = false;

    const makeWatcher = (promise, label) =>
      promise
        .then((v) => {
          resolved = true;
          return label;
        })
        .catch(() => null);

    // Error loop yang bisa reject race
    const errorWatcher = new Promise(async (resolve, reject) => {
      const deadline = Date.now() + TIMEOUT;
      while (!resolved && Date.now() < deadline) {
        await new Promise((r) => setTimeout(r, 2000));
        if (resolved) break;
        if (await this.checkForError()) {
          return reject(
            new Error(
              "MICROSOFT_ERROR_PAGE: Terdeteksi saat proses Save Payment.",
            ),
          );
        }
      }
      resolve(null);
    });

    const result = await Promise.race([
      makeWatcher(
        this.page.waitForSelector('span[data-automation-id="error-message"]', {
          state: "visible",
          timeout: TIMEOUT,
        }),
        "error",
      ),
      makeWatcher(
        this.page.waitForSelector('button:has-text("Use this address")', {
          state: "visible",
          timeout: TIMEOUT,
        }),
        "address",
      ),
      makeWatcher(
        this.page.waitForFunction(
          () =>
            window.location.href.includes("billing") ||
            document.body.innerText.includes("Check your info"),
          { timeout: TIMEOUT },
        ),
        "success",
      ),
      errorWatcher, // ← ini yang bisa reject
    ]);

    resolved = true; // pastikan error loop berhenti
    console.log(`[DEBUG] Payment result: ${result}`);

    if (result === "error") {
      const msg = await this.page
        .locator('span[data-automation-id="error-message"]')
        .first()
        .textContent()
        .catch(() => "Unknown payment error");
      throw new Error(`PAYMENT_DECLINED: ${msg?.trim()}`);
    }

    if (result === "address") {
      await this.page
        .locator('button:has-text("Use this address")')
        .click()
        .catch(() => {});
      await this.humanDelay(1000, 2000);
    }

    if (result === null) {
      console.warn(
        "[WARN] Payment result timeout — tidak ada sinyal jelas dari halaman",
      );
    }

    console.log("[INFO] Payment step finished");
  }

  async clickPostTrialNextButton() {
    console.log("[STEP 15] Clicking final Next/Get Started button after Trial");

    const possibleNames = [
      "Next",
      "Selanjutnya",
      "Get started",
      "Get Started",
      "Mulai",
    ];

    await this.page.evaluate(() => {
      document
        .querySelectorAll('[data-testid="spinner"], .css-100, .ms-Spinner')
        .forEach((el) => el.remove());
    });

    await this.clickButtonWithPossibleNames(possibleNames);
    await this.waitForPage();
  }

  async clickStartTrialButton() {
    console.log("[STEP 14] Clicking Start Trial button");
    const possibleNames = [
      "Start trial",
      "Mulai uji coba",
      "Try now",
      "Coba sekarang",
      "Start",
    ];
    await this.clickButtonWithPossibleNames(possibleNames);
  }

  async extractDomainEmail() {
    console.log("[STEP 16] Finalizing account data...");

    // Use pre-extracted data from fillPassword step
    if (this.extractedDomainEmail && this.extractedDomainPassword) {
      console.log(
        "[STEP 16] Using pre-extracted data:",
        this.extractedDomainEmail,
      );
      return {
        domainEmail: this.extractedDomainEmail,
        domainPassword: this.extractedDomainPassword,
      };
    }

    // Fallback if not found earlier (optional success page check)
    const emailLocator = this.page.locator("#displayName");
    const found = await emailLocator
      .waitFor({ state: "visible", timeout: 20000 })
      .then(() => true)
      .catch(() => false);

    if (!found) {
      // Jika tidak ketemu di akhir but we have extracted earlier, still return whatever we had
      return {
        domainEmail: this.extractedDomainEmail || "",
        domainPassword: this.extractedDomainPassword || "",
      };
    }

    const rawText = (await emailLocator.textContent())?.trim() || "";
    const emailRegex =
      /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.(onmicrosoft\.[a-z]{2,}|onmschina\.cn)/i;
    const emailMatch = rawText.match(emailRegex);

    const domainEmail = emailMatch?.[0] || this.extractedDomainEmail || "";
    const domainPassword = this.accountConfig.microsoftAccount.password;

    console.log("[STEP 16] Final Domain Email:", domainEmail);
    return { domainEmail, domainPassword };
  }

  async checkForError() {
    try {
      // Periksa apakah teks error ini ada di element apapun di halaman
      const errorData = await this.page.evaluate(() => {
        const text = document.body.innerText;
        const hasError =
          (text.includes("Something went wrong") &&
            text.includes("Error Code")) ||
          text.includes("715-123280") ||
          (text.includes("Something happened") &&
            !text.includes("Something happened to be"));

        return {
          hasError,
          text: text.substring(0, 500), // Ambil sedikit cuplikan untuk log
        };
      });

      if (errorData.hasError) {
        console.log("[ERROR] Microsoft error page detected immediately.");
        return true;
      }
    } catch (err) {
      // Jika page sudah tertutup atau crash, anggap saja tidak ada error yang bisa dicek
    }

    return false;
  }

  async waitWithCheck(locator, timeout = 60000) {
    let done = false;
    let errorFound = null;
    let intervalId = null;

    const errorLoop = (async () => {
      const deadline = Date.now() + timeout;
      while (!done && Date.now() < deadline) {
        await new Promise((r) => setTimeout(r, 2000));
        if (done) break;
        if (await this.checkForError()) {
          errorFound = new Error(
            "MICROSOFT_ERROR_PAGE: Halaman error terdeteksi.",
          );
          done = true;
          return;
        }
      }
    })();

    const errorInterrupt = new Promise((_, reject) => {
      intervalId = setInterval(() => {
        if (errorFound) {
          clearInterval(intervalId);
          reject(errorFound);
        }
      }, 200);
    });

    try {
      await Promise.race([
        locator.waitFor({ state: "visible", timeout }),
        errorInterrupt,
      ]);
    } finally {
      done = true;
      clearInterval(intervalId); // ← selalu cleanup
      await errorLoop;
    }

    if (errorFound) throw errorFound;
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

  async runStep(name, fn, delay = null) {
    console.log(`[STEP] ${name}`);
    this._currentStep = name;
    await fn();
    if (await this.checkForError()) {
      throw new Error(
        `MICROSOFT_ERROR_PAGE: Terdeteksi setelah step "${name}"`,
      );
    }
    if (delay) await this.humanDelay(...delay);
  }

  async run() {
    this._currentStep = "Initializing";
    try {
      await this.runStep(
        "Connecting to browser",
        () => this.connect(),
        [1000, 3000],
      );
      await this.runStep(
        "Opening Microsoft page",
        () => this.openMicrosoftPage(),
        [400, 800],
      );
      await this.runStep(
        "Building cart",
        () => this.clickBuildCartNextButton(),
        [300, 600],
      );
      await this.runStep("Filling email", () => this.fillEmail(), [1000, 2500]);
      await this.runStep(
        "Confirming email",
        () => this.clickCollectEmailNextButton(),
        [400, 800],
      );
      await this.runStep(
        "Setup account button",
        () => this.clickConfirmEmailSetupAccountButton(),
        [400, 800],
      );
      await this.runStep(
        "Filling basic info",
        () => this.fillBasicInfo(),
        [1500, 3500],
      );
      await this.runStep(
        "Confirming address (Stage 1)",
        () => this.clickUseThisAddressButton(),
        [300, 600],
      );
      await this.runStep(
        "Filling password",
        () => this.fillPassword(),
        [400, 800],
      );
      await this.runStep(
        "Handling sign in",
        () => this.handleOptionalSignIn(),
        [400, 800],
      );
      await this.runStep(
        "Going to payment page",
        () => this.goToPaymentPage(),
        [400, 800],
      );
      await this.runStep(
        "Filling payment details",
        () => this.fillPaymentDetails(),
        [400, 800],
      );
      await this.runStep("Saving payment", () => this.clickSavePaymentButton());

      // Stage 2 address + trial tidak perlu checkForError karena sudah ada internal check
      this._currentStep = "Confirming address (Stage 2)";
      await this.clickUseThisAddressButton();
      await this.humanDelay(300, 600);

      this._currentStep = "Clicking Start Trial";
      await this.clickStartTrialButton();
      await this.humanDelay(800, 1500);

      this._currentStep = "Finishing trial setup";
      await this.clickPostTrialNextButton();
      await this.humanDelay(800, 1500);

      this._currentStep = "Extracting domain email";
      const { domainEmail, domainPassword } = await this.extractDomainEmail();

      console.log("Automation completed successfully");
      return { success: true, domainEmail, domainPassword };
    } catch (error) {
      const step = this._currentStep;
      console.error(`Automation error at step [${step}]:`, error);
      return {
        success: false,
        domainEmail: "",
        domainPassword: "",
        error: `[${new Date().toISOString()}] Step: ${step} - ${error.message}`,
      };
    }
  }
}

module.exports = MicrosoftBot;
