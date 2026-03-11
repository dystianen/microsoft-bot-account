const { chromium } = require("playwright-core");
const fs = require("fs");
const config = require("./config");

const SPINNER_SELECTOR = '[data-testid="spinner"], .ms-Spinner, [class*="spinner" i]';

// Safety net — sangat besar, hanya untuk mencegah hang selamanya
const HARD_TIMEOUT = 1.5 * 60 * 1000; // 1 menit 30 detik

class MicrosoftBot {
  constructor(wsUrl, accountConfig) {
    this.wsUrl = wsUrl;
    this.browser = null;
    this.context = null;
    this.page = null;
    this.accountConfig = accountConfig;
  }

  // ─── Core helpers ────────────────────────────────────────────────────────────

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

  async runWithMonitor(promise, timeout = HARD_TIMEOUT) {
    let isDone = false;
    let errorMsg = null;

    const checkLoop = async () => {
      while (!isDone) {
        await this.page.waitForTimeout(2000).catch(() => { isDone = true; });
        if (isDone) break;

        const detectedError = await this.checkForError();
        if (detectedError) {
          errorMsg = detectedError;
          isDone = true;
          break;
        }
      }
    };

    const result = await Promise.race([
      promise,
      checkLoop(),
    ]).finally(() => {
      isDone = true;
    });

    if (errorMsg) {
      throw new Error(`MICROSOFT_ERROR: ${errorMsg}`);
    }

    return result;
  }

  async waitForSpinnerGone(extraDelay = 0) {
    const spinner = this.page.locator(SPINNER_SELECTOR).first();
    const spinnerVisible = await spinner.isVisible().catch(() => false);

    if (spinnerVisible) {
      console.log("[WAIT] Spinner detected, waiting until hidden...");
      try {
        await this.runWithMonitor(
          spinner.waitFor({ state: "hidden", timeout: HARD_TIMEOUT })
        );
      } catch (e) {
        if (e.message.includes("MICROSOFT_ERROR")) throw e;
        console.log("[WAIT] Spinner still visible or check failed, continuing...");
      }
      console.log("[WAIT] Spinner gone.");
    }

    const postSpinnerError = await this.checkForError();
    if (postSpinnerError) {
      throw new Error(`MICROSOFT_ERROR: ${postSpinnerError} (Detected after spinner)`);
    }

    if (extraDelay > 0) {
      await this.humanDelay(extraDelay, extraDelay + 300);
    }
  }

  async waitForVisible(locator) {
    await this.waitForSpinnerGone();
    await this.runWithMonitor(locator.waitFor({ state: "visible", timeout: HARD_TIMEOUT }));
  }

  async clickButtonWithPossibleNames(names) {
    await this.waitForSpinnerGone();

    const pattern = new RegExp(
      names
        .map(n => n.replace(/\s+/g, "\\s*"))
        .join("|"),
      "i"
    );

    const button = this.page.getByRole("button", { name: pattern }).first();

    await this.runWithMonitor(button.waitFor({ state: "visible", timeout: HARD_TIMEOUT }));

    await this.randomMouseMove();
    await this.humanDelay(200, 500);

    try {
      await button.click({ timeout: 8000, force: true });
    } catch {
      console.log("[INFO] Playwright click blocked, fallback to JS click...");
      await button.evaluate(el => el.click());
    }

    const clickedText = await button.textContent().catch(() => "unknown");
    console.log(`[INFO] Clicked: "${clickedText?.trim()}"`);
  }

  getGenericLocator(keyword, elementType = "input") {
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

  async selectDropdownByText(selector, text) {
    await this.waitForSpinnerGone();

    // Pastikan dropdown sebelumnya sudah tertutup
    await this.page.waitForSelector(".ms-Dropdown-items", {
      state: "detached",
      timeout: 5000,
    }).catch(() => { });

    const dropdown = this.page.locator(selector).first();

    await this.runWithMonitor(dropdown.waitFor({ state: "visible", timeout: HARD_TIMEOUT }));

    await dropdown.scrollIntoViewIfNeeded();
    await this.randomMouseMove();

    // buka dropdown
    await dropdown.click();

    // tunggu dropdown container muncul
    const dropdownItems = this.page.locator(".ms-Dropdown-items");
    await this.runWithMonitor(dropdownItems.waitFor({
      state: "visible",
      timeout: HARD_TIMEOUT,
    }));

    // tunggu option muncul
    const options = this.page.locator(".ms-Dropdown-item");
    await this.runWithMonitor(options.first().waitFor({
      state: "visible",
      timeout: HARD_TIMEOUT,
    }));

    // support text array atau string
    const searchList = Array.isArray(text)
      ? text.map(t => (t || "").toString().trim())
      : [(text || "").toString().trim()];

    let targetOption = null;

    for (const search of searchList) {
      const option = this.page
        .locator(".ms-Dropdown-item", { hasText: search })
        .first();

      if (await option.count()) {
        targetOption = option;
        break;
      }
    }

    if (!targetOption) {
      console.warn(`[DROPDOWN] Option not found for: ${text}`);
      return false;
    }

    const displayText = await targetOption.textContent().catch(() => text);
    console.log(`[DROPDOWN] Clicking: "${displayText?.trim()}"`);

    await targetOption.scrollIntoViewIfNeeded();
    await targetOption.click();

    // tunggu dropdown tertutup
    await this.page.waitForSelector(".ms-Dropdown-items", {
      state: "detached",
      timeout: 5000,
    }).catch(() => { });

    return true;
  }

  async waitForPage(selector) {
    await this.waitForSpinnerGone();
    if (selector) {
      await this.runWithMonitor(this.page.waitForSelector(selector, {
        state: "attached",
        timeout: HARD_TIMEOUT,
      }));
    } else {
      await this.runWithMonitor(this.page.waitForLoadState("domcontentloaded", {
        timeout: HARD_TIMEOUT,
      }));
    }
  }

  // ─── Steps ───────────────────────────────────────────────────────────────────

  async connect() {
    console.log("[STEP 1] Connecting to browser");

    this.browser = await Promise.race([
      chromium.connectOverCDP(this.wsUrl),
      new Promise((_, reject) =>
        setTimeout(() => reject(new Error("CDP connection timeout after 30s")), 30000),
      ),
    ]);

    const contexts = this.browser.contexts();
    this.context =
      contexts.length > 0 ? contexts[0] : await this.browser.newContext();

    const pages = this.context.pages();
    this.page = pages.length > 0 ? pages[0] : await this.context.newPage();
    this.profileId = this.wsUrl.split("/").pop();
  }

  async openMicrosoftPage() {
    console.log("[STEP 2] Opening Microsoft page");

    const url = this.accountConfig.microsoftUrl || config.microsoftUrl;
    await this.page.goto(url, {
      waitUntil: "domcontentloaded",
      timeout: HARD_TIMEOUT,
    });

    await this.waitForSpinnerGone();
  }

  async clickBuildCartNextButton() {
    console.log("[STEP 4] Clicking Next button");
    await this.clickButtonWithPossibleNames(["Next", "Selanjutnya", "Continue", "Berikutnya"]);
  }

  async fillEmail() {
    const email = this.accountConfig.microsoftAccount.email;
    console.log("[STEP 5] Filling email:", email);

    const emailInput = this.getGenericLocator("email");
    await this.waitForVisible(emailInput);
    await this.randomMouseMove();
    await emailInput.click();
    await emailInput.pressSequentially(email, {
      delay: Math.floor(Math.random() * 30) + 50,
    });
    await this.humanDelay(100, 300);
  }

  async clickCollectEmailNextButton() {
    console.log("[STEP 6] Clicking Next for email");
    await this.clickButtonWithPossibleNames(["Next", "Selanjutnya", "Berikutnya"]);
    // Handle CAPTCHA jika muncul
    await this.handleCaptchaIfPresent();

    console.log("[INFO] Waiting for Setup button...");
    const setupBtn = this.getGenericButton("Setup");

    const start = Date.now();
    const interval = setInterval(() => {
      console.log(
        `[INFO] Still waiting Setup... ${Math.round((Date.now() - start) / 1000)}s`,
      );
    }, 15000);

    try {
      await this.waitWithCheck(setupBtn, HARD_TIMEOUT);
      this._setupBtnReady = true;
    } finally {
      clearInterval(interval);
    }
  }

  /**
   * Detect CAPTCHA popup dan klik Next-nya, lalu tunggu user solve manual.
   * Kalau tidak ada CAPTCHA, langsung lanjut.
   */
  async handleCaptchaIfPresent() {
    const captchaIndicators = [
      'text="Melindungi akun Anda"',
      'text="Protecting your account"',
      'text="Pecahkan teka-teki"',
      'text="solve the puzzle"',
    ];

    const combinedLocator = this.page
      .locator(captchaIndicators.join(", "))
      .first();

    const hasCaptcha = await combinedLocator
      .waitFor({ state: "visible", timeout: 8000 })
      .then(() => true)
      .catch(() => false);

    if (hasCaptcha) {
      console.log("[CAPTCHA] CAPTCHA detected — aborting...");
      throw new Error("CAPTCHA_DETECTED: CAPTCHA detected.");
    }
  }

  async clickConfirmEmailSetupAccountButton() {
    console.log("[STEP 7] Clicking Setup Account button");
    await this.clickButtonWithPossibleNames([
      "Setup Account",
      "Setup",
      "Set up",
      "Atur Akun",
      "Siapkan Akun",
      "Mulai",
    ]);
    this._setupBtnReady = false;
  }

  async fillBasicInfo() {
    console.log("[STEP 8] Filling basic info");

    await this.waitWithCheck(this.getGenericLocator("first"), HARD_TIMEOUT);

    const fieldDefs = [
      { keyword: "first", value: this.accountConfig.microsoftAccount.firstName },
      { keyword: "last", value: this.accountConfig.microsoftAccount.lastName },
      { keyword: "company", value: this.accountConfig.microsoftAccount.companyName },
      { keyword: "phone", value: this.accountConfig.microsoftAccount.phone },
      { keyword: "job", value: this.accountConfig.microsoftAccount.jobTitle },
    ];

    for (const { keyword, value } of fieldDefs) {
      const locator = this.getGenericLocator(keyword);
      await this.waitForVisible(locator);
      await locator.click();
      await locator.pressSequentially(value, {
        delay: Math.floor(Math.random() * 40) + 60,
      });
      await this.humanDelay(400, 800);
    }

    const addressLocator = this.getGenericLocator("address");
    await this.waitForVisible(addressLocator);
    await addressLocator.click();
    await addressLocator.pressSequentially(
      this.accountConfig.microsoftAccount.address,
      { delay: Math.floor(Math.random() * 30) + 50 },
    );
    await this.humanDelay(100, 300);

    const cityLocator = this.getGenericLocator("city");
    await this.waitForVisible(cityLocator);
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
          { delay: Math.floor(Math.random() * 30) + 50 },
        );
        console.log("Postal code filled");
        await this.humanDelay(150, 400);
      } catch {
        console.log("Postal code field found but could not fill, skipping...");
      }
    } else {
      console.log("Postal code not provided or field not found, skipping...");
    }

    await this.selectDropdownByText(
      'div[role="combobox"][id*="size" i], div[role="combobox"][data-testid*="size" i], select[id*="size" i]',
      this.accountConfig.microsoftAccount.companySize,
    );
    await this.humanDelay(600, 1200);

    // Region / State
    const regionInput = this.page
      .locator('input[id*="region" i], input[id*="state" i]')
      .first();

    const regionIsInput = await this.waitForSpinnerGone()
      .then(() => regionInput.waitFor({ state: "visible", timeout: 8000 }))
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
      await this.selectDropdownByText(
        'div[role="combobox"][id*="region" i], div[role="combobox"][id*="state" i], select[id*="region" i]',
        this.accountConfig.microsoftAccount.state || "Alabama",
      );
    }
    await this.humanDelay(600, 1200);

    await this.selectDropdownByText(
      'div[role="combobox"][id*="website" i], div[role="combobox"][data-testid*="website" i], select[id*="website" i]',
      ["No", "Tidak"]
    );
    await this.humanDelay(600, 1200);

    // Partner checkbox
    try {
      let partnerCheckbox = this.page.locator("#partner-checkbox");
      if ((await partnerCheckbox.count()) === 0) {
        partnerCheckbox = this.page.locator(
          'input[type="checkbox"][aria-label*="share my information" i]',
        );
      }
      if ((await partnerCheckbox.count()) > 0) {
        await partnerCheckbox.waitFor({ state: "visible", timeout: 10000 });
        if (!(await partnerCheckbox.isChecked())) {
          await this.randomMouseMove();
          await partnerCheckbox.check({ force: true });
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
    await this.clickButtonWithPossibleNames(["Next", "Selanjutnya", "Berikutnya", "Continue"]);
  }

  async clickUseThisAddressButton() {
    console.log("[STEP 10] Checking for address confirmation button...");

    await this.waitForSpinnerGone();

    const combinedLocator = this.page
      .locator([
        'button:has-text("Use this address")',
        'button:has-text("Use address")',
        'button:has-text("Gunakan alamat ini")',
        'button[aria-label*="Use this address" i]',
        'button[aria-label*="Use address" i]',
        'button[aria-label*="Gunakan alamat ini" i]',
      ].join(", "))
      .first();

    // Button ini optional — pakai isVisible() bukan waitFor agar tidak blocking
    const found = await combinedLocator.isVisible().catch(() => false);

    if (!found) {
      console.log("[STEP 10] Address confirmation button not found, skipping...");
      return;
    }

    // Cek via aria-label dulu, fallback ke textContent
    const ariaLabel = await combinedLocator.getAttribute("aria-label").catch(() => "");
    const textContent = await combinedLocator.textContent().catch(() => "");
    const buttonText = (ariaLabel || textContent).trim();

    if (/gunakan alamat ini/i.test(buttonText)) {
      const radio = this.page.locator('input[type="radio"]').first();
      const radioVisible = await radio.isVisible().catch(() => false);
      if (radioVisible) {
        await radio.click();
        await this.humanDelay(200, 400);
      }
    }

    await this.randomMouseMove();
    await combinedLocator.click({ force: true });
    console.log(`[STEP 10] Clicked: "${buttonText}"`);
    await this.humanDelay(200, 500);
  }

  async fillPassword() {
    console.log("[STEP 11] Filling password");

    await this.waitForSpinnerGone();

    try {
      const domainInput = this.page
        .locator('input.ms-TextField-field[maxlength="27"]')
        .first();
      await this.runWithMonitor(domainInput.waitFor({ state: "visible", timeout: HARD_TIMEOUT }));
      await this.page
        .waitForFunction(
          (el) => el && el.value && el.value.length > 3,
          await domainInput.elementHandle(),
          { timeout: 15000 },
        )
        .catch(() => { });

      const prefix = await domainInput.inputValue();
      if (prefix) {
        this.extractedDomainEmail = `${prefix}.onmicrosoft.com`;
        this.extractedDomainPassword = this.accountConfig.microsoftAccount.password;
        console.log(`[INFO] Extracted Domain Email: ${this.extractedDomainEmail}`);
      }
    } catch (e) {
      console.log("[WARN] Could not extract domain prefix:", e.message);
    }

    const passwordLocator = this.page
      .locator(
        'input[type="password"]:not([id*="retype" i]):not([id*="confirm" i]):not([data-testid*="cpwd" i])',
      )
      .first();
    const confirmPasswordLocator = this.page
      .locator('input[type="password"]')
      .nth(1);

    await this.waitForVisible(passwordLocator);
    await this.randomMouseMove();
    await passwordLocator.click({ force: true }).catch(() => { });
    await passwordLocator.pressSequentially(
      this.accountConfig.microsoftAccount.password,
      { delay: Math.floor(Math.random() * 100) + 100 },
    );
    await this.humanDelay(100, 300);

    await confirmPasswordLocator.click({ force: true }).catch(() => { });
    await confirmPasswordLocator.pressSequentially(
      this.accountConfig.microsoftAccount.password,
      { delay: Math.floor(Math.random() * 40) + 60 },
    );
    await this.humanDelay(200, 500);

    await this.clickButtonWithPossibleNames([
      "Next",
      "Selanjutnya",
      "Berikutnya",
      "Finish",
      "Selesai",
    ]);
  }

  async handleOptionalSignIn() {
    console.log("[STEP 11.5] Checking for optional Sign In prompt...");

    try {
      await this.waitForSpinnerGone();

      if (await this.checkForError()) {
        throw new Error("MICROSOFT_ERROR_PAGE: Terdeteksi saat pengecekan Sign In.");
      }

      const signInBtn = this.getGenericButton("Sign In");

      const visible = await signInBtn.isVisible().catch(() => false);

      if (!visible) {
        console.log("No Sign In button detected, skipping...");
        return;
      }

      console.log("Sign In detected, clicking...");
      await this.randomMouseMove();

      const [popup] = await Promise.all([
        this.page.waitForEvent("popup").catch(() => null),
        signInBtn.click(),
      ]);

      if (!popup) {
        console.log("No popup detected after Sign In click.");
        return;
      }

      await popup.waitForLoadState("domcontentloaded");

      const yesBtn = popup.locator(
        'button:has-text("Yes"), input[value="Yes"], #idSIButton9'
      );

      const yesVisible = await yesBtn.isVisible().catch(() => false);

      if (yesVisible) {
        await yesBtn.click();
        console.log("Clicked Yes on Stay signed in prompt.");
      }

      await popup.waitForLoadState("networkidle").catch(() => { });
      console.log("Sign In popup handled successfully");

    } catch (e) {
      if (e.message.includes("MICROSOFT_ERROR_PAGE")) throw e;
      console.log("Optional Sign In handler skipped:", e.message);
    }
  }

  async goToPaymentPage() {
    console.log("[STEP 12] Waiting until payment page appears");

    await this.page.waitForLoadState("domcontentloaded", { timeout: HARD_TIMEOUT }).catch(() => { });
    await this.waitForSpinnerGone(500);

    const deadline = Date.now() + HARD_TIMEOUT;
    while (Date.now() < deadline) {
      await this.waitForSpinnerGone();

      // Deteksi via URL atau elemen form kartu — lebih reliable dari teks
      const found = await Promise.any([
        this.page.waitForURL(/payment|billing|checkout/i, { timeout: 3000 }),
        this.page.locator('input[id*="card" i], input[id*="accounttoken" i], input[aria-label*="Nomor kartu" i], input[aria-label*="card number" i]').first().waitFor({ state: "visible", timeout: 3000 }),
      ]).then(() => true).catch(() => false);

      if (found) {
        console.log("Payment page detected");
        return;
      }

      if (await this.checkForError()) {
        throw new Error("MICROSOFT_ERROR_PAGE: Terdeteksi saat menunggu halaman pembayaran.");
      }

      console.log("[STEP 12] Payment page not yet visible, retrying...");
      await this.humanDelay(1000, 2000);
    }

    throw new Error("Timeout waiting for payment page");
  }

  async fillPaymentDetails() {
    console.log("[STEP 13] Filling VCC payment details");

    await this.waitForSpinnerGone();

    const cardLocator = this.page
      .locator(
        'input[id*="accounttoken" i], input[id*="card" i], input[data-testid*="card" i]',
      )
      .first();
    await this.waitForVisible(cardLocator);

    console.log("Typing card number...");
    await cardLocator.click();
    await cardLocator.pressSequentially(this.accountConfig.payment.cardNumber, {
      delay: Math.floor(Math.random() * 30) + 50,
    });
    await this.humanDelay(150, 400);

    console.log("Typing CVV...");
    const cvvLocator = this.page
      .locator('input[id*="cvv" i], input[data-testid*="cvv" i], input[name*="cvv" i]')
      .first();
    await cvvLocator.click();
    await cvvLocator.pressSequentially(this.accountConfig.payment.cvv, {
      delay: Math.floor(Math.random() * 30) + 50,
    });
    await this.humanDelay(150, 400);

    let expMonth = this.accountConfig.payment.expMonth.toString();
    if (expMonth.length === 1) expMonth = "0" + expMonth;

    console.log("Selecting expiry month:", expMonth);
    await this.selectDropdownByText(
      'div[role="combobox"][id*="month" i], div[role="combobox"][data-testid*="month" i], select[id*="month" i]',
      expMonth,
    );
    await this.humanDelay(150, 400);

    console.log("Selecting expiry year:", this.accountConfig.payment.expYear);
    await this.selectDropdownByText(
      'div[role="combobox"][id*="year" i], div[role="combobox"][data-testid*="year" i], select[id*="year" i]',
      this.accountConfig.payment.expYear,
    );
    await this.humanDelay(200, 500);

    console.log("VCC details filled");
  }

  async clickSavePaymentButton() {
    await this.clickButtonWithPossibleNames(["Save", "Simpan", "Next", "Selanjutnya", "Berikutnya"]);

    console.log("[INFO] Waiting for payment response...");

    const TIMEOUT = 60000;
    let resolved = false;

    const makeWatcher = (promise, label) =>
      promise
        .then(() => { resolved = true; return label; })
        .catch(() => null);

    const errorWatcher = new Promise(async (resolve, reject) => {
      const deadline = Date.now() + TIMEOUT;
      while (!resolved && Date.now() < deadline) {
        await this.page.waitForTimeout(2000).catch(() => { });
        if (resolved) break;
        if (await this.checkForError()) {
          return reject(
            new Error("MICROSOFT_ERROR_PAGE: Terdeteksi saat proses Save Payment."),
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
      errorWatcher,
    ]);

    resolved = true;
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
        .catch(() => { });
      await this.humanDelay(1000, 2000);
    }

    if (result === null) {
      console.warn("[WARN] Payment result timeout — tidak ada sinyal jelas dari halaman");
    }

    console.log("[INFO] Payment step finished");
  }

  async clickStartTrialButton() {
    console.log("[STEP 14] Clicking Start Trial button");

    // Tunggu spinner hilang dulu
    await this.waitForSpinnerGone(800);

    // Handle checkbox jika ada
    try {
      const checkbox = this.page.locator('input[type="checkbox"]').first();
      if (await checkbox.count()) {
        const checked =
          (await checkbox.getAttribute("aria-checked")) === "true" ||
          (await checkbox.isChecked());
        if (!checked) {
          console.log("[INFO] Checking agreement checkbox...");
          await this.randomMouseMove();
          await checkbox.click({ force: true });
          await this.humanDelay(300, 700);
        }
      }
    } catch (e) {
      console.log("[INFO] Checkbox handling skipped:", e.message);
    }

    // Tunggu tombol enabled (poll terus sampai HARD_TIMEOUT)
    console.log("[INFO] Waiting for Start Trial button to be enabled...");
    await this.page.waitForFunction(
      () => {
        const btn = [...document.querySelectorAll("button")].find((b) =>
          /start trial|try now|coba sekarang|mulai uji coba/i.test(b.textContent),
        );
        return (
          btn &&
          !btn.disabled &&
          btn.getAttribute("aria-disabled") !== "true" &&
          !btn.classList.contains("is-disabled")
        );
      },
      { timeout: HARD_TIMEOUT },
    ).catch(() => console.log("[WARN] Could not confirm button enabled, proceeding..."));

    await this.clickButtonWithPossibleNames([
      "Start trial",
      "Mulai uji coba",
      "Try now",
      "Coba sekarang",
      "Start",
    ]);

    console.log("[INFO] Start Trial clicked");

    await Promise.race([
      this.page.waitForNavigation({ timeout: HARD_TIMEOUT }).catch(() => { }),
      this.page.waitForLoadState("networkidle").catch(() => { }),
    ]);
  }

  async clickPostTrialNextButton() {
    console.log("[STEP 15] Clicking final Next/Get Started button");
    await this.waitForSpinnerGone(800);

    await this.page.evaluate(() => {
      document
        .querySelectorAll('[data-testid="spinner"], .css-100, .ms-Spinner')
        .forEach((el) => el.remove());
    });

    await this.humanDelay(300, 700);

    await this.clickButtonWithPossibleNames([
      "Next",
      "Selanjutnya",
      "Berikutnya",
      "Get started",
      "Get Started",
      "Mulai",
      "Mulai percobaan"
    ]);

    console.log("[INFO] Next/Get Started clicked");
    await this.waitForPage();
  }

  async extractDomainEmail() {
    console.log("[STEP 16] Finalizing account data...");

    if (this.extractedDomainEmail && this.extractedDomainPassword) {
      console.log("[STEP 16] Using pre-extracted data:", this.extractedDomainEmail);
      return {
        domainEmail: this.extractedDomainEmail,
        domainPassword: this.extractedDomainPassword,
      };
    }

    const emailLocator = this.page.locator("#displayName");
    const found = await emailLocator
      .waitFor({ state: "visible", timeout: HARD_TIMEOUT })
      .then(() => true)
      .catch(() => false);

    if (!found) {
      return {
        domainEmail: this.extractedDomainEmail || "",
        domainPassword: this.extractedDomainPassword || "",
      };
    }

    const rawText = (await emailLocator.textContent())?.trim() || "";
    const emailRegex =
      /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.(onmicrosoft\.[a-z]{2,}|onmschina\.cn)/i;
    const domainEmail = rawText.match(emailRegex)?.[0] || this.extractedDomainEmail || "";
    const domainPassword = this.accountConfig.microsoftAccount.password;

    console.log("[STEP 16] Final Domain Email:", domainEmail);
    return { domainEmail, domainPassword };
  }

  // ─── Error detection ─────────────────────────────────────────────────────────

  async checkForError() {
    try {
      // 1. Cek pesan error validasi di field (biasanya merah di bawah input)
      const fieldError = this.page.locator('[data-automation-id="error-message"]').first();
      if (await fieldError.isVisible().catch(() => false)) {
        const msg = (await fieldError.textContent().catch(() => "")).trim();
        return `Field Validation Error: ${msg}`;
      }

      // 2. Cek teks body untuk indikasi error Microsoft (Optimized evaluate)
      const detailedError = await this.page.evaluate(() => {
        const text = document.body.innerText;
        const lowerText = text.toLowerCase();

        // List marker error yang sering muncul
        const markers = [
          "something went wrong",
          "error code",
          "terjadi sesuatu",
          "Terjadi kesalahan",
          "Melindungi akun Anda",
          "715-123280", // Kode blokir umum
          "incorrectly formatted postal code",
          "something happened",
          "we are sorry, but we could not complete this",
          "try a different way",
          "We're checking to make sure we can offer you Microsoft products and services."
        ];

        const foundMarker = markers.find(m => lowerText.includes(m));
        if (!foundMarker) return null;

        // Pengecualian protektif agar tidak false positive
        if (lowerText.includes("something happened") && lowerText.includes("something happened to be")) {
          return null;
        }

        // Coba cari elemen yang mengandung teks error untuk ambil context lebih banyak
        // Jika ada element dengan class atau ID 'error', 'errorMessage', dsb.
        const errorContainer = document.querySelector('[role="alert"], [class*="error" i], [id*="error" i]');
        if (errorContainer && errorContainer.innerText.length > 5) {
          return errorContainer.innerText.trim();
        }

        // Fallback: Ambil potongan teks di sekitar marker atau baris yang mengandung marker
        const lines = text.split('\n');
        const errorLine = lines.find(l => l.toLowerCase().includes(foundMarker));
        return errorLine ? errorLine.trim() : `Indicator detected: ${foundMarker}`;
      }).catch(() => null);

      if (detailedError) {
        console.log(`[ERROR] Microsoft error detected: ${detailedError}`);
        return detailedError;
      }
    } catch (err) {
      // Ignore silence check errors
    }
    return null;
  }

  async waitWithCheck(locator, timeout = HARD_TIMEOUT) {
    return await this.runWithMonitor(
      locator.waitFor({ state: "visible", timeout }),
      timeout
    );
  }

  // ─── Cleanup & orchestration ─────────────────────────────────────────────────

  async cleanup() {
    try {
      await this.browser.close();
    } catch (e) {
      console.error("Error closing browser:", e);
    }

    if (config.profilePath && fs.existsSync(config.profilePath)) {
      try {
        fs.rmSync(config.profilePath, { recursive: true, force: true });
        console.log("[CLEANUP] Profile folder deleted:", config.profilePath);
      } catch (e) {
        console.warn("[CLEANUP] Could not delete profile folder:", e.message);
      }
    }
  }

  async runStep(name, fn, delay = null) {
    console.log(`[STEP] ${name}`);
    this._currentStep = name;
    await fn();
    const stepError = await this.checkForError();
    if (stepError) {
      throw new Error(`MICROSOFT_ERROR: ${stepError} (Detected after step "${name}")`);
    }
    if (delay) await this.humanDelay(...delay);
  }

  async run() {
    this._currentStep = "Initializing";
    try {
      await this.runStep("Connecting to browser", () => this.connect(), [1000, 3000]);
      await this.runStep("Opening Microsoft page", () => this.openMicrosoftPage(), [400, 800]);
      await this.runStep("Building cart", () => this.clickBuildCartNextButton(), [300, 600]);
      await this.runStep("Filling email", () => this.fillEmail(), [1000, 2500]);
      await this.runStep("Confirming email", () => this.clickCollectEmailNextButton(), [400, 800]);
      await this.runStep("Setup account button", () => this.clickConfirmEmailSetupAccountButton(), [400, 800]);
      await this.runStep("Filling basic info", () => this.fillBasicInfo(), [1500, 3500]);
      await this.runStep("Confirming address (Stage 1)", () => this.clickUseThisAddressButton(), [300, 600]);
      await this.runStep("Filling password", () => this.fillPassword(), [400, 800]);
      await this.runStep("Handling sign in", () => this.handleOptionalSignIn(), [400, 800]);
      await this.runStep("Going to payment page", () => this.goToPaymentPage(), [400, 800]);
      await this.runStep("Filling payment details", () => this.fillPaymentDetails(), [400, 800]);
      await this.runStep("Saving payment", () => this.clickSavePaymentButton());

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