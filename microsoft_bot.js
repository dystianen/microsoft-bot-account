const { chromium } = require("playwright-core");
const fs = require("fs");
const config = require("./config");
const remoteLogger = require("./remote_logger");

const SPINNER_SELECTOR =
  '[data-testid="spinner"], .ms-Spinner, [class*="spinner" i], .loader';

// Safety net — sangat besar, hanya untuk mencegah hang selamanya
const HARD_TIMEOUT = 1.5 * 60 * 1000; // 1 menit 30 detik

class MicrosoftBot {
  constructor(wsUrl, accountConfig, onPaymentSaved) {
    this.wsUrl = wsUrl;
    this.browser = null;
    this.context = null;
    this.page = null;
    this.accountConfig = accountConfig;
    this.onPaymentSaved = onPaymentSaved;
    this._paymentSavedTriggered = false;
    this.currentStep = 0;
    this._setupBtnReady = false;
  }

  async _logStep(stepNum, msg) {
    this.currentStep = stepNum;
    const email = this.accountConfig.microsoftAccount.email;
    // JANGAN di-await agar bot tidak berhenti/hang jika antrean log Telegram menumpuk
    remoteLogger
      .logStep(email, stepNum, msg)
      .catch((e) => console.error(`[LOG ERROR] ${e.message}`));
  }

  async triggerPaymentSaved() {
    if (this._paymentSavedTriggered) return;
    this._paymentSavedTriggered = true;
    console.log("[INFO] Triggering onPaymentSaved callback...");
    if (typeof this.onPaymentSaved === "function") {
      await this.onPaymentSaved().catch((e) =>
        console.error("[CALLBACK ERROR] onPaymentSaved failed:", e.message),
      );
    }
  }

  // ─── Core helpers ────────────────────────────────────────────────────────────

  async humanDelay(min = 1000, max = 3000) {
    const delay = Math.floor(Math.random() * (max - min + 1) + min);
    await this.page.waitForTimeout(delay);
  }

  async humanScroll() {
    try {
      const direction = Math.random() > 0.5 ? 1 : -1;
      const distance = Math.floor(Math.random() * 300) + 100;
      await this.page.mouse.wheel(0, direction * distance);
      await this.humanDelay(500, 1200);
    } catch (e) {
      // Ignore scroll errors
    }
  }

  // Clipboard paste — lebih cepat & tidak terpengaruh proxy/lag
  async humanPaste(locator, text) {
    if (!text) return;
    await locator.click({ force: true }).catch(() => {});
    await this.humanDelay(200);
    await locator.fill(""); // clear dulu
    await this.humanDelay(100);
    await this.page.evaluate((val) => {
      const el = document.activeElement;
      if (el) {
        const nativeInputValueSetter = Object.getOwnPropertyDescriptor(
          window.HTMLInputElement.prototype,
          "value",
        )?.set;
        if (nativeInputValueSetter) {
          nativeInputValueSetter.call(el, val);
          el.dispatchEvent(new Event("input", { bubbles: true }));
          el.dispatchEvent(new Event("change", { bubbles: true }));
        } else {
          el.value = val;
          el.dispatchEvent(new Event("input", { bubbles: true }));
        }
      }
    }, text);
    // Verifikasi hasil paste
    const current = await locator.inputValue().catch(() => "");
    if (current !== text) {
      // Fallback ke fill biasa jika paste gagal
      await locator.fill(text);
    }
  }

  async humanType(locator, text) {
    if (!text) return;
    await locator.click({ force: true }).catch(() => {});
    await this.humanDelay(300);
    await locator.pressSequentially(text, {
      delay: Math.floor(Math.random() * 50) + 40,
    });
  }

  async humanClick(locator, options = {}) {
    await this.randomMouseMove();
    await locator.hover({ force: true }).catch(() => {});
    await this.humanDelay(500);
    await locator.click({ force: true, ...options });
    await this.humanDelay(300);
  }

  async randomMouseMove() {
    try {
      const { width, height } = this.page.viewportSize() || {
        width: 1280,
        height: 720,
      };

      // Target localized area instead of just anywhere to be more realistic
      const x = Math.floor(Math.random() * width);
      const y = Math.floor(Math.random() * height);

      // Use more steps for smoother/slower movement
      const steps = Math.floor(Math.random() * 15) + 10;
      await this.page.mouse.move(x, y, { steps });

      // Occasionally "flutter" the mouse
      if (Math.random() > 0.8) {
        await this.page.mouse.move(x + 5, y + 5, { steps: 5 });
      }
    } catch (e) {}
  }

  async runWithMonitor(promise, timeout = HARD_TIMEOUT) {
    let isDone = false;
    let errorMsg = null;

    const checkLoop = async () => {
      while (!isDone) {
        await this.page.waitForTimeout(2500).catch(() => {
          isDone = true;
        });
        if (isDone) break;

        const detectedError = await this.checkForError();
        if (detectedError) {
          // ✅ Double-check: tunggu 2 detik lagi, lalu cek ulang sebelum throw
          // Ini mencegah false-positive saat teks error muncul sementara (transisi halaman)
          console.log(
            `[MONITOR] Possible error detected: "${detectedError}", re-checking in 2s...`,
          );
          await this.page.waitForTimeout(2000).catch(() => {});
          if (isDone) break; // Task selesai duluan, abaikan
          const recheck = await this.checkForError();
          if (recheck) {
            errorMsg = recheck;
            isDone = true;
            break;
          } else {
            console.log(`[MONITOR] False positive cleared, continuing...`);
          }
        }
      }
    };

    const result = await Promise.race([promise, checkLoop()]).finally(() => {
      isDone = true;
    });

    if (errorMsg) {
      throw new Error(`MICROSOFT_ERROR: ${errorMsg}`);
    }

    return result;
  }

  async waitForSpinnerGone(extraDelay = 0) {
    // 1. Tunggu sebentar karena spinner sering kali baru muncul beberapa saat setelah aksi klik
    await this.page.waitForTimeout(500).catch(() => {});

    const spinner = this.page.locator(SPINNER_SELECTOR).first();
    const spinnerVisible = await spinner.isVisible().catch(() => false);

    if (spinnerVisible) {
      console.log("[WAIT] Spinner detected, waiting until hidden...");
      try {
        await this.runWithMonitor(
          spinner.waitFor({ state: "hidden", timeout: HARD_TIMEOUT }),
        );
      } catch (e) {
        if (e.message.includes("MICROSOFT_ERROR")) throw e;
        console.log(
          "[WAIT] Spinner still visible or check failed, continuing...",
        );
      }
      console.log("[WAIT] Spinner gone.");
      // Grace period agar DOM stabil setelah spinner hilang
      await this.page.waitForTimeout(800).catch(() => {});
    }

    // 2. Selesai spinner baru cek deteksi error (seperti instruksi USER)
    const postSpinnerError = await this.checkForError();
    if (postSpinnerError) {
      throw new Error(`MICROSOFT_ERROR: ${postSpinnerError}`);
    }

    if (extraDelay > 0) {
      await this.humanDelay(extraDelay + 300);
    }
  }

  async waitForVisible(locator) {
    await this.waitForSpinnerGone();
    await this.runWithMonitor(
      locator.waitFor({ state: "visible", timeout: HARD_TIMEOUT }),
    );
  }

  async clickButtonWithPossibleNames(names, timeout = HARD_TIMEOUT) {
    await this.waitForSpinnerGone();

    // ✅ Partial keyword matching — pecah tiap name jadi kata-katanya
    const keywords = names.flatMap((n) => n.trim().toLowerCase().split(/\s+/));
    // Deduplicate
    const uniqueKeywords = [...new Set(keywords)];

    const found = await this.page.evaluate((keywords) => {
      const candidates = [
        ...document.querySelectorAll(
          'button, [role="button"], a[role="button"], input[type="button"], input[type="submit"]',
        ),
      ];

      const el = candidates.find((b) => {
        const text = (
          b.textContent ||
          b.value ||
          b.getAttribute("aria-label") ||
          ""
        )
          .trim()
          .toLowerCase();

        // ✅ Cukup ada 1 keyword yang cocok, teks tidak terlalu panjang
        return (
          text.length > 0 &&
          text.length < 60 &&
          keywords.some((kw) => text.includes(kw))
        );
      });

      if (!el) return null;

      el.click();
      return el.textContent?.trim() || el.value || "unknown";
    }, uniqueKeywords);

    if (found) {
      console.log(`[INFO] Clicked: "${found}"`);
      return true;
    }

    // Fallback: Playwright dengan pattern original
    console.log("[WARN] JS click not found, fallback to Playwright...");
    const pattern = new RegExp(
      names
        .map((n) =>
          n.replace(/[.*+?^${}()|[\]\\]/g, "\\$&").replace(/\s+/g, "\\s*"),
        )
        .join("|"),
      "i",
    );

    const button = this.page.getByRole("button", { name: pattern }).first();

    try {
      await this.runWithMonitor(button.waitFor({ state: "visible", timeout }));
      await this.humanClick(button, { timeout: 8000 });
      const clickedText = await button.textContent().catch(() => "unknown");
      console.log(`[INFO] Clicked: "${clickedText?.trim()}"`);
      return true;
    } catch (err) {
      // Last ditch effort: search in all frames
      console.log("[DEBUG] Searching for button in frames...");
      for (const frame of this.page.frames()) {
        try {
          const frameButton = frame
            .getByRole("button", { name: pattern })
            .first();
          if (await frameButton.isVisible().catch(() => false)) {
            console.log(
              `[INFO] Found and clicking button in frame: ${frame.url()}`,
            );
            await frameButton.click();
            return true;
          }
        } catch (fErr) {}
      }

      const allButtons = await this.page.evaluate(() =>
        [
          ...document.querySelectorAll(
            'button, [role="button"], a[role="button"]',
          ),
        ]
          .map((b) => b.textContent?.trim())
          .filter(Boolean),
      );
      console.error(`[ERROR] Button not found. Available buttons:`, allButtons);
      console.error(`[ERROR] Looking for keywords:`, uniqueKeywords);
      throw err;
    }
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

    const dropdown = this.page.locator(selector).first();
    await this.runWithMonitor(
      dropdown.waitFor({ state: "visible", timeout: HARD_TIMEOUT }),
    );

    await dropdown.scrollIntoViewIfNeeded();

    const tagName = await dropdown.evaluate((el) => el.tagName.toLowerCase());

    if (tagName === "select") {
      const searchList = Array.isArray(text) ? text : [text];
      for (const t of searchList) {
        try {
          await dropdown.selectOption({ label: t });
          console.log(`[DROPDOWN] Selected via native: "${t}"`);
          return true;
        } catch {
          continue;
        }
      }
    }

    await dropdown.click();

    const dropdownItems = this.page.locator(".ms-Dropdown-items");
    await this.runWithMonitor(
      dropdownItems.waitFor({ state: "visible", timeout: HARD_TIMEOUT }),
    );

    const searchList = Array.isArray(text)
      ? text.map((t) => (t || "").toString().trim())
      : [(text || "").toString().trim()];

    // Build selector string once — DO NOT resolve the locator to an element yet
    let optionSelector = null;
    for (const search of searchList) {
      const escaped = search.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      const candidate = this.page
        .locator(".ms-Dropdown-item", { hasText: new RegExp(escaped, "i") })
        .first();
      if (await candidate.count()) {
        optionSelector = { hasText: new RegExp(escaped, "i") };
        break;
      }
    }

    if (!optionSelector) {
      console.warn(`[DROPDOWN] Option not found for: ${text}`);
      await this.page.keyboard.press("Escape");
      return false;
    }

    // Retry loop: re-resolve fresh locator each attempt to avoid stale DOM refs
    const MAX_RETRIES = 3;
    for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
      try {
        // Fresh locator on every attempt — this is the core fix
        const freshOption = this.page
          .locator(".ms-Dropdown-item", optionSelector)
          .first();

        await freshOption.waitFor({ state: "attached", timeout: 2000 });
        await freshOption.scrollIntoViewIfNeeded();

        const displayText = await freshOption.textContent().catch(() => text);
        console.log(
          `[DROPDOWN] Clicking: "${displayText?.trim()}" (attempt ${attempt + 1})`,
        );

        try {
          await this.humanClick(freshOption, { timeout: 3000 });
        } catch {
          console.log("[DROPDOWN] Normal click blocked, using JS click...");
          await freshOption.evaluate((el) => el.click());
        }

        await this.page
          .waitForSelector(".ms-Dropdown-items", {
            state: "detached",
            timeout: 5000,
          })
          .catch(() => {});

        return true;
      } catch (err) {
        if (attempt === MAX_RETRIES - 1) throw err;
        console.warn(
          `[DROPDOWN] Attempt ${attempt + 1} failed (${err.message}), retrying...`,
        );
        await this.page.waitForTimeout(150);
      }
    }

    return false;
  }

  async waitForPage(selector) {
    await this.waitForSpinnerGone();
    if (selector) {
      await this.runWithMonitor(
        this.page.waitForSelector(selector, {
          state: "attached",
          timeout: HARD_TIMEOUT,
        }),
      );
    } else {
      await this.runWithMonitor(
        this.page.waitForLoadState("domcontentloaded", {
          timeout: HARD_TIMEOUT,
        }),
      );
    }
    // Added random delay after every major page load to simulate human orientation time
    await this.humanDelay(2500);
  }

  // ─── Steps ───────────────────────────────────────────────────────────────────

  async connect() {
    await this._logStep(1, "🌐 Menghubungkan ke browser...");

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
    this.profileId = this.wsUrl.split("/").pop();

    // Check IP vs billing address country (Anti-Fraud)
    try {
      console.log("[INFO] Verifying IP location...");
      const ipInfoResponse = await this.page.evaluate(async () => {
        try {
          const res = await fetch("https://ipapi.co/json/");
          return await res.json();
        } catch {
          return null;
        }
      });

      if (ipInfoResponse && ipInfoResponse.country_name) {
        const ipCountry = ipInfoResponse.country_name.toLowerCase();
        const billingAddress =
          this.accountConfig.basicInfo?.address ||
          this.accountConfig.payment?.address ||
          "";
        console.log(
          `[INFO] Current IP location: ${ipInfoResponse.city}, ${ipInfoResponse.country_name} (${ipInfoResponse.ip})`,
        );

        // Simple heuristic: check if billing address mentioned country matches IP country
        // (This can be refined if we have a strict country code in config)
        if (
          billingAddress &&
          !billingAddress.toLowerCase().includes(ipCountry)
        ) {
          console.warn(
            `[ANTI-FRAUD WARNING] Location mismatch! IP is in ${ipCountry}, but billing address might be elsewhere.`,
          );
          console.warn(`Billing info provided: ${billingAddress}`);
        }
      }
    } catch (e) {
      console.log("[WARN] Could not verify IP location, continuing anyway.");
    }
  }

  async openMicrosoftPage() {
    await this._logStep(2, "🌍 Membuka halaman Microsoft...");

    const url = this.accountConfig.microsoftUrl || config.microsoftUrl;
    // Speed up initial navigation — wait for commit then poll for elements
    await this.page.goto(url, {
      waitUntil: "commit",
      timeout: HARD_TIMEOUT,
    });
  }

  async clickTryForFreeOnTargetCard() {
    const targetPlan = this.accountConfig.targetPlan || "E3";
    await this._logStep(3, `Memilih paket trial: ${targetPlan}`);

    const cards = this.page.locator(
      'div[ocr-component-name="card-plan-detail"]',
    );
    // Poll fast for cards without waiting for domcontentloaded
    const cardsVisible = await cards
      .first()
      .waitFor({ state: "visible", timeout: 90000 })
      .then(() => true)
      .catch(() => false);

    if (!cardsVisible) {
      console.log(
        "[INFO] No cards visible, checking if we're scanning global buttons...",
      );
    } else {
      const count = await cards.count();
      let targetCard = null;

      // Greedily search for title in cards
      for (let i = 0; i < count; i++) {
        const card = cards.nth(i);
        const title = await card
          .locator(".oc-product-title")
          .first()
          .textContent()
          .catch(() => "");
        if (title.toUpperCase().includes(targetPlan.toUpperCase())) {
          targetCard = card;
          break;
        }
      }

      const cardToUse =
        targetCard || (count >= 2 ? cards.nth(1) : cards.first());
      const tryFreeBtn = cardToUse
        .locator('a:has-text("Try for free"), a:has-text("Coba gratis")')
        .first();

      if ((await tryFreeBtn.count()) > 0) {
        console.log(
          `[INFO] Clicking "Try for free" (Target: ${targetPlan})...`,
        );
        const [popup] = await Promise.all([
          this.page
            .context()
            .waitForEvent("page", { timeout: 30000 })
            .catch(() => null),
          this.humanClick(tryFreeBtn),
        ]);

        if (popup) {
          this.page = popup;
          console.log(
            "[INFO] Switched to new tab. Waiting for content settle...",
          );
          // Wait for full load and a bit extra for hydration
          await this.page
            .waitForLoadState("load", { timeout: 30000 })
            .catch(() => {});
          await this.waitForSpinnerGone();

          // Wait specifically for any button to ensure JS is likely ready
          await this.page
            .locator('button, [role="button"], a.btn')
            .first()
            .waitFor({ state: "visible", timeout: 15000 })
            .catch(() => {});
          await this.humanDelay(1500); // Small grace period for event listeners to attach
          return;
        }
      }
    }

    // Fallback global search if cards not found or button not in card
    console.log("[INFO] Scanning for global 'Try for free' button...");
    const globalBtn = this.page
      .locator(
        'a:has-text("Try for free"), a:has-text("Coba gratis"), button:has-text("Try for free")',
      )
      .first();
    const [popupGlobal] = await Promise.all([
      this.page
        .context()
        .waitForEvent("page", { timeout: 30000 })
        .catch(() => null),
      this.humanClick(globalBtn).catch(() => {}),
    ]);

    if (popupGlobal) {
      this.page = popupGlobal;
      console.log(
        "[INFO] Switched to new tab (global click). Waiting for content settle...",
      );
      await this.page
        .waitForLoadState("load", { timeout: 30000 })
        .catch(() => {});
      await this.waitForSpinnerGone();
      await this.page
        .locator('button, [role="button"], a.btn')
        .first()
        .waitFor({ state: "visible", timeout: 15000 })
        .catch(() => {});
      await this.humanDelay(1500);
    }
  }

  async clickProductNextButton() {
    await this._logStep(4, "Mengklik tombol Selanjutnya...");

    // Pilih "1 month" jika opsi durasi langganan muncul (mendukung multi-bahasa)
    try {
      const oneMonthSelectors = [
        'label:has-text("1 month")',
        'label:has-text("1 bulan")',
        'label:has-text("1 mes")',
        'label:has-text("1 mois")',
        'label:has-text("1 Monat")',
        'label:has-text("1 mese")',
        'label:has-text("1 mês")',
        'span:has-text("1 month")',
        'span:has-text("1 bulan")',
        'span:has-text("1 mes")',
        'span:has-text("1 mois")',
        'span:has-text("1 Monat")',
        '[aria-label*="1 month" i]',
        '[aria-label*="1 bulan" i]',
        'input[value*="month" i]',
        'input[value*="bulan" i]',
      ].join(", ");

      const oneMonthOption = this.page.locator(oneMonthSelectors).first();

      const isVisible = await oneMonthOption
        .isVisible({ timeout: 8000 })
        .catch(() => false);
      if (isVisible) {
        console.log(
          "[STEP 4] Subscription length option detected. Selecting 1 month...",
        );
        await this.randomMouseMove();
        await oneMonthOption.click({ force: true });
        await this.humanDelay(1500);
      } else {
        console.log(
          "[STEP 4] 1 month option not detected or not visible, proceeding.",
        );
      }
    } catch (e) {
      console.log("[STEP 4] 1 month selection logic skipped:", e.message);
    }

    await this.clickButtonWithPossibleNames([
      "Next",
      "Selanjutnya",
      "Continue",
      "Berikutnya",
    ]);
  }

  async fillEmail() {
    const email = this.accountConfig.microsoftAccount.email;
    await this._logStep(5, `Mengisi email: ${email}`);

    const emailInput = this.getGenericLocator("email");
    await this.waitForVisible(emailInput);
    await this.randomMouseMove();
    await emailInput.click();
    await this.humanType(emailInput, email);

    // Verifikasi isi field sudah benar — guard untuk koneksi proxy lambat
    // jika belum lengkap, pakai insertText (seperti copas)
    const MAX_RETRIES = 3;
    for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
      await this.humanDelay(700);
      const currentValue = await emailInput.inputValue().catch(() => "");
      if (currentValue.trim() === email.trim()) {
        console.log(`[STEP 5] Email verified in input field.`);
        break;
      }
      console.warn(
        `[STEP 5] Email mismatch (attempt ${attempt}/${MAX_RETRIES}): expected "${email}", got "${currentValue}". Retrying with insertText...`,
      );
      await this.humanType(emailInput, "");
      await this.humanDelay(400);
      await emailInput.focus();
      await this.humanType(emailInput, email);
    }

    await this.humanDelay(1500);
  }

  async submitEmailAndWaitForSetup() {
    await this._logStep(6, "Submit email & menunggu tombol Setup...");
    await this.clickButtonWithPossibleNames([
      "Next",
      "Selanjutnya",
      "Berikutnya",
    ]);

    console.log("[INFO] Waiting for Setup button...");

    const setupKeywords = [
      // Setup Account variants
      "Set up account",
      "Setup Account",
      "Setup",
      "Set up",
      "Siapkan akun",
      "Atur Akun",
      "Siapkan Akun",
      "Atur",
      "Siapkan",

      // "Create new account"
      "Create new account",
      "Create account",
      "Buat akun baru",
      "Buat akun",

      // Bahasa lain
      "Crear cuenta nueva",
      "Crear cuenta",
      "Créer un compte",
      "Neues Konto erstellen",
      "Crea nuovo account",
      "Criar nova conta",

      "Mulai",
    ];

    const setupBtnSelector = [
      'button[id*="Setup" i], button[data-testid*="Setup" i], button[data-bi-id*="Setup" i]',
      'a[data-bi-id*="Setup" i]',
      ...setupKeywords.map((kw) => `button:has-text("${kw}")`),
      ...setupKeywords.map((kw) => `a:has-text("${kw}")`),
    ].join(", ");

    const setupBtn = this.page.locator(setupBtnSelector).first();

    const start = Date.now();
    const interval = setInterval(() => {
      console.log(
        `[INFO] Still waiting Setup... ${Math.round((Date.now() - start) / 1000)}s`,
      );
    }, 15000);

    try {
      // ✅ Retry loop: cek apakah halaman berpindah ke step lain (skip Setup)
      const deadline = Date.now() + HARD_TIMEOUT;
      let found = false;
      while (Date.now() < deadline && !found) {
        found = await setupBtn.isVisible({ timeout: 5000 }).catch(() => false);
        if (!found) {
          // Cek apakah sudah langsung masuk ke form basic info (Setup skipped)
          const basicInfoVisible = await this.page
            .locator('input[id*="first" i], input[id*="firstName" i]')
            .first()
            .isVisible({ timeout: 2000 })
            .catch(() => false);
          if (basicInfoVisible) {
            console.log(
              "[INFO] Basic info form detected, Setup step skipped by Microsoft.",
            );
            this._setupBtnReady = false;
            return;
          }
          const err = await this.checkForError();
          if (err) throw new Error(`MICROSOFT_ERROR: ${err}`);
          await this.humanDelay(2000);
        }
      }
      if (!found) throw new Error("Timeout waiting for Setup button");
      this._setupBtnReady = true;
    } finally {
      clearInterval(interval);
    }
  }

  async clickSetupAccountButton() {
    if (this._setupBtnReady === false) {
      console.log(
        "[STEP 7] Setup already skipped or detected as basic info form, skipping click.",
      );
      return;
    }

    await this._logStep(7, "Mengklik tombol Setup Account...");
    await this.clickButtonWithPossibleNames([
      // Setup Account variants
      "Set up account",
      "Setup Account",
      "Setup",
      "Set up",
      "Siapkan akun",
      "Atur Akun",
      "Siapkan Akun",
      "Atur",
      "Siapkan",

      // "Create new account"
      "Create new account",
      "Create account",
      "Buat akun baru",
      "Buat akun",

      // Bahasa lain
      "Crear cuenta nueva",
      "Crear cuenta",
      "Créer un compte",
      "Neues Konto erstellen",
      "Crea nuovo account",
      "Criar nova conta",

      "Mulai",
    ]);
    this._setupBtnReady = false;
  }

  async fillBasicInfo() {
    await this._logStep(8, "Mengisi informasi dasar akun...");

    await this.waitWithCheck(this.getGenericLocator("first"), HARD_TIMEOUT);

    // Row 1: First name — pakai paste agar tidak kena proxy lag
    const firstLocator = this.getGenericLocator("first");
    await this.waitForVisible(firstLocator);
    await this.humanPaste(
      firstLocator,
      this.accountConfig.microsoftAccount.firstName,
    );
    await this.humanDelay(400);

    // Last name
    const lastLocator = this.getGenericLocator("last");
    await this.waitForVisible(lastLocator);
    await this.humanPaste(
      lastLocator,
      this.accountConfig.microsoftAccount.lastName,
    );
    await this.humanDelay(500);

    // Company
    const companyLocator = this.getGenericLocator("company");
    await this.waitForVisible(companyLocator);
    await this.humanPaste(
      companyLocator,
      this.accountConfig.microsoftAccount.companyName,
    );
    await this.humanDelay(400);

    // Company size
    await this.humanScroll();
    await this.selectDropdownByText(
      'div[role="combobox"][id*="size" i], div[role="combobox"][data-testid*="size" i], select[id*="size" i]',
      this.accountConfig.microsoftAccount.companySize,
    );
    await this.humanDelay(1200, 2000);

    // Phone — pakai paste
    const phoneLocator = this.getGenericLocator("phone");
    await this.waitForVisible(phoneLocator);
    await this.humanPaste(
      phoneLocator,
      this.accountConfig.microsoftAccount.phone,
    );
    await this.humanDelay(300);

    // Job — pakai paste
    const jobLocator = this.getGenericLocator("job");
    await this.waitForVisible(jobLocator);
    await this.humanPaste(
      jobLocator,
      this.accountConfig.microsoftAccount.jobTitle,
    );
    await this.humanDelay(500);

    // Address 1 — pakai paste
    const addressLocator = this.getGenericLocator("address");
    await this.waitForVisible(addressLocator);
    await this.humanPaste(
      addressLocator,
      this.accountConfig.microsoftAccount.address,
    );
    await this.humanDelay(400);

    // City — pakai paste
    const cityLocator = this.getGenericLocator("city");
    await this.waitForVisible(cityLocator);
    await this.humanPaste(
      cityLocator,
      this.accountConfig.microsoftAccount.city,
    );
    await this.humanDelay(400);

    const regionInput = this.page
      .locator('input[id*="region" i], input[id*="state" i]')
      .first();

    const regionIsInput = await regionInput
      .waitFor({ state: "visible", timeout: 5000 })
      .then(() => true)
      .catch(() => false);

    await this.humanDelay(1010);
    if (regionIsInput) {
      await this.humanPaste(
        regionInput,
        this.accountConfig.microsoftAccount.state,
      );
      console.log("Region filled as input");
    } else {
      await this.selectDropdownByText(
        'div[role="combobox"][id*="region" i], div[role="combobox"][id*="state" i], select[id*="region" i], select[id*="state" i]',
        this.accountConfig.microsoftAccount.state,
      );
    }

    await this.humanDelay(613);

    // 🔥 POSTAL — pakai paste
    const zipLocator = this.page
      .locator(
        'input[id*="postal" i], input[id*="zip" i], input[data-testid*="postal" i], input[data-testid*="zip" i]',
      )
      .first();

    if (
      this.accountConfig.microsoftAccount.postalCode &&
      (await zipLocator.count()) > 0
    ) {
      try {
        await this.humanPaste(
          zipLocator,
          this.accountConfig.microsoftAccount.postalCode,
        );
        console.log("Postal filled");
        await this.humanDelay(1000);
      } catch {
        console.log("Postal found but failed to fill");
      }
    } else {
      console.log("Postal not found / not provided");
    }

    // Country
    // await this.selectDropdownByText(
    //   'div[role="combobox"][id*="country" i], select[id*="country" i]',
    //   this.accountConfig.microsoftAccount.country || "United States",
    // ).catch(() => { });
    // await this.humanDelay(800, 1500);

    // Website (jangan pakai "Select one")
    await this.humanDelay(714);
    await this.humanScroll();
    await this.selectDropdownByText(
      'div[role="combobox"][id*="website" i], div[role="combobox"][data-testid*="website" i], select[id*="website" i]',
      ["No", "Tidak"],
    );
    await this.humanDelay(2000, 3000);

    // Checkbox
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
          console.log("Checkbox checked");
        }
      }
    } catch (err) {
      console.log("Checkbox error:", err.message);
    }

    await this.humanDelay(1310);

    await this.clickButtonWithPossibleNames([
      "Next",
      "Selanjutnya",
      "Berikutnya",
      "Continue",
    ]);
  }

  async confirmAddressIfPrompted() {
    await this._logStep(10, "Mengecek konfirmasi alamat...");

    await this.waitForSpinnerGone();

    const combinedLocator = this.page
      .locator(
        [
          'button:has-text("Use this address")',
          'button:has-text("Use address")',
          'button:has-text("Gunakan alamat ini")',
          'button[aria-label*="Use this address" i]',
          'button[aria-label*="Use address" i]',
          'button[aria-label*="Gunakan alamat ini" i]',
        ].join(", "),
      )
      .first();

    const found = await combinedLocator.isVisible().catch(() => false);

    if (!found) {
      console.log(
        "[STEP 10] Address confirmation button not found, skipping...",
      );
      return;
    }

    // Selalu pilih radio button pertama (atas) jika ada
    const firstRadio = this.page.locator('input[type="radio"]').first();
    const radioVisible = await firstRadio.isVisible().catch(() => false);
    if (radioVisible) {
      await firstRadio.click();
      await this.humanDelay(400);
    }

    await this.randomMouseMove();
    await combinedLocator.click({ force: true });

    const buttonText = await combinedLocator.textContent().catch(() => "");
    console.log(`[STEP 10] Clicked: "${buttonText.trim()}"`);
    await this.humanDelay(500);
  }

  async fillPassword() {
    await this._logStep(11, "Mengisi password dan konfirmasi domain...");

    await this.waitForSpinnerGone();
    try {
      const inputs = this.page.locator("input.ms-TextField-field");
      const count = await inputs.count();
      let username = "";
      let prefix = "";

      for (let i = 0; i < count; i++) {
        const input = inputs.nth(i);
        const id = await input.getAttribute("id").catch(() => "");
        const placeholder = await input
          .getAttribute("placeholder")
          .catch(() => "");
        const val = await input.inputValue();

        if (id?.includes("username") || placeholder?.includes("username")) {
          username = val;
        } else if ((await input.getAttribute("maxlength")) === "27") {
          prefix = val;
        }
      }

      // Fallback if ID/placeholder check fails
      if (!username && count >= 2) {
        username = await inputs.nth(0).inputValue();
        prefix = await inputs.nth(1).inputValue();
      }

      if (username && prefix) {
        this.extractedDomainEmail = `${username}@${prefix}.onmicrosoft.com`;
        this.extractedDomainPassword =
          this.accountConfig.microsoftAccount.password;
        console.log(
          `[INFO] Extracted Domain Email: ${this.extractedDomainEmail}`,
        );
      } else if (prefix) {
        this.extractedDomainEmail = `${prefix}.onmicrosoft.com`;
      }
    } catch (e) {
      console.log("[WARN] Could not extract domain info:", e.message);
    }

    const passwordLocator = this.page
      .locator(
        'input[type="password"]:not([id*="retype" i]):not([id*="confirm" i]):not([data-testid*="cpwd" i])',
      )
      .first();

    await this.waitForVisible(passwordLocator);
    await this.randomMouseMove();

    // ✅ Password: gunakan humanPaste (copas) agar lebih stabil terhadap lag/proxy
    // Sama seperti yang dilakukan pada pengisian biodata sebelumnya
    await passwordLocator.click({ force: true }).catch(() => {});
    await this.humanDelay(300);
    await this.humanPaste(
      passwordLocator,
      this.accountConfig.microsoftAccount.password,
    );

    await this.humanDelay(700);
    const confirmPasswordLocator = this.page
      .locator('input[type="password"]')
      .nth(1);

    const confirmVisible = await confirmPasswordLocator
      .isVisible({ timeout: 5000 })
      .catch(() => false);
    if (confirmVisible) {
      await confirmPasswordLocator.click({ force: true }).catch(() => {});
      await this.humanPaste(
        confirmPasswordLocator,
        this.accountConfig.microsoftAccount.password,
      );
    }
    await this.humanDelay(500);

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
      // Tunggu halaman benar-benar settle setelah submit password
      await this.page.waitForLoadState("domcontentloaded");
      await this.waitForSpinnerGone();
      await this.humanDelay(1500); // beri waktu DOM stabil

      if (await this.checkForError()) {
        throw new Error(
          "MICROSOFT_ERROR_PAGE: Terdeteksi saat pengecekan Sign In.",
        );
      }

      const signInBtn = this.page
        .locator(
          [
            'button:has-text("Sign In")',
            'button:has-text("Sign-In")',
            'button:has-text("Masuk")',
            'a:has-text("Sign In")',
            'a:has-text("Masuk")',
            '[data-bi-id*="signin" i]',
            'button[id*="signin" i]',
          ].join(", "),
        )
        .first();

      const paymentPageLocator = this.page
        .locator(
          'input[id*="card" i], input[id*="accounttoken" i], input[aria-label*="card number" i], input[aria-label*="Nomor kartu" i]',
        )
        .first();

      // Race: Sign In button vs Payment page — prioritaskan deteksi elemen fisik daripada URL
      const winner = await Promise.race([
        signInBtn
          .waitFor({ state: "visible", timeout: 12000 })
          .then(() => "signin")
          .catch(() => null),

        paymentPageLocator
          .waitFor({ state: "visible", timeout: 12000 })
          .then(() => "payment")
          .catch(() => null),

        this.page
          .waitForURL(/payment|billing|checkout/i, { timeout: 12000 })
          .then(() => "payment_url")
          .catch(() => null),
      ]);

      console.log(`[STEP 11.5] Race result: ${winner}`);

      if (winner === "payment") {
        console.log("[STEP 11.5] Payment field detected, skipping Sign In.");
        return;
      }

      if (winner === "payment_url") {
        // Jika hanya URL yang match, cek lagi apakah tombol Sign In sebenarnya ada
        const signVisible = await signInBtn.isVisible().catch(() => false);
        if (signVisible) {
          console.log(
            "[STEP 11.5] URL match payment but Sign In button is visible. Prioritizing Sign In.",
          );
        } else {
          console.log(
            "[STEP 11.5] Payment URL detected and no Sign In button found, skipping.",
          );
          return;
        }
      }

      if (
        !winner ||
        (!winner.includes("signin") &&
          !(await signInBtn.isVisible().catch(() => false)))
      ) {
        console.log(
          "[STEP 11.5] No Sign In or Payment page detected, skipping.",
        );
        return;
      }

      // Proceed to click Sign In
      console.log("[STEP 11.5] Sign In detected, clicking...");
      await this.randomMouseMove();

      const [popup] = await Promise.all([
        this.page.waitForEvent("popup").catch(() => null),
        signInBtn.click(),
      ]);

      if (!popup) {
        console.log("[STEP 11.5] No popup after Sign In click, continuing...");
        return;
      }

      await popup.waitForLoadState("domcontentloaded");
      const yesBtn = popup.locator(
        'button:has-text("Yes"), input[value="Yes"], #idSIButton9',
      );
      const yesVisible = await yesBtn
        .waitFor({ state: "visible", timeout: 15000 })
        .then(() => true)
        .catch(() => false);

      if (yesVisible) {
        await yesBtn.click();
        console.log("[STEP 11.5] Clicked Yes on Stay signed in prompt.");
      }

      await popup.waitForLoadState("networkidle").catch(() => {});
      console.log("[STEP 11.5] Sign In popup handled successfully.");
    } catch (e) {
      if (e.message.includes("MICROSOFT_ERROR_PAGE")) throw e;
      console.log("[STEP 11.5] Optional Sign In handler skipped:", e.message);
    }
  }

  async goToPaymentPage() {
    await this._logStep(12, "Menunggu halaman pembayaran muncul...");

    await this.page
      .waitForLoadState("domcontentloaded", { timeout: HARD_TIMEOUT })
      .catch(() => {});
    await this.waitForSpinnerGone(500);

    const deadline = Date.now() + HARD_TIMEOUT;
    while (Date.now() < deadline) {
      await this.waitForSpinnerGone();

      // Deteksi via URL atau elemen form kartu — lebih reliable dari teks
      const found = await Promise.any([
        this.page.waitForURL(/payment|billing|checkout/i, { timeout: 3000 }),
        this.page
          .locator(
            'input[id*="card" i], input[id*="accounttoken" i], input[aria-label*="Nomor kartu" i], input[aria-label*="card number" i]',
          )
          .first()
          .waitFor({ state: "visible", timeout: 3000 }),
      ])
        .then(() => true)
        .catch(() => false);

      if (found) {
        console.log("Payment page detected");
        return;
      }

      if (await this.checkForError()) {
        throw new Error(
          "MICROSOFT_ERROR_PAGE: Terdeteksi saat menunggu halaman pembayaran.",
        );
      }

      console.log("[STEP 12] Payment page not yet visible, retrying...");
      await this.humanDelay(1000);
    }

    throw new Error("Timeout waiting for payment page");
  }

  async fillPaymentDetails() {
    await this._logStep(13, "Mengisi detail pembayaran VCC...");

    await this.waitForSpinnerGone();

    const cardLocator = this.page
      .locator(
        'input[id*="accounttoken" i], input[id*="card" i], input[data-testid*="card" i]',
      )
      .first();
    await this.waitForVisible(cardLocator);

    console.log("Typing card number...");
    await cardLocator.click();
    await this.humanType(cardLocator, this.accountConfig.payment.cardNumber);
    await this.humanDelay(1500);

    console.log("Typing CVV...");
    const cvvLocator = this.page
      .locator(
        'input[id*="cvv" i], input[data-testid*="cvv" i], input[name*="cvv" i]',
      )
      .first();
    await cvvLocator.click();
    await this.humanType(cvvLocator, this.accountConfig.payment.cvv);
    await this.humanDelay(1000);

    let expMonth = this.accountConfig.payment.expMonth.toString();
    if (expMonth.length === 1) expMonth = "0" + expMonth;

    console.log("Selecting expiry month:", expMonth);
    await this.selectDropdownByText(
      'div[role="combobox"][id*="month" i], div[role="combobox"][data-testid*="month" i], select[id*="month" i]',
      expMonth,
    );
    await this.humanDelay(400);

    console.log("Selecting expiry year:", this.accountConfig.payment.expYear);
    await this.selectDropdownByText(
      'div[role="combobox"][id*="year" i], div[role="combobox"][data-testid*="year" i], select[id*="year" i]',
      this.accountConfig.payment.expYear,
    );
    await this.humanDelay(500);

    console.log("VCC details filled");
  }

  async submitPaymentAndWaitResult() {
    await this.clickButtonWithPossibleNames([
      "Save",
      "Simpan",
      "Next",
      "Selanjutnya",
      "Berikutnya",
    ]);

    console.log("[INFO] Waiting for payment response...");

    const waitForPaymentOutcome = async (timeout = 60000) => {
      let resolved = false;

      const makeWatcher = (promise, label) =>
        promise
          .then(() => {
            resolved = true;
            return label;
          })
          .catch(() => null);

      // Selector address: EN + ID
      const ADDRESS_SELECTOR = [
        'button:has-text("Use this address")',
        'button:has-text("Use address")',
        'button:has-text("Gunakan alamat ini")',
        'button[aria-label*="Use this address" i]',
        'button[aria-label*="Gunakan alamat ini" i]',
      ].join(", ");

      const errorWatcher = new Promise(async (resolve, reject) => {
        const deadline = Date.now() + timeout;
        while (!resolved && Date.now() < deadline) {
          await this.page.waitForTimeout(2000).catch(() => {});
          if (resolved) break;
          const err = await this.checkForError();
          if (err) {
            // Unify: instead of rejecting with a different label,
            // resolve as "error" and store the message for consistent handling
            this._lastPaymentMonitorError = err;
            return resolve("error");
          }
        }
        resolve(null);
      });

      const result = await Promise.race([
        makeWatcher(
          this.page.waitForSelector(
            'span[data-automation-id="error-message"]',
            {
              state: "visible",
              timeout,
            },
          ),
          "error",
        ),
        makeWatcher(
          this.page.waitForSelector(ADDRESS_SELECTOR, {
            state: "visible",
            timeout,
          }),
          "address",
        ),
        makeWatcher(
          this.page.waitForFunction(
            () => {
              const text = document.body.innerText.toLowerCase();
              return (
                text.includes("check your info") ||
                text.includes("review your order") ||
                text.includes("ordersummary") ||
                text.includes("tinjau pesanan") ||
                text.includes("periksa info") ||
                text.includes("ringkasan pesanan") ||
                text.includes("setup your account") ||
                text.includes("siapkan akun") ||
                // Avoid too generic "mulai" unless it's a specific page pattern
                (text.includes("mulai") &&
                  (text.includes("pesanan") ||
                    text.includes("data") ||
                    text.includes("akun"))) ||
                window.location.href.includes("ordersummary") ||
                window.location.href.includes("setup-account") ||
                window.location.href.includes("review")
              );
            },
            { timeout },
          ),
          "success",
        ),
        errorWatcher,
      ]);

      resolved = true;
      return result;
    };

    let result = await waitForPaymentOutcome(60000);
    console.log(`[DEBUG] Payment result: ${result}`);

    // Kalau ada address confirmation — klik, lalu tunggu outcome sebenarnya
    if (result === "address") {
      console.log("[INFO] Address confirmation prompt detected, clicking...");
      const ADDRESS_SELECTOR = [
        'button:has-text("Use this address")',
        'button:has-text("Use address")',
        'button:has-text("Gunakan alamat ini")',
        'button[aria-label*="Use this address" i]',
        'button[aria-label*="Gunakan alamat ini" i]',
      ].join(", ");

      try {
        await this.page
          .locator(ADDRESS_SELECTOR)
          .first()
          .click({ force: true });
        console.log("[INFO] Address confirmed, waiting for payment outcome...");
      } catch (e) {
        console.warn("[WARN] Could not click address button:", e.message);
      }

      await this.humanDelay(1000);

      // Tunggu lagi setelah klik address — cek apakah success atau card error
      result = await waitForPaymentOutcome(30000);
      console.log(`[DEBUG] Payment result (post-address): ${result}`);
    }

    if (result === "success") {
      console.log("[INFO] Payment successfully saved signal detected.");
      await this.triggerPaymentSaved();
    } else if (result === "error") {
      let errorText = "";

      // Prioritize specific error from monitor if it caught more detail
      if (this._lastPaymentMonitorError) {
        errorText = this._lastPaymentMonitorError
          .replace(/Field Validation Error:|MICROSOFT_ERROR_PAGE:/i, "")
          .trim();
        this._lastPaymentMonitorError = null; // reset
      }

      if (!errorText) {
        const msg = await this.page
          .locator('span[data-automation-id="error-message"]')
          .first()
          .textContent()
          .catch(() => "Unknown payment error");
        errorText = msg?.trim() || "Unknown payment error";
      }

      console.error(`[ERROR] Payment error detected: ${errorText}`);
      throw new Error(`PAYMENT_DECLINED: ${errorText}`);
    } else if (result === null) {
      console.warn(
        "[WARN] Payment result timeout - long loading time, trigger payment saved",
      );
      await this.triggerPaymentSaved();
    }

    console.log("[INFO] Payment step finished");
  }

  async acceptTrialAndStart() {
    await this._logStep(14, "Menyetujui trial dan memulai...");

    // ✅ Tunggu halaman transisi dan spinner benar-benar hilang
    await this.page
      .waitForLoadState("domcontentloaded", { timeout: 30000 })
      .catch(() => {});
    await this.waitForSpinnerGone(1500);
    await this.page
      .waitForLoadState("networkidle", { timeout: 10000 })
      .catch(() => {});

    // Handle checkbox (Agreement)
    try {
      const checkboxSelectors = [
        'input[type="checkbox"]',
        '[role="checkbox"]',
        ".ms-Checkbox-input",
        "#agreement-checkbox",
      ];

      const checkbox = this.page.locator(checkboxSelectors.join(", ")).first();
      if (await checkbox.count()) {
        const isChecked = await checkbox
          .evaluate((el) => {
            return (
              el.checked ||
              el.getAttribute("aria-checked") === "true" ||
              el.classList.contains("is-checked")
            );
          })
          .catch(() => false);

        if (!isChecked) {
          console.log("[INFO] Checking agreement checkbox...");
          await this.randomMouseMove();
          await checkbox.click({ force: true }).catch(() => {});
          await this.humanDelay(800);
        }
      }
    } catch (e) {
      console.log("[INFO] Checkbox handling skipped/failed:", e.message);
    }

    // Tunggu tombol enabled — pakai partial keyword yang lebih luas
    console.log("[INFO] Waiting for Start Trial/Order button to be enabled...");
    await this.runWithMonitor(
      this.page.waitForFunction(
        () => {
          const keywords = [
            "start",
            "trial",
            "mulai",
            "coba",
            "try",
            "now",
            "uji",
            "order",
            "pesan",
            "checkout",
            "buy",
            "beli",
            "place",
            "setup",
            "subscribe",
            "langganan",
            "confirm",
            "konfirmasi",
          ];

          const candidates = [
            ...document.querySelectorAll(
              'button, [role="button"], a[role="button"], input[type="submit"]',
            ),
          ];
          const btn = candidates.find((b) => {
            const text = (
              b.textContent ||
              b.value ||
              b.getAttribute("aria-label") ||
              ""
            )
              .trim()
              .toLowerCase();
            return (
              keywords.some((kw) => text.includes(kw)) &&
              text.length > 0 &&
              text.length < 60
            );
          });

          if (!btn) return false;

          const isEnabled =
            !btn.disabled &&
            btn.getAttribute("aria-disabled") !== "true" &&
            !btn.classList.contains("is-disabled") &&
            !btn.classList.contains("ms-Button--disabled");

          // Pastikan juga terlihat (tidak hidden)
          const style = window.getComputedStyle(btn);
          const isVisible =
            style.display !== "none" &&
            style.visibility !== "hidden" &&
            style.opacity !== "0";

          return isEnabled && isVisible;
        },
        { timeout: 45000 }, // Cukup 45 detik untuk nunggu tombol muncul/enabled
      ),
      45000,
    ).catch(() =>
      console.log(
        "[WARN] Could not confirm button enabled, proceeding anyway...",
      ),
    );

    await this.clickButtonWithPossibleNames([
      "Start trial",
      "Mulai uji coba",
      "Try now",
      "Coba sekarang",
      "Mulai percobaan",
      "Start free trial",
      "Start",
      "Mulai",
      "Place order",
      "Pesan sekarang",
      "Order now",
      "Checkout",
      "Selesaikan pesanan",
      "Confirm",
      "Konfirmasi",
    ]);

    console.log("[INFO] Start Trial clicked");

    await this.runWithMonitor(
      Promise.race([
        this.page.waitForNavigation({ timeout: HARD_TIMEOUT }).catch(() => {}),
        this.page.waitForLoadState("networkidle").catch(() => {}),
      ]),
    );
  }

  async clickGetStartedButton() {
    await this._logStep(15, "Klik tombol Get Started terakhir...");
    await this.waitForSpinnerGone(800);

    await this.page.evaluate(() => {
      document
        .querySelectorAll('[data-testid="spinner"], .css-100, .ms-Spinner')
        .forEach((el) => el.remove());
    });

    await this.humanDelay(700);

    await this.clickButtonWithPossibleNames([
      "Next",
      "Selanjutnya",
      "Berikutnya",
      "Get started",
      "Get Started",
      "Mulai",
      "Mulai percobaan",
    ]);

    console.log("[INFO] Next/Get Started clicked");
    await this.waitForPage();
  }

  async extractFinalDomainAccount() {
    await this._logStep(16, "Finalisasi data akun...");

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
    const domainEmail =
      rawText.match(emailRegex)?.[0] || this.extractedDomainEmail || "";
    const domainPassword = this.accountConfig.microsoftAccount.password;

    console.log("[STEP 16] Final Domain Email:", domainEmail);
    return { domainEmail, domainPassword };
  }

  // ─── Error detection ─────────────────────────────────────────────────────────

  async checkForError() {
    try {
      // 1. Check title & URL for obvious error states
      const title = await this.page.title().catch(() => "");
      const url = this.page.url().toLowerCase();

      if (
        /error|sorry|happened|wrong|failed|terjadi kesalahan/i.test(title) ||
        url.includes("error")
      ) {
        // Double check text content to avoid false positives from "error reporting" pages etc.
        const bodyText = await this.page.textContent("body").catch(() => "");
        if (
          /something went wrong|something happened|terjadi kesalahan|terjadi sesuatu/i.test(
            bodyText,
          )
        ) {
          return `Error Page Detected: ${title || url}`;
        }
      }

      // 2. Cek keberadaan iframe Arkose/Captcha secara eksplisit
      const captchaIndicators = [
        'button:has-text("solve the puzzle")',
        'h2:has-text("Protecting your account")',
      ];

      for (const selector of captchaIndicators) {
        if (
          await this.page
            .locator(selector)
            .first()
            .isVisible()
            .catch(() => false)
        ) {
          return "CAPTCHA_DETECTED: Microsoft/Arkose puzzle visible.";
        }
      }

      // 3. Cek pesan error validasi di field (data-automation-id common in Fluent UI)
      const fieldError = this.page
        .locator(
          '[data-automation-id="error-message"], [role="alert"].error, .ms-MessageBar--error',
        )
        .first();
      if (await fieldError.isVisible().catch(() => false)) {
        const msg = (await fieldError.textContent().catch(() => "")).trim();
        if (msg) return `Validation/UI Error: ${msg}`;
      }

      // 4. Cek teks di SEMUA frame (termasuk iframe tersembunyi)
      const markers = [
        "something went wrong",
        "something happened",
        "there's a problem",
        "there was a problem",
        "we're sorry",
        "we are sorry",
        "try again later",
        "try again shortly",
        "request can't be completed",
        "request cannot be completed",
        "terjadi sesuatu",
        "Terjadi kesalahan",
        "Sesuatu telah terjadi",
        "maaf, ada masalah",
        "Mohon maaf",
        "Melindungi akun Anda",
        "try a different way",
        "Protecting your account",
        "Please solve the puzzle",
        "so we know you're not a robot",
        "Selesaikan teka-teki",
        "agar kami tahu Anda bukan robot",
        "error code",
        "correlation id",
        "715-123280",
      ];

      for (const frame of this.page.frames()) {
        try {
          // evaluate textContent catches things innerText might miss (hidden/shadow)
          const frameText = await frame
            .evaluate(() => document.body?.textContent || "")
            .catch(() => "");
          if (!frameText) continue;

          const lowerFrameText = frameText.toLowerCase();
          const found = markers.find((m) =>
            lowerFrameText.includes(m.toLowerCase()),
          );
          if (found) {
            // Ambil snippet text di sekitar marker (max 150 karakter)
            const index = lowerFrameText.indexOf(found.toLowerCase());
            const snippet = frameText
              .substring(index, index + 150)
              .replace(/\s+/g, " ")
              .trim();

            console.log(
              `[ERROR] Marker "${found}" detected in frame: ${frame.url()}. Context: ${snippet}`,
            );
            return snippet || found;
          }
        } catch (e) {
          /* skip inaccessible frames */
        }
      }
    } catch (err) {
      // Ignore errors during check
    }
    return null;
  }

  async waitWithCheck(locator, timeout = HARD_TIMEOUT) {
    return await this.runWithMonitor(
      locator.waitFor({ state: "visible", timeout }),
      timeout,
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

  async executeStep(name, fn, delay = null) {
    console.log(`[STEP] ${name}`);
    this._currentStep = name;
    await fn();
    // Tunggu spinner hilang
    await this.waitForSpinnerGone();
    // ✅ Double-check error: beri jeda 1.5 detik lalu cek ulang
    // Ini mencegah false-positive dari teks error yang muncul sementara saat transisi halaman
    const firstCheck = await this.checkForError();
    if (firstCheck) {
      console.log(
        `[executeStep] Possible error after "${name}": "${firstCheck}", re-checking in 1.5s...`,
      );
      await this.humanDelay(1500);
      const recheck = await this.checkForError();
      if (recheck) {
        throw new Error(
          `MICROSOFT_ERROR: ${recheck} (Detected after step "${name}")`,
        );
      }
      console.log(`[executeStep] False positive cleared after re-check.`);
    }
    if (delay) await this.humanDelay(...delay);
  }

  async run() {
    this._currentStep = "Initializing";
    try {
      await this.executeStep(
        "Connecting to browser",
        () => this.connect(),
        [1000, 3000],
      );
      await this.executeStep(
        "Opening Microsoft page",
        () => this.openMicrosoftPage(),
        [400, 800],
      );
      await this.executeStep(
        "Clicking Try for free for target plan",
        () => this.clickTryForFreeOnTargetCard(),
        [500, 1000],
      );
      await this.executeStep(
        "Clicking product page Next",
        () => this.clickProductNextButton(),
        [300, 600],
      );
      await this.executeStep(
        "Filling email",
        () => this.fillEmail(),
        [1000, 2500],
      );
      await this.executeStep(
        "Submitting email & waiting for Setup",
        () => this.submitEmailAndWaitForSetup(),
        [400, 800],
      );
      await this.executeStep(
        "Clicking Setup Account button",
        () => this.clickSetupAccountButton(),
        [400, 800],
      );
      await this.executeStep(
        "Filling basic info",
        () => this.fillBasicInfo(),
        [1500, 3500],
      );
      await this.executeStep(
        "Confirming address (pre-password)",
        () => this.confirmAddressIfPrompted(),
        [300, 600],
      );
      await this.executeStep(
        "Filling password",
        () => this.fillPassword(),
        [400, 800],
      );
      await this.executeStep(
        "Handling optional Sign In",
        () => this.handleOptionalSignIn(),
        [400, 800],
      );
      await this.executeStep(
        "Navigating to payment page",
        async () => {
          await this.humanScroll();
          await this.randomMouseMove();
          await this.goToPaymentPage();
        },
        [400, 800],
      );
      await this.executeStep(
        "Filling VCC payment details",
        () => this.fillPaymentDetails(),
        [400, 800],
      );
      await this.executeStep("Submitting payment & waiting result", () =>
        this.submitPaymentAndWaitResult(),
      );
      await this.executeStep(
        "Confirming address (post-payment)",
        () => this.confirmAddressIfPrompted(),
        [300, 600],
      );

      if (this.accountConfig.stopPoint === "vcc_success") {
        console.log(
          "[INFO] Stop point reached: vcc_success. Finalizing account data...",
        );
        this._currentStep = "Extracting final domain account (early stop)";
        const { domainEmail, domainPassword } =
          await this.extractFinalDomainAccount();
        await this.triggerPaymentSaved();
        return { success: true, domainEmail, domainPassword };
      }

      await this.executeStep(
        "Accepting trial & clicking Start",
        () => this.acceptTrialAndStart(),
        [800, 1500],
      );
      await this.executeStep(
        "Clicking Get Started",
        () => this.clickGetStartedButton(),
        [800, 1500],
      );

      this._currentStep = "Extracting final domain account";
      const { domainEmail, domainPassword } =
        await this.extractFinalDomainAccount();

      console.log("Automation completed successfully");

      // Fallback: Pastikan saldo berkurang jika sampai tahap ini tapi sinyal tadi terlewat
      await this.triggerPaymentSaved();

      return { success: true, domainEmail, domainPassword };
    } catch (error) {
      const step = this._currentStep;
      console.error(`Automation error at step [${step}]:`, error);
      return {
        success: false,
        domainEmail: "",
        domainPassword: "",
        error: `Step - ${step}\nError: ${error.message.trim()}`,
      };
    }
  }
}

module.exports = MicrosoftBot;
