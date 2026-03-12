const TelegramBot = require("node-telegram-bot-api");
const config = require("./config");
const { processSingleAccount } = require("./index");
const connectDB = require("./db");
const { SuccessAccount, VCC, UserConfig } = require("./models");

// Wrap in an async function to allow awaiting connection
async function startBot() {
  try {
    // Connect to MongoDB
    await connectDB();

    const token = config.telegram.token;

    if (!token) {
      console.error("Please set TELEGRAM_BOT_TOKEN in .env and restart.");
      process.exit(1);
    }

    const bot = new TelegramBot(token, { polling: true });

    // Move the initialization of everything that depends on 'bot' inside
    initializeBotHandlers(bot);

    console.log("Telegram Bot with MongoDB (ms365bot) started.");
  } catch (err) {
    console.error("Failed to start bot:", err.message);
    process.exit(1);
  }
}

// Separate handlers to keep it clean
function initializeBotHandlers(bot) {

  // Sequential message queue for Telegram
  let _msgQueue = Promise.resolve();
  async function safeSendMessage(chatId, text, options = {}) {
    _msgQueue = _msgQueue.then(async () => {
      try {
        await bot.sendMessage(chatId, text, options);
      } catch (err) {
        console.error("[Telegram] Error sending message:", err.message);
        try {
          await bot.sendMessage(chatId, text.replace(/<[^>]*>?/gm, ""));
        } catch (inner) {
          console.error("[Telegram] Final fallback failed:", inner.message);
        }
      }
      await new Promise((r) => setTimeout(r, 500));
    });
    return _msgQueue;
  }

  function escapeHTML(str) {
    if (!str) return "";
    return str.toString().replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  // Memory storage for temporary interactive steps and pending accounts
  const sessions = {};

  async function getUserConfig(telegram_id) {
    let userConf = await UserConfig.findOne({ telegram_id: telegram_id.toString() });
    if (!userConf) {
      userConf = new UserConfig({
        telegram_id: telegram_id.toString(),
        microsoftUrl: config.microsoftUrl,
        concurrencyLimit: config.concurrencyLimit,
        proxyUsername: config.proxy.username,
        proxyPassword: config.proxy.password,
      });
      await userConf.save();
    }
    return userConf;
  }

  const mainMenu = {
    reply_markup: {
      keyboard: [
        [{ text: "➕ Add Account" }, { text: "💳 Add VCC" }],
        [{ text: "🚀 Generate" }, { text: "📊 Check VCC" }],
        [{ text: "📜 History" }, { text: "⚙️ Config" }],
        [{ text: "🗑️ Delete Success" }, { text: "🧹 Reset Session" }],
      ],
      resize_keyboard: true,
    },
  };

  bot.onText(/\/start/, async (msg) => {
    const chatId = msg.chat.id;
    await getUserConfig(chatId);
    sessions[chatId] = { accounts: [], step: "IDLE", running: false };

    bot.sendMessage(
      chatId,
      "Welcome to Microsoft Bot! 🤖 (MongoDB Powered)\n\nAdd accounts to the session, and they will be cleared once processed or manually reset.",
      mainMenu,
    );
  });

  bot.onText(/➕ Add Account/, (msg) => {
    const chatId = msg.chat.id;
    sessions[chatId] = sessions[chatId] || { accounts: [], step: "IDLE" };
    sessions[chatId].step = "WAIT_ACCOUNT";
    bot.sendMessage(
      chatId,
      "Send Microsoft Account data in this format (one per line):\n\n`email|firstName|lastName|companyName|companySize|phone|jobTitle|address|city|state|postalCode|country|password`",
      { parse_mode: "Markdown" },
    );
  });

  bot.onText(/💳 Add VCC/, async (msg) => {
    const chatId = msg.chat.id;
    sessions[chatId] = sessions[chatId] || { accounts: [], step: "IDLE" };
    sessions[chatId].step = "WAIT_VCC";
    const userConf = await getUserConfig(chatId);
    bot.sendMessage(
      chatId,
      `Send VCC data in this format (one per line):\n\n\`cardNumber|cvv|expMonth|expYear\` (Default saldo: ${userConf.maxAccountsPerPayment})`,
      { parse_mode: "Markdown" },
    );
  });

  bot.onText(/📊 Check VCC/, async (msg) => {
    const chatId = msg.chat.id;
    const vccs = await VCC.find({ telegram_id: chatId.toString(), saldo: { $gt: 0 }, status: "active" });

    if (vccs.length === 0) {
      return bot.sendMessage(chatId, "No active VCCs found in database.", mainMenu);
    }

    let summary = `💳 <b>Available VCCs:</b>\n\n`;
    vccs.forEach((vcc, idx) => {
      const maskedCard = `****${vcc.cardNumber.slice(-4)}`;
      summary += `${idx + 1}. <code>${maskedCard}</code> - Saldo: <b>${vcc.saldo}</b>\n`;
    });

    bot.sendMessage(chatId, summary, { parse_mode: "HTML", ...mainMenu });
  });

  bot.onText(/🧹 Reset Session/, (msg) => {
    const chatId = msg.chat.id;
    sessions[chatId] = { accounts: [], step: "IDLE", running: false };
    bot.sendMessage(chatId, "Temporary session and pending accounts have been cleared.", mainMenu);
  });

  bot.onText(/🗑️ Delete Success/, async (msg) => {
    const chatId = msg.chat.id;
    await SuccessAccount.deleteMany({ telegram_id: chatId.toString() });
    bot.sendMessage(chatId, "All success records for your account have been deleted from DB.", mainMenu);
  });

  bot.onText(/📜 History/, async (msg) => {
    const chatId = msg.chat.id;
    const history = await SuccessAccount.find({ telegram_id: chatId.toString() }).sort({ createdAt: -1 }).limit(100);

    if (history.length === 0) {
      return bot.sendMessage(chatId, "No history found in database.", mainMenu);
    }

    let summary = `📜 <b>Last 100 Success Accounts:</b>\n\n`;
    history.forEach((item, idx) => {
      summary += `${idx + 1}. <code>${item.domainEmail}</code> | <code>${item.domainPassword}</code>\n`;
    });

    bot.sendMessage(chatId, summary, { parse_mode: "HTML" });
  });

  bot.onText(/⚙️ Config/, async (msg) => {
    const chatId = msg.chat.id;
    const userConf = await getUserConfig(chatId);

    const options = {
      reply_markup: {
        inline_keyboard: [
          [{ text: "Set Microsoft URL", callback_data: `set_url` }],
          [{ text: `Concurrency: ${userConf.concurrencyLimit}`, callback_data: `set_concurrency` }],
          [{ text: `Max Per VCC: ${userConf.maxAccountsPerPayment}`, callback_data: `set_max_vcc` }],
          [{ text: "Set Proxy Username", callback_data: "set_proxy_user" }],
          [{ text: "Set Proxy Password", callback_data: "set_proxy_pass" }],
        ],
      },
    };

    bot.sendMessage(
      chatId,
      `⚙️ <b>Current Configuration:</b>\n\n` +
      `URL: <code>${userConf.microsoftUrl}</code>\n` +
      `Concurrency: ${userConf.concurrencyLimit}\n` +
      `Max Accounts/VCC: ${userConf.maxAccountsPerPayment}\n` +
      `Proxy User: <code>${userConf.proxyUsername}</code>`,
      { parse_mode: "HTML", ...options },
    );
  });

  bot.on("callback_query", async (callbackQuery) => {
    const message = callbackQuery.message;
    const chatId = message.chat.id;
    const data = callbackQuery.data;

    // Initialize session if it doesn't exist (e.g. after restart)
    sessions[chatId] = sessions[chatId] || { accounts: [], step: "IDLE", running: false };

    if (data === "set_url") {
      sessions[chatId].step = "SET_URL";
      bot.sendMessage(chatId, "Please send the new Microsoft Signup URL.");
    } else if (data === "set_concurrency") {
      sessions[chatId].step = "SET_CONCURRENCY";
      bot.sendMessage(chatId, "Please send the new concurrency limit (number).");
    } else if (data === "set_max_vcc") {
      sessions[chatId].step = "SET_MAX_VCC";
      bot.sendMessage(chatId, "Please send the maximum accounts allowed per VCC (number).");
    } else if (data === "set_proxy_user") {
      sessions[chatId].step = "SET_PROXY_USER";
      bot.sendMessage(chatId, "Please send the new Proxy Username.");
    } else if (data === "set_proxy_pass") {
      sessions[chatId].step = "SET_PROXY_PASS";
      bot.sendMessage(chatId, "Please send the new Proxy Password.");
    }
  });

  bot.onText(/🚀 Generate/, async (msg) => {
    const chatId = msg.chat.id;
    const session = sessions[chatId] || { accounts: [], running: false };

    if (session.running) {
      return bot.sendMessage(chatId, "⚠️ Automation is already running!");
    }

    if (session.accounts.length === 0) {
      return bot.sendMessage(chatId, "Please add accounts to the session first.");
    }

    const userConf = await getUserConfig(chatId);
    const vccs = await VCC.find({ telegram_id: chatId.toString(), saldo: { $gt: 0 }, status: "active" });

    if (vccs.length === 0) {
      return bot.sendMessage(chatId, "No active VCCs with balance found in database.");
    }

    session.running = true;
    sessions[chatId] = session;

    console.log(`[Batch] Starting for user ${chatId}. Concurrency: ${userConf.concurrencyLimit}, Max per VCC: ${userConf.maxAccountsPerPayment}`);
    bot.sendMessage(chatId, `🚀 Starting batch for ${session.accounts.length} accounts using ${vccs.length} VCCs (Concurrency: ${userConf.concurrencyLimit})...`);

    const runQueue = async () => {
      // Semaphore untuk batasi browser terbuka bersamaan
      let activeWorkers = 0;
      const maxWorkers = userConf.concurrencyLimit;

      const processAccount = async (accountData, currentIdx) => {
        // Cari VCC yang tersedia
        let vcc = await VCC.findOne({
          telegram_id: chatId.toString(),
          saldo: { $gt: 0 },
          status: "active",
        });

        if (!vcc) {
          await safeSendMessage(chatId, "❌ No more active VCCs with balance found.");
          return;
        }

        await safeSendMessage(chatId, `⏳ [${currentIdx}] Processing: ${escapeHTML(accountData.email)} using VCC ending in ${vcc.cardNumber.slice(-4)}`);

        const pairedData = {
          microsoftAccount: accountData,
          payment: {
            cardNumber: vcc.cardNumber,
            cvv: vcc.cvv,
            expMonth: vcc.expMonth,
            expYear: vcc.expYear,
          },
          telegram_id: chatId,
          microsoftUrl: userConf.microsoftUrl,
          proxyUsername: userConf.proxyUsername,
          proxyPassword: userConf.proxyPassword,
        };

        try {
          const onPaymentSaved = async () => {
            try {
              const updatedVcc = await VCC.findByIdAndUpdate(
                vcc._id,
                { $inc: { saldo: -1 } },
                { new: true }
              );
              if (!updatedVcc) return;
              console.log(`[DB] VCC ****${vcc.cardNumber.slice(-4)} saldo: ${vcc.saldo} → ${updatedVcc.saldo}`);
              if (updatedVcc.saldo <= 0) {
                await VCC.findByIdAndUpdate(vcc._id, { status: "inactive" });
                console.log(`[DB] VCC ****${vcc.cardNumber.slice(-4)} set inactive`);
              }
              vcc.saldo = updatedVcc.saldo;
            } catch (err) {
              console.error(`[DB] onPaymentSaved error:`, err.message);
            }
          };

          const result = await processSingleAccount(
            pairedData,
            currentIdx - 1,
            currentIdx,
            onPaymentSaved,
          );

          if (result.status === "SUCCESS") {
            let message = `✅ <b>Success [${currentIdx}]</b>\n`;
            message += `Email: <code>${escapeHTML(accountData.email)}</code>\n`;
            message += `Domain: <code>${escapeHTML(result.domainEmail)}</code>\n`;
            message += `Password: <code>${escapeHTML(result.domainPassword)}</code>\n`;
            message += `VCC Balance: ${vcc.saldo}`;
            await safeSendMessage(chatId, message, { parse_mode: "HTML" });
          } else {
            let message = `❌ <b>Failed [${currentIdx}] for ${escapeHTML(accountData.email)}</b>\n`;
            if (result.log && result.log.includes("CAPTCHA_DETECTED")) {
              message += `🚨 <b>CAPTCHA DETECTED!</b>\n`;
            }
            message += `Log: ${escapeHTML(result.log || "Unknown error")}`;
            await safeSendMessage(chatId, message, { parse_mode: "HTML" });
          }
        } catch (err) {
          await safeSendMessage(chatId, `❌ Error: ${escapeHTML(err.message)}`);
        }
      };

      try {
        let globalIdx = 0;
        const pendingPromises = new Set();

        // Loop utama — terus berjalan selama ada akun DI antrian
        while (true) {
          // Kalau ada slot kosong dan ada akun di antrian, ambil dan jalankan
          if (activeWorkers < maxWorkers && session.accounts.length > 0) {
            const accountData = session.accounts.shift();
            globalIdx++;
            const currentIdx = globalIdx;

            activeWorkers++;
            const promise = processAccount(accountData, currentIdx).finally(() => {
              activeWorkers--;
              pendingPromises.delete(promise);
            });
            pendingPromises.add(promise);

            // Jika masih ada akun lagi (baik sekarang atau yang akan ditambah),
            // tunggu 60 detik sebelum mengambil akun berikutnya
            if (session.accounts.length > 0 || activeWorkers < maxWorkers) {
              console.log(`[Queue] Waiting 60s before next account...`);
              await new Promise((r) => setTimeout(r, 60000));
            }
          } else if (activeWorkers === 0 && session.accounts.length === 0) {
            // Semua selesai
            break;
          } else {
            // Semua slot penuh atau antrian kosong tapi masih ada yang jalan
            // Tunggu sebentar lalu cek lagi (mendukung add akun saat running)
            await new Promise((r) => setTimeout(r, 2000));
          }
        }

        // Tunggu semua worker selesai
        if (pendingPromises.size > 0) {
          await Promise.all(Array.from(pendingPromises));
        }
      } catch (err) {
        console.error("[Queue Error]", err);
        await safeSendMessage(chatId, `⚠️ Queue error: ${escapeHTML(err.message)}`);
      } finally {
        session.running = false;
        bot.sendMessage(chatId, "🏁 Finished processing session accounts.", mainMenu);
      }
    };

    runQueue();
  });

  bot.on("message", async (msg) => {
    const chatId = msg.chat.id;
    const text = msg.text;

    if (!text || text.startsWith("/")) return;
    const session = sessions[chatId];
    if (!session || session.step === "IDLE") return;

    if (session.step === "WAIT_ACCOUNT") {
      const lines = text.split("\n");
      let added = 0;
      for (const line of lines) {
        const parts = line.split("|").map((s) => s.trim());
        if (parts.length >= 13) {
          session.accounts.push({
            email: parts[0],
            firstName: parts[1],
            lastName: parts[2],
            companyName: parts[3],
            companySize: parts[4],
            phone: parts[5],
            jobTitle: parts[6],
            address: parts[7],
            city: parts[8],
            state: parts[9],
            postalCode: parts[10],
            country: parts[11],
            password: parts[12],
          });
          added++;
        }
      }
      if (added > 0) {
        bot.sendMessage(chatId, `Successfully added ${added} accounts to memory. (Not in DB)`, mainMenu);
        session.step = "IDLE";
      } else {
        bot.sendMessage(chatId, "Invalid format. Use pipe-separated format.");
      }
    } else if (session.step === "WAIT_VCC") {
      const lines = text.split("\n");
      let added = 0;
      const userConf = await getUserConfig(chatId);

      for (const line of lines) {
        const parts = line.split("|").map((s) => s.trim());
        if (parts.length >= 4) {
          const existing = await VCC.findOne({ cardNumber: parts[0] });
          if (!existing) {
            const vcc = new VCC({
              cardNumber: parts[0],
              cvv: parts[1],
              expMonth: parts[2],
              expYear: parts[3],
              saldo: userConf.maxAccountsPerPayment, // Use database config
              telegram_id: chatId.toString(),
            });
            await vcc.save();
            added++;
          }
        }
      }
      if (added > 0) {
        bot.sendMessage(chatId, `Successfully added ${added} VCCs to DB.`, mainMenu);
        session.step = "IDLE";
      } else {
        bot.sendMessage(chatId, "Invalid format or VCC already exists.");
      }
    } else if (session.step === "SET_URL") {
      const userConf = await getUserConfig(chatId);
      userConf.microsoftUrl = text.trim();
      await userConf.save();
      bot.sendMessage(chatId, "Microsoft URL updated.", mainMenu);
      session.step = "IDLE";
    } else if (session.step === "SET_CONCURRENCY") {
      const num = parseInt(text);
      if (!isNaN(num)) {
        const userConf = await getUserConfig(chatId);
        userConf.concurrencyLimit = num;
        await userConf.save();
        bot.sendMessage(chatId, "Concurrency limit updated.", mainMenu);
        session.step = "IDLE";
      }
    } else if (session.step === "SET_MAX_VCC") {
      const num = parseInt(text);
      if (!isNaN(num)) {
        const userConf = await getUserConfig(chatId);
        userConf.maxAccountsPerPayment = num;
        await userConf.save();
        bot.sendMessage(chatId, "Max accounts per VCC updated.", mainMenu);
        session.step = "IDLE";
      }
    } else if (session.step === "SET_PROXY_USER") {
      const userConf = await getUserConfig(chatId);
      userConf.proxyUsername = text.trim();
      await userConf.save();
      bot.sendMessage(chatId, "Proxy Username updated.", mainMenu);
      session.step = "IDLE";
    } else if (session.step === "SET_PROXY_PASS") {
      const userConf = await getUserConfig(chatId);
      userConf.proxyPassword = text.trim();
      await userConf.save();
      bot.sendMessage(chatId, "Proxy Password updated.", mainMenu);
      session.step = "IDLE";
    }
  });
}

// Start the bot
startBot();
