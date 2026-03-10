const TelegramBot = require("node-telegram-bot-api");
const config = require("./config");
const { processSingleAccount } = require("./index");
const fs = require("fs");

const token = config.telegram.token;

if (!token) {
  console.error("Please set TELEGRAM_BOT_TOKEN in .env and restart.");
  process.exit(1);
}

const bot = new TelegramBot(token, { polling: true });

// Memory storage for user data
const sessions = {};

console.log("Telegram Bot is running...");

const mainMenu = {
  reply_markup: {
    keyboard: [
      [{ text: "➕ Add Account" }, { text: "💳 Add Payment" }],
      [{ text: "🚀 Generate" }, { text: "📊 Status" }],
      [{ text: "📜 History" }, { text: "🧹 Reset" }],
    ],
    resize_keyboard: true,
  },
};

bot.onText(/\/start/, (msg) => {
  const chatId = msg.chat.id;
  sessions[chatId] = {
    accounts: [],
    payments: [],
    step: "IDLE",
  };

  bot.sendMessage(
    chatId,
    "Welcome to Microsoft Bot! 🤖\n\nUse the menu below to manage your data:",
    mainMenu,
  );
});

bot.onText(/\/add_account|➕ Add Account/, (msg) => {
  const chatId = msg.chat.id;
  if (!sessions[chatId])
    sessions[chatId] = { accounts: [], payments: [], step: "IDLE" };

  sessions[chatId].step = "WAIT_ACCOUNT";
  bot.sendMessage(
    chatId,
    "Send Microsoft Account data in this format (one per line or one block):\n\n`email|firstName|lastName|companyName|companySize|phone|jobTitle|address|city|state|postalCode|country|password`",
    { parse_mode: "Markdown" },
  );
});

bot.onText(/\/add_payment|💳 Add Payment/, (msg) => {
  const chatId = msg.chat.id;
  if (!sessions[chatId])
    sessions[chatId] = { accounts: [], payments: [], step: "IDLE" };

  sessions[chatId].step = "WAIT_PAYMENT";
  bot.sendMessage(
    chatId,
    "Send Payment data in this format:\n\n`cardNumber|cvv|expMonth|expYear`",
    { parse_mode: "Markdown" },
  );
});

bot.onText(/\/status|📊 Status/, (msg) => {
  const chatId = msg.chat.id;
  const session = sessions[chatId];
  if (!session) return bot.sendMessage(chatId, "Please /start first.");

  bot.sendMessage(
    chatId,
    `Current Queue:\nAccounts: ${session.accounts.length}\nPayments: ${session.payments.length}`,
    mainMenu,
  );
});

bot.onText(/\/reset|🧹 Reset/, (msg) => {
  const chatId = msg.chat.id;
  sessions[chatId] = { accounts: [], payments: [], step: "IDLE" };
  bot.sendMessage(chatId, "All data has been cleared.", mainMenu);
});

bot.onText(/\/history|📜 History/, (msg) => {
  const chatId = msg.chat.id;
  const { HISTORY_FILE, EXCEL_FILE } = require("./index");

  if (!fs.existsSync(HISTORY_FILE)) {
    return bot.sendMessage(chatId, "No history found yet.", mainMenu);
  }

  try {
    const history = JSON.parse(fs.readFileSync(HISTORY_FILE, "utf8"));
    if (history.length === 0) {
      return bot.sendMessage(chatId, "History is empty.", mainMenu);
    }

    let summary = `📜 <b>Total Successful Accounts: ${history.length}</b>\n\n`;
    const last5 = history.slice(-5); // Show last 5
    last5.forEach((item, idx) => {
      summary += `${idx + 1}. <code>${item.domainEmail}</code>\n`;
    });

    if (history.length > 5) {
      summary += `\n<i>...and ${history.length - 5} others.</i>`;
    }

    bot.sendMessage(chatId, summary, { parse_mode: "HTML" });

    // Send the Excel file too
    if (fs.existsSync(EXCEL_FILE)) {
      bot.sendDocument(chatId, EXCEL_FILE, {
        caption: "Full Success Report (Excel)",
      });
    }
  } catch (e) {
    bot.sendMessage(chatId, "Error reading history.");
  }
});

bot.onText(/\/generate|🚀 Generate/, async (msg) => {
  const chatId = msg.chat.id;
  const session = sessions[chatId];

  if (
    !session ||
    session.accounts.length === 0 ||
    session.payments.length === 0
  ) {
    return bot.sendMessage(
      chatId,
      "Please add at least one account and one payment method first.",
    );
  }

  if (session.running) {
    return bot.sendMessage(
      chatId,
      "⚠️ Automation is already running! You can add more accounts now, but they will only be processed in the next run.",
    );
  }

  bot.sendMessage(
    chatId,
    `Starting automation for ${session.accounts.length} accounts...`,
  );

  const maxPerPayment = config.maxAccountsPerPayment || 3;
  const paired = [];
  const paymentUsage = session.payments.map(() => 0);

  // Round-robin distribution
  session.accounts.forEach((acc, index) => {
    let startIdx = index % session.payments.length;
    for (let i = 0; i < session.payments.length; i++) {
      const pIndex = (startIdx + i) % session.payments.length;
      if (paymentUsage[pIndex] < maxPerPayment) {
        paired.push({
          microsoftAccount: acc,
          payment: session.payments[pIndex],
        });
        paymentUsage[pIndex]++;
        break;
      }
    }
  });

  if (paired.length === 0) {
    return bot.sendMessage(
      chatId,
      "No accounts could be paired with payments.",
    );
  }

  session.running = true; // Kunci proses
  const batchSize = config.concurrencyLimit || 2;
  
  bot.sendMessage(
    chatId,
    `Paired ${paired.length} accounts. Running in batches of ${batchSize} (VCC Safety Mode)...`,
    mainMenu,
  );

  try {
    for (let i = 0; i < paired.length; i += batchSize) {
      const batch = paired.slice(i, i + batchSize);
      const batchPromises = [];

      for (let j = 0; j < batch.length; j++) {
        const account = batch[j];
        const globalIndex = i + j;

        const taskPromise = (async () => {
          bot.sendMessage(
            chatId,
            `Starting [${globalIndex + 1}/${paired.length}]: ${account.microsoftAccount.email}...`,
          );

          try {
            const result = await processSingleAccount(
              account,
              globalIndex,
              paired.length,
            );

            // If domain email is missing, mark as FAILED
            if (result.status === "SUCCESS" && !result.domainEmail) {
              result.status = "FAILED";
              result.log = "Confirmation page loaded but Domain Email not found.";
            }

            let statusEmoji = result.status === "SUCCESS" ? "✅" : "❌";

            // Escape special HTML characters from log
            const safeLog = (result.log || "Unknown error")
              .replace(/&/g, "&amp;")
              .replace(/</g, "&lt;")
              .replace(/>/g, "&gt;")
              .substring(0, 1000);

            let message = `${statusEmoji} <b>Result for ${account.microsoftAccount.email}</b>\n\n`;
            message += `<b>Status:</b> ${result.status}\n`;

            if (result.status === "SUCCESS") {
              message += `<b>Domain Email:</b> <code>${result.domainEmail}</code>\n`;
              message += `<b>Domain Password:</b> <code>${result.domainPassword}</code>\n`;
            } else {
              message += `<b>Log:</b> ${safeLog}\n`;
            }

            await bot
              .sendMessage(chatId, message, { parse_mode: "HTML" })
              .catch((e) => {
                return bot.sendMessage(
                  chatId,
                  `❌ Result for ${account.microsoftAccount.email}\nStatus: ${result.status}\nError: ${safeLog.substring(0, 200)}`,
                );
              });
          } catch (err) {
            bot.sendMessage(
              chatId,
              `❌ System Error for ${account.microsoftAccount.email}: ${err.message.substring(0, 200)}`,
            );
          }
        })();

        batchPromises.push(taskPromise);

        if (j < batch.length - 1) {
          await new Promise((resolve) => setTimeout(resolve, 5000));
        }
      }

      await Promise.all(batchPromises);

      if (i + batchSize < paired.length) {
        bot.sendMessage(
          chatId,
          `Batch finished. Waiting for next batch to start...`,
        );
        await new Promise((resolve) => setTimeout(resolve, 2000));
      }
    }
  } finally {
    session.running = false; // Buka kunci proses
    // Hapus akun yang sudah diproses dari antrean
    session.accounts = session.accounts.slice(paired.length);
  }

  bot.sendMessage(chatId, "🏁 All tasks finished!", mainMenu);
});

bot.on("message", (msg) => {
  const chatId = msg.chat.id;
  const text = msg.text;

  if (!text) return;
  if (
    text.startsWith("/") ||
    text.includes("Add Account") ||
    text.includes("Add Payment") ||
    text.includes("Generate") ||
    text.includes("Status") ||
    text.includes("History") ||
    text.includes("Reset")
  )
    return;

  const session = sessions[chatId];
  if (!session || session.step === "IDLE") return;

  if (session.step === "WAIT_ACCOUNT") {
    let addedCount = 0;

    // Check if input is JSON
    if (text.trim().startsWith("[") || text.trim().startsWith("{")) {
      try {
        const jsonData = JSON.parse(text);
        const accounts = Array.isArray(jsonData) ? jsonData : [jsonData];

        accounts.forEach((acc) => {
          if (acc.email && acc.password) {
            session.accounts.push({
              email: acc.email,
              firstName: acc.firstName || "",
              lastName: acc.lastName || "",
              companyName: acc.companyName || "",
              phone: acc.phone || "",
              jobTitle: acc.jobTitle || "",
              address: acc.address || "",
              city: acc.city || "",
              state: acc.state || "",
              postalCode: acc.postalCode || "",
              country: acc.country || "",
              password: acc.password,
              companySize: acc.companySize || "1 person",
            });
            addedCount++;
          }
        });
      } catch (e) {
        console.error("JSON Parse Error:", e.message);
      }
    }

    // Fallback to pipe-separated format if no accounts added via JSON
    if (addedCount === 0) {
      const lines = text.split("\n");
      lines.forEach((line) => {
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
          addedCount++;
        }
      });
    }

    if (addedCount > 0) {
      bot.sendMessage(
        chatId,
        `Successfully added ${addedCount} accounts. Total: ${session.accounts.length}`,
        mainMenu,
      );
      session.step = "IDLE";
    } else {
      bot.sendMessage(
        chatId,
        "Format invalid. Please use pipe-separated format or a JSON array of accounts.",
        { parse_mode: "Markdown", ...mainMenu },
      );
    }
  } else if (session.step === "WAIT_PAYMENT") {
    const lines = text.split("\n");
    let addedCount = 0;
    lines.forEach((line) => {
      const parts = line.split("|").map((s) => s.trim());
      if (parts.length >= 4) {
        session.payments.push({
          cardNumber: parts[0],
          cvv: parts[1],
          expMonth: parts[2],
          expYear: parts[3],
        });
        addedCount++;
      }
    });

    if (addedCount > 0) {
      bot.sendMessage(
        chatId,
        `Successfully added ${addedCount} payment methods. Total: ${session.payments.length}`,
        mainMenu,
      );
      session.step = "IDLE";
    } else {
      bot.sendMessage(
        chatId,
        "Format invalid. Please use: `cardNumber|cvv|expMonth|expYear`",
        { parse_mode: "Markdown", ...mainMenu },
      );
    }
  }
});
