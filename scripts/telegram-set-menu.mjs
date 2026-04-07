import { loadConfig } from "../lib/config.mjs";

const config = loadConfig(process.cwd());

const botToken = String(process.env.TELEGRAM_BOT_TOKEN || "").trim();
const miniAppUrl = String(process.env.TELEGRAM_MINI_APP_URL || "").trim();
const menuText = String(process.env.TELEGRAM_MENU_TEXT || "Mo mini app").trim();
const chatId = String(process.env.TELEGRAM_CHAT_ID || "").trim();

if (!botToken) {
  throw new Error("Thieu TELEGRAM_BOT_TOKEN trong .env.");
}

if (!miniAppUrl) {
  throw new Error("Thieu TELEGRAM_MINI_APP_URL trong .env.");
}

if (!/^https:\/\//i.test(miniAppUrl)) {
  throw new Error("TELEGRAM_MINI_APP_URL phai la HTTPS de Telegram mo duoc Mini App.");
}

const apiBaseUrl = `https://api.telegram.org/bot${botToken}`;

const menuButton = {
  type: "web_app",
  text: menuText || "Mo mini app",
  web_app: {
    url: miniAppUrl,
  },
};

const payload = {
  menu_button: menuButton,
};

if (chatId) {
  payload.chat_id = Number.isFinite(Number(chatId)) ? Number(chatId) : chatId;
}

console.log(`Dang cau hinh menu Telegram cho ${chatId ? `chat ${chatId}` : "mac dinh toan bot"}...`);

const response = await fetch(`${apiBaseUrl}/setChatMenuButton`, {
  method: "POST",
  headers: {
    "Content-Type": "application/json",
    Accept: "application/json",
  },
  body: JSON.stringify(payload),
});

const result = await response.json();
if (!response.ok || result.ok !== true) {
  throw new Error(result.description || `Telegram API loi ${response.status}.`);
}

console.log("Da cau hinh menu button thanh cong.");
console.log(`Mini App URL: ${miniAppUrl}`);
console.log(`Menu text: ${menuButton.text}`);
console.log(`Server web hien tai doc .env tu: ${config.cwd}`);
