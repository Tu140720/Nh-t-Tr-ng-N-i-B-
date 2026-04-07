import tls from "node:tls";

export async function sendMail(config, message) {
  if (!config?.host || !config?.user || !config?.pass || !config?.from) {
    throw new Error("Thieu cau hinh SMTP de gui ma xac thuc qua email.");
  }

  if (!config.secure) {
    throw new Error("He thong chi ho tro SMTP secure (TLS) cho chuc nang gui ma xac thuc.");
  }

  const socket = tls.connect({
    host: config.host,
    port: config.port || 465,
    servername: config.host,
  });

  socket.setEncoding("utf8");

  const response = createResponseReader(socket);

  await waitForConnect(socket);
  await response.read([220]);
  await sendCommand(socket, response, `EHLO localhost`);
  await sendCommand(socket, response, `AUTH LOGIN`, [334]);
  await sendCommand(socket, response, Buffer.from(config.user).toString("base64"), [334]);
  await sendCommand(socket, response, Buffer.from(config.pass).toString("base64"), [235]);
  await sendCommand(socket, response, `MAIL FROM:<${config.from}>`);
  await sendCommand(socket, response, `RCPT TO:<${message.to}>`);
  await sendCommand(socket, response, `DATA`, [354]);

  const data = [
    `From: ${config.from}`,
    `To: ${message.to}`,
    `Subject: ${encodeSubject(message.subject)}`,
    "MIME-Version: 1.0",
    "Content-Type: text/plain; charset=utf-8",
    "",
    normalizeBody(message.text),
    ".",
  ].join("\r\n");

  socket.write(`${data}\r\n`);
  await response.read([250]);
  await sendCommand(socket, response, `QUIT`, [221]);
  socket.end();
}

function createResponseReader(socket) {
  let buffer = "";
  const pending = [];

  socket.on("data", (chunk) => {
    buffer += chunk;
    flush();
  });

  socket.on("error", (error) => {
    while (pending.length > 0) {
      pending.shift().reject(error);
    }
  });

  function flush() {
    while (pending.length > 0) {
      const match = readCompleteResponse(buffer);
      if (!match) {
        return;
      }
      buffer = buffer.slice(match.length);
      pending.shift().resolve(match);
    }
  }

  return {
    read(expectedCodes = [250]) {
      return new Promise((resolve, reject) => {
        pending.push({
          resolve: (message) => {
            const code = Number.parseInt(message.slice(0, 3), 10);
            if (!expectedCodes.includes(code)) {
              reject(new Error(`SMTP loi ${code}: ${message.trim()}`));
              return;
            }
            resolve(message);
          },
          reject,
        });
        flush();
      });
    },
  };
}

function readCompleteResponse(buffer) {
  const lines = buffer.split("\r\n");
  if (lines.length < 2) {
    return "";
  }

  let consumed = 0;
  for (const line of lines.slice(0, -1)) {
    consumed += line.length + 2;
    if (/^\d{3}\s/.test(line)) {
      return buffer.slice(0, consumed);
    }
    if (!/^\d{3}-/.test(line)) {
      return buffer.slice(0, consumed);
    }
  }

  return "";
}

async function sendCommand(socket, response, command, expectedCodes = [250]) {
  socket.write(`${command}\r\n`);
  await response.read(expectedCodes);
}

function waitForConnect(socket) {
  return new Promise((resolve, reject) => {
    socket.once("secureConnect", resolve);
    socket.once("error", reject);
  });
}

function encodeSubject(value) {
  return `=?UTF-8?B?${Buffer.from(String(value || ""), "utf8").toString("base64")} ?=`;
}

function normalizeBody(value) {
  return String(value || "")
    .replace(/\r?\n/g, "\r\n")
    .replace(/^\./gm, "..");
}
