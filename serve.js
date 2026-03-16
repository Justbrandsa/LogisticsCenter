const http = require("http");
const fs = require("fs");
const os = require("os");
const path = require("path");

let Pool = null;
let driverLoadError = null;
let nodemailer = null;
let mailerLoadError = null;

try {
  ({ Pool } = require("pg"));
} catch (error) {
  driverLoadError = error;
}

try {
  nodemailer = require("nodemailer");
} catch (error) {
  mailerLoadError = error;
}

const HOST = "0.0.0.0";
const PORT = Number(process.env.PORT || 4173);
const ROOT = __dirname;
const MAX_BODY_BYTES = 1024 * 1024;
const MIME = {
  ".css": "text/css; charset=UTF-8",
  ".html": "text/html; charset=UTF-8",
  ".js": "application/javascript; charset=UTF-8",
  ".json": "application/json; charset=UTF-8",
};
const RPC_DEFINITIONS = Object.freeze({
  get_login_state: { params: [] },
  bootstrap_admin: { params: ["p_name", "p_password"] },
  login_user: { params: ["p_name", "p_password"] },
  logout_user: { params: ["p_token"] },
  run_daily_rollover: { params: [] },
  get_app_snapshot: { params: ["p_token"] },
  create_user_account: {
    params: ["p_token", "p_name", "p_password", "p_role", "p_phone"],
  },
  toggle_user_active: { params: ["p_token", "p_user_id"] },
  delete_user_account: { params: ["p_token", "p_user_id"] },
  create_supplier: { params: ["p_token", "p_name"] },
  delete_supplier: { params: ["p_token", "p_supplier_id"] },
  create_location: {
    params: ["p_token", "p_supplier_id", "p_name", "p_address", "p_lat", "p_lng", "p_notes"],
  },
  delete_location: { params: ["p_token", "p_location_id"] },
  create_order: {
    params: [
      "p_token",
      "p_driver_user_id",
      "p_location_id",
      "p_entry_type",
      "p_factory_order_number",
      "p_inhouse_order_number",
      "p_allow_duplicate",
      "p_notice",
    ],
  },
  complete_order: { params: ["p_token", "p_order_id"] },
  delete_order: { params: ["p_token", "p_order_id"] },
});

const database = createDatabase();
const mailer = createMailer();
const server = http.createServer((request, response) => {
  void routeRequest(request, response);
});

server.listen(PORT, HOST, () => {
  const networkUrls = Object.values(os.networkInterfaces())
    .flat()
    .filter((entry) => entry && entry.family === "IPv4" && !entry.internal)
    .map((entry) => `http://${entry.address}:${PORT}`);

  const status = database.getStatus();
  const mailStatus = mailer.getStatus();

  console.log(`Route Ledger available at http://127.0.0.1:${PORT}`);
  console.log(
    status.configured
      ? "Neon database connection is configured."
      : `Neon database connection is not configured yet: ${status.reason}`,
  );
  console.log(
    mailStatus.configured
      ? `CSV mail-out is configured via ${mailStatus.provider} for ${mailStatus.from} -> ${mailStatus.to}.`
      : `CSV mail-out is not configured yet: ${mailStatus.reason}`,
  );
  networkUrls.forEach((url) => {
    console.log(`Route Ledger network URL: ${url}`);
  });
});

process.on("SIGINT", () => {
  void database.close().finally(() => process.exit(0));
});

process.on("SIGTERM", () => {
  void database.close().finally(() => process.exit(0));
});

async function routeRequest(request, response) {
  const requestUrl = new URL(request.url || "/", `http://${request.headers.host || "127.0.0.1"}`);
  const cleanPath = decodeURIComponent(requestUrl.pathname || "/");

  try {
    if (request.method === "GET" && cleanPath === "/api/status") {
      sendJson(response, 200, {
        ...database.getStatus(),
        mailConfigured: mailer.getStatus().configured,
        mailReason: mailer.getStatus().reason,
        mailProvider: mailer.getStatus().provider,
        mailFrom: mailer.getStatus().from,
        mailTo: mailer.getStatus().to,
      });
      return;
    }

    if (request.method === "POST" && cleanPath.startsWith("/api/rpc/")) {
      const functionName = cleanPath.slice("/api/rpc/".length);
      const payload = await readJsonBody(request);
      const data = await database.call(functionName, payload?.parameters || {});
      sendJson(response, 200, { data });
      return;
    }

    if (request.method === "POST" && cleanPath === "/api/export/email") {
      const payload = await readJsonBody(request);
      const result = await mailer.sendSnapshotCsv(payload?.token);
      sendJson(response, 200, result);
      return;
    }

    if (request.method === "POST" && cleanPath === "/api/export/email/test") {
      const payload = await readJsonBody(request);
      const result = await mailer.sendTestEmail(payload?.token);
      sendJson(response, 200, result);
      return;
    }

    if (request.method !== "GET" && request.method !== "HEAD") {
      sendJson(response, 405, { error: "Method not allowed." });
      return;
    }

    await serveStaticFile(cleanPath, request.method === "HEAD", response);
  } catch (error) {
    const statusCode = Number.isInteger(error?.statusCode) ? error.statusCode : 500;
    const message = statusCode >= 500 ? "Server error." : normalizeErrorMessage(error);

    if (statusCode >= 500) {
      console.error(error);
    }

    sendJson(response, statusCode, { error: message });
  }
}

function createDatabase() {
  const config = loadDatabaseConfig();
  const status = {
    configured: false,
    reason: "",
  };

  if (driverLoadError) {
    status.reason = "Run npm install to add the PostgreSQL driver.";
  } else if (!config.connectionString) {
    status.reason = "Add a Neon connection string in neon-config.js or DATABASE_URL.";
  } else {
    status.configured = true;
    status.reason = "";
  }

  const pool = status.configured
    ? new Pool({
        connectionString: config.connectionString,
        ssl: shouldUseTls(config.connectionString) ? { rejectUnauthorized: false } : undefined,
      })
    : null;

  return {
    getStatus() {
      return { ...status };
    },
    async call(functionName, parameters) {
      if (!pool) {
        throw createHttpError(503, status.reason || "Database connection is not configured.");
      }

      const definition = RPC_DEFINITIONS[functionName];
      if (!definition) {
        throw createHttpError(404, "Unknown RPC function.");
      }

      const values = definition.params.map((parameterName) =>
        Object.prototype.hasOwnProperty.call(parameters, parameterName) ? parameters[parameterName] : null,
      );
      const query = buildRpcQuery(functionName, definition.params.length);
      let result;

      try {
        result = await pool.query(query, values);
      } catch (error) {
        if (typeof error?.code === "string") {
          throw createHttpError(400, normalizeErrorMessage(error));
        }
        throw error;
      }

      return result.rows[0]?.data ?? null;
    },
    async close() {
      if (pool) {
        await pool.end();
      }
    },
  };
}

function createMailer() {
  const config = loadMailConfig();
  const status = {
    configured: false,
    reason: "",
    from: config.from,
    to: config.to,
    provider: config.provider,
  };

  if (config.provider === "microsoft-graph") {
    if (!config.tenantId || !config.clientId || !config.clientSecret) {
      status.reason = "Add Microsoft Graph mail settings in mail-config.js or MAIL_* environment variables.";
    } else {
      status.configured = true;
    }
  } else if (config.provider === "smtp") {
    if (mailerLoadError) {
      status.reason = "Run npm install to add the mailer dependency.";
    } else if (!config.transport.auth?.user || !config.transport.auth?.pass) {
      status.reason = "Add SMTP settings in mail-config.js or SMTP_* environment variables.";
    } else if (!config.transport.service && !config.transport.host) {
      status.reason = "SMTP host or service is required for CSV email delivery.";
    } else {
      status.configured = true;
    }
  } else {
    status.reason = "Unsupported mail provider. Use microsoft-graph or smtp.";
  }

  const transporter = status.configured && config.provider === "smtp"
    ? nodemailer.createTransport(config.transport)
    : null;

  async function getAuthorizedMailContext(token) {
    if (!status.configured) {
      throw createHttpError(503, status.reason || "Email delivery is not configured.");
    }

    if (!token) {
      throw createHttpError(400, "Session token is required.");
    }

    const snapshot = await database.call("get_app_snapshot", { p_token: token });
    const currentUser = snapshot?.user || null;

    if (!currentUser) {
      throw createHttpError(403, "You must be signed in to send email.");
    }

    if (!["admin", "sales"].includes(currentUser.role)) {
      throw createHttpError(403, "Only admin or sales users can send email.");
    }

    return {
      currentUser,
      snapshot,
    };
  }

  async function sendMessage(message) {
    if (config.provider === "microsoft-graph") {
      await sendViaMicrosoftGraph(config, message);
      return;
    }

    await transporter.sendMail({
      from: message.from,
      to: message.to,
      subject: message.subject,
      text: message.text,
      attachments: Array.isArray(message.attachments) ? message.attachments : [],
    });
  }

  return {
    getStatus() {
      return { ...status };
    },
    async sendSnapshotCsv(token) {
      const { currentUser, snapshot } = await getAuthorizedMailContext(token);

      const orders = Array.isArray(snapshot?.orders) ? snapshot.orders : [];
      const dateStamp = new Date().toISOString().slice(0, 10);
      const filename = `route-ledger-${dateStamp}.csv`;
      const subject = `Route Ledger CSV export ${dateStamp}`;

      const text = [
        "Attached is the latest Route Ledger CSV export.",
        "",
        `Sent by: ${currentUser.name}`,
        `Entries included: ${orders.length}`,
      ].join("\n");
      const csv = buildOrdersCsv(orders);

      await sendMessage({
        from: status.from,
        to: status.to,
        subject,
        text,
        attachments: [
          {
            filename,
            content: csv,
            contentType: "text/csv; charset=UTF-8",
          },
        ],
      });

      return {
        ok: true,
        sentTo: status.to,
        sentFrom: status.from,
        count: orders.length,
      };
    },
    async sendTestEmail(token) {
      const { currentUser } = await getAuthorizedMailContext(token);
      const timestamp = new Date().toISOString();
      const subject = `Route Ledger test email ${timestamp}`;
      const text = [
        "This is a Route Ledger email test.",
        "",
        `Sent by: ${currentUser.name}`,
        `Provider: ${status.provider}`,
        `From: ${status.from}`,
        `To: ${status.to}`,
        `Timestamp: ${timestamp}`,
      ].join("\n");

      await sendMessage({
        from: status.from,
        to: status.to,
        subject,
        text,
      });

      return {
        ok: true,
        sentTo: status.to,
        sentFrom: status.from,
        subject,
      };
    },
  };
}

function loadDatabaseConfig() {
  const filePath = path.join(ROOT, "neon-config.js");
  let fileConfig = {};

  if (fs.existsSync(filePath)) {
    try {
      const resolvedPath = require.resolve(filePath);
      delete require.cache[resolvedPath];
      const loadedConfig = require(resolvedPath);
      fileConfig = typeof loadedConfig === "string" ? { connectionString: loadedConfig } : loadedConfig || {};
    } catch (error) {
      console.error("Failed to load neon-config.js", error);
    }
  }

  const connectionString = [
    process.env.NEON_DATABASE_URL,
    process.env.DATABASE_URL,
    fileConfig.connectionString,
  ]
    .find((value) => typeof value === "string" && value.trim())
    ?.trim();

  return {
    connectionString: connectionString || "",
  };
}

function loadMailConfig() {
  const filePath = path.join(ROOT, "mail-config.js");
  let fileConfig = {};

  if (fs.existsSync(filePath)) {
    try {
      const resolvedPath = require.resolve(filePath);
      delete require.cache[resolvedPath];
      fileConfig = require(resolvedPath) || {};
    } catch (error) {
      console.error("Failed to load mail-config.js", error);
    }
  }

  const provider = String(
    firstNonEmpty([
      process.env.MAIL_PROVIDER,
      fileConfig.provider,
      "microsoft-graph",
    ]) || "microsoft-graph",
  ).toLowerCase();
  const from = firstNonEmpty([
    process.env.MAIL_FROM,
    process.env.SMTP_FROM,
    fileConfig.from,
    "artwork3@giftwrap.co.za",
  ]);
  const to = firstNonEmpty([
    process.env.MAIL_TO,
    process.env.SMTP_TO,
    fileConfig.to,
    "admin3@giftwrap.co.za",
  ]);
  if (provider === "smtp") {
    const service = firstNonEmpty([
      process.env.SMTP_SERVICE,
      fileConfig.service,
    ]);
    const host = firstNonEmpty([
      process.env.SMTP_HOST,
      fileConfig.host,
    ]);
    const portValue = firstNonEmpty([
      process.env.SMTP_PORT,
      fileConfig.port,
    ]);
    const secureValue = firstNonEmpty([
      process.env.SMTP_SECURE,
      fileConfig.secure,
    ]);
    const user = firstNonEmpty([
      process.env.SMTP_USER,
      fileConfig.auth?.user,
    ]);
    const pass = firstNonEmpty([
      process.env.SMTP_PASS,
      fileConfig.auth?.pass,
    ]);

    const transport = {
      auth: {
        user: user || "",
        pass: pass || "",
      },
    };

    if (service) {
      transport.service = service;
    } else {
      transport.host = host || "";
      transport.port = Number(portValue || 587);
      transport.secure = normalizeBoolean(secureValue, transport.port === 465);
    }

    return {
      provider,
      transport,
      from,
      to,
    };
  }

  return {
    provider,
    from,
    to,
    tenantId: firstNonEmpty([
      process.env.MAIL_TENANT_ID,
      fileConfig.tenantId,
    ]) || "",
    clientId: firstNonEmpty([
      process.env.MAIL_CLIENT_ID,
      fileConfig.clientId,
    ]) || "",
    clientSecret: firstNonEmpty([
      process.env.MAIL_CLIENT_SECRET,
      fileConfig.clientSecret,
    ]) || "",
  };
}

function shouldUseTls(connectionString) {
  try {
    const parsed = new URL(connectionString);
    const sslMode = (parsed.searchParams.get("sslmode") || "").toLowerCase();

    if (sslMode === "disable") {
      return false;
    }

    return parsed.protocol.startsWith("postgres");
  } catch (error) {
    return true;
  }
}

function buildRpcQuery(functionName, parameterCount) {
  if (!RPC_DEFINITIONS[functionName]) {
    throw createHttpError(404, "Unknown RPC function.");
  }

  if (parameterCount === 0) {
    return `select public.${functionName}() as data`;
  }

  const placeholders = Array.from({ length: parameterCount }, (_, index) => `$${index + 1}`).join(", ");
  return `select public.${functionName}(${placeholders}) as data`;
}

function buildOrdersCsv(orders) {
  const rows = [
    [
      "Reference",
      "Driver",
      "Pickup location",
      "Collection or delivery",
      "Factory order number",
      "In-house order number",
      "Created by",
      "Status",
      "Notice",
      "Created at",
      "Completed at",
    ],
    ...orders.map((order) => [
      order.reference || "",
      order.driverName || "",
      order.locationName || "",
      order.entryType || "",
      order.factoryOrderNumber || "",
      order.inhouseOrderNumber || "",
      order.createdByName || "",
      order.status || "",
      buildOrderNoticeText(order),
      order.createdAt || "",
      order.completedAt || "",
    ]),
  ];

  return rows.map((columns) => columns.map(escapeCsvValue).join(",")).join("\n");
}

function buildOrderNoticeText(order) {
  const lines = [];
  const notice = String(order?.notes || "").trim();
  const carryOverCount = Number(order?.carryOverCount || 0);

  if (notice) {
    lines.push(notice);
  }

  if (carryOverCount > 0) {
    const scheduledFor = String(order?.scheduledFor || "").trim();
    const originalScheduledFor = String(order?.originalScheduledFor || "").trim();
    const dayLabel = carryOverCount === 1 ? "day" : "days";

    if (scheduledFor && originalScheduledFor) {
      lines.push(`Rolled to ${scheduledFor} from ${originalScheduledFor} (${carryOverCount} ${dayLabel}).`);
    } else {
      lines.push(`Rolled to the next day (${carryOverCount} ${dayLabel}).`);
    }
  }

  return lines.join(" | ");
}

async function sendViaMicrosoftGraph(config, message) {
  const token = await getMicrosoftGraphAccessToken(config);
  const attachments = Array.isArray(message.attachments)
    ? message.attachments.map((attachment) => ({
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: attachment.filename,
        contentType: attachment.contentType || "application/octet-stream",
        contentBytes: Buffer.from(attachment.content || "", "utf8").toString("base64"),
      }))
    : [];
  const payload = {
    message: {
      subject: message.subject,
      body: {
        contentType: "Text",
        content: message.text,
      },
      toRecipients: [
        {
          emailAddress: {
            address: message.to,
          },
        },
      ],
    },
    saveToSentItems: true,
  };

  if (attachments.length) {
    payload.message.attachments = attachments;
  }

  const response = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(message.from)}/sendMail`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    const payload = await safeReadJson(response);
    throw createHttpError(502, payload?.error?.message || "Microsoft Graph mail delivery failed.");
  }
}

async function getMicrosoftGraphAccessToken(config) {
  const body = new URLSearchParams({
    client_id: config.clientId,
    client_secret: config.clientSecret,
    grant_type: "client_credentials",
    scope: "https://graph.microsoft.com/.default",
  });
  const response = await fetch(`https://login.microsoftonline.com/${encodeURIComponent(config.tenantId)}/oauth2/v2.0/token`, {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body,
  });

  const payload = await safeReadJson(response);
  if (!response.ok) {
    throw createHttpError(502, payload?.error_description || "Failed to acquire a Microsoft Graph access token.");
  }

  return payload?.access_token || "";
}

async function readJsonBody(request) {
  const chunks = [];
  let size = 0;

  for await (const chunk of request) {
    size += chunk.length;
    if (size > MAX_BODY_BYTES) {
      throw createHttpError(413, "Request body is too large.");
    }
    chunks.push(chunk);
  }

  if (chunks.length === 0) {
    return {};
  }

  try {
    return JSON.parse(Buffer.concat(chunks).toString("utf8"));
  } catch (error) {
    throw createHttpError(400, "Request body must be valid JSON.");
  }
}

async function safeReadJson(response) {
  try {
    return await response.json();
  } catch (error) {
    return null;
  }
}

function firstNonEmpty(values) {
  return values.find((value) => {
    if (typeof value === "boolean" || typeof value === "number") {
      return true;
    }

    return typeof value === "string" && value.trim();
  });
}

function normalizeBoolean(value, fallback) {
  if (typeof value === "boolean") {
    return value;
  }

  if (typeof value === "string") {
    const normalized = value.trim().toLowerCase();
    if (["true", "1", "yes", "on"].includes(normalized)) {
      return true;
    }
    if (["false", "0", "no", "off"].includes(normalized)) {
      return false;
    }
  }

  return fallback;
}

function escapeCsvValue(value) {
  return `"${String(value || "").replaceAll('"', '""')}"`;
}

async function serveStaticFile(cleanPath, headOnly, response) {
  const relativePath = cleanPath === "/" ? "index.html" : cleanPath.replace(/^\/+/, "");
  const filePath = path.resolve(ROOT, relativePath);
  const relativeToRoot = path.relative(ROOT, filePath);

  if (relativeToRoot.startsWith("..") || path.isAbsolute(relativeToRoot)) {
    sendText(response, 403, "Forbidden");
    return;
  }

  let contents;
  try {
    contents = await fs.promises.readFile(filePath);
  } catch (error) {
    sendText(response, 404, "Not found");
    return;
  }

  const extension = path.extname(filePath).toLowerCase();
  response.writeHead(200, {
    "Content-Type": MIME[extension] || "application/octet-stream",
  });
  response.end(headOnly ? undefined : contents);
}

function sendJson(response, statusCode, payload) {
  response.writeHead(statusCode, {
    "Cache-Control": "no-store",
    "Content-Type": "application/json; charset=UTF-8",
  });
  response.end(JSON.stringify(payload));
}

function sendText(response, statusCode, body) {
  response.writeHead(statusCode, {
    "Content-Type": "text/plain; charset=UTF-8",
  });
  response.end(body);
}

function createHttpError(statusCode, message) {
  const error = new Error(message);
  error.statusCode = statusCode;
  return error;
}

function normalizeErrorMessage(error) {
  if (!error) {
    return "Something went wrong.";
  }

  if (typeof error === "string") {
    return error;
  }

  return error.message || "Something went wrong.";
}
