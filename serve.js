const http = require("http");
const fs = require("fs");
const os = require("os");
const path = require("path");

let Pool = null;
let driverLoadError = null;
let nodemailer = null;
let mailerLoadError = null;
let QRCode = null;
let qrLoadError = null;

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

try {
  QRCode = require("qrcode");
} catch (error) {
  qrLoadError = error;
}

const HOST = "0.0.0.0";
const PORT = Number(process.env.PORT || 4173);
const ROOT = __dirname;
const MAX_BODY_BYTES = 1024 * 1024;
const LIVE_RELOAD_FILES = new Set(["index.html", "app.js", "styles.css"]);
const MIME = {
  ".css": "text/css; charset=UTF-8",
  ".html": "text/html; charset=UTF-8",
  ".js": "application/javascript; charset=UTF-8",
  ".json": "application/json; charset=UTF-8",
};
const UUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
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
  update_user_account: {
    params: ["p_token", "p_user_id", "p_name", "p_role", "p_phone", "p_password"],
  },
  toggle_user_active: { params: ["p_token", "p_user_id"] },
  delete_user_account: { params: ["p_token", "p_user_id"] },
  create_supplier: { params: ["p_token", "p_name", "p_contact_person", "p_contact_number", "p_factory"] },
  update_supplier: { params: ["p_token", "p_supplier_id", "p_name", "p_contact_person", "p_contact_number", "p_factory"] },
  delete_supplier: { params: ["p_token", "p_supplier_id"] },
  create_location: {
    params: ["p_token", "p_location_type", "p_name", "p_address", "p_lat", "p_lng", "p_contact_person", "p_contact_number"],
  },
  update_location: {
    params: ["p_token", "p_location_id", "p_location_type", "p_name", "p_address", "p_lat", "p_lng", "p_contact_person", "p_contact_number"],
  },
  delete_location: { params: ["p_token", "p_location_id"] },
  create_stock_item: {
    params: ["p_token", "p_name", "p_sku", "p_quote_number", "p_invoice_number", "p_sales_order_number", "p_po_number", "p_unit", "p_notes"],
  },
  update_stock_item: {
    params: ["p_token", "p_stock_item_id", "p_name", "p_sku", "p_quote_number", "p_invoice_number", "p_sales_order_number", "p_po_number", "p_unit", "p_notes"],
  },
  delete_stock_item: { params: ["p_token", "p_stock_item_id"] },
  record_stock_movement: {
    params: ["p_token", "p_stock_item_id", "p_movement_type", "p_quantity", "p_supplier_name", "p_driver_user_id", "p_notes"],
  },
  update_stock_movement: {
    params: ["p_token", "p_stock_movement_id", "p_stock_item_id", "p_movement_type", "p_quantity", "p_supplier_name", "p_driver_user_id", "p_notes"],
  },
  create_artwork_request: { params: ["p_token", "p_stock_item_id", "p_requested_quantity", "p_notes", "p_sent_to"] },
  create_order: {
    params: [
      "p_token",
      "p_driver_user_id",
      "p_location_id",
      "p_entry_type",
      "p_quote_number",
      "p_sales_order_number",
      "p_invoice_number",
      "p_po_number",
      "p_allow_duplicate",
      "p_notice",
      "p_move_to_factory",
      "p_factory_destination_location_id",
    ],
  },
  assign_order: { params: ["p_token", "p_order_id", "p_driver_user_id", "p_allow_duplicate"] },
  set_order_priority: { params: ["p_token", "p_order_id", "p_priority"] },
  set_order_flag: { params: ["p_token", "p_order_id", "p_flag_type", "p_note"] },
  complete_order: { params: ["p_token", "p_order_id"] },
  delete_order: { params: ["p_token", "p_order_id"] },
});

const database = createDatabase();
const mailer = createMailer();
const liveReloadClients = new Set();

let liveReloadWatcherStarted = false;
let liveReloadTimer = null;
let liveReloadPendingFile = "";

function startServer() {
  startLiveReloadWatcher();

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
}

if (require.main === module) {
  startServer();
}

async function routeRequest(request, response) {
  const requestUrl = new URL(request.url || "/", `http://${request.headers.host || "127.0.0.1"}`);
  const cleanPath = decodeURIComponent(requestUrl.pathname || "/");

  try {
    if (request.method === "GET" && cleanPath === "/__live-reload") {
      handleLiveReloadStream(request, response);
      return;
    }

    if (request.method === "GET" && cleanPath === "/api/status") {
      sendJson(response, 200, {
        ...database.getStatus(),
        mailConfigured: mailer.getStatus().configured,
        mailReason: mailer.getStatus().reason,
        mailProvider: mailer.getStatus().provider,
        mailFrom: mailer.getStatus().from,
        mailTo: mailer.getStatus().to,
        artworkTo: mailer.getStatus().artworkTo,
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

    if (request.method === "POST" && cleanPath === "/api/artwork/request") {
      const payload = await readJsonBody(request);
      const result = await mailer.sendArtworkRequest(payload?.token, payload || {});
      sendJson(response, 200, result);
      return;
    }

    if (
      request.method === "POST"
      && (cleanPath === "/api/barcode/svg" || cleanPath === "/api/qr/svg")
    ) {
      const payload = await readJsonBody(request);
      const text = String(payload?.text || "").trim();
      const svg = await createQrSvg(text);
      sendSvg(response, 200, svg);
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
    const message = statusCode === 500 ? "Server error." : normalizeErrorMessage(error);

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
    artworkTo: config.artworkTo,
    provider: config.provider,
  };

  if (config.provider === "microsoft-graph") {
    const graphConfigIssue = getMicrosoftGraphConfigIssue(config);
    if (graphConfigIssue) {
      status.reason = graphConfigIssue;
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

  async function getAuthorizedMailContext(token, allowedRoles, deniedMessage) {
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

    if (!Array.isArray(allowedRoles) || !allowedRoles.includes(currentUser.role)) {
      throw createHttpError(403, deniedMessage || "You do not have permission to send email.");
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
      const { currentUser, snapshot } = await getAuthorizedMailContext(
        token,
        ["admin", "sales"],
        "Only admin or sales users can send email.",
      );

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
      const { currentUser } = await getAuthorizedMailContext(
        token,
        ["admin", "sales"],
        "Only admin or sales users can send email.",
      );
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
    async sendArtworkRequest(token, payload) {
      const { currentUser, snapshot } = await getAuthorizedMailContext(
        token,
        ["admin", "logistics"],
        "Only admin or logistics users can send artwork requests.",
      );

      const stockItemId = String(payload?.stockItemId || "").trim();
      const requestedQuantity = Number(payload?.requestedQuantity || 0);
      const notes = String(payload?.notes || "").trim();
      const stockItems = Array.isArray(snapshot?.stockItems) ? snapshot.stockItems : [];
      const stockItem = stockItems.find((item) => item.id === stockItemId) || null;

      if (!stockItem) {
        throw createHttpError(400, "Stock item not found.");
      }

      if (!Number.isInteger(requestedQuantity) || requestedQuantity <= 0) {
        throw createHttpError(400, "Requested quantity must be greater than zero.");
      }

      const destination = status.artworkTo || status.from;
      const subject = `Artwork request: ${stockItem.name}${stockItem.sku ? ` (${stockItem.sku})` : ""}`;
      const text = [
        "Artwork has been requested for stock preparation.",
        "",
        `Requested by: ${currentUser.name}`,
        `Item: ${stockItem.name}`,
        `SKU: ${stockItem.sku || "Not set"}`,
        `Requested quantity: ${requestedQuantity}`,
        `Unit: ${stockItem.unit || "units"}`,
        `Current on hand: ${Number(stockItem.onHandQuantity || 0)}`,
        `Notes: ${notes || "None"}`,
      ].join("\n");

      await sendMessage({
        from: status.from,
        to: destination,
        subject,
        text,
      });

      await database.call("create_artwork_request", {
        p_token: token,
        p_stock_item_id: stockItem.id,
        p_requested_quantity: requestedQuantity,
        p_notes: notes,
        p_sent_to: destination,
      });

      return {
        ok: true,
        sentTo: destination,
        sentFrom: status.from,
        subject,
      };
    },
  };
}

function getMicrosoftGraphConfigIssue(config) {
  if (!config.tenantId || !config.clientId || !config.clientSecret) {
    return "Add Microsoft Graph mail settings in mail-config.js or MAIL_* environment variables.";
  }

  if (
    looksLikeTemplateValue(config.tenantId) ||
    looksLikeTemplateValue(config.clientId) ||
    looksLikeTemplateValue(config.clientSecret)
  ) {
    return "Replace the Microsoft Graph placeholder values in mail-config.js or the MAIL_* environment variables.";
  }

  if (UUID_PATTERN.test(String(config.clientSecret).trim())) {
    return "MAIL_CLIENT_SECRET looks like an Azure secret ID, not the secret value. Replace it in mail-config.js or the MAIL_CLIENT_SECRET environment variable.";
  }

  return "";
}

function looksLikeTemplateValue(value) {
  const normalized = String(value || "").trim().toUpperCase();
  return normalized.startsWith("YOUR_") || normalized.includes("MICROSOFT_APP_CLIENT_SECRET");
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
  const artworkTo = firstNonEmpty([
    process.env.MAIL_ARTWORK_TO,
    fileConfig.artworkTo,
    from,
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
      artworkTo,
    };
  }

  return {
    provider,
    from,
    to,
    artworkTo,
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
      "Move to factory",
      "Quote number",
      "Sales order number",
      "Invoice number",
      "PO number",
      "Created by",
      "Status",
      "Notice",
      "Created at",
      "Completed at",
    ],
    ...orders.map((order) => [
      order.reference || "",
      order.driverName || "Unassigned",
      order.locationName || "",
      order.entryType || "",
      getMoveToFactoryLabel(order) || "No",
      order.quoteNumber || order.inhouseOrderNumber || "",
      order.salesOrderNumber || order.factoryOrderNumber || "",
      order.invoiceNumber || "",
      order.poNumber || "",
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
  const driverFlag = getOrderFlagNoticeText(order);
  const moveToFactory = getMoveToFactoryText(order);
  const carryOverCount = Number(order?.carryOverCount || 0);

  if (driverFlag) {
    lines.push(driverFlag);
  }

  if (notice) {
    lines.push(notice);
  }

  if (moveToFactory) {
    lines.push(moveToFactory);
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

function getOrderFlagNoticeText(order) {
  const label = getOrderFlagLabel(order);
  if (!label) {
    return "";
  }

  const note = String(order?.driverFlagNote || "").trim();
  const flaggedAt = String(order?.driverFlaggedAt || "").trim();
  const flaggedBy = String(order?.driverFlaggedByName || "").trim();
  const parts = [`Driver follow-up: ${label}.`];

  if (note) {
    parts.push(note);
  }

  if (flaggedAt || flaggedBy) {
    parts.push(`Logged ${[flaggedAt, flaggedBy ? `by ${flaggedBy}` : ""].filter(Boolean).join(" ")}.`);
  }

  return parts.join(" ");
}

function getOrderFlagLabel(order) {
  const labels = {
    not_collected: "Not collected",
    not_ready: "Not yet ready",
  };

  return labels[String(order?.driverFlagType || "").trim()] || "";
}

function getMoveToFactoryLabel(order) {
  if (!order?.moveToFactory) {
    return "";
  }

  const destinationName = String(order?.factoryDestinationName || "").trim();
  return destinationName ? `Yes: ${destinationName}` : "Yes";
}

function getMoveToFactoryText(order) {
  const label = getMoveToFactoryLabel(order);
  if (!label) {
    return "";
  }

  const destinationName = String(order?.factoryDestinationName || "").trim();
  return destinationName
    ? `Move collected stock to ${destinationName}.`
    : "Move collected stock to a factory.";
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
    if (String(payload?.error || "").trim().toLowerCase() === "invalid_client") {
      const looksLikeSecretId = UUID_PATTERN.test(String(config.clientSecret || "").trim());
      throw createHttpError(
        502,
        looksLikeSecretId
          ? "Microsoft Graph rejected the configured client secret. The configured value looks like the Azure secret ID, not the secret value. Update MAIL_CLIENT_SECRET or mail-config.js."
          : "Microsoft Graph rejected the configured client secret. Update MAIL_CLIENT_SECRET or mail-config.js.",
      );
    }
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

async function createQrSvg(text) {
  if (!QRCode || qrLoadError) {
    throw createHttpError(503, "QR generation is not configured.");
  }

  if (!text) {
    throw createHttpError(400, "QR text is required.");
  }

  if (text.length > 1600) {
    throw createHttpError(400, "QR text is too long.");
  }

  return QRCode.toString(text, {
    type: "svg",
    errorCorrectionLevel: "M",
    margin: 1,
    width: 320,
    color: {
      dark: "#173c34",
      light: "#0000",
    },
  });
}

function startLiveReloadWatcher() {
  if (liveReloadWatcherStarted) {
    return;
  }

  liveReloadWatcherStarted = true;

  try {
    fs.watch(ROOT, (_eventType, filename) => {
      const normalized = normalizeLiveReloadFileName(filename);
      if (!normalized || !LIVE_RELOAD_FILES.has(normalized)) {
        return;
      }

      scheduleLiveReload(normalized);
    });
  } catch (error) {
    console.warn(`Live reload watcher could not start: ${normalizeErrorMessage(error)}`);
  }
}

function normalizeLiveReloadFileName(filename) {
  if (typeof filename !== "string" || !filename.trim()) {
    return "";
  }

  return path.basename(filename).toLowerCase();
}

function scheduleLiveReload(fileName) {
  liveReloadPendingFile = fileName || liveReloadPendingFile;

  if (liveReloadTimer) {
    clearTimeout(liveReloadTimer);
  }

  liveReloadTimer = setTimeout(() => {
    const changedFile = liveReloadPendingFile;
    liveReloadPendingFile = "";
    liveReloadTimer = null;
    broadcastLiveReload(changedFile);
  }, 120);
}

function broadcastLiveReload(fileName) {
  if (!liveReloadClients.size) {
    return;
  }

  const payload = JSON.stringify({
    changedAt: Date.now(),
    file: fileName || "",
  });

  liveReloadClients.forEach((client) => {
    try {
      client.write(`event: reload\ndata: ${payload}\n\n`);
    } catch (error) {
      liveReloadClients.delete(client);
      if (!client.writableEnded) {
        client.end();
      }
    }
  });
}

function handleLiveReloadStream(request, response) {
  response.writeHead(200, {
    "Cache-Control": "no-store",
    Connection: "keep-alive",
    "Content-Type": "text/event-stream; charset=UTF-8",
    "X-Accel-Buffering": "no",
  });
  response.flushHeaders?.();
  response.write(": connected\n\n");

  const keepAliveTimer = setInterval(() => {
    if (!response.writableEnded) {
      response.write(": keepalive\n\n");
    }
  }, 15000);

  liveReloadClients.add(response);

  request.on("close", () => {
    clearInterval(keepAliveTimer);
    liveReloadClients.delete(response);
    if (!response.writableEnded) {
      response.end();
    }
  });
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
    "Cache-Control": "no-store",
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

function sendSvg(response, statusCode, body) {
  response.writeHead(statusCode, {
    "Cache-Control": "no-store",
    "Content-Type": "image/svg+xml; charset=UTF-8",
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

module.exports = {
  routeRequest,
  startServer,
};
