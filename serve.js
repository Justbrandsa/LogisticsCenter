const http = require("http");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { loadProjectEnv } = require("./load-project-env");

loadProjectEnv(__dirname);

const { createLocalDatabase } = require("./local-database");

let nodemailer = null;
let mailerLoadError = null;
let QRCode = null;
let qrLoadError = null;

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

const HOST = String(process.env.HOST || "0.0.0.0").trim() || "0.0.0.0";
const PORT = Number(process.env.PORT || 4173);
const ROOT = __dirname;
const TIME_ZONE = "Africa/Johannesburg";
const MAX_BODY_BYTES = 1024 * 1024;
const LIVE_RELOAD_FILES = new Set(["index.html", "app.js", "styles.css"]);
const DELETE_LOG_EMAIL_POLL_MS = 10 * 60 * 1000;
const DELETE_LOG_CRON_PATH = "/api/jobs/order-delete-log-email";
const DEFAULT_MAIL_SENDER_NAME = "Logistics Centre";
const DEFAULT_ADMIN_ACTION_NOTIFICATION_EMAIL = "admin3@giftwrap.co.za";
const DEFAULT_ROLLOVER_TEST_EMAIL = "artwork3@giftwrap.co.za";
const DEFAULT_DROPPED_OFFICE_SS_EMAIL = "Sheryl-ann@giftwrap.co.za";
const DEFAULT_DROPPED_OFFICE_SB_EMAIL = "reception@giftwrap.co.za";
const DEFAULT_DROPPED_OFFICE_MOR_MAR_EMAIL = "promo22@giftwrap.co.za";
const DEFAULT_DROPPED_OFFICE_ORDER_EMAIL = "orders@giftwrapshop.co.za";
const DEFAULT_DROPPED_OFFICE_FALLBACK_EMAIL = "order@giftwrapshop.co.za";
const DROPPED_OFFICE_SS_PREFIXES = Object.freeze(["SS", "PSS"]);
const DROPPED_OFFICE_SB_PREFIXES = Object.freeze(["SB", "PSB"]);
const DROPPED_OFFICE_MOR_MAR_PREFIXES = Object.freeze(["MOR", "MAR", "PMOR", "PMAR"]);
const DROPPED_OFFICE_ORDER_PREFIXES = Object.freeze(["ORDER", "SO", "BAR"]);
const ROLLOVER_EMAIL_FUNCTIONS = new Set(["get_app_snapshot", "run_daily_rollover"]);
const GEOCODE_SEARCH_URL = "https://nominatim.openstreetmap.org/search";
const GEOCODE_CACHE_FILENAME = "geocode-cache.json";
const GEOCODE_CACHE_LIMIT = 500;
const GEOCODE_MIN_INTERVAL_MS = 1100;
const GEOCODE_USER_AGENT = "LogisticsCenter/1.0 (self-hosted address lookup)";
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
    params: ["p_token", "p_name", "p_sku", "p_quote_number", "p_invoice_number", "p_sales_order_number", "p_po_number", "p_unit", "p_initial_quantity", "p_notes"],
  },
  update_stock_item: {
    params: ["p_token", "p_stock_item_id", "p_name", "p_sku", "p_quote_number", "p_invoice_number", "p_sales_order_number", "p_po_number", "p_unit", "p_notes"],
  },
  record_driver_position: {
    params: ["p_token", "p_lat", "p_lng"],
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
      "p_branding",
      "p_stock_description",
      "p_stock_item_names",
      "p_delivery_address",
      "p_delivery_location_id",
      "p_save_delivery_location",
      "p_delivery_location_name",
      "p_scheduled_for",
      "p_priority",
      "p_allow_duplicate",
      "p_notice",
      "p_move_to_factory",
      "p_factory_destination_location_id",
    ],
  },
  update_order: {
    params: [
      "p_token",
      "p_order_id",
      "p_driver_user_id",
      "p_location_id",
      "p_entry_type",
      "p_quote_number",
      "p_sales_order_number",
      "p_invoice_number",
      "p_po_number",
      "p_branding",
      "p_stock_description",
      "p_delivery_address",
      "p_delivery_location_id",
      "p_save_delivery_location",
      "p_delivery_location_name",
      "p_scheduled_for",
      "p_priority",
      "p_allow_duplicate",
      "p_notice",
      "p_move_to_factory",
      "p_factory_destination_location_id",
    ],
  },
  assign_order: { params: ["p_token", "p_order_id", "p_driver_user_id", "p_allow_duplicate"] },
  set_order_priority: { params: ["p_token", "p_order_id", "p_priority"] },
  clear_all_order_priorities: { params: ["p_token"] },
  clear_order_rollovers: { params: ["p_token"] },
  set_order_flag: { params: ["p_token", "p_order_id", "p_flag_type", "p_note"] },
  pick_up_order: { params: ["p_token", "p_order_id"] },
  complete_order: { params: ["p_token", "p_order_id", "p_completion_type"] },
  delete_order: { params: ["p_token", "p_order_id"] },
});

const database = createDatabase();
const mailer = createMailer();
const geocoder = createGeocoder();
const liveReloadClients = new Set();

let liveReloadWatcherStarted = false;
let liveReloadTimer = null;
let liveReloadPendingFile = "";
let deleteLogNotificationTimer = null;
let deleteLogNotificationPromise = null;

function getStockUnitLabel(record) {
  const rawUnit = String(record?.unit || "").trim();
  if (!rawUnit) {
    return "units";
  }

  if (/^\d+(?:\s*units?)?$/i.test(rawUnit)) {
    return "units";
  }

  return rawUnit;
}

function isDeliveryOrder(order) {
  return String(order?.entryType || "").trim().toLowerCase() === "delivery";
}

function getOrderCompletionLabel(order) {
  const completionType = String(order?.completionType || "").trim();
  if (completionType === "office" && isDeliveryOrder(order)) {
    return "Delivered to client";
  }

  const labels = {
    office: "Dropped at office",
    factory: "Dropped at factory",
  };

  return labels[completionType] || "";
}

function getOrderCompletionText(order) {
  const label = getOrderCompletionLabel(order);
  if (!label) {
    return "";
  }

  const completedAt = String(order?.completedAt || "").trim();
  const completedBy = String(order?.completedByName || "").trim();
  const parts = [`Completion: ${label}.`];

  if (completedAt || completedBy) {
    parts.push(`Logged ${[completedAt, completedBy ? `by ${completedBy}` : ""].filter(Boolean).join(" ")}.`);
  }

  return parts.join(" ");
}

function startServer() {
  startLiveReloadWatcher();
  startOrderDeleteLogPolling();

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
        ? `Local database ready at ${status.storagePath || "the project data folder"}.`
        : `Local database is not available yet: ${status.reason}`,
    );
    if (status.seededFromBundledSnapshot) {
      console.warn("Temporary runtime database was seeded from the bundled project snapshot at startup.");
    }
    if (status.warning) {
      console.warn(`Local database warning: ${status.warning}`);
    }
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
    stopOrderDeleteLogPolling();
    void database.close().finally(() => process.exit(0));
  });

  process.on("SIGTERM", () => {
    stopOrderDeleteLogPolling();
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
      const mailStatus = mailer.getStatus();
      sendJson(response, 200, {
        ...database.getStatus(),
        mailConfigured: mailStatus.configured,
        mailReason: mailStatus.reason,
        mailProvider: mailStatus.provider,
        mailFrom: mailStatus.from,
        mailFromDisplay: mailStatus.fromDisplay,
        mailSenderName: mailStatus.senderName,
        mailTo: mailStatus.to,
        artworkTo: mailStatus.artworkTo,
      });
      return;
    }

    if (request.method === "POST" && cleanPath === "/api/mail/settings") {
      const payload = await readJsonBody(request);
      const result = await getMailSettings(payload?.token);
      sendJson(response, 200, result);
      return;
    }

    if (request.method === "POST" && cleanPath === "/api/mail/settings/update") {
      const payload = await readJsonBody(request);
      const result = await updateMailSettings(payload?.token, payload || {});
      sendJson(response, 200, result);
      return;
    }

    if (request.method === "POST" && cleanPath === "/api/admin/data/export") {
      const payload = await readJsonBody(request);
      const result = await exportDatabaseData(payload?.token);
      sendJson(response, 200, result);
      return;
    }

    if (request.method === "POST" && cleanPath === "/api/admin/data/import") {
      const payload = await readJsonBody(request);
      const result = await importDatabaseData(payload?.token, payload?.data, payload?.source);
      sendJson(response, 200, result);
      return;
    }

    if ((request.method === "GET" || request.method === "POST") && cleanPath === DELETE_LOG_CRON_PATH) {
      if (!isAuthorizedJobRequest(request)) {
        sendJson(response, 401, { error: "Unauthorized." });
        return;
      }

      const result = await processOrderDeleteLogNotifications({
        source: "cron",
        allowConcurrentWait: true,
      });
      sendJson(response, 200, result);
      return;
    }

    if (request.method === "POST" && cleanPath.startsWith("/api/rpc/")) {
      const functionName = cleanPath.slice("/api/rpc/".length);
      const payload = await readJsonBody(request);
      const parameters = payload?.parameters || {};
      const driverTransferEmailContext = await buildDriverTransferEmailContext(functionName, parameters);
      const adminActionEmailContext = await buildAdminActionEmailContext(functionName, parameters);
      const data = await database.call(functionName, parameters);
      const droppedOfficeEmailContext = await buildDroppedOfficeEmailContext(functionName, parameters, data);
      let responseData = data;
      const warnings = [];
      await maybeSendCarryOverEmail(functionName, data);
      if (driverTransferEmailContext) {
        try {
          await mailer.sendDriverTransferEmail(driverTransferEmailContext);
        } catch (error) {
          warnings.push(`Admin email could not be sent: ${normalizeErrorMessage(error)}`);
          console.error("Failed to send driver transfer email.", error);
        }
      }
      if (adminActionEmailContext && Number(data?.updatedOrders || 0) > 0) {
        try {
          await mailer.sendAdminActionNotification({
            ...adminActionEmailContext,
            affectedCount: Number(data?.updatedOrders || adminActionEmailContext.affectedCount || 0),
          });
        } catch (error) {
          warnings.push(`Admin action email could not be sent: ${normalizeErrorMessage(error)}`);
          console.error("Failed to send admin action email.", error);
        }
      }
      if (droppedOfficeEmailContext) {
        try {
          await mailer.sendDroppedOfficeEmail(droppedOfficeEmailContext);
        } catch (error) {
          warnings.push(`Dropped-at-office email could not be sent: ${normalizeErrorMessage(error)}`);
          console.error("Failed to send dropped-at-office email.", error);
        }
      }
      const warning = warnings.join(" ");
      if (warning) {
        responseData = appendWarningToRpcResult(data, warning);
      }
      sendJson(response, 200, { data: responseData });
      return;
    }

    if (request.method === "POST" && cleanPath === "/api/artwork/request") {
      const payload = await readJsonBody(request);
      const result = await mailer.sendArtworkRequest(payload?.token, payload || {});
      sendJson(response, 200, result);
      return;
    }

    if (request.method === "POST" && cleanPath === "/api/geocode/address") {
      const payload = await readJsonBody(request);
      const result = await geocoder.lookupAddress(payload?.address, {
        force: normalizeBoolean(payload?.force, false),
      });
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

    if (request.method === "POST" && cleanPath === "/api/export/email/rollover-test") {
      const payload = await readJsonBody(request);
      const result = await mailer.sendCarryOverTestEmail(payload?.token);
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

async function requireAdminSession(token, deniedMessage = "Only admin users can do that.") {
  return requireRoleSession(token, ["admin"], deniedMessage);
}

async function requireRoleSession(token, allowedRoles, deniedMessage = "You do not have permission to do that.") {
  const cleanToken = String(token || "").trim();
  if (!cleanToken) {
    throw createHttpError(400, "Session token is required.");
  }

  const currentUser = await database.getUserByToken(cleanToken);
  if (!currentUser) {
    throw createHttpError(403, "You must be signed in.");
  }

  if (!Array.isArray(allowedRoles) || !allowedRoles.includes(currentUser.role)) {
    throw createHttpError(403, deniedMessage);
  }

  return currentUser;
}

async function exportDatabaseData(token) {
  const currentUser = await requireAdminSession(token, "Only admin users can export data.");
  return {
    exportedAt: new Date().toISOString(),
    exportedBy: currentUser.name,
    storage: database.getStatus().storagePath || "",
    counts: database.getTableCounts(),
    data: database.exportAllData(),
  };
}

async function importDatabaseData(token, data, source = "") {
  const currentUser = await requireAdminSession(token, "Only admin users can import data.");
  const result = await database.replaceAllData(data, {
    source: String(source || "").trim() || `admin import by ${currentUser.name}`,
  });
  return {
    ...result,
    importedAt: new Date().toISOString(),
    importedBy: currentUser.name,
  };
}

async function getMailSettings(token) {
  await requireRoleSession(token, ["admin", "maintenance"], "Only admin or maintenance users can manage email settings.");
  return mailer.getManagementSettings();
}

async function updateMailSettings(token, payload = {}) {
  const currentUser = await requireRoleSession(token, ["admin", "maintenance"], "Only admin or maintenance users can manage email settings.");
  await database.call("update_mail_settings", {
    p_token: token,
    p_disabled: normalizeBoolean(payload?.disabled, false),
    p_sender_name: payload?.senderName,
    p_to: payload?.to,
    p_artwork_to: payload?.artworkTo,
    p_admin_action_to: payload?.adminActionTo,
    p_rollover_test_to: payload?.rolloverTestTo,
    p_dropped_office_ss_to: payload?.droppedOfficeSsTo,
    p_dropped_office_sb_to: payload?.droppedOfficeSbTo,
    p_dropped_office_mor_mar_to: payload?.droppedOfficeMorMarTo,
    p_dropped_office_order_to: payload?.droppedOfficeOrderTo,
    p_dropped_office_fallback_to: payload?.droppedOfficeFallbackTo,
  });

  return {
    ...mailer.getManagementSettings(),
    updatedAt: new Date().toISOString(),
    updatedBy: currentUser.name,
  };
}

async function buildDriverTransferEmailContext(functionName, parameters) {
  if (functionName !== "assign_order") {
    return null;
  }

  const token = String(parameters?.p_token || "").trim();
  const orderId = String(parameters?.p_order_id || "").trim();
  const nextDriverUserId = String(parameters?.p_driver_user_id || "").trim();
  if (!token || !orderId || !nextDriverUserId) {
    return null;
  }

  const snapshot = await database.call("get_app_snapshot", { p_token: token });
  const currentUser = snapshot?.user || null;
  if (!currentUser || currentUser.role !== "driver") {
    return null;
  }

  const orders = Array.isArray(snapshot?.orders) ? snapshot.orders : [];
  const drivers = Array.isArray(snapshot?.users) ? snapshot.users : [];
  const order = orders.find((entry) => entry.id === orderId && entry.status === "active") || null;
  const nextDriver = drivers.find((driver) => driver.id === nextDriverUserId) || null;

  if (!order || !nextDriver) {
    return null;
  }

  return {
    initiatedByName: currentUser.name || "Driver",
    fromDriverName: order.driverName || currentUser.name || "Unknown driver",
    toDriverName: nextDriver.name || "Unknown driver",
    transferredAt: new Date().toISOString(),
    order,
  };
}

async function buildAdminActionEmailContext(functionName, parameters) {
  if (!["clear_all_order_priorities", "clear_order_rollovers"].includes(functionName)) {
    return null;
  }

  const token = String(parameters?.p_token || "").trim();
  if (!token) {
    return null;
  }

  const currentUser = await database.getUserByToken(token);
  if (!currentUser || currentUser.role !== "admin") {
    return null;
  }

  const orders = await database.listOrdersForMailExport();
  const isPriorityAction = functionName === "clear_all_order_priorities";
  const affectedOrders = orders.filter((order) => (
    order.status === "active"
    && (isPriorityAction
      ? getOrderPriority(order) === "high"
      : Number(order.carryOverCount || 0) > 0)
  ));

  if (!affectedOrders.length) {
    return null;
  }

  const actionLabel = isPriorityAction ? "Clear all priority" : "Clear rollover";
  const detailLabel = isPriorityAction ? "Priority entries cleared" : "Rollover entries cleared";

  return {
    actionLabel,
    detailLabel,
    initiatedByName: currentUser.name || "Admin",
    initiatedByRole: currentUser.role || "admin",
    triggeredAt: new Date().toISOString(),
    affectedCount: affectedOrders.length,
    affectedOrders: affectedOrders.slice(0, 25),
    remainingCount: Math.max(affectedOrders.length - 25, 0),
  };
}

async function buildDroppedOfficeEmailContext(functionName, parameters, result) {
  if (functionName !== "complete_order") {
    return null;
  }

  const orderId = String(parameters?.p_order_id || "").trim();
  const completionType = String(result?.completionType || "").trim().toLowerCase();
  if (!orderId || completionType !== "office") {
    return null;
  }

  const orders = await database.listOrdersForMailExport();
  const order = orders.find((entry) => entry.id === orderId) || null;
  if (!order) {
    return null;
  }

  if (String(order?.status || "").trim().toLowerCase() !== "completed") {
    return null;
  }

  if (String(order?.completionType || "").trim().toLowerCase() !== "office") {
    return null;
  }

  if (isDeliveryOrder(order)) {
    return null;
  }

  return {
    order,
    completedAt: order.completedAt || new Date().toISOString(),
    completedByName: String(order?.completedByName || "").trim(),
  };
}

function appendWarningToRpcResult(data, warning) {
  if (!warning) {
    return data;
  }

  if (data && typeof data === "object" && !Array.isArray(data)) {
    return {
      ...data,
      warning,
    };
  }

  return {
    value: data,
    warning,
  };
}

function createDatabase() {
  return createLocalDatabase(ROOT);
}

function createGeocoder() {
  const cacheFilePath = getRuntimeDataFilePath(GEOCODE_CACHE_FILENAME);
  const cache = loadGeocodeCache(cacheFilePath);
  let lastRequestAt = 0;
  let requestQueue = Promise.resolve();

  const runRateLimitedRequest = async (task) => {
    const scheduledTask = requestQueue.catch(() => undefined).then(async () => {
      const waitMs = Math.max(0, GEOCODE_MIN_INTERVAL_MS - (Date.now() - lastRequestAt));
      if (waitMs > 0) {
        await delay(waitMs);
      }

      lastRequestAt = Date.now();
      return task();
    });

    requestQueue = scheduledTask.then(() => undefined, () => undefined);
    return scheduledTask;
  };

  return {
    async lookupAddress(address, options = {}) {
      const normalizedAddress = normalizeGeocodeAddress(address);
      if (!normalizedAddress) {
        throw createHttpError(400, "Address is required for coordinate lookup.");
      }

      if (!options.force) {
        const cached = getGeocodeCacheEntry(cache, normalizedAddress);
        if (cached) {
          if (cached.notFound) {
            cache.delete(normalizedAddress.toLowerCase());
            persistGeocodeCache(cache, cacheFilePath);
          } else {
            return {
              lat: cached.lat,
              lng: cached.lng,
              displayName: cached.displayName,
              cached: true,
            };
          }
        }
      }

      try {
        const queryVariants = buildGeocodeQueryVariants(normalizedAddress);
        let match = null;

        for (const query of queryVariants) {
          match = await runRateLimitedRequest(() => fetchGeocodeMatch(query));
          if (match) {
            break;
          }
        }

        if (!match) {
          throw createHttpError(404, "No coordinates were found for that address.");
        }

        const lat = Number(match.lat);
        const lng = Number(match.lon);
        if (!Number.isFinite(lat) || !Number.isFinite(lng)) {
          throw createHttpError(502, "Address lookup returned invalid coordinates.");
        }

        const result = {
          lat,
          lng,
          displayName: String(match.display_name || "").trim(),
          updatedAt: new Date().toISOString(),
        };
        setGeocodeCacheEntry(cache, normalizedAddress, result);
        persistGeocodeCache(cache, cacheFilePath);

        return {
          lat: result.lat,
          lng: result.lng,
          displayName: result.displayName,
          cached: false,
        };
      } catch (error) {
        if (Number.isInteger(error?.statusCode)) {
          throw error;
        }

        throw createHttpError(503, normalizeErrorMessage(error));
      }
    },
  };
}

function getRuntimeDataFilePath(fileName) {
  const status = database.getStatus();
  const preferredDir = String(status?.storageDir || "").trim();
  const fallbackDir = path.join(os.tmpdir(), "logistics-center-data");
  return path.join(preferredDir || fallbackDir, fileName);
}

function normalizeGeocodeAddress(value) {
  return String(value || "").trim().replace(/\s+/g, " ");
}

async function fetchGeocodeMatch(query) {
  const url = new URL(GEOCODE_SEARCH_URL);
  url.searchParams.set("q", query);
  url.searchParams.set("format", "jsonv2");
  url.searchParams.set("limit", "1");
  url.searchParams.set("addressdetails", "0");

  let response;
  try {
    response = await fetch(url, {
      headers: {
        Accept: "application/json",
        "Accept-Language": "en-ZA,en;q=0.9",
        "User-Agent": GEOCODE_USER_AGENT,
      },
    });
  } catch (error) {
    throw createHttpError(503, `Address lookup needs internet access. ${normalizeErrorMessage(error)}`);
  }

  const payload = await safeReadJson(response);
  if (response.status === 429) {
    throw createHttpError(429, "Address lookup is being rate-limited right now. Please wait a moment and try again.");
  }

  if (!response.ok) {
    throw createHttpError(502, payload?.error || "Address lookup failed.");
  }

  return Array.isArray(payload) ? (payload[0] || null) : null;
}

function buildGeocodeQueryVariants(address) {
  const normalizedAddress = normalizeGeocodeAddress(address);
  const parts = normalizedAddress
    .split(",")
    .map((part) => normalizeGeocodeAddress(part))
    .filter(Boolean);
  const baseQueries = [];
  const queries = [];

  const pushBaseQuery = (value) => {
    const query = normalizeGeocodeAddress(value);
    if (!query || baseQueries.includes(query) || baseQueries.length >= 4) {
      return;
    }
    baseQueries.push(query);
  };

  const pushQuery = (value) => {
    const query = normalizeGeocodeAddress(value);
    if (!query || queries.includes(query) || queries.length >= 10) {
      return;
    }
    queries.push(query);
  };

  pushBaseQuery(normalizedAddress);

  if (parts.length > 1 && looksLikePostalCode(parts[parts.length - 1])) {
    pushBaseQuery(parts.slice(0, -1).join(", "));
  }

  if (parts.length > 2) {
    pushBaseQuery(parts.slice(0, -1).join(", "));
    pushBaseQuery(parts.slice(0, -2).join(", "));
  }

  baseQueries.forEach((query) => {
    pushQuery(query);

    if (query.includes("&")) {
      pushQuery(query.replace(/\s*&\s*/g, " and "));
    }

    if (!/south africa/i.test(query)) {
      pushQuery(`${query}, South Africa`);

      if (query.includes("&")) {
        pushQuery(`${query.replace(/\s*&\s*/g, " and ")}, South Africa`);
      }
    }
  });

  return queries;
}

function looksLikePostalCode(value) {
  return /^\d{4,6}$/.test(String(value || "").trim());
}

function loadGeocodeCache(filePath) {
  try {
    const raw = fs.readFileSync(filePath, "utf8");
    const parsed = JSON.parse(raw);
    const entries = Object.entries(parsed || {})
      .map(([key, value]) => [normalizeGeocodeAddress(key).toLowerCase(), normalizeGeocodeCacheEntry(value)])
      .filter((entry) => entry[0] && entry[1]);
    return new Map(entries);
  } catch (error) {
    return new Map();
  }
}

function normalizeGeocodeCacheEntry(value) {
  if (!value || typeof value !== "object") {
    return null;
  }

  if (value.notFound) {
    return {
      notFound: true,
      updatedAt: String(value.updatedAt || "").trim(),
    };
  }

  const lat = Number(value.lat);
  const lng = Number(value.lng);
  if (!Number.isFinite(lat) || !Number.isFinite(lng)) {
    return null;
  }

  return {
    lat,
    lng,
    displayName: String(value.displayName || "").trim(),
    updatedAt: String(value.updatedAt || "").trim(),
  };
}

function getGeocodeCacheEntry(cache, address) {
  const key = normalizeGeocodeAddress(address).toLowerCase();
  if (!key) {
    return null;
  }

  const entry = cache.get(key) || null;
  if (!entry) {
    return null;
  }

  cache.delete(key);
  cache.set(key, entry);
  return entry;
}

function setGeocodeCacheEntry(cache, address, entry) {
  const key = normalizeGeocodeAddress(address).toLowerCase();
  if (!key || !entry) {
    return;
  }

  cache.delete(key);
  cache.set(key, entry);

  while (cache.size > GEOCODE_CACHE_LIMIT) {
    const oldestKey = cache.keys().next().value;
    if (!oldestKey) {
      break;
    }
    cache.delete(oldestKey);
  }
}

function persistGeocodeCache(cache, filePath) {
  try {
    fs.mkdirSync(path.dirname(filePath), { recursive: true });
    fs.writeFileSync(
      filePath,
      JSON.stringify(Object.fromEntries(cache.entries()), null, 2),
      "utf8",
    );
  } catch (error) {
    console.warn(`Failed to persist geocode cache: ${normalizeErrorMessage(error)}`);
  }
}

function delay(ms) {
  return new Promise((resolve) => {
    setTimeout(resolve, ms);
  });
}

async function maybeSendCarryOverEmail(functionName, payload) {
  if (!ROLLOVER_EMAIL_FUNCTIONS.has(functionName)) {
    return;
  }

  if (!mailer.getStatus().configured) {
    return;
  }

  const rollover = functionName === "get_app_snapshot" ? payload?.rollover : payload;
  const updatedOrders = Number(rollover?.updatedOrders || 0);

  if (updatedOrders <= 0) {
    return;
  }

  try {
    await mailer.sendCarryOverEmail(rollover);
  } catch (error) {
    console.error("Failed to send carry-over email.", error);
  }
}

function startOrderDeleteLogPolling() {
  if (deleteLogNotificationTimer) {
    return;
  }

  const runPoll = (source) => {
    void processOrderDeleteLogNotifications({ source }).catch((error) => {
      console.error("Failed to process delete-log notifications.", error);
    });
  };

  runPoll("startup");
  deleteLogNotificationTimer = setInterval(() => {
    runPoll("interval");
  }, DELETE_LOG_EMAIL_POLL_MS);

  if (typeof deleteLogNotificationTimer.unref === "function") {
    deleteLogNotificationTimer.unref();
  }
}

function stopOrderDeleteLogPolling() {
  if (!deleteLogNotificationTimer) {
    return;
  }

  clearInterval(deleteLogNotificationTimer);
  deleteLogNotificationTimer = null;
}

function isAuthorizedJobRequest(request) {
  const secret = String(process.env.CRON_SECRET || "").trim();
  if (!secret) {
    return true;
  }

  const authorization = String(request?.headers?.authorization || "").trim();
  return authorization === `Bearer ${secret}`;
}

async function processOrderDeleteLogNotifications(options = {}) {
  const {
    source = "manual",
    allowConcurrentWait = false,
  } = options;

  if (deleteLogNotificationPromise) {
    return allowConcurrentWait
      ? deleteLogNotificationPromise
      : { ok: true, skipped: true, reason: "Delete-log notification job already running." };
  }

  deleteLogNotificationPromise = (async () => {
    if (!database.getStatus().configured) {
      return { ok: true, skipped: true, reason: "Database connection is not configured." };
    }

    const pendingEntries = await database.listPendingOrderDeleteNotifications();
    if (!pendingEntries.length) {
      return { ok: true, skipped: true, reason: "No pending delete-log entries.", count: 0, source };
    }

    if (!mailer.getStatus().configured) {
      return {
        ok: true,
        skipped: true,
        reason: mailer.getStatus().reason || "Email delivery is not configured.",
        count: pendingEntries.length,
        source,
      };
    }

    const logIds = pendingEntries
      .map((entry) => String(entry?.id || "").trim())
      .filter(Boolean);

    try {
      const mailResult = await mailer.sendOrderDeleteLogEmail(pendingEntries);
      await database.markOrderDeleteNotificationsSent(logIds);
      return {
        ok: true,
        count: pendingEntries.length,
        source,
        sentTo: mailResult?.sentTo || "",
        subject: mailResult?.subject || "",
      };
    } catch (error) {
      await database.markOrderDeleteNotificationFailure(logIds, normalizeErrorMessage(error));
      throw error;
    }
  })();

  try {
    return await deleteLogNotificationPromise;
  } finally {
    deleteLogNotificationPromise = null;
  }
}

function createMailer() {
  function getRuntime() {
    const config = getRuntimeMailConfig(
      loadMailConfig(),
      typeof database.getMailSettingsOverrides === "function"
        ? database.getMailSettingsOverrides()
        : null,
    );
    const status = buildMailStatus(config);
    const transporter = status.configured && config.provider === "smtp"
      ? nodemailer.createTransport(config.transport)
      : null;

    return {
      config,
      status,
      transporter,
    };
  }

  async function getAuthorizedMailContext(token, allowedRoles, deniedMessage) {
    const runtime = getRuntime();
    if (!runtime.status.configured) {
      throw createHttpError(503, runtime.status.reason || "Email delivery is not configured.");
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
      runtime,
      currentUser,
      snapshot,
    };
  }

  async function sendMessage(runtime, message) {
    const senderAddress = String(message.fromAddress || runtime.config.from || "").trim();
    const senderName = String(message.senderName || runtime.status.senderName || "").trim();

    if (runtime.config.provider === "microsoft-graph") {
      await sendViaMicrosoftGraph(runtime.config, {
        ...message,
        fromAddress: senderAddress,
        senderName,
      });
      return;
    }

    await runtime.transporter.sendMail({
      from: formatMailSender(senderAddress, senderName),
      to: message.to,
      subject: message.subject,
      text: message.text,
      html: message.html,
      attachments: Array.isArray(message.attachments) ? message.attachments : [],
    });
  }

  return {
    getStatus() {
      return { ...getRuntime().status };
    },
    getManagementSettings() {
      const runtime = getRuntime();
      return {
        ...runtime.status,
        disabled: Boolean(runtime.config.disabled),
        fromAddress: runtime.config.from,
        senderName: runtime.config.senderName,
        to: runtime.config.to,
        artworkTo: runtime.config.artworkTo,
        adminActionTo: runtime.config.adminActionTo,
        rolloverTestTo: runtime.config.rolloverTestTo,
        droppedOfficeSsTo: runtime.config.droppedOfficeSsTo,
        droppedOfficeSbTo: runtime.config.droppedOfficeSbTo,
        droppedOfficeMorMarTo: runtime.config.droppedOfficeMorMarTo,
        droppedOfficeOrderTo: runtime.config.droppedOfficeOrderTo,
        droppedOfficeFallbackTo: runtime.config.droppedOfficeFallbackTo,
      };
    },
    async sendSnapshotCsv(token) {
      const { runtime, currentUser, snapshot } = await getAuthorizedMailContext(
        token,
        ["admin", "sales"],
        "Only admin or sales users can send email.",
      );
      const { status } = runtime;

      const dateStamp = String(snapshot?.today || "").trim() || getCurrentLocalDateValue();
      const orders = filterOrdersForScheduledDate(snapshot?.orders, dateStamp);
      const filename = `route-ledger-${dateStamp}.csv`;
      const subject = `Route Ledger CSV export ${dateStamp}`;

      const text = [
        "Attached is the latest Route Ledger CSV export.",
        "",
        `Sent by: ${currentUser.name}`,
        `Entries included: ${orders.length}`,
      ].join("\n");
      const csv = buildOrdersCsv(orders);

      await sendMessage(runtime, {
        fromAddress: status.fromAddress,
        senderName: status.senderName,
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
      const { runtime, currentUser } = await getAuthorizedMailContext(
        token,
        ["admin", "sales", "maintenance"],
        "Only admin, sales, or maintenance users can send email.",
      );
      const { status } = runtime;
      const timestamp = new Date().toISOString();
      const subject = `Route Ledger test email ${timestamp}`;
      const text = [
        "This is a Route Ledger email test.",
        "",
        `Sent by: ${currentUser.name}`,
        `Provider: ${status.provider}`,
        `From: ${status.fromDisplay}`,
        `To: ${status.to}`,
        `Timestamp: ${timestamp}`,
      ].join("\n");

      await sendMessage(runtime, {
        fromAddress: status.fromAddress,
        senderName: status.senderName,
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
    async sendCarryOverTestEmail(token) {
      const { runtime, currentUser, snapshot } = await getAuthorizedMailContext(
        token,
        ["admin", "sales", "maintenance"],
        "Only admin, sales, or maintenance users can send email.",
      );
      const rollover = buildCarryOverTestPayload(snapshot);
      const result = await this.sendCarryOverEmail(rollover, {
        allowEmpty: true,
        requestedByName: currentUser.name,
        subjectPrefix: "TEST",
        to: runtime.status.rolloverTestTo,
      });

      return {
        ...result,
        requestedBy: currentUser.name,
      };
    },
    async sendCarryOverEmail(rollover, options = {}) {
      const runtime = getRuntime();
      const { status } = runtime;
      if (!status.configured) {
        throw createHttpError(503, status.reason || "Email delivery is not configured.");
      }

      const carriedOrders = Array.isArray(rollover?.carriedOrders)
        ? rollover.carriedOrders.filter(Boolean)
        : [];
      const count = Number(rollover?.updatedOrders || carriedOrders.length || 0);
      const allowEmpty = Boolean(options.allowEmpty);

      if (count <= 0 && !allowEmpty) {
        return {
          ok: true,
          skipped: true,
          count: 0,
        };
      }

      const dateStamp = String(rollover?.today || "").trim() || getCurrentLocalDateValue();
      const csvOrders = Array.isArray(options.orders)
        ? filterOrdersForScheduledDate(options.orders, dateStamp)
        : await database.listOrdersForMailExport(dateStamp);
      const csvFilename = `route-ledger-${dateStamp}.csv`;
      const csv = buildOrdersCsv(csvOrders);
      const emailTo = String(options.to || "").trim() || status.to;
      const summary = buildCarryOverEmailSummary(rollover, csvOrders, {
        isTest: Boolean(options.subjectPrefix),
        requestedByName: options.requestedByName,
        recipient: emailTo,
      });
      const itemLabel = count === 1 ? "item" : "items";
      const subjectBase = count > 0
        ? `Route Ledger carry-over ${dateStamp} (${count} ${itemLabel})`
        : `Route Ledger carry-over ${dateStamp} (no items)`;
      const subject = options.subjectPrefix
        ? `${String(options.subjectPrefix).trim()} ${subjectBase}`.trim()
        : subjectBase;
      const text = buildCarryOverEmailText(summary);
      const html = buildCarryOverEmailHtml(summary);

      await sendMessage(runtime, {
        fromAddress: status.fromAddress,
        senderName: status.senderName,
        to: emailTo,
        subject,
        text,
        html,
        attachments: [
          {
            filename: csvFilename,
            content: csv,
            contentType: "text/csv; charset=UTF-8",
          },
        ],
      });

      return {
        ok: true,
        sentTo: emailTo,
        sentFrom: status.from,
        subject,
        count,
        csvCount: csvOrders.length,
      };
    },
    async sendDriverTransferEmail(transfer) {
      const runtime = getRuntime();
      const { status } = runtime;
      if (!status.configured) {
        throw createHttpError(503, status.reason || "Email delivery is not configured.");
      }

      const order = transfer?.order || null;
      if (!order) {
        throw createHttpError(400, "Transfer details are required.");
      }

      const transferredAt = formatDateTime(transfer?.transferredAt || new Date().toISOString());
      const subject = `Driver transfer: ${getOrderPrimaryDisplay(order) || "Route Ledger entry"} -> ${transfer?.toDriverName || "New driver"}`;
      const text = [
        "A driver transferred an active Route Ledger entry.",
        "",
        `Transferred by: ${transfer?.initiatedByName || "Driver"}`,
        `From driver: ${transfer?.fromDriverName || "Unknown driver"}`,
        `To driver: ${transfer?.toDriverName || "Unknown driver"}`,
        `Entry: ${getOrderPrimaryDisplay(order) || "Unknown"}`,
        `Other references: ${getOrderOtherReferenceLines(order).join(" | ") || "None"}`,
        `Stock required: ${String(order.stockDescription || "").trim() || "Not set"}`,
        `Branding: ${String(order.branding || "").trim() || "None"}`,
        `Entry type: ${getOrderEntryTypeLabel(order.entryType) || "Not set"}`,
        `Priority: ${capitalize(getOrderPriority(order)) || "Medium"}`,
        `Pickup location: ${String(order.locationName || "").trim() || "Unknown"}`,
        `Pickup address: ${String(order.locationAddress || "").trim() || "Unknown"}`,
        `Delivery address: ${String(order.deliveryAddress || "").trim() || "Not set"}`,
        `Destination: ${getCollectionDestinationLabel(order) || "Not set"}`,
        `Move to factory: ${getMoveToFactoryLabel(order) || "No"}`,
        `Driver issue: ${getOrderDriverIssue(order) || "None"}`,
        `Notes: ${String(order.notes || "").trim() || "None"}`,
        `Schedule: ${getOrderScheduleSummary(order) || "Not scheduled"}`,
        `Transferred at: ${transferredAt || "Just now"}`,
      ].join("\n");

      await sendMessage(runtime, {
        fromAddress: status.fromAddress,
        senderName: status.senderName,
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
    async sendAdminActionNotification(notification) {
      const runtime = getRuntime();
      const { status } = runtime;
      if (!status.configured) {
        throw createHttpError(503, status.reason || "Email delivery is not configured.");
      }

      const affectedCount = Number(notification?.affectedCount || 0);
      if (affectedCount <= 0) {
        return {
          ok: true,
          skipped: true,
          count: 0,
        };
      }

      const actionLabel = String(notification?.actionLabel || "Admin action").trim();
      const detailLabel = String(notification?.detailLabel || "Affected entries").trim();
      const initiatedByName = String(notification?.initiatedByName || "Admin").trim() || "Admin";
      const initiatedByRole = capitalize(String(notification?.initiatedByRole || "admin").trim()) || "Admin";
      const triggeredAt = formatDateTime(notification?.triggeredAt || new Date().toISOString());
      const affectedOrders = Array.isArray(notification?.affectedOrders)
        ? notification.affectedOrders.filter(Boolean)
        : [];
      const remainingCount = Number(notification?.remainingCount || 0);
      const itemLabel = affectedCount === 1 ? "entry" : "entries";
      const subject = `Route Ledger admin action: ${actionLabel.toLowerCase()} (${affectedCount} ${itemLabel})`;
      const lines = [
        "An admin used a bulk Route Ledger action.",
        "",
        `Action: ${actionLabel}`,
        `Performed by: ${initiatedByName} (${initiatedByRole})`,
        `Time: ${triggeredAt || "Just now"}`,
        `${detailLabel}: ${affectedCount}`,
      ];

      if (affectedOrders.length) {
        lines.push("");
        lines.push("Entries:");
        affectedOrders.forEach((order) => {
          lines.push(`- ${formatAdminActionOrderLine(order)}`);
        });
      }

      if (remainingCount > 0) {
        lines.push(`- Plus ${remainingCount} more entr${remainingCount === 1 ? "y" : "ies"}`);
      }

      await sendMessage(runtime, {
        fromAddress: status.fromAddress,
        senderName: status.senderName,
        to: status.adminActionTo,
        subject,
        text: lines.join("\n"),
      });

      return {
        ok: true,
        sentTo: status.adminActionTo,
        sentFrom: status.from,
        subject,
        count: affectedCount,
      };
    },
    async sendDroppedOfficeEmail(notification) {
      const runtime = getRuntime();
      const { status } = runtime;
      if (!status.configured) {
        throw createHttpError(503, status.reason || "Email delivery is not configured.");
      }

      const order = notification?.order || null;
      if (!order) {
        throw createHttpError(400, "Dropped-at-office order details are required.");
      }

      const route = resolveDroppedOfficeEmailRoute(runtime.config, order);
      const recipient = String(route.to || "").trim();
      if (!recipient) {
        return {
          ok: true,
          skipped: true,
          reason: "No dropped-at-office inbox is configured.",
        };
      }

      const completedAt = formatDateTime(notification?.completedAt || order?.completedAt || new Date().toISOString());
      const completedByName = String(notification?.completedByName || order?.completedByName || "").trim() || "Unknown user";
      const subject = `Dropped at office: ${getOrderPrimaryDisplay(order) || "Route Ledger entry"}`;
      const identifierText = route.identifier
        ? route.identifierLabel
          ? `${route.identifierLabel} ${route.identifier}`
          : route.identifier
        : "None";
      const text = [
        "A Route Ledger entry was marked as dropped at the office.",
        "",
        `Sent to: ${recipient}`,
        `Routing rule: ${route.ruleLabel}`,
        `Identifier used: ${identifierText}`,
        `Entry: ${getOrderPrimaryDisplay(order) || "Unknown"}`,
        `Other references: ${getOrderOtherReferenceLines(order).join(" | ") || "None"}`,
        `Entry type: ${getOrderEntryTypeLabel(order.entryType) || "Not set"}`,
        `Completed by: ${completedByName}`,
        `Completed at: ${completedAt || "Just now"}`,
        `Driver: ${String(order.driverName || "").trim() || "Unassigned"}`,
        `Pickup location: ${String(order.locationName || "").trim() || "Unknown"}`,
        `Pickup address: ${String(order.locationAddress || "").trim() || "Unknown"}`,
        `Delivery address: ${String(order.deliveryAddress || "").trim() || "Not set"}`,
        `Destination: ${getCollectionDestinationLabel(order) || "Office"}`,
        `Move to factory: ${getMoveToFactoryLabel(order) || "No"}`,
        `Stock required: ${String(order.stockDescription || "").trim() || "Not set"}`,
        `Branding: ${String(order.branding || "").trim() || "None"}`,
        `Notes: ${String(order.notes || "").trim() || "None"}`,
        `Schedule: ${getOrderScheduleSummary(order) || "Not scheduled"}`,
      ].join("\n");

      await sendMessage(runtime, {
        fromAddress: status.fromAddress,
        senderName: status.senderName,
        to: recipient,
        subject,
        text,
      });

      return {
        ok: true,
        sentTo: recipient,
        sentFrom: status.from,
        subject,
        rule: route.ruleKey,
      };
    },
    async sendOrderDeleteLogEmail(entries) {
      const runtime = getRuntime();
      const { status } = runtime;
      if (!status.configured) {
        throw createHttpError(503, status.reason || "Email delivery is not configured.");
      }

      const items = Array.isArray(entries) ? entries.filter(Boolean) : [];
      const affectedCount = items.length;
      if (!affectedCount) {
        return {
          ok: true,
          skipped: true,
          count: 0,
        };
      }

      const previewItems = items.slice(0, 25);
      const remainingCount = Math.max(affectedCount - previewItems.length, 0);
      const itemLabel = affectedCount === 1 ? "entry" : "entries";
      const subject = `Logictics Centre delete log (${affectedCount} ${itemLabel})`;
      const lines = [
        "One or more entries were deleted and captured in the server delete log.",
        "",
        `Deleted entries: ${affectedCount}`,
      ];

      if (previewItems.length) {
        lines.push("");
        lines.push("Deleted items:");
        previewItems.forEach((entry) => {
          lines.push(`- ${formatDeletedOrderLogLine(entry)}`);
        });
      }

      if (remainingCount > 0) {
        lines.push(`- Plus ${remainingCount} more entr${remainingCount === 1 ? "y" : "ies"}`);
      }

      await sendMessage(runtime, {
        fromAddress: status.fromAddress,
        senderName: status.senderName,
        to: status.adminActionTo,
        subject,
        text: lines.join("\n"),
      });

      return {
        ok: true,
        sentTo: status.adminActionTo,
        sentFrom: status.from,
        subject,
        count: affectedCount,
      };
    },
    async sendArtworkRequest(token, payload) {
      const { runtime, currentUser, snapshot } = await getAuthorizedMailContext(
        token,
        ["admin", "logistics"],
        "Only admin or logistics users can send artwork requests.",
      );
      const { status } = runtime;

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

      const destination = status.artworkTo || status.fromAddress;
      const subject = `Artwork request: ${stockItem.name}${stockItem.sku ? ` (${stockItem.sku})` : ""}`;
      const text = [
        "Artwork has been requested for stock preparation.",
        "",
        `Requested by: ${currentUser.name}`,
        `Item: ${stockItem.name}`,
        `SKU: ${stockItem.sku || "Not set"}`,
        `Requested quantity: ${requestedQuantity}`,
        `Unit: ${getStockUnitLabel(stockItem)}`,
        `Current on hand: ${Number(stockItem.onHandQuantity || 0)}`,
        `Notes: ${notes || "None"}`,
      ].join("\n");

      await sendMessage(runtime, {
        fromAddress: status.fromAddress,
        senderName: status.senderName,
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

  const disabled = normalizeBoolean(
    firstNonEmpty([
      process.env.MAIL_DISABLED,
      fileConfig.disabled,
    ]),
    false,
  );
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
  const senderName = firstNonEmpty([
    process.env.MAIL_FROM_NAME,
    fileConfig.senderName,
    DEFAULT_MAIL_SENDER_NAME,
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
  const adminActionTo = firstNonEmpty([
    process.env.MAIL_ADMIN_ACTION_TO,
    fileConfig.adminActionTo,
    to,
    DEFAULT_ADMIN_ACTION_NOTIFICATION_EMAIL,
  ]);
  const rolloverTestTo = firstNonEmpty([
    process.env.MAIL_ROLLOVER_TEST_TO,
    fileConfig.rolloverTestTo,
    DEFAULT_ROLLOVER_TEST_EMAIL,
  ]);
  const droppedOfficeSsTo = firstNonEmpty([
    process.env.MAIL_DROPPED_OFFICE_SS_TO,
    fileConfig.droppedOfficeSsTo,
    DEFAULT_DROPPED_OFFICE_SS_EMAIL,
  ]);
  const droppedOfficeSbTo = firstNonEmpty([
    process.env.MAIL_DROPPED_OFFICE_SB_TO,
    fileConfig.droppedOfficeSbTo,
    DEFAULT_DROPPED_OFFICE_SB_EMAIL,
  ]);
  const droppedOfficeMorMarTo = firstNonEmpty([
    process.env.MAIL_DROPPED_OFFICE_MOR_MAR_TO,
    fileConfig.droppedOfficeMorMarTo,
    DEFAULT_DROPPED_OFFICE_MOR_MAR_EMAIL,
  ]);
  const droppedOfficeOrderTo = firstNonEmpty([
    process.env.MAIL_DROPPED_OFFICE_ORDER_TO,
    fileConfig.droppedOfficeOrderTo,
    DEFAULT_DROPPED_OFFICE_ORDER_EMAIL,
  ]);
  const droppedOfficeFallbackTo = firstNonEmpty([
    process.env.MAIL_DROPPED_OFFICE_FALLBACK_TO,
    fileConfig.droppedOfficeFallbackTo,
    DEFAULT_DROPPED_OFFICE_FALLBACK_EMAIL,
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
      disabled,
      provider,
      transport,
      from,
      senderName,
      to,
      artworkTo,
      adminActionTo,
      rolloverTestTo,
      droppedOfficeSsTo,
      droppedOfficeSbTo,
      droppedOfficeMorMarTo,
      droppedOfficeOrderTo,
      droppedOfficeFallbackTo,
    };
  }

  return {
    disabled,
    provider,
    from,
    senderName,
    to,
    artworkTo,
    adminActionTo,
    rolloverTestTo,
    droppedOfficeSsTo,
    droppedOfficeSbTo,
    droppedOfficeMorMarTo,
    droppedOfficeOrderTo,
    droppedOfficeFallbackTo,
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

function getRuntimeMailConfig(baseConfig, overrides = null) {
  const safeOverrides = overrides && typeof overrides === "object" ? overrides : {};
  const normalized = {
    ...baseConfig,
    disabled: typeof safeOverrides.disabled === "boolean" ? safeOverrides.disabled : Boolean(baseConfig.disabled),
    senderName: String(safeOverrides.senderName || baseConfig.senderName || DEFAULT_MAIL_SENDER_NAME).trim() || DEFAULT_MAIL_SENDER_NAME,
    to: String(safeOverrides.to || baseConfig.to || "").trim(),
    artworkTo: String(safeOverrides.artworkTo || baseConfig.artworkTo || baseConfig.from || "").trim(),
    adminActionTo: String(safeOverrides.adminActionTo || baseConfig.adminActionTo || baseConfig.to || DEFAULT_ADMIN_ACTION_NOTIFICATION_EMAIL).trim(),
    rolloverTestTo: String(safeOverrides.rolloverTestTo || baseConfig.rolloverTestTo || DEFAULT_ROLLOVER_TEST_EMAIL).trim(),
    droppedOfficeSsTo: String(safeOverrides.droppedOfficeSsTo || baseConfig.droppedOfficeSsTo || DEFAULT_DROPPED_OFFICE_SS_EMAIL).trim(),
    droppedOfficeSbTo: String(safeOverrides.droppedOfficeSbTo || baseConfig.droppedOfficeSbTo || DEFAULT_DROPPED_OFFICE_SB_EMAIL).trim(),
    droppedOfficeMorMarTo: String(safeOverrides.droppedOfficeMorMarTo || baseConfig.droppedOfficeMorMarTo || DEFAULT_DROPPED_OFFICE_MOR_MAR_EMAIL).trim(),
    droppedOfficeOrderTo: String(safeOverrides.droppedOfficeOrderTo || baseConfig.droppedOfficeOrderTo || DEFAULT_DROPPED_OFFICE_ORDER_EMAIL).trim(),
    droppedOfficeFallbackTo: String(safeOverrides.droppedOfficeFallbackTo || baseConfig.droppedOfficeFallbackTo || DEFAULT_DROPPED_OFFICE_FALLBACK_EMAIL).trim(),
  };

  if (!normalized.artworkTo) {
    normalized.artworkTo = normalized.from;
  }

  if (!normalized.adminActionTo) {
    normalized.adminActionTo = normalized.to || DEFAULT_ADMIN_ACTION_NOTIFICATION_EMAIL;
  }

  if (!normalized.rolloverTestTo) {
    normalized.rolloverTestTo = DEFAULT_ROLLOVER_TEST_EMAIL;
  }

  if (!normalized.droppedOfficeSsTo) {
    normalized.droppedOfficeSsTo = DEFAULT_DROPPED_OFFICE_SS_EMAIL;
  }

  if (!normalized.droppedOfficeSbTo) {
    normalized.droppedOfficeSbTo = DEFAULT_DROPPED_OFFICE_SB_EMAIL;
  }

  if (!normalized.droppedOfficeMorMarTo) {
    normalized.droppedOfficeMorMarTo = DEFAULT_DROPPED_OFFICE_MOR_MAR_EMAIL;
  }

  if (!normalized.droppedOfficeOrderTo) {
    normalized.droppedOfficeOrderTo = DEFAULT_DROPPED_OFFICE_ORDER_EMAIL;
  }

  if (!normalized.droppedOfficeFallbackTo) {
    normalized.droppedOfficeFallbackTo = DEFAULT_DROPPED_OFFICE_FALLBACK_EMAIL;
  }

  return normalized;
}

function buildMailStatus(config) {
  const status = {
    configured: false,
    reason: "",
    disabled: Boolean(config.disabled),
    from: config.from,
    fromAddress: config.from,
    fromDisplay: formatMailSender(config.from, config.senderName),
    senderName: config.senderName,
    to: config.to,
    artworkTo: config.artworkTo,
    adminActionTo: config.adminActionTo,
    rolloverTestTo: config.rolloverTestTo,
    provider: config.provider,
  };

  if (config.disabled) {
    status.reason = "Email delivery is temporarily disabled.";
  } else if (config.provider === "microsoft-graph") {
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

  return status;
}

function formatMailSender(address, senderName = "") {
  const normalizedAddress = String(address || "").trim();
  const normalizedName = String(senderName || "").trim();

  if (!normalizedAddress) {
    return normalizedName;
  }

  if (!normalizedName) {
    return normalizedAddress;
  }

  return `${normalizedName} <${normalizedAddress}>`;
}

function getOrderQuoteNumber(order) {
  return String(order?.quoteNumber || "").trim();
}

function getOrderSalesOrderNumber(order) {
  const salesOrderNumber = String(order?.salesOrderNumber || "").trim();
  if (!salesOrderNumber) {
    return "";
  }

  const orderNumber = String(order?.orderNumber || "").trim();
  if (orderNumber && salesOrderNumber.toLowerCase() === `legacy-${orderNumber}`.toLowerCase()) {
    return "";
  }

  return salesOrderNumber;
}

function getOrderPrimaryDisplay(order) {
  const quoteNumber = getOrderQuoteNumber(order);
  if (quoteNumber) {
    return `Inhouse ${quoteNumber}`;
  }

  const orderNumber = String(order?.orderNumber || "").trim();
  return orderNumber ? `Entry ${orderNumber}` : "Entry";
}

function getOrderOtherReferenceLines(order) {
  const lines = [];
  const salesOrderNumber = getOrderSalesOrderNumber(order);
  const invoiceNumber = String(order?.invoiceNumber || "").trim();
  const poNumber = String(order?.poNumber || "").trim();

  if (salesOrderNumber) {
    lines.push(`SO ${salesOrderNumber}`);
  }

  if (invoiceNumber) {
    lines.push(`Invoice ${invoiceNumber}`);
  }

  if (poNumber) {
    lines.push(`PO ${poNumber}`);
  }

  return lines;
}

function getOrderReferenceSummary(order) {
  const primary = getOrderPrimaryDisplay(order);
  return [primary, ...getOrderOtherReferenceLines(order)].filter(Boolean).join(" | ");
}

function getDroppedOfficeReferenceCandidates(order) {
  return [
    { label: "Inhouse order", value: getOrderQuoteNumber(order) },
    { label: "Sales order", value: getOrderSalesOrderNumber(order) },
    { label: "Invoice", value: String(order?.invoiceNumber || "").trim() },
    { label: "PO", value: String(order?.poNumber || "").trim() },
  ].filter((entry) => entry.value);
}

function referenceStartsWithAny(value, prefixes) {
  const normalizedValue = String(value || "").trim().toUpperCase();
  return prefixes.some((prefix) => normalizedValue.startsWith(prefix));
}

function resolveDroppedOfficeEmailRoute(config, order) {
  const safeConfig = config && typeof config === "object" ? config : {};
  const recipients = {
    ss: String(safeConfig.droppedOfficeSsTo || DEFAULT_DROPPED_OFFICE_SS_EMAIL).trim() || DEFAULT_DROPPED_OFFICE_SS_EMAIL,
    sb: String(safeConfig.droppedOfficeSbTo || DEFAULT_DROPPED_OFFICE_SB_EMAIL).trim() || DEFAULT_DROPPED_OFFICE_SB_EMAIL,
    morMar: String(safeConfig.droppedOfficeMorMarTo || DEFAULT_DROPPED_OFFICE_MOR_MAR_EMAIL).trim() || DEFAULT_DROPPED_OFFICE_MOR_MAR_EMAIL,
    order: String(safeConfig.droppedOfficeOrderTo || DEFAULT_DROPPED_OFFICE_ORDER_EMAIL).trim() || DEFAULT_DROPPED_OFFICE_ORDER_EMAIL,
    fallback: String(safeConfig.droppedOfficeFallbackTo || DEFAULT_DROPPED_OFFICE_FALLBACK_EMAIL).trim() || DEFAULT_DROPPED_OFFICE_FALLBACK_EMAIL,
  };
  const references = getDroppedOfficeReferenceCandidates(order);

  for (const reference of references) {
    if (referenceStartsWithAny(reference.value, DROPPED_OFFICE_SS_PREFIXES)) {
      return {
        to: recipients.ss,
        ruleKey: "ss",
        ruleLabel: "SS/PSS reference",
        identifier: reference.value,
        identifierLabel: reference.label,
      };
    }
    if (referenceStartsWithAny(reference.value, DROPPED_OFFICE_SB_PREFIXES)) {
      return {
        to: recipients.sb,
        ruleKey: "sb",
        ruleLabel: "SB/PSB reference",
        identifier: reference.value,
        identifierLabel: reference.label,
      };
    }
    if (referenceStartsWithAny(reference.value, DROPPED_OFFICE_MOR_MAR_PREFIXES)) {
      return {
        to: recipients.morMar,
        ruleKey: "mor-mar",
        ruleLabel: "MAR/MOR/PMAR/PMOR reference",
        identifier: reference.value,
        identifierLabel: reference.label,
      };
    }
    if (referenceStartsWithAny(reference.value, DROPPED_OFFICE_ORDER_PREFIXES)) {
      return {
        to: recipients.order,
        ruleKey: "order",
        ruleLabel: "Order/SO/BAR reference",
        identifier: reference.value,
        identifierLabel: reference.label,
      };
    }
  }

  const fallbackReference = references[0] || null;
  return {
    to: recipients.fallback,
    ruleKey: fallbackReference ? "fallback" : "no-identifier",
    ruleLabel: fallbackReference ? "Fallback inbox" : "No identifier fallback",
    identifier: fallbackReference?.value || "",
    identifierLabel: fallbackReference?.label || "",
  };
}

function getOrderPriority(order) {
  const priority = String(order?.priority || "medium").trim().toLowerCase();
  return ["high", "medium", "low"].includes(priority) ? priority : "medium";
}

function getOrderDriverIssue(order) {
  const label = getOrderFlagLabel(order);
  const note = String(order?.driverFlagNote || "").trim();
  return [label, note].filter(Boolean).join(": ");
}

function getOrderPickupCsvValue(order) {
  const pickedUpBy = String(order?.pickedUpByName || "").trim();
  const pickedUpAt = formatDateTime(order?.pickedUpAt);

  if (!pickedUpBy && !pickedUpAt) {
    return "";
  }

  if (pickedUpBy && pickedUpAt) {
    return `${pickedUpBy} (${pickedUpAt})`;
  }

  return pickedUpBy || pickedUpAt;
}

function getOrderDriverHandoffCsvValue(order) {
  const pickedUpByUserId = String(order?.pickedUpByUserId || "").trim();
  const currentDriverUserId = String(order?.driverUserId || "").trim();
  const pickedUpByName = String(order?.pickedUpByName || "").trim();
  const currentDriverName = String(order?.driverName || "").trim();

  if (!pickedUpByUserId || !currentDriverUserId || pickedUpByUserId === currentDriverUserId) {
    return "";
  }

  return [pickedUpByName || "Picked-up driver", currentDriverName || "Assigned driver"].join(" -> ");
}

function formatAdminActionOrderLine(order) {
  const reference = getOrderPrimaryDisplay(order);
  const locationName = String(order?.locationName || "Unknown location").trim() || "Unknown location";
  const driverName = String(order?.driverName || "").trim() || "Unassigned";
  const schedule = getOrderScheduleSummary(order) || "Not scheduled";
  const references = getOrderOtherReferenceLines(order).join(" | ");
  const parts = [
    reference,
    locationName,
    `Driver: ${driverName}`,
    `Schedule: ${schedule}`,
  ];

  if (references) {
    parts.push(references);
  }

  return parts.join(" | ");
}

function getCollectionDestinationLabel(order) {
  if (String(order?.entryType || "").trim().toLowerCase() !== "collection") {
    return "";
  }

  if (order?.moveToFactory) {
    return String(order?.factoryDestinationName || "").trim() || "Factory";
  }

  return "Office";
}

function formatDateOnly(value) {
  if (!value) {
    return "";
  }

  const date = new Date(`${value}T00:00:00`);
  return new Intl.DateTimeFormat("en-ZA", {
    day: "numeric",
    month: "short",
    year: "numeric",
    timeZone: TIME_ZONE,
  }).format(date);
}

function getCurrentLocalDateValue() {
  const parts = new Intl.DateTimeFormat("en", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    timeZone: TIME_ZONE,
  }).formatToParts(new Date());
  const year = parts.find((part) => part.type === "year")?.value || "";
  const month = parts.find((part) => part.type === "month")?.value || "";
  const day = parts.find((part) => part.type === "day")?.value || "";
  return [year, month, day].filter(Boolean).join("-");
}

function formatDateTime(value) {
  if (!value) {
    return "";
  }

  const date = new Date(value);
  return new Intl.DateTimeFormat("en-ZA", {
    day: "numeric",
    month: "short",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
    timeZone: TIME_ZONE,
  }).format(date);
}

function capitalize(value) {
  const text = String(value || "").trim();
  if (!text) {
    return "";
  }

  return text.charAt(0).toUpperCase() + text.slice(1);
}

function shouldMarkOrderAsRolledOver(order) {
  return Boolean(order?.driverUserId) && Number(order?.carryOverCount || 0) > 0;
}

function getOrderScheduledForValue(order) {
  return String(order?.scheduledFor || "").trim();
}

function filterOrdersForScheduledDate(orders, scheduledFor) {
  const targetDate = String(scheduledFor || "").trim();
  const items = Array.isArray(orders) ? orders.filter(Boolean) : [];
  if (!targetDate) {
    return items;
  }

  return items.filter((order) => getOrderScheduledForValue(order) === targetDate);
}

function getOrderScheduleSummary(order) {
  const scheduledFor = formatDateOnly(order?.scheduledFor);
  const originalScheduledFor = formatDateOnly(order?.originalScheduledFor);
  const carryOverCount = Number(order?.carryOverCount || 0);
  const dayLabel = carryOverCount === 1 ? "day" : "days";

  if (!shouldMarkOrderAsRolledOver(order)) {
    return scheduledFor;
  }

  if (scheduledFor && originalScheduledFor && carryOverCount > 0) {
    return `${scheduledFor} (Rolled from ${originalScheduledFor}, ${carryOverCount} ${dayLabel})`;
  }

  if (scheduledFor && carryOverCount > 0) {
    return `${scheduledFor} (Rolled over ${carryOverCount} ${dayLabel})`;
  }

  return scheduledFor;
}

function buildOrderCsvRow(order) {
  return [
    getOrderQuoteNumber(order),
    order.driverName || "Unassigned",
    getOrderPickupCsvValue(order),
    getOrderDriverHandoffCsvValue(order),
    order.locationName || "",
    order.locationAddress || "",
    order.deliveryAddress || "",
    getOrderEntryTypeLabel(order.entryType),
    capitalize(getOrderPriority(order)),
    getOrderOtherReferenceLines(order).join(" | "),
    order.stockDescription || "",
    order.branding || "",
    order.moveToFactory ? "Yes" : "No",
    getCollectionDestinationLabel(order),
    getOrderDriverIssue(order),
    order.notes || "",
    getOrderScheduleSummary(order),
    capitalize(order.status || "active"),
    getOrderCompletionLabel(order),
    order.createdByName || "",
    formatDateTime(order.createdAt),
    formatDateTime(order.completedAt),
  ];
}

function buildOrdersCsv(orders) {
  const lineBreak = "\r\n";
  const rows = [
    [
      "Inhouse order",
      "Assigned driver",
      "Picked up by",
      "Driver handoff",
      "Pickup location",
      "Pickup address",
      "Delivery address",
      "Entry type",
      "Priority",
      "Other references",
      "Stock required",
      "Branding",
      "Move to factory",
      "Destination",
      "Driver issue",
      "Extra notes",
      "Schedule",
      "Status",
      "Completion",
      "Created by",
      "Created",
      "Completed",
    ],
    ...orders.map(buildOrderCsvRow),
  ];

  return `\uFEFFsep=,${lineBreak}${rows
    .map((columns) => columns.map(escapeCsvValue).join(","))
    .join(lineBreak)}`;
}

function buildOrderNoticeText(order) {
  const lines = [];
  const stockDescription = String(order?.stockDescription || "").trim();
  const branding = String(order?.branding || "").trim();
  const notice = String(order?.notes || "").trim();
  const driverFlag = getOrderFlagNoticeText(order);
  const moveToFactory = getMoveToFactoryText(order);
  const completion = getOrderCompletionText(order);

  if (driverFlag) {
    lines.push(driverFlag);
  }

  if (stockDescription) {
    lines.push(`Stock: ${stockDescription}.`);
  }

  if (branding) {
    lines.push(`Branding: ${branding}.`);
  }

  if (notice) {
    lines.push(notice);
  }

  if (moveToFactory) {
    lines.push(moveToFactory);
  }

  if (completion) {
    lines.push(completion);
  }

  if (shouldMarkOrderAsRolledOver(order)) {
    const carryOverCount = Number(order?.carryOverCount || 0);
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

function buildCarryOverTestPayload(snapshot) {
  const today = String(snapshot?.today || "").trim() || getCurrentLocalDateValue();
  const carriedOrders = filterOrdersForScheduledDate(snapshot?.orders, today)
    .filter((order) =>
      String(order?.status || "").trim().toLowerCase() === "active"
      && shouldMarkOrderAsRolledOver(order),
    )
    .sort(compareCarryOverEmailOrders);

  return {
    today,
    updatedOrders: carriedOrders.length,
    carriedOrders,
  };
}

function compareCarryOverEmailOrders(left, right) {
  const leftUnassigned = left?.driverUserId ? 1 : 0;
  const rightUnassigned = right?.driverUserId ? 1 : 0;
  if (leftUnassigned !== rightUnassigned) {
    return leftUnassigned - rightUnassigned;
  }

  const leftDriver = String(left?.driverName || "").trim().toLowerCase();
  const rightDriver = String(right?.driverName || "").trim().toLowerCase();
  if (leftDriver !== rightDriver) {
    return leftDriver.localeCompare(rightDriver);
  }

  const leftLocation = String(left?.locationName || "").trim().toLowerCase();
  const rightLocation = String(right?.locationName || "").trim().toLowerCase();
  if (leftLocation !== rightLocation) {
    return leftLocation.localeCompare(rightLocation);
  }

  return Number(left?.orderNumber || 0) - Number(right?.orderNumber || 0);
}

function buildCarryOverEmailSummary(rollover, csvOrders, options = {}) {
  const carriedOrders = Array.isArray(rollover?.carriedOrders)
    ? rollover.carriedOrders.filter(Boolean).sort(compareCarryOverEmailOrders)
    : [];
  const count = Number(rollover?.updatedOrders || carriedOrders.length || 0);
  const scheduledFor = String(rollover?.today || "").trim();
  const assignedCount = carriedOrders.filter((order) => order?.driverUserId).length;
  const unassignedCount = Math.max(carriedOrders.length - assignedCount, 0);

  return {
    scheduledFor,
    count,
    assignedCount,
    unassignedCount,
    csvCount: Array.isArray(csvOrders) ? csvOrders.length : 0,
    groups: groupCarryOverOrdersByDriver(carriedOrders),
    isTest: Boolean(options.isTest),
    requestedByName: String(options.requestedByName || "").trim(),
    recipient: String(options.recipient || "").trim(),
  };
}

function buildCarryOverEmailText(summary) {
  const lines = [
    summary?.isTest
      ? "This is a test send of the Route Ledger carry-over email."
      : "Open Route Ledger items were rolled forward to the next day.",
    "",
    `Scheduled date: ${summary?.scheduledFor || "Unknown"}`,
    `Rolled items: ${Number(summary?.count || 0)}`,
    `Assigned: ${Number(summary?.assignedCount || 0)}`,
    `Unassigned: ${Number(summary?.unassignedCount || 0)}`,
    `CSV attached: ${Number(summary?.csvCount || 0)} entries`,
  ];

  if (summary?.requestedByName) {
    lines.splice(2, 0, `Requested by: ${summary.requestedByName}`);
  }

  if (summary?.recipient) {
    lines.splice(summary?.requestedByName ? 3 : 2, 0, `Sent to: ${summary.recipient}`);
  }

  if (!Array.isArray(summary?.groups) || !summary.groups.length) {
    lines.push("", "No rolled-over active items are currently scheduled for this date.");
    return lines.join("\n");
  }

  summary.groups.forEach((group) => {
    lines.push("");
    lines.push(`Driver: ${group.driverName} (${group.count})`);
    group.locationGroups.forEach((locationGroup) => {
      lines.push(`Location: ${locationGroup.locationName} (${locationGroup.count})`);
      locationGroup.orders.forEach((order) => {
        lines.push(`- ${getCarryOverEmailOrderTitle(order)}`);
        buildCarryOverEmailOrderDetails(order).forEach((detail) => {
          lines.push(`  ${detail.label}: ${detail.value}`);
        });
      });
    });
  });

  return lines.join("\n");
}

function buildCarryOverEmailHtml(summary) {
  const groups = Array.isArray(summary?.groups) ? summary.groups : [];
  const summaryCards = [
    { label: "Scheduled date", value: summary?.scheduledFor || "Unknown" },
    { label: "Rolled items", value: String(Number(summary?.count || 0)) },
    { label: "Assigned", value: String(Number(summary?.assignedCount || 0)) },
    { label: "Unassigned", value: String(Number(summary?.unassignedCount || 0)) },
    { label: "CSV attached", value: `${Number(summary?.csvCount || 0)} entries` },
  ];
  const intro = summary?.isTest
    ? "This is a test send of the Route Ledger carry-over email."
    : "Open Route Ledger items were rolled forward to the next day.";
  const metaLines = [
    summary?.requestedByName ? `<p style="margin:0 0 6px;">Requested by: <strong>${escapeHtml(summary.requestedByName)}</strong></p>` : "",
    summary?.recipient ? `<p style="margin:0 0 18px;">Sent to: <strong>${escapeHtml(summary.recipient)}</strong></p>` : "",
  ].join("");

  return `<!doctype html>
<html lang="en">
  <body style="margin:0;padding:24px;background:#f4f1ea;color:#1f2937;font-family:Segoe UI,Arial,sans-serif;">
    <div style="max-width:860px;margin:0 auto;background:#ffffff;border:1px solid #e5dccb;border-radius:20px;overflow:hidden;">
      <div style="padding:28px 32px;background:linear-gradient(135deg,#1d3557 0%,#355070 100%);color:#ffffff;">
        <p style="margin:0 0 8px;font-size:12px;letter-spacing:0.12em;text-transform:uppercase;opacity:0.78;">Route Ledger</p>
        <h1 style="margin:0;font-size:28px;line-height:1.2;">Carry-over summary</h1>
        <p style="margin:12px 0 0;font-size:15px;line-height:1.6;max-width:640px;">${escapeHtml(intro)}</p>
      </div>
      <div style="padding:28px 32px;">
        ${metaLines}
        <div style="margin:0 0 24px;font-size:14px;color:#475569;">The full daily CSV export is attached to this email.</div>
        <div style="font-size:0;margin:0 -8px 24px;">
          ${summaryCards.map((card) => `
            <div style="display:inline-block;vertical-align:top;width:156px;min-width:156px;margin:0 8px 12px;background:#f8fafc;border:1px solid #dbe4ee;border-radius:16px;padding:14px 16px;box-sizing:border-box;">
              <div style="font-size:12px;text-transform:uppercase;letter-spacing:0.08em;color:#64748b;margin:0 0 8px;">${escapeHtml(card.label)}</div>
              <div style="font-size:20px;font-weight:700;color:#0f172a;">${escapeHtml(card.value)}</div>
            </div>
          `).join("")}
        </div>
        ${
          groups.length
            ? groups.map((group) => `
              <section style="margin:0 0 28px;">
                <div style="margin:0 0 12px;padding:12px 16px;border-radius:14px;background:#efe8dc;color:#3f3a2f;font-size:16px;font-weight:700;">
                  ${escapeHtml(group.driverName)} (${group.count})
                </div>
                ${group.locationGroups.map((locationGroup) => `
                  <section style="margin:0 0 16px;">
                    <div style="margin:0 0 10px;padding:10px 14px;border-radius:12px;background:#f8f3ea;color:#5b5547;font-size:14px;font-weight:700;">
                      ${escapeHtml(locationGroup.locationName)} (${locationGroup.count})
                    </div>
                    ${locationGroup.orders.map((order) => renderCarryOverEmailHtmlCard(order)).join("")}
                  </section>
                `).join("")}
              </section>
            `).join("")
            : `
              <div style="padding:18px 20px;border:1px dashed #c8b79d;border-radius:16px;background:#fbf8f3;color:#5b5547;">
                No rolled-over active items are currently scheduled for this date.
              </div>
            `
        }
      </div>
    </div>
  </body>
</html>`;
}

function renderCarryOverEmailHtmlCard(order) {
  const details = buildCarryOverEmailOrderDetails(order);
  return `
    <article style="margin:0 0 14px;padding:18px 20px;border:1px solid #e5e7eb;border-radius:16px;background:#ffffff;">
      <div style="margin:0 0 10px;font-size:18px;font-weight:700;color:#0f172a;">
        ${escapeHtml(getCarryOverEmailOrderTitle(order))}
      </div>
      ${details.map((detail) => `
        <p style="margin:0 0 6px;font-size:14px;line-height:1.6;color:#334155;">
          <strong style="color:#0f172a;">${escapeHtml(detail.label)}:</strong> ${escapeHtml(detail.value)}
        </p>
      `).join("")}
    </article>
  `;
}

function groupCarryOverOrdersByDriver(orders) {
  const groups = [];
  const groupedByDriver = new Map();

  orders.forEach((order) => {
    const driverName = String(order?.driverName || "").trim() || "Unassigned";
    let group = groupedByDriver.get(driverName);

    if (!group) {
      group = {
        driverName,
        count: 0,
        locationGroups: [],
        locationsByName: new Map(),
      };
      groupedByDriver.set(driverName, group);
      groups.push(group);
    }

    group.count += 1;

    const locationName = String(order?.locationName || "").trim() || "Unknown location";
    let locationGroup = group.locationsByName.get(locationName);

    if (!locationGroup) {
      locationGroup = {
        locationName,
        count: 0,
        orders: [],
      };
      group.locationsByName.set(locationName, locationGroup);
      group.locationGroups.push(locationGroup);
    }

    locationGroup.count += 1;
    locationGroup.orders.push(order);
  });

  return groups.map((group) => ({
    driverName: group.driverName,
    count: group.count,
    locationGroups: group.locationGroups,
  }));
}

function getCarryOverEmailOrderTitle(order) {
  const reference = getOrderPrimaryDisplay(order);
  const customerName = String(order?.customerName || "").trim();
  const locationName = String(order?.locationName || "").trim();
  return [reference || "Route Ledger entry", customerName || "", locationName || ""]
    .filter(Boolean)
    .join(" | ");
}

function formatDeletedOrderLogLine(entry) {
  const reference = getOrderPrimaryDisplay(entry);
  const locationName = String(entry?.locationName || "Unknown location").trim() || "Unknown location";
  const driverName = String(entry?.driverName || "").trim() || "Unassigned";
  const deletedByName = String(entry?.deletedByName || "Unknown user").trim() || "Unknown user";
  const deletedByRole = capitalize(String(entry?.deletedByRole || "").trim()) || "User";
  const deletedAt = formatDateTime(entry?.deletedAt) || "Just now";
  const otherReferences = getOrderOtherReferenceLines(entry).join(" | ");
  const parts = [
    reference,
    locationName,
    `Driver: ${driverName}`,
    `Deleted by: ${deletedByName} (${deletedByRole})`,
    `Deleted at: ${deletedAt}`,
  ];

  if (otherReferences) {
    parts.push(otherReferences);
  }

  return parts.join(" | ");
}

function buildCarryOverEmailOrderDetails(order) {
  const details = [];
  const entryType = getOrderEntryTypeLabel(order?.entryType);
  const locationName = String(order?.locationName || "").trim();
  const deliveryAddress = String(order?.deliveryAddress || "").trim();
  const collectionDestination = getCollectionDestinationLabel(order);
  const priority = capitalize(getOrderPriority(order));
  const quoteNumber = getOrderQuoteNumber(order);
  const salesOrderNumber = getOrderSalesOrderNumber(order);
  const invoiceNumber = String(order?.invoiceNumber || "").trim();
  const poNumber = String(order?.poNumber || "").trim();
  const scheduleText = buildCarryOverScheduleText(order);
  const noticeText = buildCarryOverEmailNoticeText(order);

  if (entryType) {
    details.push({ label: "Type", value: entryType });
  }

  if (locationName) {
    details.push({ label: "Location", value: locationName });
  }

  if (deliveryAddress) {
    details.push({ label: "Delivery address", value: deliveryAddress });
  }

  if (collectionDestination) {
    details.push({ label: "Destination", value: collectionDestination });
  }

  if (priority) {
    details.push({ label: "Priority", value: priority });
  }

  if (quoteNumber) {
    details.push({ label: "Inhouse order", value: quoteNumber });
  }

  if (salesOrderNumber) {
    details.push({ label: "Sales order", value: salesOrderNumber });
  }

  if (invoiceNumber) {
    details.push({ label: "Invoice", value: invoiceNumber });
  }

  if (poNumber) {
    details.push({ label: "PO", value: poNumber });
  }

  if (scheduleText) {
    details.push({ label: "Schedule", value: scheduleText });
  }

  if (noticeText) {
    details.push({ label: "Notes", value: noticeText });
  }

  return details;
}

function buildCarryOverScheduleText(order) {
  const scheduledFor = String(order?.scheduledFor || "").trim();
  const originalScheduledFor = String(order?.originalScheduledFor || "").trim();
  const carryOverCount = Number(order?.carryOverCount || 0);
  const dayLabel = carryOverCount === 1 ? "day" : "days";
  const parts = [];

  if (scheduledFor) {
    parts.push(`Scheduled for ${scheduledFor}`);
  }

  if (carryOverCount > 0) {
    parts.push(`Carry-over total: ${carryOverCount} ${dayLabel}`);
  }

  if (originalScheduledFor && originalScheduledFor !== scheduledFor) {
    parts.push(`Original date: ${originalScheduledFor}`);
  }

  return parts.join(" | ");
}

function buildCarryOverEmailNoticeText(order) {
  const parts = [];
  const flagNotice = getOrderFlagNoticeText(order);
  const note = String(order?.notes || "").trim();
  const moveToFactory = getMoveToFactoryText(order);

  if (flagNotice) {
    parts.push(flagNotice);
  }

  if (note) {
    parts.push(`Notice: ${note}`);
  }

  if (moveToFactory) {
    parts.push(moveToFactory);
  }

  return parts.join(" ");
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function getOrderEntryTypeLabel(entryType) {
  const normalized = String(entryType || "").trim().toLowerCase();

  if (!normalized) {
    return "";
  }

  return normalized === "collection"
    ? "Collection"
    : normalized === "delivery"
      ? "Delivery"
      : normalized;
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

function getEmailRecipientList(value) {
  if (Array.isArray(value)) {
    return value
      .map((entry) => String(entry || "").trim())
      .filter(Boolean);
  }

  return String(value || "")
    .split(/[,\n;]+/)
    .map((entry) => entry.trim())
    .filter(Boolean);
}

async function sendViaMicrosoftGraph(config, message) {
  const token = await getMicrosoftGraphAccessToken(config);
  const recipients = getEmailRecipientList(message.to);
  const fromAddress = String(message.fromAddress || config.from || "").trim();
  const senderName = String(message.senderName || config.senderName || "").trim();
  if (!recipients.length) {
    throw createHttpError(400, "At least one email recipient is required.");
  }
  if (!fromAddress) {
    throw createHttpError(400, "A sender mailbox is required.");
  }
  const bodyContent = String(message.html || message.text || "");
  const bodyType = message.html ? "HTML" : "Text";
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
        contentType: bodyType,
        content: bodyContent,
      },
      toRecipients: recipients.map((recipient) => ({
        emailAddress: {
          address: recipient,
        },
      })),
      from: {
        emailAddress: {
          address: fromAddress,
          name: senderName || undefined,
        },
      },
    },
    saveToSentItems: true,
  };

  if (attachments.length) {
    payload.message.attachments = attachments;
  }

  const response = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(fromAddress)}/sendMail`, {
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
