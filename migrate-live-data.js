const crypto = require("crypto");
const fs = require("fs");
const path = require("path");
const { loadProjectEnv } = require("./load-project-env");
const { createLocalDatabase } = require("./local-database");

loadProjectEnv(__dirname);

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) {
    printUsage();
    return;
  }

  const sourceUrl = normalizeBaseUrl(args["source-url"]);
  if (!sourceUrl) {
    printUsage();
    throw new Error("Provide --source-url.");
  }

  const hasToken = Boolean(String(args.token || "").trim());
  const hasLogin = Boolean(String(args.name || "").trim() && String(args.password || "").trim());
  if (!hasToken && !hasLogin) {
    printUsage();
    throw new Error("Provide either --token or both --name and --password.");
  }

  const token = hasToken
    ? String(args.token || "").trim()
    : await loginUser(sourceUrl, String(args.name || "").trim(), String(args.password || ""));

  const fallbackPassword = String(args["fallback-password"] || args.password || "").trim();
  const exportPayload = await exportLiveData(sourceUrl, token, {
    fallbackPassword,
  });

  const backupPath = resolveBackupPath(args["backup-file"]);
  fs.mkdirSync(path.dirname(backupPath), { recursive: true });
  fs.writeFileSync(backupPath, JSON.stringify(exportPayload, null, 2), "utf8");
  console.log(`Backup saved: ${backupPath}`);
  console.log(`Export mode: ${exportPayload.exportMode || "full-export"}`);
  console.log(`Export counts: ${formatCounts(exportPayload.counts)}`);
  if (exportPayload.warning) {
    console.log(`Warning: ${exportPayload.warning}`);
  }

  if (args["export-only"]) {
    console.log("Export-only mode enabled. No local import was performed.");
    return;
  }

  const database = createLocalDatabase(process.cwd());
  const status = database.getStatus();
  if (!status.configured) {
    throw new Error(status.reason || "The local database is not available.");
  }

  try {
    const result = await database.replaceAllData(exportPayload.data, {
      source: String(exportPayload.storage || sourceUrl).trim() || sourceUrl,
    });
    console.log(`Imported into: ${status.storagePath}`);
    console.log(`Imported counts: ${formatCounts(result.localCounts)}`);
  } finally {
    await database.close();
  }
}

async function loginUser(baseUrl, name, password) {
  const payload = await postJson(baseUrl, "/api/rpc/login_user", {
    parameters: {
      p_name: name,
      p_password: password,
    },
  });
  const token = String(payload?.data?.token || "").trim();
  if (!token) {
    throw new Error("Login succeeded without returning a session token.");
  }
  return token;
}

async function exportLiveData(baseUrl, token, options = {}) {
  try {
    const payload = await postJson(baseUrl, "/api/admin/data/export", { token });
    if (!payload?.data || typeof payload.data !== "object") {
      throw new Error("The source server did not return export data.");
    }

    return {
      ...payload,
      exportMode: "full-export",
      warning: "",
    };
  } catch (error) {
    const message = String(error?.message || error || "").trim();
    const allowFallback = /method not allowed|request failed with status 404|unknown|not found/i.test(message);
    if (!allowFallback) {
      throw error;
    }

    if (!String(options.fallbackPassword || "").trim()) {
      throw new Error("The live site does not support full export yet, and snapshot rescue needs --fallback-password or the admin password.");
    }

    const snapshotPayload = await postJson(baseUrl, "/api/rpc/get_app_snapshot", {
      parameters: {
        p_token: token,
      },
    });
    const snapshot = snapshotPayload?.data || {};
    const importData = transformSnapshotToImportData(snapshot, {
      fallbackPassword: String(options.fallbackPassword || "").trim(),
    });

    return {
      exportedAt: new Date().toISOString(),
      exportedBy: snapshot?.user?.name || "admin",
      storage: `${baseUrl} (snapshot fallback)`,
      counts: countTables(importData),
      data: importData,
      exportMode: "snapshot-fallback",
      warning: "The live site did not expose the full export route, so the migration used the available admin snapshot. Imported accounts were assigned the fallback password used for this rescue run.",
    };
  }
}

async function postJson(baseUrl, requestPath, payload) {
  const response = await fetch(new URL(requestPath, `${baseUrl}/`), {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload || {}),
  });

  const text = await response.text();
  let parsed = null;

  try {
    parsed = text ? JSON.parse(text) : null;
  } catch (error) {
    parsed = null;
  }

  if (!response.ok) {
    throw new Error(parsed?.error || text || `Request failed with status ${response.status}.`);
  }

  return parsed;
}

function transformSnapshotToImportData(snapshot, options = {}) {
  const fallbackPassword = String(options.fallbackPassword || "").trim();
  if (!fallbackPassword) {
    throw new Error("A fallback password is required for snapshot-based user migration.");
  }

  const users = Array.isArray(snapshot?.users) ? snapshot.users : [];
  const suppliers = Array.isArray(snapshot?.suppliers) ? snapshot.suppliers : [];
  const locations = Array.isArray(snapshot?.locations) ? snapshot.locations : [];
  const orders = Array.isArray(snapshot?.orders) ? snapshot.orders : [];
  const stockItems = Array.isArray(snapshot?.stockItems) ? snapshot.stockItems : [];
  const stockMovements = Array.isArray(snapshot?.stockMovements) ? snapshot.stockMovements : [];
  const artworkRequests = Array.isArray(snapshot?.artworkRequests) ? snapshot.artworkRequests : [];
  const fallbackActorId = String(snapshot?.user?.id || users[0]?.id || "").trim();

  if (!fallbackActorId) {
    throw new Error("The live snapshot did not include a usable admin user id.");
  }

  return {
    app_users: users.map((user) => ({
      id: String(user.id || "").trim(),
      name: String(user.name || "").trim(),
      role: String(user.role || "admin").trim() || "admin",
      password_hash: hashPassword(fallbackPassword),
      active: Boolean(user.active),
      phone: String(user.phone || "").trim() || null,
      vehicle: String(user.vehicle || "").trim() || null,
      last_known_lat: toNullableNumber(user.lastKnownLat),
      last_known_lng: toNullableNumber(user.lastKnownLng),
      last_known_recorded_at: String(user.lastKnownRecordedAt || "").trim() || null,
      created_at: normalizeTimestamp(user.createdAt),
      updated_at: normalizeTimestamp(user.createdAt),
    })),
    suppliers: suppliers.map((supplier) => ({
      id: String(supplier.id || "").trim(),
      name: String(supplier.name || "").trim(),
      contact_person: String(supplier.contactPerson || "").trim(),
      contact_number: String(supplier.contactNumber || "").trim(),
      factory: Boolean(supplier.factory),
      created_by: fallbackActorId,
      created_at: normalizeTimestamp(supplier.createdAt),
      updated_at: normalizeTimestamp(supplier.createdAt),
    })),
    locations: locations.map((location) => ({
      id: String(location.id || "").trim(),
      supplier_id: String(location.supplierId || "").trim() || null,
      location_type: String(location.locationType || "supplier").trim() || "supplier",
      name: String(location.name || "").trim(),
      address: String(location.address || "").trim(),
      lat: toNullableNumber(location.lat),
      lng: toNullableNumber(location.lng),
      contact_person: String(location.contactPerson || "").trim(),
      contact_number: String(location.contactNumber || "").trim(),
      notes: String(location.notes || "").trim(),
      created_by: fallbackActorId,
      created_at: normalizeTimestamp(location.createdAt),
      updated_at: normalizeTimestamp(location.createdAt),
    })),
    orders: orders.map((order) => ({
      id: String(order.id || "").trim(),
      order_number: toNullableInteger(order.orderNumber),
      driver_user_id: String(order.driverUserId || "").trim() || null,
      location_id: String(order.locationId || "").trim(),
      entry_type: String(order.entryType || "delivery").trim() || "delivery",
      factory_order_number: String(order.salesOrderNumber || order.factoryOrderNumber || "").trim(),
      inhouse_order_number: String(order.quoteNumber || order.inhouseOrderNumber || "").trim(),
      invoice_number: String(order.invoiceNumber || "").trim(),
      po_number: String(order.poNumber || "").trim(),
      customer_name: String(order.customerName || "").trim(),
      delivery_address: String(order.deliveryAddress || "").trim(),
      priority: String(order.priority || "medium").trim() || "medium",
      notes: String(order.notes || "").trim(),
      driver_flag_type: String(order.driverFlagType || "").trim() || null,
      driver_flag_note: String(order.driverFlagNote || "").trim(),
      driver_flagged_at: String(order.driverFlaggedAt || "").trim() || null,
      driver_flagged_by_user_id: String(order.driverFlaggedByUserId || "").trim() || null,
      picked_up_at: String(order.pickedUpAt || "").trim() || null,
      picked_up_by_user_id: String(order.pickedUpByUserId || "").trim() || null,
      move_to_factory: Boolean(order.moveToFactory),
      factory_destination_location_id: String(order.factoryDestinationLocationId || "").trim() || null,
      status: String(order.status || "active").trim() || "active",
      scheduled_for: normalizeDate(order.scheduledFor),
      original_scheduled_for: normalizeNullableDate(order.originalScheduledFor),
      carry_over_count: Number(order.carryOverCount || 0),
      created_by_user_id: String(order.createdByUserId || fallbackActorId).trim() || fallbackActorId,
      created_at: normalizeTimestamp(order.createdAt),
      completed_at: String(order.completedAt || "").trim() || null,
      completion_type: String(order.completionType || "").trim() || null,
      completed_by_user_id: String(order.completedByUserId || "").trim() || null,
      updated_at: normalizeTimestamp(order.completedAt || order.createdAt),
      branding: String(order.branding || "").trim(),
      stock_description: String(order.stockDescription || "").trim(),
    })),
    order_delete_log: [],
    stock_items: stockItems.map((item) => ({
      id: String(item.id || "").trim(),
      name: String(item.name || "").trim(),
      sku: String(item.sku || "").trim(),
      quote_number: String(item.quoteNumber || "").trim(),
      invoice_number: String(item.invoiceNumber || "").trim(),
      sales_order_number: String(item.salesOrderNumber || "").trim(),
      po_number: String(item.poNumber || "").trim(),
      unit: String(item.unit || "units").trim() || "units",
      notes: String(item.notes || "").trim(),
      created_source: String(item.createdSource || "manual").trim() || "manual",
      created_by_user_id: fallbackActorId,
      created_at: normalizeTimestamp(item.createdAt),
      updated_at: normalizeTimestamp(item.updatedAt || item.createdAt),
    })),
    stock_movements: stockMovements.map((movement) => ({
      id: String(movement.id || "").trim(),
      stock_item_id: String(movement.stockItemId || "").trim(),
      movement_type: String(movement.movementType || "in").trim() || "in",
      quantity: Number(movement.quantity || 0),
      supplier_name: String(movement.supplierName || "").trim(),
      driver_user_id: String(movement.driverUserId || "").trim() || null,
      notes: String(movement.notes || "").trim(),
      created_by_user_id: String(movement.createdByUserId || fallbackActorId).trim() || fallbackActorId,
      created_at: normalizeTimestamp(movement.createdAt),
    })),
    artwork_requests: artworkRequests.map((request) => ({
      id: String(request.id || "").trim(),
      stock_item_id: String(request.stockItemId || "").trim(),
      requested_quantity: Number(request.requestedQuantity || 0),
      notes: String(request.notes || "").trim(),
      sent_to: String(request.sentTo || "").trim(),
      requested_by_user_id: String(request.requestedByUserId || fallbackActorId).trim() || fallbackActorId,
      sent_at: normalizeTimestamp(request.sentAt),
    })),
    app_sessions: [],
  };
}

function hashPassword(password) {
  const salt = crypto.randomBytes(16).toString("hex");
  const digest = crypto.scryptSync(password, salt, 64).toString("hex");
  return `scrypt:${salt}:${digest}`;
}

function countTables(data) {
  const tables = data && typeof data === "object" ? data : {};
  return Object.fromEntries(
    [
      "app_users",
      "suppliers",
      "locations",
      "orders",
      "order_delete_log",
      "stock_items",
      "stock_movements",
      "artwork_requests",
      "app_sessions",
    ].map((tableName) => [tableName, Array.isArray(tables[tableName]) ? tables[tableName].length : 0]),
  );
}

function normalizeTimestamp(value) {
  const text = String(value || "").trim();
  return text || new Date().toISOString();
}

function normalizeDate(value) {
  const text = normalizeNullableDate(value);
  if (!text) {
    return new Date().toISOString().slice(0, 10);
  }
  return text;
}

function normalizeNullableDate(value) {
  const text = String(value || "").trim();
  if (!text) {
    return null;
  }
  const match = text.match(/^(\d{4}-\d{2}-\d{2})/);
  return match ? match[1] : null;
}

function toNullableNumber(value) {
  if (value === null || value === undefined || String(value).trim() === "") {
    return null;
  }
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function toNullableInteger(value) {
  if (value === null || value === undefined || String(value).trim() === "") {
    return null;
  }
  const number = Number(value);
  return Number.isInteger(number) ? number : null;
}

function parseArgs(argv) {
  const result = {};

  for (let index = 0; index < argv.length; index += 1) {
    const token = String(argv[index] || "");
    if (!token.startsWith("--")) {
      continue;
    }

    const key = token.slice(2);
    const next = argv[index + 1];
    if (typeof next === "string" && !next.startsWith("--")) {
      result[key] = next;
      index += 1;
      continue;
    }

    result[key] = true;
  }

  return result;
}

function normalizeBaseUrl(value) {
  const text = String(value || "").trim();
  if (!text) {
    return "";
  }

  return text.replace(/\/+$/, "");
}

function resolveBackupPath(value) {
  const text = String(value || "").trim();
  if (text) {
    return path.resolve(process.cwd(), text);
  }

  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  return path.resolve(process.cwd(), "data", `migration-backup-${stamp}.json`);
}

function formatCounts(counts) {
  const entries = counts && typeof counts === "object" ? Object.entries(counts) : [];
  if (!entries.length) {
    return "none";
  }

  return entries
    .map(([tableName, count]) => `${tableName}=${Number(count) || 0}`)
    .join(", ");
}

function printUsage() {
  console.log("Usage:");
  console.log("  node migrate-live-data.js --source-url <url> --token <admin-session-token>");
  console.log("  node migrate-live-data.js --source-url <url> --name <admin-name> --password <admin-password>");
  console.log("");
  console.log("Options:");
  console.log("  --backup-file <path>  Save the exported JSON to a custom file.");
  console.log("  --export-only         Download the JSON backup without importing locally.");
  console.log("  --fallback-password <password>  Password to assign when snapshot fallback is used.");
  console.log("");
  console.log("The current local database target follows LOGISTICS_DB_PATH or LOGISTICS_DATA_DIR when set.");
}

main().catch((error) => {
  console.error(String(error?.message || error || "Migration failed."));
  process.exitCode = 1;
});
