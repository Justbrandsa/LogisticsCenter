const crypto = require("crypto");
const fs = require("fs");
const os = require("os");
const path = require("path");

let DatabaseSync = null;
let sqliteLoadError = null;
let bcrypt = null;
let createLibsqlClient = null;
let libsqlLoadError = null;

try {
  ({ DatabaseSync } = require("node:sqlite"));
} catch (error) {
  sqliteLoadError = error;
}

try {
  bcrypt = require("bcryptjs");
} catch (error) {
  bcrypt = null;
}

try {
  ({ createClient: createLibsqlClient } = require("@libsql/client"));
} catch (error) {
  libsqlLoadError = error;
}

const TIME_ZONE = "Africa/Johannesburg";
const SESSION_TTL_MS = 14 * 24 * 60 * 60 * 1000;
const USER_ROLES = new Set(["admin", "sales", "driver", "logistics", "maintenance"]);
const LOCATION_TYPES = new Set(["supplier", "factory", "both", "client"]);
const ORDER_PRIORITIES = new Set(["high", "medium", "low"]);
const ORDER_STATUSES = new Set(["active", "completed"]);
const ORDER_ENTRY_TYPES = new Set(["collection", "delivery"]);
const ORDER_FLAG_TYPES = new Set(["not_collected", "not_ready"]);
const ORDER_COMPLETION_TYPES = new Set(["office", "factory"]);
const STOCK_MOVEMENT_TYPES = new Set(["in", "out"]);
const INHOUSE_ORDER_PREFIXES = Object.freeze([
  "SS",
  "SB",
  "MAR",
  "MOR",
  "ORDER",
  "SO",
  "PSS",
  "PSB",
  "PMAR",
  "PMOR",
  "BAR",
]);
const INHOUSE_ORDER_PREFIX_LABEL = INHOUSE_ORDER_PREFIXES.join(", ");
const APP_SETTING_KEYS = Object.freeze({
  mailDisabled: "mail_disabled_override",
  mailSenderName: "mail_sender_name_override",
  mailTo: "mail_to_override",
  mailArtworkTo: "mail_artwork_to_override",
  mailAdminActionTo: "mail_admin_action_to_override",
  mailRolloverTestTo: "mail_rollover_test_to_override",
  mailDroppedOfficeSsTo: "mail_dropped_office_ss_to_override",
  mailDroppedOfficeSbTo: "mail_dropped_office_sb_to_override",
  mailDroppedOfficeMorMarTo: "mail_dropped_office_mor_mar_to_override",
  mailDroppedOfficeOrderTo: "mail_dropped_office_order_to_override",
  mailDroppedOfficeFallbackTo: "mail_dropped_office_fallback_to_override",
});
const MAIL_SETTING_KEY_LIST = Object.freeze(Object.values(APP_SETTING_KEYS));
const DEFAULT_DATABASE_FILENAME = "route-ledger.sqlite";
const TEMP_DATA_DIRECTORY_NAME = "logistics-center-data";
const DATABASE_PATH_ENV_KEYS = ["LOGISTICS_DB_PATH", "LOCAL_DB_PATH", "SQLITE_DB_PATH", "DATABASE_PATH"];
const DATA_DIR_ENV_KEYS = [
  "LOGISTICS_DATA_DIR",
  "LOCAL_DATA_DIR",
  "SQLITE_DATA_DIR",
  "PERSISTENT_DATA_DIR",
  "RENDER_DISK_PATH",
  "RAILWAY_VOLUME_MOUNT_PATH",
];
const REQUIRE_PERSISTENT_STORAGE_ENV_KEYS = [
  "LOGISTICS_REQUIRE_PERSISTENT_STORAGE",
  "REQUIRE_PERSISTENT_STORAGE",
];
const ALLOW_TEMPORARY_STORAGE_ENV_KEYS = [
  "LOGISTICS_ALLOW_TEMPORARY_STORAGE",
  "ALLOW_TEMPORARY_STORAGE",
];
const TURSO_DATABASE_URL_ENV_KEYS = ["TURSO_DATABASE_URL", "LIBSQL_DATABASE_URL"];
const TURSO_AUTH_TOKEN_ENV_KEYS = ["TURSO_AUTH_TOKEN", "LIBSQL_AUTH_TOKEN"];
const TURSO_STATE_TABLE_NAME = "logistics_center_state";
const TURSO_STATE_ID = "route-ledger";

function createLocalDatabase(rootDir) {
  if (sqliteLoadError || !DatabaseSync) {
    const reason = sqliteLoadError
      ? `SQLite is unavailable in this Node runtime: ${normalizeErrorMessage(sqliteLoadError)}`
      : "SQLite is unavailable in this Node runtime.";
    return createUnavailableDatabase(reason);
  }

  try {
    const tursoConfig = getTursoConfig();
    if (tursoConfig) {
      if (!createLibsqlClient) {
        return createUnavailableDatabase(`Turso Cloud is configured, but the libSQL client could not load: ${normalizeErrorMessage(libsqlLoadError)}`);
      }

      return new TursoBackedDatabase(rootDir, tursoConfig);
    }

    return new LocalDatabase(rootDir);
  } catch (error) {
    return createUnavailableDatabase(`Failed to start the local database: ${normalizeErrorMessage(error)}`);
  }
}

function getTursoConfig() {
  const databaseUrl = getFirstConfiguredEnvironmentValue(TURSO_DATABASE_URL_ENV_KEYS);
  if (!databaseUrl) {
    return null;
  }

  const authToken = getFirstConfiguredEnvironmentValue(TURSO_AUTH_TOKEN_ENV_KEYS);
  return {
    url: databaseUrl.value,
    urlKey: databaseUrl.key,
    authToken: authToken?.value || "",
    authTokenKey: authToken?.key || "",
  };
}

function createUnavailableDatabase(reason) {
  const message = String(reason || "").trim() || "Local database is unavailable.";
  const hosting = detectHostingEnvironment();
  const persistentRequirement = getFirstEnabledEnvironmentFlag(REQUIRE_PERSISTENT_STORAGE_ENV_KEYS);
  return {
    getStatus() {
      return {
        configured: false,
        reason: message,
        storage: "local-sqlite",
        storagePath: "",
        storageDir: "",
        storagePersistent: false,
        storageTemporary: false,
        storageLabel: "",
        storageConfiguredBy: "",
        storageHost: hosting.id || "local",
        persistentStorageRequired: Boolean(persistentRequirement),
        warning: "",
      };
    },
    async call() {
      throw createHttpError(503, message);
    },
    async listOrdersForMailExport() {
      throw createHttpError(503, message);
    },
    async listPendingOrderDeleteNotifications() {
      throw createHttpError(503, message);
    },
    async markOrderDeleteNotificationsSent() {
      throw createHttpError(503, message);
    },
    async markOrderDeleteNotificationFailure() {
      throw createHttpError(503, message);
    },
    async clearAllOrderPriorities() {
      throw createHttpError(503, message);
    },
    async clearOrderRollovers() {
      throw createHttpError(503, message);
    },
    async getUserByToken() {
      throw createHttpError(503, message);
    },
    exportAllData() {
      throw createHttpError(503, message);
    },
    getTableCounts() {
      throw createHttpError(503, message);
    },
    replaceAllData() {
      throw createHttpError(503, message);
    },
    getMailSettingsOverrides() {
      return {
        disabled: null,
        senderName: "",
        to: "",
        artworkTo: "",
        adminActionTo: "",
        rolloverTestTo: "",
        droppedOfficeSsTo: "",
        droppedOfficeSbTo: "",
        droppedOfficeMorMarTo: "",
        droppedOfficeOrderTo: "",
        droppedOfficeFallbackTo: "",
      };
    },
    async close() {},
  };
}

function resolveDatabaseLocation(rootDir, options = {}) {
  const hosting = detectHostingEnvironment();
  const allowTemporaryStorage = Boolean(options.allowTemporaryStorage);
  const preferTemporaryStorage = Boolean(options.preferTemporaryStorage);
  const explicitPersistentRequirement = allowTemporaryStorage
    ? null
    : getFirstEnabledEnvironmentFlag(REQUIRE_PERSISTENT_STORAGE_ENV_KEYS);
  const temporaryStorageAllowance = getFirstEnabledEnvironmentFlag(ALLOW_TEMPORARY_STORAGE_ENV_KEYS);
  const persistentRequirement = explicitPersistentRequirement
    || (hosting.id === "vercel" && !temporaryStorageAllowance && !allowTemporaryStorage
      ? { key: "VERCEL", value: "persistent-storage-required-by-default" }
      : null);
  const configuredCandidate = getConfiguredDatabaseCandidate(rootDir);
  const candidates = configuredCandidate
    ? [configuredCandidate]
    : [];

  if (!preferTemporaryStorage) {
    candidates.push({
      label: "project data folder",
      configuredBy: "project-data",
      databasePath: path.join(rootDir, "data", DEFAULT_DATABASE_FILENAME),
      temporary: false,
    });
  }
  if (!persistentRequirement) {
    candidates.push({
      label: "temporary runtime folder",
      configuredBy: "temporary-runtime",
      databasePath: path.join(os.tmpdir(), TEMP_DATA_DIRECTORY_NAME, DEFAULT_DATABASE_FILENAME),
      temporary: true,
      warning: buildTemporaryStorageWarning(hosting),
    });
  }

  const failures = [];

  for (const candidate of candidates) {
    try {
      ensureWritableDatabaseLocation(candidate.databasePath);
      if (persistentRequirement && isTemporaryDatabasePath(candidate.databasePath)) {
        throw new Error("Configured database path is inside temporary runtime storage.");
      }
      return {
        label: candidate.label,
        configuredBy: String(candidate.configuredBy || "").trim(),
        dataDir: path.dirname(candidate.databasePath),
        databasePath: candidate.databasePath,
        temporary: Boolean(candidate.temporary || isTemporaryDatabasePath(candidate.databasePath)),
        hosting,
        persistentStorageRequired: Boolean(persistentRequirement),
        warning: String(candidate.warning || "").trim(),
      };
    } catch (error) {
      const detail = `${candidate.label} (${candidate.databasePath}): ${normalizeErrorMessage(error)}`;
      if (candidate.required) {
        throw new Error(detail);
      }
      failures.push(detail);
    }
  }

  if (persistentRequirement) {
    const requirementMessage = persistentRequirement.key === "VERCEL"
      ? "Persistent storage is required on Vercel because its writable filesystem is temporary. Move this app to a host with a persistent disk/volume, or set LOGISTICS_ALLOW_TEMPORARY_STORAGE=true only for a disposable demo."
      : `Persistent storage is required by ${persistentRequirement.key}=true.`;
    throw new Error(
      [
        requirementMessage,
        "No writable persistent database location was found.",
        failures.join(" "),
      ].filter(Boolean).join(" "),
    );
  }

  throw new Error(`No writable database location was found. ${failures.join(" ")}`.trim());
}

function getConfiguredDatabaseCandidate(rootDir) {
  const configuredFilePath = getFirstConfiguredEnvironmentValue(DATABASE_PATH_ENV_KEYS);
  if (configuredFilePath) {
    return {
      label: `configured database path (${configuredFilePath.key})`,
      configuredBy: configuredFilePath.key,
      databasePath: resolveFromRoot(rootDir, configuredFilePath.value),
      required: true,
      temporary: false,
    };
  }

  const configuredDataDir = getFirstConfiguredEnvironmentValue(DATA_DIR_ENV_KEYS);
  if (!configuredDataDir) {
    return null;
  }

  return {
    label: `configured data directory (${configuredDataDir.key})`,
    configuredBy: configuredDataDir.key,
    databasePath: path.join(resolveFromRoot(rootDir, configuredDataDir.value), DEFAULT_DATABASE_FILENAME),
    required: true,
    temporary: false,
  };
}

function resolveFromRoot(rootDir, value) {
  const target = String(value || "").trim();
  if (!target) {
    return "";
  }

  return path.isAbsolute(target)
    ? path.normalize(target)
    : path.resolve(rootDir, target);
}

function isTemporaryDatabasePath(databasePath) {
  const tempDir = path.resolve(os.tmpdir());
  const targetPath = path.resolve(databasePath);
  const relativePath = path.relative(tempDir, targetPath);
  return Boolean(relativePath) && !relativePath.startsWith("..") && !path.isAbsolute(relativePath);
}

function ensureWritableDatabaseLocation(databasePath) {
  const directory = path.dirname(databasePath);
  ensureWritableDirectory(directory);

  if (fs.existsSync(databasePath)) {
    fs.accessSync(databasePath, fs.constants.R_OK | fs.constants.W_OK);
  }
}

function ensureWritableDirectory(directoryPath) {
  fs.mkdirSync(directoryPath, { recursive: true });
  const probePath = path.join(
    directoryPath,
    `.write-test-${process.pid}-${Date.now()}-${Math.random().toString(16).slice(2)}.tmp`,
  );
  fs.writeFileSync(probePath, "ok", "utf8");
  fs.unlinkSync(probePath);
}

function getFirstConfiguredEnvironmentValue(keys) {
  for (const key of keys) {
    const value = process.env[key];
    const text = String(value || "").trim();
    if (text) {
      return { key, value: text };
    }
  }

  return null;
}

function getFirstEnabledEnvironmentFlag(keys) {
  for (const key of keys) {
    if (isTruthyEnvironmentValue(process.env[key])) {
      return {
        key,
        value: String(process.env[key]).trim(),
      };
    }
  }

  return null;
}

function isTruthyEnvironmentValue(value) {
  const text = String(value || "").trim().toLowerCase();
  if (!text) {
    return false;
  }

  return !["0", "false", "no", "off"].includes(text);
}

function detectHostingEnvironment() {
  if (String(process.env.VERCEL || "").trim()) {
    return { id: "vercel", label: "Vercel" };
  }
  if (String(process.env.RENDER || "").trim()) {
    return { id: "render", label: "Render" };
  }
  if (
    String(process.env.RAILWAY_ENVIRONMENT_ID || "").trim()
    || String(process.env.RAILWAY_PROJECT_ID || "").trim()
    || String(process.env.RAILWAY_SERVICE_ID || "").trim()
    || String(process.env.RAILWAY_VOLUME_MOUNT_PATH || "").trim()
  ) {
    return { id: "railway", label: "Railway" };
  }
  return { id: "", label: "" };
}

function buildTemporaryStorageWarning(hosting) {
  const baseWarning = "This host is using temporary runtime storage because the deployed app folder is read-only. Data can reset when the server restarts, scales, or redeploys.";

  if (hosting.id === "vercel") {
    return `${baseWarning} Vercel only provides a read-only deployment filesystem plus temporary /tmp scratch space, so SQLite writes here are not durable. Move this app to a host with a persistent disk or volume for long-term storage.`;
  }

  if (hosting.id === "render") {
    return `${baseWarning} Attach a persistent disk and mount it over this app's data folder, or point LOGISTICS_DATA_DIR at the disk mount path.`;
  }

  if (hosting.id === "railway") {
    return `${baseWarning} Attach a Railway volume and mount it at /app/data, or point LOGISTICS_DATA_DIR at the mounted volume path.`;
  }

  return `${baseWarning} For long-term SQLite storage, use a single-instance host with a persistent disk or volume and point LOGISTICS_DATA_DIR or LOGISTICS_DB_PATH at it.`;
}

function seedTemporaryDatabaseFromProject(rootDir, databasePath) {
  const projectDatabasePath = path.join(rootDir, "data", DEFAULT_DATABASE_FILENAME);
  const targetExists = fs.existsSync(databasePath) && Number(fs.statSync(databasePath).size || 0) > 0;
  if (targetExists || !fs.existsSync(projectDatabasePath)) {
    return false;
  }

  fs.accessSync(projectDatabasePath, fs.constants.R_OK);
  fs.mkdirSync(path.dirname(databasePath), { recursive: true });

  const suffixes = ["", "-wal", "-shm"];
  suffixes.forEach((suffix) => {
    const sourcePath = `${projectDatabasePath}${suffix}`;
    const targetPath = `${databasePath}${suffix}`;
    if (!fs.existsSync(sourcePath)) {
      return;
    }

    fs.copyFileSync(sourcePath, targetPath);
  });

  return true;
}

class TursoBackedDatabase {
  constructor(rootDir, config) {
    this.rootDir = rootDir;
    this.config = config;
    this.client = createLibsqlClient({
      url: config.url,
      authToken: config.authToken || undefined,
    });
    this.local = new LocalDatabase(rootDir, {
      allowTemporaryStorage: true,
      preferTemporaryStorage: true,
    });
    const localStatus = this.local.getStatus();
    this.status = {
      ...localStatus,
      storage: "turso-cloud",
      storagePath: config.url,
      storageDir: "Turso Cloud",
      storagePersistent: true,
      storageTemporary: false,
      storageLabel: "Turso Cloud",
      storageConfiguredBy: config.urlKey,
      storageHost: "turso",
      persistentStorageRequired: true,
      warning: "",
      localCachePath: localStatus.storagePath,
      localCacheTemporary: localStatus.storageTemporary,
      seededFromBundledSnapshot: localStatus.seededFromBundledSnapshot,
      tursoConnected: false,
      tursoLastSyncedAt: "",
    };
    this.persistQueue = Promise.resolve();
    this.readyPromise = this.initialize().catch((error) => {
      this.status.configured = false;
      this.status.reason = `Failed to connect to Turso Cloud: ${normalizeErrorMessage(error)}`;
      throw error;
    });
  }

  async initialize() {
    this.assertTursoCredentials();
    await this.ensureTursoStateTable();
    const restored = await this.restoreFromTurso();
    if (!restored) {
      await this.persistCurrentState("initial-seed");
    }

    this.status.configured = true;
    this.status.reason = "";
    this.status.tursoConnected = true;
  }

  assertTursoCredentials() {
    const url = String(this.config.url || "").trim();
    const isRemote = /^(libsql|https|http):\/\//i.test(url);
    if (isRemote && !String(this.config.authToken || "").trim()) {
      throw new Error("TURSO_AUTH_TOKEN is required when TURSO_DATABASE_URL points at Turso Cloud.");
    }
  }

  async ensureReady() {
    await this.readyPromise;
  }

  getStatus() {
    return { ...this.status };
  }

  async call(functionName, parameters = {}) {
    await this.ensureReady();
    const result = await this.local.call(functionName, parameters);
    if (shouldPersistTursoAfterCall(functionName)) {
      await this.persistCurrentState(functionName);
    }
    return result;
  }

  async listOrdersForMailExport(scheduledFor = "") {
    await this.ensureReady();
    return this.local.listOrdersForMailExport(scheduledFor);
  }

  async listPendingOrderDeleteNotifications(limit = 100) {
    await this.ensureReady();
    return this.local.listPendingOrderDeleteNotifications(limit);
  }

  async markOrderDeleteNotificationsSent(logIds) {
    await this.ensureReady();
    const result = await this.local.markOrderDeleteNotificationsSent(logIds);
    await this.persistCurrentState("mark-order-delete-notifications-sent");
    return result;
  }

  async markOrderDeleteNotificationFailure(logIds, errorMessage) {
    await this.ensureReady();
    const result = await this.local.markOrderDeleteNotificationFailure(logIds, errorMessage);
    await this.persistCurrentState("mark-order-delete-notification-failure");
    return result;
  }

  async clearAllOrderPriorities(token) {
    await this.ensureReady();
    const result = await this.local.clearAllOrderPriorities(token);
    await this.persistCurrentState("clear-all-order-priorities");
    return result;
  }

  async clearOrderRollovers(token) {
    await this.ensureReady();
    const result = await this.local.clearOrderRollovers(token);
    await this.persistCurrentState("clear-order-rollovers");
    return result;
  }

  async getUserByToken(token) {
    await this.ensureReady();
    const user = await this.local.getUserByToken(token);
    if (user) {
      await this.persistCurrentState("session-refresh");
    }
    return user;
  }

  exportAllData() {
    return this.local.exportAllData();
  }

  getTableCounts() {
    return this.local.getTableCounts();
  }

  async replaceAllData(data, options = {}) {
    await this.ensureReady();
    const result = this.local.replaceAllData(data, options);
    await this.persistCurrentState(options?.source || "replace-all-data");
    return result;
  }

  getMailSettingsOverrides() {
    return this.local.getMailSettingsOverrides();
  }

  async close() {
    await this.persistQueue.catch(() => {});
    if (typeof this.client?.close === "function") {
      this.client.close();
    }
    await this.local.close();
  }

  async ensureTursoStateTable() {
    await this.client.execute(`
      create table if not exists ${TURSO_STATE_TABLE_NAME} (
        id text primary key,
        data text not null,
        counts text not null default '{}',
        source text not null default '',
        updated_at text not null
      )
    `);
  }

  async restoreFromTurso() {
    const result = await this.client.execute({
      sql: `select data, updated_at from ${TURSO_STATE_TABLE_NAME} where id = ? limit 1`,
      args: [TURSO_STATE_ID],
    });
    const row = Array.isArray(result.rows) ? result.rows[0] : null;
    const rawData = row?.data ? String(row.data) : "";
    if (!rawData) {
      return false;
    }

    const parsed = JSON.parse(rawData);
    this.local.replaceAllData(parsed, {
      source: `turso-cloud ${row?.updated_at || ""}`.trim(),
    });
    this.status.tursoLastSyncedAt = String(row?.updated_at || "").trim();
    return true;
  }

  async persistCurrentState(source = "app") {
    this.persistQueue = this.persistQueue.then(async () => {
      const data = this.local.exportAllData();
      const counts = this.local.getTableCounts();
      const updatedAt = nowIso();
      await this.client.execute({
        sql: `
          insert into ${TURSO_STATE_TABLE_NAME} (
            id,
            data,
            counts,
            source,
            updated_at
          )
          values (?, ?, ?, ?, ?)
          on conflict(id) do update set
            data = excluded.data,
            counts = excluded.counts,
            source = excluded.source,
            updated_at = excluded.updated_at
        `,
        args: [
          TURSO_STATE_ID,
          JSON.stringify(data),
          JSON.stringify(counts),
          String(source || "").trim() || "app",
          updatedAt,
        ],
      });
      this.status.tursoLastSyncedAt = updatedAt;
      this.status.tursoConnected = true;
    });

    return this.persistQueue;
  }
}

function shouldPersistTursoAfterCall(functionName) {
  return String(functionName || "").trim() !== "get_login_state";
}

class LocalDatabase {
  constructor(rootDir, options = {}) {
    this.rootDir = rootDir;
    const storageLocation = resolveDatabaseLocation(rootDir, options);
    this.dataDir = storageLocation.dataDir;
    this.databasePath = storageLocation.databasePath;
    this.statementCache = new Map();
    this.status = {
      configured: false,
      reason: "",
      storage: "local-sqlite",
      storagePath: this.databasePath,
      storageDir: this.dataDir,
      storagePersistent: !storageLocation.temporary,
      storageTemporary: storageLocation.temporary,
      storageLabel: storageLocation.label,
      storageConfiguredBy: storageLocation.configuredBy,
      storageHost: storageLocation.hosting.id || "local",
      persistentStorageRequired: storageLocation.persistentStorageRequired,
      warning: storageLocation.warning,
      seededFromBundledSnapshot: false,
    };

    fs.mkdirSync(this.dataDir, { recursive: true });
    if (storageLocation.temporary) {
      this.status.seededFromBundledSnapshot = seedTemporaryDatabaseFromProject(rootDir, this.databasePath);
    }
    this.db = new DatabaseSync(this.databasePath);
    this.db.exec("PRAGMA foreign_keys = ON;");
    this.db.exec("PRAGMA journal_mode = WAL;");
    this.db.exec("PRAGMA busy_timeout = 5000;");
    this.db.exec("PRAGMA synchronous = NORMAL;");
    this.db.exec("PRAGMA temp_store = MEMORY;");
    this.initializeSchema();
    this.ensureSchemaMigrations();
    this.ensureOfficeLocationForExistingUsers();

    this.status.configured = true;
  }

  getStatus() {
    return { ...this.status };
  }

  async call(functionName, parameters = {}) {
    try {
      switch (functionName) {
        case "get_login_state":
          return this.getLoginState();
        case "bootstrap_admin":
          return this.bootstrapAdmin(parameters);
        case "login_user":
          return this.loginUser(parameters);
        case "logout_user":
          return this.logoutUser(parameters);
        case "run_daily_rollover":
          return this.runDailyRollover();
        case "get_app_snapshot":
          return this.getAppSnapshot(parameters);
        case "record_driver_position":
          return this.recordDriverPosition(parameters);
        case "update_mail_settings":
          return this.updateMailSettings(parameters);
        case "create_user_account":
          return this.createUserAccount(parameters);
        case "update_user_account":
          return this.updateUserAccount(parameters);
        case "toggle_user_active":
          return this.toggleUserActive(parameters);
        case "delete_user_account":
          return this.deleteUserAccount(parameters);
        case "create_supplier":
          return this.createSupplier(parameters);
        case "update_supplier":
          return this.updateSupplier(parameters);
        case "delete_supplier":
          return this.deleteSupplier(parameters);
        case "create_location":
          return this.createLocation(parameters);
        case "update_location":
          return this.updateLocation(parameters);
        case "delete_location":
          return this.deleteLocation(parameters);
        case "create_stock_item":
          return this.createStockItem(parameters);
        case "update_stock_item":
          return this.updateStockItem(parameters);
        case "delete_stock_item":
          return this.deleteStockItem(parameters);
        case "record_stock_movement":
          return this.recordStockMovement(parameters);
        case "update_stock_movement":
          return this.updateStockMovement(parameters);
        case "create_artwork_request":
          return this.createArtworkRequest(parameters);
        case "create_order":
          return this.createOrder(parameters);
        case "update_order":
          return this.updateOrder(parameters);
        case "assign_order":
          return this.assignOrder(parameters);
        case "set_order_priority":
          return this.setOrderPriority(parameters);
        case "clear_all_order_priorities":
          return this.clearAllOrderPriorities(parameters?.p_token);
        case "clear_order_rollovers":
          return this.clearOrderRollovers(parameters?.p_token);
        case "set_order_flag":
          return this.setOrderFlag(parameters);
        case "pick_up_order":
          return this.pickUpOrder(parameters);
        case "complete_order":
          return this.completeOrder(parameters);
        case "delete_order":
          return this.deleteOrder(parameters);
        default:
          throw createHttpError(404, "Unknown RPC function.");
      }
    } catch (error) {
      if (error?.statusCode) {
        throw error;
      }
      throw createHttpError(400, normalizeErrorMessage(error));
    }
  }

  async listOrdersForMailExport(scheduledFor = "") {
    const targetDate = normalizeOptionalDate(scheduledFor);
    return this.selectOrderRows(targetDate ? "where o.scheduled_for = ?" : "", targetDate ? [targetDate] : []);
  }

  async listPendingOrderDeleteNotifications(limit = 100) {
    const safeLimit = Math.max(1, Math.min(Number(limit) || 100, 250));
    const rows = this.all(
      `
        select
          id,
          order_id as orderId,
          order_number as orderNumber,
          reference,
          quote_number as quoteNumber,
          sales_order_number as salesOrderNumber,
          invoice_number as invoiceNumber,
          po_number as poNumber,
          entry_type as entryType,
          priority,
          delivery_address as deliveryAddress,
          delivery_location_id as deliveryLocationId,
          delivery_location_name as deliveryLocationName,
          delivery_location_address as deliveryLocationAddress,
          branding,
          stock_description as stockDescription,
          notes,
          move_to_factory as moveToFactory,
          factory_destination_location_id as factoryDestinationLocationId,
          factory_destination_name as factoryDestinationName,
          factory_destination_address as factoryDestinationAddress,
          location_id as locationId,
          location_name as locationName,
          location_address as locationAddress,
          driver_user_id as driverUserId,
          driver_name as driverName,
          created_by_user_id as createdByUserId,
          created_by_name as createdByName,
          created_at as createdAt,
          scheduled_for as scheduledFor,
          original_scheduled_for as originalScheduledFor,
          carry_over_count as carryOverCount,
          status,
          deleted_by_user_id as deletedByUserId,
          deleted_by_name as deletedByName,
          deleted_by_role as deletedByRole,
          deleted_at as deletedAt,
          notification_attempts as notificationAttempts,
          last_notification_error as lastNotificationError,
          notification_sent_at as notificationSentAt
        from order_delete_log
        where notification_sent_at is null
        order by deleted_at asc, id asc
        limit ?
      `,
      [safeLimit],
    );
    return rows.map((row) => this.buildDeleteLogRow(row));
  }

  async markOrderDeleteNotificationsSent(logIds) {
    const ids = normalizeIdList(logIds);
    if (!ids.length) {
      return 0;
    }

    const now = nowIso();
    const sql = `
      update order_delete_log
      set notification_sent_at = ?,
          notification_attempts = notification_attempts + 1,
          last_notification_error = ''
      where id in (${buildPlaceholders(ids.length)})
    `;
    const result = this.run(sql, [now, ...ids]);
    return Number(result?.changes || 0);
  }

  async markOrderDeleteNotificationFailure(logIds, errorMessage) {
    const ids = normalizeIdList(logIds);
    if (!ids.length) {
      return 0;
    }

    const message = normalizeOptionalText(errorMessage) || "Delete-log email failed.";
    const sql = `
      update order_delete_log
      set notification_attempts = notification_attempts + 1,
          last_notification_error = ?
      where id in (${buildPlaceholders(ids.length)})
    `;
    const result = this.run(sql, [message, ...ids]);
    return Number(result?.changes || 0);
  }

  async clearAllOrderPriorities(token) {
    const actor = this.requireUser(token, ["admin"]);
    const result = this.run(
      `
        update orders
        set priority = 'medium',
            updated_at = ?
        where status = 'active'
          and priority = 'high'
      `,
      [nowIso()],
    );
    return {
      ok: true,
      updatedOrders: Number(result?.changes || 0),
      clearedBy: actor.id,
    };
  }

  async clearOrderRollovers(token) {
    const actor = this.requireUser(token, ["admin"]);
    const result = this.run(
      `
        update orders
        set carry_over_count = 0,
            original_scheduled_for = scheduled_for,
            updated_at = ?
        where status = 'active'
          and carry_over_count > 0
      `,
      [nowIso()],
    );
    return {
      ok: true,
      updatedOrders: Number(result?.changes || 0),
      clearedBy: actor.id,
    };
  }

  updateMailSettings(parameters = {}) {
    const actor = this.requireUser(parameters?.p_token, ["admin", "maintenance"]);
    const disabled = Boolean(parameters?.p_disabled);
    const senderName = normalizeOptionalText(parameters?.p_sender_name);
    const to = normalizeEmailDestinationText(parameters?.p_to, "General inbox");
    const artworkTo = normalizeEmailDestinationText(parameters?.p_artwork_to, "Artwork inbox");
    const adminActionTo = normalizeEmailDestinationText(parameters?.p_admin_action_to, "Admin action inbox");
    const rolloverTestTo = normalizeEmailDestinationText(parameters?.p_rollover_test_to, "Carry-over test inbox");
    const droppedOfficeSsTo = normalizeEmailDestinationText(parameters?.p_dropped_office_ss_to, "Dropped-at-office SS inbox");
    const droppedOfficeSbTo = normalizeEmailDestinationText(parameters?.p_dropped_office_sb_to, "Dropped-at-office SB inbox");
    const droppedOfficeMorMarTo = normalizeEmailDestinationText(parameters?.p_dropped_office_mor_mar_to, "Dropped-at-office MOR/MAR inbox");
    const droppedOfficeOrderTo = normalizeEmailDestinationText(parameters?.p_dropped_office_order_to, "Dropped-at-office Order inbox");
    const droppedOfficeFallbackTo = normalizeEmailDestinationText(parameters?.p_dropped_office_fallback_to, "Dropped-at-office fallback inbox");
    const updatedAt = nowIso();

    this.withTransaction(() => {
      this.writeAppSetting(APP_SETTING_KEYS.mailDisabled, disabled ? "true" : "false", actor.id, updatedAt, { allowEmpty: false });
      this.writeAppSetting(APP_SETTING_KEYS.mailSenderName, senderName, actor.id, updatedAt);
      this.writeAppSetting(APP_SETTING_KEYS.mailTo, to, actor.id, updatedAt);
      this.writeAppSetting(APP_SETTING_KEYS.mailArtworkTo, artworkTo, actor.id, updatedAt);
      this.writeAppSetting(APP_SETTING_KEYS.mailAdminActionTo, adminActionTo, actor.id, updatedAt);
      this.writeAppSetting(APP_SETTING_KEYS.mailRolloverTestTo, rolloverTestTo, actor.id, updatedAt);
      this.writeAppSetting(APP_SETTING_KEYS.mailDroppedOfficeSsTo, droppedOfficeSsTo, actor.id, updatedAt);
      this.writeAppSetting(APP_SETTING_KEYS.mailDroppedOfficeSbTo, droppedOfficeSbTo, actor.id, updatedAt);
      this.writeAppSetting(APP_SETTING_KEYS.mailDroppedOfficeMorMarTo, droppedOfficeMorMarTo, actor.id, updatedAt);
      this.writeAppSetting(APP_SETTING_KEYS.mailDroppedOfficeOrderTo, droppedOfficeOrderTo, actor.id, updatedAt);
      this.writeAppSetting(APP_SETTING_KEYS.mailDroppedOfficeFallbackTo, droppedOfficeFallbackTo, actor.id, updatedAt);
    });

    return {
      ok: true,
      updatedBy: actor.id,
      updatedAt,
    };
  }

  async getUserByToken(token) {
    const cleanToken = normalizeOptionalText(token);
    if (!cleanToken) {
      return null;
    }

    const now = Date.now();
    const nowText = new Date(now).toISOString();
    const row = this.get(
      `
        select
          u.id,
          u.name,
          u.role,
          u.active
        from app_sessions s
        join app_users u on u.id = s.user_id
        where s.token = ?
          and s.expires_at > ?
          and u.active = 1
        order by s.created_at desc
        limit 1
      `,
      [cleanToken, nowText],
    );

    if (!row) {
      return null;
    }

    this.refreshSessionExpiry(cleanToken, now);

    return {
      id: row.id,
      name: row.name,
      role: row.role,
      active: Boolean(row.active),
    };
  }

  getMailSettingsOverrides() {
    const rows = this.all(
      `
        select key, value
        from app_settings
        where key in (${buildPlaceholders(MAIL_SETTING_KEY_LIST.length)})
      `,
      MAIL_SETTING_KEY_LIST,
    );
    const rowMap = Object.fromEntries(rows.map((row) => [row.key, row.value]));

    return {
      disabled: Object.prototype.hasOwnProperty.call(rowMap, APP_SETTING_KEYS.mailDisabled)
        ? normalizeStoredBoolean(rowMap[APP_SETTING_KEYS.mailDisabled], false)
        : null,
      senderName: rowMap[APP_SETTING_KEYS.mailSenderName] || "",
      to: rowMap[APP_SETTING_KEYS.mailTo] || "",
      artworkTo: rowMap[APP_SETTING_KEYS.mailArtworkTo] || "",
      adminActionTo: rowMap[APP_SETTING_KEYS.mailAdminActionTo] || "",
      rolloverTestTo: rowMap[APP_SETTING_KEYS.mailRolloverTestTo] || "",
      droppedOfficeSsTo: rowMap[APP_SETTING_KEYS.mailDroppedOfficeSsTo] || "",
      droppedOfficeSbTo: rowMap[APP_SETTING_KEYS.mailDroppedOfficeSbTo] || "",
      droppedOfficeMorMarTo: rowMap[APP_SETTING_KEYS.mailDroppedOfficeMorMarTo] || "",
      droppedOfficeOrderTo: rowMap[APP_SETTING_KEYS.mailDroppedOfficeOrderTo] || "",
      droppedOfficeFallbackTo: rowMap[APP_SETTING_KEYS.mailDroppedOfficeFallbackTo] || "",
    };
  }

  writeAppSetting(key, value, updatedByUserId, updatedAt, options = {}) {
    const allowEmpty = options.allowEmpty !== false;
    const normalizedKey = String(key || "").trim();
    const normalizedValue = value === null || value === undefined ? "" : String(value);

    if (!normalizedKey) {
      throw new Error("App setting key is required.");
    }

    if (!allowEmpty && normalizedValue === "") {
      throw new Error(`App setting ${normalizedKey} cannot be empty.`);
    }

    if (allowEmpty && normalizedValue === "") {
      this.run("delete from app_settings where key = ?", [normalizedKey]);
      return;
    }

    this.run(
      `
        insert into app_settings (
          key,
          value,
          updated_by_user_id,
          updated_at
        )
        values (?, ?, ?, ?)
        on conflict(key) do update set
          value = excluded.value,
          updated_by_user_id = excluded.updated_by_user_id,
          updated_at = excluded.updated_at
      `,
      [normalizedKey, normalizedValue, updatedByUserId || null, updatedAt || nowIso()],
    );
  }

  async close() {
    if (this.db) {
      this.db.close();
    }
  }

  getTableCounts() {
    return {
      app_users: Number(this.get("select count(*) as count from app_users")?.count || 0),
      suppliers: Number(this.get("select count(*) as count from suppliers")?.count || 0),
      locations: Number(this.get("select count(*) as count from locations")?.count || 0),
      orders: Number(this.get("select count(*) as count from orders")?.count || 0),
      order_delete_log: Number(this.get("select count(*) as count from order_delete_log")?.count || 0),
      stock_items: Number(this.get("select count(*) as count from stock_items")?.count || 0),
      stock_movements: Number(this.get("select count(*) as count from stock_movements")?.count || 0),
      artwork_requests: Number(this.get("select count(*) as count from artwork_requests")?.count || 0),
      app_sessions: Number(this.get("select count(*) as count from app_sessions")?.count || 0),
      app_settings: Number(this.get("select count(*) as count from app_settings")?.count || 0),
    };
  }

  exportAllData() {
    return {
      app_users: this.all("select * from app_users order by created_at asc, id asc"),
      suppliers: this.all("select * from suppliers order by created_at asc, id asc"),
      locations: this.all("select * from locations order by created_at asc, id asc"),
      orders: this.all("select * from orders order by created_at asc, order_number asc, id asc"),
      order_delete_log: this.all("select * from order_delete_log order by deleted_at asc, id asc"),
      stock_items: this.all("select * from stock_items order by created_at asc, id asc"),
      stock_movements: this.all("select * from stock_movements order by created_at asc, id asc"),
      artwork_requests: this.all("select * from artwork_requests order by sent_at asc, id asc"),
      app_sessions: this.all("select * from app_sessions order by created_at asc, token asc"),
      app_settings: this.all("select * from app_settings order by key asc"),
    };
  }

  replaceAllData(rawData = {}, options = {}) {
    const data = normalizeImportData(rawData);
    return this.withTransaction(() => {
      this.run("delete from app_sessions");
      this.run("delete from app_settings");
      this.run("delete from artwork_requests");
      this.run("delete from stock_movements");
      this.run("delete from stock_items");
      this.run("delete from order_delete_log");
      this.run("delete from orders");
      this.run("delete from locations");
      this.run("delete from suppliers");
      this.run("delete from app_users");

      data.app_users.forEach((row) => {
        this.run(
          `
            insert into app_users (
              id,
              name,
              role,
              password_hash,
              active,
              phone,
              vehicle,
              last_known_lat,
              last_known_lng,
              last_known_recorded_at,
              created_at,
              updated_at
            )
            values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
          `,
          [
            requireImportedText(row.id, "Imported user id is required."),
            importedText(row.name),
            importedText(row.role),
            importedText(row.password_hash),
            importedBoolean(row.active),
            importedNullableText(row.phone),
            importedNullableText(row.vehicle),
            importedNullableNumber(row.last_known_lat),
            importedNullableNumber(row.last_known_lng),
            importedNullableTimestamp(row.last_known_recorded_at),
            importedTimestamp(row.created_at),
            importedTimestamp(row.updated_at),
          ],
        );
      });

      data.suppliers.forEach((row) => {
        this.run(
          `
            insert into suppliers (
              id,
              name,
              contact_person,
              contact_number,
              factory,
              created_by,
              created_at,
              updated_at
            )
            values (?, ?, ?, ?, ?, ?, ?, ?)
          `,
          [
            requireImportedText(row.id, "Imported supplier id is required."),
            importedText(row.name),
            importedText(row.contact_person),
            importedText(row.contact_number),
            importedBoolean(row.factory),
            requireImportedText(row.created_by, "Imported supplier created_by is required."),
            importedTimestamp(row.created_at),
            importedTimestamp(row.updated_at),
          ],
        );
      });

      data.locations.forEach((row) => {
        this.run(
          `
            insert into locations (
              id,
              supplier_id,
              location_type,
              name,
              address,
              lat,
              lng,
              contact_person,
              contact_number,
              notes,
              created_by,
              created_at,
              updated_at
            )
            values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
          `,
          [
            requireImportedText(row.id, "Imported location id is required."),
            importedNullableText(row.supplier_id),
            importedText(row.location_type) || "supplier",
            importedText(row.name),
            importedText(row.address),
            importedNullableNumber(row.lat),
            importedNullableNumber(row.lng),
            importedText(row.contact_person),
            importedText(row.contact_number),
            importedText(row.notes),
            requireImportedText(row.created_by, "Imported location created_by is required."),
            importedTimestamp(row.created_at),
            importedTimestamp(row.updated_at),
          ],
        );
      });

      data.orders.forEach((row) => {
        this.run(
          `
            insert into orders (
              id,
              order_number,
              driver_user_id,
              location_id,
              entry_type,
              factory_order_number,
              inhouse_order_number,
              invoice_number,
              po_number,
              customer_name,
              delivery_address,
              delivery_location_id,
              priority,
              notes,
              driver_flag_type,
              driver_flag_note,
              driver_flagged_at,
              driver_flagged_by_user_id,
              picked_up_at,
              picked_up_by_user_id,
              move_to_factory,
              factory_destination_location_id,
              status,
              scheduled_for,
              original_scheduled_for,
              carry_over_count,
              created_by_user_id,
              created_at,
              completed_at,
              completion_type,
              completed_by_user_id,
              updated_at,
              branding,
              stock_description
            )
            values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
          `,
          [
            requireImportedText(row.id, "Imported order id is required."),
            importedInteger(row.order_number),
            importedNullableText(row.driver_user_id),
            requireImportedText(row.location_id, "Imported order location_id is required."),
            importedText(row.entry_type) || "delivery",
            importedText(row.factory_order_number),
            importedText(row.inhouse_order_number),
            importedText(row.invoice_number),
            importedText(row.po_number),
            importedText(row.customer_name),
            importedText(row.delivery_address),
            importedNullableText(row.delivery_location_id),
            importedText(row.priority) || "medium",
            importedText(row.notes),
            importedNullableText(row.driver_flag_type),
            importedText(row.driver_flag_note),
            importedNullableTimestamp(row.driver_flagged_at),
            importedNullableText(row.driver_flagged_by_user_id),
            importedNullableTimestamp(row.picked_up_at),
            importedNullableText(row.picked_up_by_user_id),
            importedBoolean(row.move_to_factory),
            importedNullableText(row.factory_destination_location_id),
            importedText(row.status) || "active",
            importedDate(row.scheduled_for),
            importedDate(row.original_scheduled_for),
            importedInteger(row.carry_over_count, 0),
            requireImportedText(row.created_by_user_id, "Imported order created_by_user_id is required."),
            importedTimestamp(row.created_at),
            importedNullableTimestamp(row.completed_at),
            importedNullableText(row.completion_type),
            importedNullableText(row.completed_by_user_id),
            importedTimestamp(row.updated_at),
            importedText(row.branding),
            importedText(row.stock_description),
          ],
        );
      });

      data.order_delete_log.forEach((row) => {
        this.run(
          `
            insert into order_delete_log (
              id,
              order_id,
              order_number,
              reference,
              quote_number,
              sales_order_number,
              invoice_number,
              po_number,
              entry_type,
              priority,
              delivery_address,
              delivery_location_id,
              delivery_location_name,
              delivery_location_address,
              branding,
              stock_description,
              notes,
              move_to_factory,
              factory_destination_location_id,
              factory_destination_name,
              factory_destination_address,
              location_id,
              location_name,
              location_address,
              driver_user_id,
              driver_name,
              created_by_user_id,
              created_by_name,
              created_at,
              scheduled_for,
              original_scheduled_for,
              carry_over_count,
              status,
              deleted_by_user_id,
              deleted_by_name,
              deleted_by_role,
              deleted_at,
              notification_attempts,
              last_notification_error,
              notification_sent_at
            )
            values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
          `,
          [
            requireImportedText(row.id, "Imported delete-log id is required."),
            requireImportedText(row.order_id, "Imported delete-log order_id is required."),
            importedNullableInteger(row.order_number),
            importedText(row.reference),
            importedText(row.quote_number),
            importedText(row.sales_order_number),
            importedText(row.invoice_number),
            importedText(row.po_number),
            importedText(row.entry_type) || "delivery",
            importedText(row.priority) || "medium",
            importedText(row.delivery_address),
            importedNullableText(row.delivery_location_id),
            importedText(row.delivery_location_name),
            importedText(row.delivery_location_address),
            importedText(row.branding),
            importedText(row.stock_description),
            importedText(row.notes),
            importedBoolean(row.move_to_factory),
            importedNullableText(row.factory_destination_location_id),
            importedText(row.factory_destination_name),
            importedText(row.factory_destination_address),
            importedNullableText(row.location_id),
            importedText(row.location_name),
            importedText(row.location_address),
            importedNullableText(row.driver_user_id),
            importedText(row.driver_name),
            importedNullableText(row.created_by_user_id),
            importedText(row.created_by_name),
            importedNullableTimestamp(row.created_at),
            importedNullableDate(row.scheduled_for),
            importedNullableDate(row.original_scheduled_for),
            importedInteger(row.carry_over_count, 0),
            importedText(row.status) || "active",
            importedNullableText(row.deleted_by_user_id),
            importedText(row.deleted_by_name),
            importedText(row.deleted_by_role),
            importedTimestamp(row.deleted_at),
            importedInteger(row.notification_attempts, 0),
            importedText(row.last_notification_error),
            importedNullableTimestamp(row.notification_sent_at),
          ],
        );
      });

      data.stock_items.forEach((row) => {
        this.run(
          `
            insert into stock_items (
              id,
              name,
              sku,
              quote_number,
              invoice_number,
              sales_order_number,
              po_number,
              unit,
              notes,
              created_source,
              created_by_user_id,
              created_at,
              updated_at
            )
            values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
          `,
          [
            requireImportedText(row.id, "Imported stock item id is required."),
            importedText(row.name),
            importedText(row.sku),
            importedText(row.quote_number),
            importedText(row.invoice_number),
            importedText(row.sales_order_number),
            importedText(row.po_number),
            importedText(row.unit) || "units",
            importedText(row.notes),
            importedText(row.created_source) || "manual",
            requireImportedText(row.created_by_user_id, "Imported stock item created_by_user_id is required."),
            importedTimestamp(row.created_at),
            importedTimestamp(row.updated_at),
          ],
        );
      });

      data.stock_movements.forEach((row) => {
        this.run(
          `
            insert into stock_movements (
              id,
              stock_item_id,
              movement_type,
              quantity,
              supplier_name,
              driver_user_id,
              notes,
              created_by_user_id,
              created_at
            )
            values (?, ?, ?, ?, ?, ?, ?, ?, ?)
          `,
          [
            requireImportedText(row.id, "Imported stock movement id is required."),
            requireImportedText(row.stock_item_id, "Imported stock movement stock_item_id is required."),
            importedText(row.movement_type) || "in",
            importedInteger(row.quantity),
            importedText(row.supplier_name),
            importedNullableText(row.driver_user_id),
            importedText(row.notes),
            requireImportedText(row.created_by_user_id, "Imported stock movement created_by_user_id is required."),
            importedTimestamp(row.created_at),
          ],
        );
      });

      data.artwork_requests.forEach((row) => {
        this.run(
          `
            insert into artwork_requests (
              id,
              stock_item_id,
              requested_quantity,
              notes,
              sent_to,
              requested_by_user_id,
              sent_at
            )
            values (?, ?, ?, ?, ?, ?, ?)
          `,
          [
            requireImportedText(row.id, "Imported artwork request id is required."),
            requireImportedText(row.stock_item_id, "Imported artwork request stock_item_id is required."),
            importedInteger(row.requested_quantity),
            importedText(row.notes),
            importedText(row.sent_to),
            requireImportedText(row.requested_by_user_id, "Imported artwork request requested_by_user_id is required."),
            importedTimestamp(row.sent_at),
          ],
        );
      });

      data.app_sessions.forEach((row) => {
        this.run(
          `
            insert into app_sessions (
              token,
              user_id,
              created_at,
              last_seen_at,
              expires_at
            )
            values (?, ?, ?, ?, ?)
          `,
          [
            requireImportedText(row.token, "Imported session token is required."),
            requireImportedText(row.user_id, "Imported session user_id is required."),
            importedTimestamp(row.created_at),
            importedTimestamp(row.last_seen_at),
            importedTimestamp(row.expires_at),
          ],
        );
      });

      data.app_settings.forEach((row) => {
        this.run(
          `
            insert into app_settings (
              key,
              value,
              updated_by_user_id,
              updated_at
            )
            values (?, ?, ?, ?)
          `,
          [
            requireImportedText(row.key, "Imported app setting key is required."),
            importedText(row.value),
            importedNullableText(row.updated_by_user_id),
            importedTimestamp(row.updated_at),
          ],
        );
      });

      return {
        source: String(options?.source || "").trim() || "unknown",
        importedCounts: countImportRows(data),
        localCounts: this.getTableCounts(),
      };
    });
  }

  initializeSchema() {
    this.db.exec(`
      create table if not exists app_users (
        id text primary key,
        name text not null,
        role text not null,
        password_hash text not null,
        active integer not null default 1,
        phone text,
        vehicle text,
        last_known_lat real,
        last_known_lng real,
        last_known_recorded_at text,
        created_at text not null,
        updated_at text not null,
        check (role in ('admin', 'sales', 'driver', 'logistics', 'maintenance')),
        check (active in (0, 1))
      );

      create unique index if not exists app_users_name_unique
        on app_users(lower(trim(name)));

      create table if not exists suppliers (
        id text primary key,
        name text not null,
        contact_person text not null default '',
        contact_number text not null default '',
        factory integer not null default 0,
        created_by text not null references app_users(id) on delete restrict,
        created_at text not null,
        updated_at text not null,
        check (factory in (0, 1))
      );

      create unique index if not exists suppliers_name_unique
        on suppliers(lower(trim(name)));

      create table if not exists locations (
        id text primary key,
        supplier_id text references suppliers(id) on delete restrict,
        location_type text not null default 'supplier',
        name text not null,
        address text not null,
        lat real,
        lng real,
        contact_person text not null default '',
        contact_number text not null default '',
        notes text not null default '',
        created_by text not null references app_users(id) on delete restrict,
        created_at text not null,
        updated_at text not null,
        check (location_type in ('supplier', 'factory', 'both', 'client'))
      );

      create unique index if not exists locations_supplier_name_unique
        on locations(coalesce(supplier_id, ''), location_type, lower(trim(name)));

      create table if not exists orders (
        id text primary key,
        order_number integer not null unique,
        driver_user_id text references app_users(id) on delete restrict,
        location_id text not null references locations(id) on delete restrict,
        entry_type text not null default 'delivery',
        factory_order_number text not null default '',
        inhouse_order_number text not null default '',
        invoice_number text not null default '',
        po_number text not null default '',
        customer_name text not null,
        delivery_address text not null default '',
        delivery_location_id text references locations(id) on delete restrict,
        priority text not null default 'medium',
        notes text not null default '',
        driver_flag_type text,
        driver_flag_note text not null default '',
        driver_flagged_at text,
        driver_flagged_by_user_id text references app_users(id) on delete set null,
        picked_up_at text,
        picked_up_by_user_id text references app_users(id) on delete set null,
        move_to_factory integer not null default 0,
        factory_destination_location_id text references locations(id) on delete restrict,
        status text not null default 'active',
        scheduled_for text not null,
        original_scheduled_for text not null,
        carry_over_count integer not null default 0,
        created_by_user_id text not null references app_users(id) on delete restrict,
        created_at text not null,
        completed_at text,
        completion_type text,
        completed_by_user_id text references app_users(id) on delete set null,
        updated_at text not null,
        branding text not null default '',
        stock_description text not null default '',
        check (entry_type in ('collection', 'delivery')),
        check (priority in ('high', 'medium', 'low')),
        check (status in ('active', 'completed')),
        check (driver_flag_type in ('not_collected', 'not_ready') or driver_flag_type is null),
        check (completion_type in ('office', 'factory') or completion_type is null),
        check (move_to_factory in (0, 1)),
        check (carry_over_count >= 0)
      );

      create index if not exists orders_driver_scheduled_idx
        on orders(driver_user_id, scheduled_for);

      create index if not exists orders_status_scheduled_idx
        on orders(status, scheduled_for);

      create index if not exists orders_factory_destination_idx
        on orders(factory_destination_location_id);

      create table if not exists order_delete_log (
        id text primary key,
        order_id text not null,
        order_number integer,
        reference text not null default '',
        quote_number text not null default '',
        sales_order_number text not null default '',
        invoice_number text not null default '',
        po_number text not null default '',
        entry_type text not null default 'delivery',
        priority text not null default 'medium',
        delivery_address text not null default '',
        delivery_location_id text,
        delivery_location_name text not null default '',
        delivery_location_address text not null default '',
        branding text not null default '',
        stock_description text not null default '',
        notes text not null default '',
        move_to_factory integer not null default 0,
        factory_destination_location_id text,
        factory_destination_name text not null default '',
        factory_destination_address text not null default '',
        location_id text,
        location_name text not null default '',
        location_address text not null default '',
        driver_user_id text,
        driver_name text not null default '',
        created_by_user_id text,
        created_by_name text not null default '',
        created_at text,
        scheduled_for text,
        original_scheduled_for text,
        carry_over_count integer not null default 0,
        status text not null default 'active',
        deleted_by_user_id text references app_users(id) on delete set null,
        deleted_by_name text not null default '',
        deleted_by_role text not null default '',
        deleted_at text not null,
        notification_attempts integer not null default 0,
        last_notification_error text not null default '',
        notification_sent_at text,
        check (move_to_factory in (0, 1))
      );

      create index if not exists order_delete_log_pending_idx
        on order_delete_log(notification_sent_at, deleted_at);

      create table if not exists stock_items (
        id text primary key,
        name text not null,
        sku text not null default '',
        quote_number text not null default '',
        invoice_number text not null default '',
        sales_order_number text not null default '',
        po_number text not null default '',
        unit text not null default 'units',
        notes text not null default '',
        created_source text not null default 'manual',
        created_by_user_id text not null references app_users(id) on delete restrict,
        created_at text not null,
        updated_at text not null,
        check (created_source in ('manual', 'order'))
      );

      create unique index if not exists stock_items_reference_name_unique
        on stock_items(
          lower(trim(name)),
          lower(trim(quote_number)),
          lower(trim(invoice_number)),
          lower(trim(sales_order_number)),
          lower(trim(po_number))
        );

      create unique index if not exists stock_items_sku_unique
        on stock_items(lower(trim(sku)))
        where trim(sku) <> '';

      create table if not exists stock_movements (
        id text primary key,
        stock_item_id text not null references stock_items(id) on delete cascade,
        movement_type text not null,
        quantity integer not null,
        supplier_name text not null default '',
        driver_user_id text references app_users(id) on delete restrict,
        notes text not null default '',
        created_by_user_id text not null references app_users(id) on delete restrict,
        created_at text not null,
        check (movement_type in ('in', 'out')),
        check (quantity > 0)
      );

      create index if not exists stock_movements_item_created_idx
        on stock_movements(stock_item_id, created_at desc);

      create index if not exists stock_movements_driver_created_idx
        on stock_movements(driver_user_id, created_at desc);

      create table if not exists artwork_requests (
        id text primary key,
        stock_item_id text not null references stock_items(id) on delete cascade,
        requested_quantity integer not null,
        notes text not null default '',
        sent_to text not null default '',
        requested_by_user_id text not null references app_users(id) on delete restrict,
        sent_at text not null,
        check (requested_quantity > 0)
      );

      create index if not exists artwork_requests_item_sent_idx
        on artwork_requests(stock_item_id, sent_at desc);

      create table if not exists app_sessions (
        token text primary key,
        user_id text not null references app_users(id) on delete cascade,
        created_at text not null,
        last_seen_at text not null,
        expires_at text not null
      );

      create index if not exists app_sessions_user_idx
        on app_sessions(user_id);

      create table if not exists app_settings (
        key text primary key,
        value text not null default '',
        updated_by_user_id text references app_users(id) on delete set null,
        updated_at text not null
      );
    `);
  }

  ensureSchemaMigrations() {
    this.ensureAppUsersRoleSchema();
    this.ensureReusableDeliveryLocationSchema();
  }

  getTableSql(tableName) {
    return String(
      this.get(
        `
          select sql
          from sqlite_master
          where type = 'table'
            and name = ?
          limit 1
        `,
        [tableName],
      )?.sql || "",
    );
  }

  hasTableColumn(tableName, columnName) {
    return this.all(`pragma table_info(${tableName})`)
      .some((column) => String(column?.name || "").toLowerCase() === String(columnName || "").toLowerCase());
  }

  ensureReusableDeliveryLocationSchema() {
    const locationsSql = this.getTableSql("locations").toLowerCase();
    if (locationsSql && !locationsSql.includes("'client'")) {
      this.rebuildLocationsForClientType();
    } else {
      this.db.exec(`
        drop index if exists locations_supplier_name_unique;
        create unique index if not exists locations_supplier_name_unique
          on locations(coalesce(supplier_id, ''), location_type, lower(trim(name)));
      `);
    }

    if (!this.hasTableColumn("orders", "delivery_location_id")) {
      this.db.exec("alter table orders add column delivery_location_id text references locations(id) on delete restrict;");
    }

    this.db.exec(`
      create index if not exists orders_delivery_location_idx
        on orders(delivery_location_id);
    `);

    if (!this.hasTableColumn("order_delete_log", "delivery_location_id")) {
      this.db.exec("alter table order_delete_log add column delivery_location_id text;");
    }
    if (!this.hasTableColumn("order_delete_log", "delivery_location_name")) {
      this.db.exec("alter table order_delete_log add column delivery_location_name text not null default '';");
    }
    if (!this.hasTableColumn("order_delete_log", "delivery_location_address")) {
      this.db.exec("alter table order_delete_log add column delivery_location_address text not null default '';");
    }
  }

  rebuildLocationsForClientType() {
    this.db.exec("PRAGMA foreign_keys = OFF;");
    try {
      this.db.exec(`
        drop table if exists locations__migrated;

        create table locations__migrated (
          id text primary key,
          supplier_id text references suppliers(id) on delete restrict,
          location_type text not null default 'supplier',
          name text not null,
          address text not null,
          lat real,
          lng real,
          contact_person text not null default '',
          contact_number text not null default '',
          notes text not null default '',
          created_by text not null references app_users(id) on delete restrict,
          created_at text not null,
          updated_at text not null,
          check (location_type in ('supplier', 'factory', 'both', 'client'))
        );

        insert into locations__migrated (
          id,
          supplier_id,
          location_type,
          name,
          address,
          lat,
          lng,
          contact_person,
          contact_number,
          notes,
          created_by,
          created_at,
          updated_at
        )
        select
          id,
          supplier_id,
          location_type,
          name,
          address,
          lat,
          lng,
          contact_person,
          contact_number,
          notes,
          created_by,
          created_at,
          updated_at
        from locations;

        drop table locations;
        alter table locations__migrated rename to locations;

        create unique index if not exists locations_supplier_name_unique
          on locations(coalesce(supplier_id, ''), location_type, lower(trim(name)));
      `);
    } finally {
      this.db.exec("PRAGMA foreign_keys = ON;");
    }
  }

  ensureAppUsersRoleSchema() {
    const tableSql = this.getTableSql("app_users").toLowerCase();

    if (tableSql.includes("'maintenance'")) {
      return;
    }

    this.db.exec("PRAGMA foreign_keys = OFF;");
    try {
      this.db.exec(`
        drop table if exists app_users__migrated;

        create table app_users__migrated (
          id text primary key,
          name text not null,
          role text not null,
          password_hash text not null,
          active integer not null default 1,
          phone text,
          vehicle text,
          last_known_lat real,
          last_known_lng real,
          last_known_recorded_at text,
          created_at text not null,
          updated_at text not null,
          check (role in ('admin', 'sales', 'driver', 'logistics', 'maintenance')),
          check (active in (0, 1))
        );

        insert into app_users__migrated (
          id,
          name,
          role,
          password_hash,
          active,
          phone,
          vehicle,
          last_known_lat,
          last_known_lng,
          last_known_recorded_at,
          created_at,
          updated_at
        )
        select
          id,
          name,
          role,
          password_hash,
          active,
          phone,
          vehicle,
          last_known_lat,
          last_known_lng,
          last_known_recorded_at,
          created_at,
          updated_at
        from app_users;

        drop table app_users;
        alter table app_users__migrated rename to app_users;

        create unique index if not exists app_users_name_unique
          on app_users(lower(trim(name)));
      `);
    } finally {
      this.db.exec("PRAGMA foreign_keys = ON;");
    }
  }

  ensureOfficeLocationForExistingUsers() {
    const firstUser = this.get(
      `
        select id
        from app_users
        order by created_at asc
        limit 1
      `,
    );

    if (firstUser?.id) {
      this.ensureOfficeLocation(firstUser.id);
    }
  }

  getLoginState() {
    const today = todayLocal();
    const weekStart = getWeekStart(today);
    const hasUsers = Boolean(this.get("select 1 as value from app_users limit 1")?.value);
    return {
      today,
      weekStart,
      weekEnd: addDays(weekStart, 6),
      hasUsers,
    };
  }

  bootstrapAdmin(parameters = {}) {
    return this.withTransaction(() => {
      if (this.get("select id from app_users limit 1")) {
        throw createHttpError(400, "The first admin account already exists.");
      }

      const name = normalizeRequiredText(parameters?.p_name, "Name is required.");
      const password = normalizeRequiredText(parameters?.p_password, "Password is required.");
      if (password.length < 4) {
        throw createHttpError(400, "Password must be at least 4 characters long.");
      }

      const now = nowIso();
      const userId = randomId();
      this.run(
        `
          insert into app_users (
            id,
            name,
            role,
            password_hash,
            active,
            phone,
            vehicle,
            created_at,
            updated_at
          )
          values (?, ?, 'admin', ?, 1, null, null, ?, ?)
        `,
        [userId, name, hashPassword(password), now, now],
      );

      this.ensureOfficeLocation(userId);
      const token = this.issueSession(userId);
      return { ok: true, token };
    });
  }

  loginUser(parameters = {}) {
    const name = normalizeRequiredText(parameters?.p_name, "Name and password are required.");
    const password = normalizeRequiredText(parameters?.p_password, "Name and password are required.");
    const user = this.get(
      `
        select *
        from app_users
        where lower(trim(name)) = lower(trim(?))
        limit 1
      `,
      [name],
    );

    if (!user || !passwordMatches(password, user.password_hash)) {
      throw createHttpError(400, "Invalid name or password.");
    }

    if (!Boolean(user.active)) {
      throw createHttpError(400, "This account is inactive.");
    }

    const token = this.issueSession(user.id);
    return { ok: true, token };
  }

  logoutUser(parameters = {}) {
    const token = normalizeOptionalText(parameters?.p_token);
    if (token) {
      this.run("delete from app_sessions where token = ?", [token]);
    }
    return { ok: true };
  }

  runDailyRollover() {
    return this.withTransaction(() => this.rollForwardOpenOrders(todayLocal()));
  }

  getAppSnapshot(parameters = {}) {
    return this.withTransaction(() => {
      const actor = this.requireUser(parameters?.p_token);
      const rollover = this.rollForwardOpenOrders(todayLocal());
      const users = this.selectSnapshotUsers(actor);
      const suppliers = this.selectSnapshotSuppliers(actor);
      const locations = this.selectSnapshotLocations(actor);
      const orders = this.selectSnapshotOrders(actor);
      const stockItems = this.selectSnapshotStockItems(actor);
      const stockMovements = this.selectSnapshotStockMovements(actor);
      const artworkRequests = this.selectSnapshotArtworkRequests(actor);

      return {
        today: todayLocal(),
        rollover,
        user: this.buildUserJson(this.selectUserById(actor.id)),
        users,
        suppliers,
        locations,
        orders,
        stockItems,
        stockMovements,
        artworkRequests,
      };
    });
  }

  recordDriverPosition(parameters = {}) {
    const actor = this.requireUser(parameters?.p_token, ["driver"]);
    const lat = toFiniteNumber(parameters?.p_lat);
    const lng = toFiniteNumber(parameters?.p_lng);

    if (!Number.isFinite(lat) || !Number.isFinite(lng)) {
      throw createHttpError(400, "Latitude and longitude are required.");
    }
    if (lat < -90 || lat > 90) {
      throw createHttpError(400, "Latitude must be between -90 and 90.");
    }
    if (lng < -180 || lng > 180) {
      throw createHttpError(400, "Longitude must be between -180 and 180.");
    }

    const recordedAt = nowIso();
    this.run(
      `
        update app_users
        set last_known_lat = ?,
            last_known_lng = ?,
            last_known_recorded_at = ?,
            updated_at = ?
        where id = ?
      `,
      [lat, lng, recordedAt, recordedAt, actor.id],
    );

    return {
      ok: true,
      recordedAt,
    };
  }

  createUserAccount(parameters = {}) {
    const actor = this.requireUser(parameters?.p_token, ["admin"]);
    const name = normalizeRequiredText(parameters?.p_name, "Name is required.");
    const password = normalizeRequiredText(parameters?.p_password, "Password is required.");
    const role = normalizeRole(parameters?.p_role);
    const phone = normalizeOptionalText(parameters?.p_phone);

    if (password.length < 4) {
      throw createHttpError(400, "Password must be at least 4 characters long.");
    }
    if (!USER_ROLES.has(role)) {
      throw createHttpError(400, "Invalid role.");
    }
    if (this.get("select id from app_users where lower(trim(name)) = lower(trim(?)) limit 1", [name])) {
      throw createHttpError(400, "That name is already in use.");
    }
    if (role === "driver" && !phone) {
      throw createHttpError(400, "Driver accounts require a phone number.");
    }

    const now = nowIso();
    this.run(
      `
        insert into app_users (
          id,
          name,
          role,
          password_hash,
          active,
          phone,
          vehicle,
          created_at,
          updated_at
        )
        values (?, ?, ?, ?, 1, ?, null, ?, ?)
      `,
      [randomId(), name, role, hashPassword(password), role === "driver" ? phone : null, now, now],
    );

    return { ok: true, createdBy: actor.id };
  }

  updateUserAccount(parameters = {}) {
    const actor = this.requireUser(parameters?.p_token, ["admin"]);
    const userId = normalizeId(parameters?.p_user_id, "User not found.");
    const target = this.selectUserById(userId);
    if (!target) {
      throw createHttpError(400, "User not found.");
    }

    const name = normalizeRequiredText(parameters?.p_name, "Name is required.");
    const role = normalizeRole(parameters?.p_role);
    const phone = normalizeOptionalText(parameters?.p_phone);
    const password = normalizeOptionalText(parameters?.p_password);

    if (!USER_ROLES.has(role)) {
      throw createHttpError(400, "Invalid role.");
    }
    if (password && password.length < 4) {
      throw createHttpError(400, "Password must be at least 4 characters long.");
    }
    if (
      this.get(
        `
          select id
          from app_users
          where id <> ?
            and lower(trim(name)) = lower(trim(?))
          limit 1
        `,
        [userId, name],
      )
    ) {
      throw createHttpError(400, "That name is already in use.");
    }
    if (role === "driver" && !phone) {
      throw createHttpError(400, "Driver accounts require a phone number.");
    }
    if (target.id === actor.id && role !== actor.role) {
      throw createHttpError(400, "You cannot change your own role.");
    }

    const now = nowIso();
    this.run(
      `
        update app_users
        set name = ?,
            role = ?,
            phone = ?,
            password_hash = ?,
            updated_at = ?
        where id = ?
      `,
      [
        name,
        role,
        role === "driver" ? phone : null,
        password ? hashPassword(password) : target.password_hash,
        now,
        userId,
      ],
    );

    return { ok: true, updatedBy: actor.id };
  }

  toggleUserActive(parameters = {}) {
    const actor = this.requireUser(parameters?.p_token, ["admin"]);
    const userId = normalizeId(parameters?.p_user_id, "User not found.");
    const target = this.selectUserById(userId);
    if (!target) {
      throw createHttpError(400, "User not found.");
    }
    if (target.id === actor.id) {
      throw createHttpError(400, "You cannot change your own active status.");
    }

    this.run(
      `
        update app_users
        set active = case when active = 1 then 0 else 1 end,
            updated_at = ?
        where id = ?
      `,
      [nowIso(), userId],
    );

    return { ok: true };
  }

  deleteUserAccount(parameters = {}) {
    const actor = this.requireUser(parameters?.p_token, ["admin"]);
    const userId = normalizeId(parameters?.p_user_id, "User not found.");
    const target = this.selectUserById(userId);

    if (!target) {
      throw createHttpError(400, "User not found.");
    }
    if (target.id === actor.id) {
      throw createHttpError(400, "You cannot delete your own account.");
    }
    if (
      this.get(
        `
          select id
          from orders
          where driver_user_id = ?
             or created_by_user_id = ?
          limit 1
        `,
        [userId, userId],
      )
    ) {
      throw createHttpError(400, "This account has order history. Disable it instead of deleting it.");
    }

    this.run("delete from app_users where id = ?", [userId]);
    return { ok: true };
  }

  createSupplier(parameters = {}) {
    const actor = this.requireUser(parameters?.p_token, ["admin"]);
    const name = normalizeRequiredText(parameters?.p_name, "Supplier name and contact number are required.");
    const contactPerson = normalizeOptionalText(parameters?.p_contact_person);
    const contactNumber = normalizeRequiredText(parameters?.p_contact_number, "Supplier name and contact number are required.");
    const factory = Boolean(parameters?.p_factory);

    if (this.get("select id from suppliers where lower(trim(name)) = lower(trim(?)) limit 1", [name])) {
      throw createHttpError(400, "That supplier already exists.");
    }

    const now = nowIso();
    this.run(
      `
        insert into suppliers (
          id,
          name,
          contact_person,
          contact_number,
          factory,
          created_by,
          created_at,
          updated_at
        )
        values (?, ?, ?, ?, ?, ?, ?, ?)
      `,
      [randomId(), name, contactPerson, contactNumber, factory ? 1 : 0, actor.id, now, now],
    );

    return { ok: true };
  }

  updateSupplier(parameters = {}) {
    const actor = this.requireUser(parameters?.p_token, ["admin"]);
    const supplierId = normalizeId(parameters?.p_supplier_id, "Supplier not found.");
    const name = normalizeRequiredText(parameters?.p_name, "Supplier name and contact number are required.");
    const contactPerson = normalizeOptionalText(parameters?.p_contact_person);
    const contactNumber = normalizeRequiredText(parameters?.p_contact_number, "Supplier name and contact number are required.");
    const factory = Boolean(parameters?.p_factory);

    if (!this.get("select id from suppliers where id = ? limit 1", [supplierId])) {
      throw createHttpError(400, "Supplier not found.");
    }
    if (
      this.get(
        `
          select id
          from suppliers
          where id <> ?
            and lower(trim(name)) = lower(trim(?))
          limit 1
        `,
        [supplierId, name],
      )
    ) {
      throw createHttpError(400, "That supplier already exists.");
    }

    this.run(
      `
        update suppliers
        set name = ?,
            contact_person = ?,
            contact_number = ?,
            factory = ?,
            updated_at = ?
        where id = ?
      `,
      [name, contactPerson, contactNumber, factory ? 1 : 0, nowIso(), supplierId],
    );

    return { ok: true, updatedBy: actor.id };
  }

  deleteSupplier(parameters = {}) {
    const actor = this.requireUser(parameters?.p_token, ["admin"]);
    const supplierId = normalizeId(parameters?.p_supplier_id, "Supplier not found.");

    if (this.get("select id from locations where supplier_id = ? limit 1", [supplierId])) {
      throw createHttpError(400, "Delete supplier locations first.");
    }

    this.run("delete from suppliers where id = ?", [supplierId]);
    return { ok: true, deletedBy: actor.id };
  }

  createLocation(parameters = {}) {
    const actor = this.requireUser(parameters?.p_token, ["admin"]);
    const locationType = normalizeLocationType(parameters?.p_location_type);
    const name = normalizeRequiredText(parameters?.p_name, "Location name and address are required.");
    const address = normalizeRequiredText(parameters?.p_address, "Location name and address are required.");
    const lat = normalizeOptionalNumber(parameters?.p_lat);
    const lng = normalizeOptionalNumber(parameters?.p_lng);
    const contactPerson = normalizeOptionalText(parameters?.p_contact_person);
    const contactNumber = normalizeOptionalText(parameters?.p_contact_number);

    if (!LOCATION_TYPES.has(locationType)) {
      throw createHttpError(400, "Location type must be supplier, factory, both, or client.");
    }
    if (normalizeNameKey(name) === "office") {
      throw createHttpError(400, "Office already exists as a built-in location. Edit it instead.");
    }
    if ((lat === null) !== (lng === null)) {
      throw createHttpError(400, "Latitude and longitude must both be provided, or both left blank.");
    }

    const now = nowIso();
    this.run(
      `
        insert into locations (
          id,
          supplier_id,
          location_type,
          name,
          address,
          lat,
          lng,
          contact_person,
          contact_number,
          notes,
          created_by,
          created_at,
          updated_at
        )
        values (?, null, ?, ?, ?, ?, ?, ?, ?, '', ?, ?, ?)
      `,
      [randomId(), locationType, name, address, lat, lng, contactPerson, contactNumber, actor.id, now, now],
    );

    return { ok: true };
  }

  updateLocation(parameters = {}) {
    const actor = this.requireUser(parameters?.p_token, ["admin"]);
    const locationId = normalizeId(parameters?.p_location_id, "Location not found.");
    const existing = this.selectLocationById(locationId);
    if (!existing) {
      throw createHttpError(400, "Location not found.");
    }

    let locationType = normalizeLocationType(parameters?.p_location_type);
    const name = normalizeRequiredText(parameters?.p_name, "Location name and address are required.");
    const address = normalizeRequiredText(parameters?.p_address, "Location name and address are required.");
    const lat = normalizeOptionalNumber(parameters?.p_lat);
    const lng = normalizeOptionalNumber(parameters?.p_lng);
    const contactPerson = normalizeOptionalText(parameters?.p_contact_person);
    const contactNumber = normalizeOptionalText(parameters?.p_contact_number);

    if (!LOCATION_TYPES.has(locationType)) {
      throw createHttpError(400, "Location type must be supplier, factory, both, or client.");
    }
    if ((lat === null) !== (lng === null)) {
      throw createHttpError(400, "Latitude and longitude must both be provided, or both left blank.");
    }

    const existingIsOffice = !existing.supplier_id && normalizeNameKey(existing.name) === "office";
    if (existingIsOffice) {
      if (normalizeNameKey(name) !== "office") {
        throw createHttpError(400, "Office is a built-in location and must stay named Office.");
      }
      locationType = "supplier";
    } else if (normalizeNameKey(name) === "office") {
      throw createHttpError(400, "Office is reserved as a built-in location.");
    }

    this.run(
      `
        update locations
        set location_type = ?,
            name = ?,
            address = ?,
            lat = ?,
            lng = ?,
            contact_person = ?,
            contact_number = ?,
            updated_at = ?
        where id = ?
      `,
      [locationType, name, address, lat, lng, contactPerson, contactNumber, nowIso(), locationId],
    );

    return { ok: true, updatedBy: actor.id };
  }

  deleteLocation(parameters = {}) {
    const actor = this.requireUser(parameters?.p_token, ["admin"]);
    const locationId = normalizeId(parameters?.p_location_id, "Location not found.");
    const location = this.selectLocationById(locationId);
    if (!location) {
      throw createHttpError(400, "Location not found.");
    }
    if (!location.supplier_id && normalizeNameKey(location.name) === "office") {
      throw createHttpError(400, "Office is a built-in location and cannot be deleted.");
    }
    if (
      this.get(
        `
          select id
          from orders
          where location_id = ?
             or factory_destination_location_id = ?
             or delivery_location_id = ?
          limit 1
        `,
        [locationId, locationId, locationId],
      )
    ) {
      throw createHttpError(400, "This location has order history and cannot be deleted.");
    }

    this.run("delete from locations where id = ?", [locationId]);
    return { ok: true, deletedBy: actor.id };
  }

  resolveDeliveryLocationForOrder({ actor, deliveryLocationId, deliveryAddress, saveDeliveryLocation, deliveryLocationName }) {
    const normalizedLocationId = normalizeOptionalText(deliveryLocationId);
    const normalizedAddress = normalizeOptionalText(deliveryAddress);

    if (normalizedLocationId) {
      const deliveryLocation = this.selectLocationById(normalizedLocationId);
      if (!deliveryLocation || deliveryLocation.location_type !== "client") {
        throw createHttpError(400, "Delivery location not found.");
      }

      return {
        deliveryLocationId: deliveryLocation.id,
        deliveryAddress: normalizeOptionalText(deliveryLocation.address) || normalizedAddress,
      };
    }

    if (!normalizedAddress) {
      throw createHttpError(400, "Delivery address is required for delivery entries.");
    }

    if (!saveDeliveryLocation) {
      return {
        deliveryLocationId: null,
        deliveryAddress: normalizedAddress,
      };
    }

    const deliveryLocation = this.findOrCreateClientDeliveryLocation({
      actorId: actor.id,
      name: deliveryLocationName,
      address: normalizedAddress,
    });

    return {
      deliveryLocationId: deliveryLocation.id,
      deliveryAddress: deliveryLocation.address || normalizedAddress,
    };
  }

  findOrCreateClientDeliveryLocation({ actorId, name, address }) {
    const normalizedAddress = normalizeRequiredText(address, "Delivery address is required for delivery entries.");
    const existingByAddress = this.get(
      `
        select *
        from locations
        where location_type = 'client'
          and lower(trim(address)) = lower(trim(?))
        limit 1
      `,
      [normalizedAddress],
    );

    if (existingByAddress) {
      return existingByAddress;
    }

    const baseLocationName = normalizeOptionalText(name) || normalizedAddress;
    let locationName = baseLocationName;
    let suffix = 2;
    while (
      this.get(
        `
          select id
          from locations
          where location_type = 'client'
            and lower(trim(name)) = lower(trim(?))
          limit 1
        `,
        [locationName],
      )
    ) {
      locationName = `${baseLocationName} ${suffix}`;
      suffix += 1;
    }

    const now = nowIso();
    const locationId = randomId();
    this.run(
      `
        insert into locations (
          id,
          supplier_id,
          location_type,
          name,
          address,
          lat,
          lng,
          contact_person,
          contact_number,
          notes,
          created_by,
          created_at,
          updated_at
        )
        values (?, null, 'client', ?, ?, null, null, '', '', 'Saved from a delivery entry.', ?, ?, ?)
      `,
      [locationId, locationName, normalizedAddress, actorId, now, now],
    );

    return this.selectLocationById(locationId);
  }

  createStockItem(parameters = {}) {
    const actor = this.requireUser(parameters?.p_token);
    if (!["admin", "logistics"].includes(actor.role)) {
      throw createHttpError(400, "Permission denied");
    }

    const name = normalizeRequiredText(parameters?.p_name, "Stock item name is required.");
    const sku = normalizeOptionalText(parameters?.p_sku);
    const quoteNumber = normalizeOptionalText(parameters?.p_quote_number);
    const invoiceNumber = normalizeOptionalText(parameters?.p_invoice_number);
    const salesOrderNumber = normalizeOptionalText(parameters?.p_sales_order_number);
    const poNumber = normalizeOptionalText(parameters?.p_po_number);
    const unit = normalizeOptionalText(parameters?.p_unit) || "units";
    const notes = normalizeOptionalText(parameters?.p_notes);
    const initialQuantity = Math.trunc(Number(parameters?.p_initial_quantity || 0));

    if (!quoteNumber && !invoiceNumber && !salesOrderNumber && !poNumber) {
      throw createHttpError(400, "Enter at least one quote, sales order, invoice, or PO number.");
    }
    if (!Number.isInteger(initialQuantity) || initialQuantity <= 0) {
      throw createHttpError(400, "Opening stock quantity must be greater than zero.");
    }
    if (this.findStockItemByReference(name, quoteNumber, invoiceNumber, salesOrderNumber, poNumber)) {
      throw createHttpError(400, "That stock item already exists for the same reference combination.");
    }
    if (sku && this.get("select id from stock_items where lower(trim(sku)) = lower(trim(?)) limit 1", [sku])) {
      throw createHttpError(400, "That stock code is already in use.");
    }

    return this.withTransaction(() => {
      const now = nowIso();
      const stockItemId = randomId();
      this.run(
        `
          insert into stock_items (
            id,
            name,
            sku,
            quote_number,
            invoice_number,
            sales_order_number,
            po_number,
            unit,
            notes,
            created_source,
            created_by_user_id,
            created_at,
            updated_at
          )
          values (?, ?, ?, ?, ?, ?, ?, ?, ?, 'manual', ?, ?, ?)
        `,
        [stockItemId, name, sku, quoteNumber, invoiceNumber, salesOrderNumber, poNumber, unit, notes, actor.id, now, now],
      );
      this.run(
        `
          insert into stock_movements (
            id,
            stock_item_id,
            movement_type,
            quantity,
            supplier_name,
            driver_user_id,
            notes,
            created_by_user_id,
            created_at
          )
          values (?, ?, 'in', ?, 'Opening stock', null, 'Opening stock logged when the item was created.', ?, ?)
        `,
        [randomId(), stockItemId, initialQuantity, actor.id, now],
      );

      return { ok: true, stockItemId };
    });
  }

  updateStockItem(parameters = {}) {
    const actor = this.requireUser(parameters?.p_token, ["admin"]);
    const stockItemId = normalizeId(parameters?.p_stock_item_id, "Stock item not found.");
    const existing = this.selectStockItemById(stockItemId);
    if (!existing) {
      throw createHttpError(400, "Stock item not found.");
    }

    const name = normalizeRequiredText(parameters?.p_name, "Stock item name is required.");
    const sku = normalizeOptionalText(parameters?.p_sku);
    const quoteNumber = normalizeOptionalText(parameters?.p_quote_number);
    const invoiceNumber = normalizeOptionalText(parameters?.p_invoice_number);
    const salesOrderNumber = normalizeOptionalText(parameters?.p_sales_order_number);
    const poNumber = normalizeOptionalText(parameters?.p_po_number);
    const unit = normalizeOptionalText(parameters?.p_unit) || "units";
    const notes = normalizeOptionalText(parameters?.p_notes);

    if (!quoteNumber && !invoiceNumber && !salesOrderNumber && !poNumber) {
      throw createHttpError(400, "Enter at least one quote, sales order, invoice, or PO number.");
    }
    if (
      this.get(
        `
          select id
          from stock_items
          where id <> ?
            and lower(trim(name)) = lower(trim(?))
            and lower(trim(quote_number)) = lower(trim(?))
            and lower(trim(invoice_number)) = lower(trim(?))
            and lower(trim(sales_order_number)) = lower(trim(?))
            and lower(trim(po_number)) = lower(trim(?))
          limit 1
        `,
        [stockItemId, name, quoteNumber, invoiceNumber, salesOrderNumber, poNumber],
      )
    ) {
      throw createHttpError(400, "That stock item already exists for the same reference combination.");
    }
    if (
      sku
      && this.get(
        `
          select id
          from stock_items
          where id <> ?
            and lower(trim(sku)) = lower(trim(?))
          limit 1
        `,
        [stockItemId, sku],
      )
    ) {
      throw createHttpError(400, "That stock code is already in use.");
    }

    this.run(
      `
        update stock_items
        set name = ?,
            sku = ?,
            quote_number = ?,
            invoice_number = ?,
            sales_order_number = ?,
            po_number = ?,
            unit = ?,
            notes = ?,
            updated_at = ?
        where id = ?
      `,
      [name, sku, quoteNumber, invoiceNumber, salesOrderNumber, poNumber, unit, notes, nowIso(), stockItemId],
    );

    return { ok: true, updatedBy: actor.id, stockItemId: existing.id };
  }

  deleteStockItem(parameters = {}) {
    const actor = this.requireUser(parameters?.p_token, ["admin"]);
    const stockItemId = normalizeId(parameters?.p_stock_item_id, "Stock item not found.");
    if (!this.selectStockItemById(stockItemId)) {
      throw createHttpError(400, "Stock item not found.");
    }

    this.run("delete from stock_items where id = ?", [stockItemId]);
    return { ok: true, deletedBy: actor.id };
  }

  recordStockMovement(parameters = {}) {
    const actor = this.requireUser(parameters?.p_token);
    if (!["admin", "logistics"].includes(actor.role)) {
      throw createHttpError(400, "Permission denied");
    }

    const stockItemId = normalizeId(parameters?.p_stock_item_id, "Stock item not found.");
    const movementType = normalizeMovementType(parameters?.p_movement_type);
    const quantity = Math.trunc(Number(parameters?.p_quantity || 0));
    const supplierName = normalizeOptionalText(parameters?.p_supplier_name);
    const driverUserId = normalizeOptionalText(parameters?.p_driver_user_id);
    const notes = normalizeOptionalText(parameters?.p_notes);

    if (!STOCK_MOVEMENT_TYPES.has(movementType)) {
      throw createHttpError(400, "Movement type must be in or out.");
    }
    if (!Number.isInteger(quantity) || quantity <= 0) {
      throw createHttpError(400, "Quantity must be greater than zero.");
    }

    const stockItem = this.selectStockItemById(stockItemId);
    if (!stockItem) {
      throw createHttpError(400, "Stock item not found.");
    }
    if (movementType === "in" && !supplierName) {
      throw createHttpError(400, "Supplier is required for stock coming in.");
    }

    let driver = null;
    if (movementType === "out") {
      driver = this.selectActiveDriver(driverUserId);
      if (!driver) {
        throw createHttpError(400, "Driver is required for stock going out.");
      }
      if (this.getStockOnHand(stockItemId) < quantity) {
        throw createHttpError(400, "Not enough stock on hand for that movement.");
      }
    }

    this.run(
      `
        insert into stock_movements (
          id,
          stock_item_id,
          movement_type,
          quantity,
          supplier_name,
          driver_user_id,
          notes,
          created_by_user_id,
          created_at
        )
        values (?, ?, ?, ?, ?, ?, ?, ?, ?)
      `,
      [
        randomId(),
        stockItemId,
        movementType,
        quantity,
        movementType === "in" ? supplierName : "",
        movementType === "out" ? driver.id : null,
        notes,
        actor.id,
        nowIso(),
      ],
    );

    return { ok: true };
  }

  updateStockMovement(parameters = {}) {
    const actor = this.requireUser(parameters?.p_token);
    if (!["admin", "logistics"].includes(actor.role)) {
      throw createHttpError(400, "Permission denied");
    }

    const movementId = normalizeId(parameters?.p_stock_movement_id, "Stock movement not found.");
    const stockItemId = normalizeId(parameters?.p_stock_item_id, "Stock item not found.");
    const movementType = normalizeMovementType(parameters?.p_movement_type);
    const quantity = Math.trunc(Number(parameters?.p_quantity || 0));
    const supplierName = normalizeOptionalText(parameters?.p_supplier_name);
    const driverUserId = normalizeOptionalText(parameters?.p_driver_user_id);
    const notes = normalizeOptionalText(parameters?.p_notes);

    if (!STOCK_MOVEMENT_TYPES.has(movementType)) {
      throw createHttpError(400, "Movement type must be in or out.");
    }
    if (!Number.isInteger(quantity) || quantity <= 0) {
      throw createHttpError(400, "Quantity must be greater than zero.");
    }

    const existing = this.selectStockMovementById(movementId);
    if (!existing) {
      throw createHttpError(400, "Stock movement not found.");
    }

    const stockItem = this.selectStockItemById(stockItemId);
    if (!stockItem) {
      throw createHttpError(400, "Stock item not found.");
    }
    if (movementType === "in" && !supplierName) {
      throw createHttpError(400, "Supplier is required for stock coming in.");
    }

    let driver = null;
    if (movementType === "out") {
      driver = this.selectActiveDriver(driverUserId);
      if (!driver) {
        throw createHttpError(400, "Driver is required for stock going out.");
      }
    }

    const existingEffect = existing.movement_type === "in" ? Number(existing.quantity) : -Number(existing.quantity);
    const newEffect = movementType === "in" ? quantity : -quantity;

    if (existing.stock_item_id === stockItemId) {
      const nextOnHand = this.getStockOnHand(existing.stock_item_id) - existingEffect + newEffect;
      if (nextOnHand < 0) {
        throw createHttpError(400, "Not enough stock on hand for that movement.");
      }
    } else {
      const oldItemOnHand = this.getStockOnHand(existing.stock_item_id) - existingEffect;
      if (oldItemOnHand < 0) {
        throw createHttpError(400, "Editing this movement would leave the original stock item below zero.");
      }
      const newItemOnHand = this.getStockOnHand(stockItemId) + newEffect;
      if (newItemOnHand < 0) {
        throw createHttpError(400, "Not enough stock on hand for that movement.");
      }
    }

    this.run(
      `
        update stock_movements
        set stock_item_id = ?,
            movement_type = ?,
            quantity = ?,
            supplier_name = ?,
            driver_user_id = ?,
            notes = ?
        where id = ?
      `,
      [
        stockItemId,
        movementType,
        quantity,
        movementType === "in" ? supplierName : "",
        movementType === "out" ? driver.id : null,
        notes,
        movementId,
      ],
    );

    return { ok: true };
  }

  createArtworkRequest(parameters = {}) {
    const actor = this.requireUser(parameters?.p_token);
    if (!["admin", "logistics"].includes(actor.role)) {
      throw createHttpError(400, "Permission denied");
    }

    const stockItemId = normalizeId(parameters?.p_stock_item_id, "Stock item not found.");
    const requestedQuantity = Math.trunc(Number(parameters?.p_requested_quantity || 0));
    const notes = normalizeOptionalText(parameters?.p_notes);
    const sentTo = normalizeOptionalText(parameters?.p_sent_to);

    if (!Number.isInteger(requestedQuantity) || requestedQuantity <= 0) {
      throw createHttpError(400, "Requested quantity must be greater than zero.");
    }
    if (!this.selectStockItemById(stockItemId)) {
      throw createHttpError(400, "Stock item not found.");
    }

    this.run(
      `
        insert into artwork_requests (
          id,
          stock_item_id,
          requested_quantity,
          notes,
          sent_to,
          requested_by_user_id,
          sent_at
        )
        values (?, ?, ?, ?, ?, ?, ?)
      `,
      [randomId(), stockItemId, requestedQuantity, notes, sentTo, actor.id, nowIso()],
    );

    return { ok: true };
  }

  createOrder(parameters = {}) {
    return this.withTransaction(() => {
      const actor = this.requireUser(parameters?.p_token);
      if (!["admin", "sales"].includes(actor.role)) {
        throw createHttpError(400, "Permission denied");
      }

      const driverUserId = normalizeOptionalText(parameters?.p_driver_user_id);
      const locationId = normalizeId(parameters?.p_location_id, "Location not found.");
      const entryType = normalizeEntryType(parameters?.p_entry_type);
      const quoteNumber = normalizeInhouseOrderNumber(parameters?.p_quote_number);
      const salesOrderNumber = normalizeOptionalText(parameters?.p_sales_order_number);
      const invoiceNumber = normalizeOptionalText(parameters?.p_invoice_number);
      const poNumber = normalizeOptionalText(parameters?.p_po_number);
      const branding = normalizeOptionalText(parameters?.p_branding);
      const stockDescription = normalizeRequiredText(parameters?.p_stock_description, "Stock description is required.");
      const deliveryAddress = normalizeOptionalText(parameters?.p_delivery_address);
      let deliveryLocationId = normalizeOptionalText(parameters?.p_delivery_location_id);
      const saveDeliveryLocation = Boolean(parameters?.p_save_delivery_location);
      const deliveryLocationName = normalizeOptionalText(parameters?.p_delivery_location_name);
      const scheduledFor = normalizeOptionalDate(parameters?.p_scheduled_for) || todayLocal();
      const priority = normalizePriority(parameters?.p_priority);
      const allowDuplicate = actor.role === "admin" && Boolean(parameters?.p_allow_duplicate);
      const notice = normalizeOptionalText(parameters?.p_notice);
      const moveToFactory = Boolean(parameters?.p_move_to_factory);
      let factoryDestinationLocationId = moveToFactory ? normalizeOptionalText(parameters?.p_factory_destination_location_id) : "";
      const stockItemNames = normalizeStockItemNames(parameters?.p_stock_item_names, stockDescription);

      if (!ORDER_ENTRY_TYPES.has(entryType)) {
        throw createHttpError(400, "Entry type must be collection or delivery.");
      }
      if (!ORDER_PRIORITIES.has(priority)) {
        throw createHttpError(400, "Choose a valid priority.");
      }
      if (compareDateOnly(scheduledFor, todayLocal()) < 0) {
        throw createHttpError(400, "Schedule date cannot be in the past.");
      }

      let driver = null;
      if (driverUserId) {
        driver = this.selectActiveDriver(driverUserId);
        if (!driver) {
          throw createHttpError(400, "Driver not found.");
        }
      }

      const location = this.selectLocationById(locationId);
      if (!location) {
        throw createHttpError(400, "Location not found.");
      }
      if (location.location_type === "client") {
        throw createHttpError(400, "Pickup location must be a supplier, factory, or both location.");
      }

      let finalDeliveryAddress = "";
      if (entryType === "delivery") {
        const deliveryDestination = this.resolveDeliveryLocationForOrder({
          actor,
          deliveryLocationId,
          deliveryAddress,
          saveDeliveryLocation,
          deliveryLocationName,
        });
        deliveryLocationId = deliveryDestination.deliveryLocationId;
        finalDeliveryAddress = deliveryDestination.deliveryAddress;
      } else {
        deliveryLocationId = null;
      }

      if (moveToFactory && entryType !== "collection") {
        throw createHttpError(400, "Only collection entries can be marked to move stock to a factory.");
      }
      if (moveToFactory && !factoryDestinationLocationId) {
        throw createHttpError(400, "Select which factory the collected stock should go to.");
      }
      if (moveToFactory) {
        const destination = this.selectLocationById(factoryDestinationLocationId);
        if (!destination || !["factory", "both"].includes(destination.location_type)) {
          throw createHttpError(400, "Factory destination not found.");
        }
      } else {
        factoryDestinationLocationId = "";
      }

      this.assertOrderAssignmentAllowed({
        actor,
        orderId: "",
        driverUserId: driver?.id || "",
        quoteNumber,
        locationId,
        scheduledFor,
        allowDuplicate,
      });

      const orderId = randomId();
      const orderNumber = this.nextOrderNumber();
      const now = nowIso();
      this.run(
        `
          insert into orders (
            id,
            order_number,
            customer_name,
            driver_user_id,
            location_id,
            entry_type,
            factory_order_number,
            inhouse_order_number,
            invoice_number,
            po_number,
            branding,
            stock_description,
            delivery_address,
            delivery_location_id,
            priority,
            notes,
            move_to_factory,
            factory_destination_location_id,
            status,
            scheduled_for,
            original_scheduled_for,
            carry_over_count,
            created_by_user_id,
            created_at,
            updated_at,
            completed_at,
            completion_type,
            completed_by_user_id,
            driver_flag_type,
            driver_flag_note,
            driver_flagged_at,
            driver_flagged_by_user_id,
            picked_up_at,
            picked_up_by_user_id
          )
          values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'active', ?, ?, 0, ?, ?, ?, null, null, null, null, '', null, null, null, null)
        `,
        [
          orderId,
          orderNumber,
          quoteNumber,
          driver ? driver.id : null,
          locationId,
          entryType,
          salesOrderNumber,
          quoteNumber,
          invoiceNumber,
          poNumber,
          branding,
          stockDescription,
          finalDeliveryAddress,
          deliveryLocationId,
          priority,
          notice,
          moveToFactory ? 1 : 0,
          moveToFactory ? factoryDestinationLocationId : null,
          scheduledFor,
          scheduledFor,
          actor.id,
          now,
          now,
        ],
      );

      let firstStockItemId = "";
      let stockItemCreatedCount = 0;
      const stockItemNotes = branding ? `Branding: ${branding}.` : "";
      stockItemNames.forEach((itemName) => {
        const existing = this.findStockItemByReference(itemName, quoteNumber, invoiceNumber, salesOrderNumber, poNumber);
        let stockItemId = existing?.id || "";

        if (!stockItemId) {
          const created = this.insertOrderStockItem({
            actorId: actor.id,
            name: itemName,
            quoteNumber,
            invoiceNumber,
            salesOrderNumber,
            poNumber,
            notes: stockItemNotes,
          });
          stockItemId = created.id;
          stockItemCreatedCount += 1;
        }

        if (!firstStockItemId) {
          firstStockItemId = stockItemId;
        }
      });

      return {
        ok: true,
        stockItemId: firstStockItemId || null,
        stockItemCreated: stockItemCreatedCount > 0,
        stockItemCount: stockItemNames.length,
      };
    });
  }

  updateOrder(parameters = {}) {
    return this.withTransaction(() => {
      const actor = this.requireUser(parameters?.p_token);
      if (!["admin", "sales"].includes(actor.role)) {
        throw createHttpError(400, "Permission denied");
      }

      const orderId = normalizeId(parameters?.p_order_id, "Order not found.");
      const existing = this.selectOrderById(orderId);
      if (!existing) {
        throw createHttpError(400, "Order not found.");
      }
      if (actor.role !== "admin" && existing.status === "completed") {
        throw createHttpError(400, "Only admins can edit completed entries.");
      }
      if (!ORDER_STATUSES.has(existing.status)) {
        throw createHttpError(400, "This entry cannot be edited.");
      }

      const driverUserId = normalizeOptionalText(parameters?.p_driver_user_id);
      const locationId = normalizeId(parameters?.p_location_id, "Location not found.");
      const entryType = normalizeEntryType(parameters?.p_entry_type);
      const quoteNumber = normalizeInhouseOrderNumber(parameters?.p_quote_number);
      const salesOrderNumber = normalizeOptionalText(parameters?.p_sales_order_number);
      const invoiceNumber = normalizeOptionalText(parameters?.p_invoice_number);
      const poNumber = normalizeOptionalText(parameters?.p_po_number);
      const branding = normalizeOptionalText(parameters?.p_branding);
      const stockDescription = normalizeRequiredText(parameters?.p_stock_description, "Stock description is required.");
      const deliveryAddress = normalizeOptionalText(parameters?.p_delivery_address);
      let deliveryLocationId = normalizeOptionalText(parameters?.p_delivery_location_id);
      const saveDeliveryLocation = Boolean(parameters?.p_save_delivery_location);
      const deliveryLocationName = normalizeOptionalText(parameters?.p_delivery_location_name);
      const scheduledFor = normalizeOptionalDate(parameters?.p_scheduled_for) || existing.scheduled_for || todayLocal();
      const priority = normalizePriority(parameters?.p_priority);
      const allowDuplicate = actor.role === "admin" && Boolean(parameters?.p_allow_duplicate);
      const notice = normalizeOptionalText(parameters?.p_notice);
      const moveToFactory = Boolean(parameters?.p_move_to_factory);
      let factoryDestinationLocationId = moveToFactory ? normalizeOptionalText(parameters?.p_factory_destination_location_id) : "";

      if (!ORDER_ENTRY_TYPES.has(entryType)) {
        throw createHttpError(400, "Entry type must be collection or delivery.");
      }
      if (existing.status === "active" && compareDateOnly(scheduledFor, todayLocal()) < 0) {
        throw createHttpError(400, "Schedule date cannot be in the past.");
      }
      if (!ORDER_PRIORITIES.has(priority)) {
        throw createHttpError(400, "Choose a valid priority.");
      }

      let driver = null;
      if (driverUserId) {
        driver = this.selectActiveDriver(driverUserId);
        if (!driver) {
          throw createHttpError(400, "Driver not found.");
        }
      }

      const location = this.selectLocationById(locationId);
      if (!location) {
        throw createHttpError(400, "Location not found.");
      }
      if (location.location_type === "client") {
        throw createHttpError(400, "Pickup location must be a supplier, factory, or both location.");
      }

      let finalDeliveryAddress = "";
      if (entryType === "delivery") {
        const deliveryDestination = this.resolveDeliveryLocationForOrder({
          actor,
          deliveryLocationId,
          deliveryAddress,
          saveDeliveryLocation,
          deliveryLocationName,
        });
        deliveryLocationId = deliveryDestination.deliveryLocationId;
        finalDeliveryAddress = deliveryDestination.deliveryAddress;
      } else {
        deliveryLocationId = null;
      }

      if (moveToFactory && entryType !== "collection") {
        throw createHttpError(400, "Only collection entries can be marked to move stock to a factory.");
      }
      if (moveToFactory && !factoryDestinationLocationId) {
        throw createHttpError(400, "Select which factory the collected stock should go to.");
      }
      if (moveToFactory) {
        const destination = this.selectLocationById(factoryDestinationLocationId);
        if (!destination || !["factory", "both"].includes(destination.location_type)) {
          throw createHttpError(400, "Factory destination not found.");
        }
      } else {
        factoryDestinationLocationId = "";
      }

      if (existing.status === "active") {
        this.assertOrderAssignmentAllowed({
          actor,
          orderId,
          driverUserId: driver?.id || "",
          quoteNumber,
          locationId,
          scheduledFor,
          allowDuplicate,
        });
      }

      const scheduleChanged = scheduledFor !== existing.scheduled_for;
      const nextPriority = (
        existing.status === "active"
        && existing.driver_user_id
        && !driverUserId
        && Number(existing.carry_over_count || 0) > 0
        && priority === "high"
      )
        ? "medium"
        : priority;

      this.run(
        `
          update orders
          set customer_name = ?,
              driver_user_id = ?,
              location_id = ?,
              entry_type = ?,
              factory_order_number = ?,
              inhouse_order_number = ?,
              invoice_number = ?,
              po_number = ?,
              branding = ?,
              stock_description = ?,
              delivery_address = ?,
              delivery_location_id = ?,
              priority = ?,
              notes = ?,
              move_to_factory = ?,
              factory_destination_location_id = ?,
              scheduled_for = ?,
              original_scheduled_for = ?,
              carry_over_count = ?,
              updated_at = ?
          where id = ?
        `,
        [
          quoteNumber,
          driver ? driver.id : null,
          locationId,
          entryType,
          salesOrderNumber,
          quoteNumber,
          invoiceNumber,
          poNumber,
          branding,
          stockDescription,
          finalDeliveryAddress,
          deliveryLocationId,
          nextPriority,
          notice,
          moveToFactory ? 1 : 0,
          moveToFactory ? factoryDestinationLocationId : null,
          scheduledFor,
          scheduleChanged ? scheduledFor : existing.original_scheduled_for,
          scheduleChanged ? 0 : Number(existing.carry_over_count || 0),
          nowIso(),
          orderId,
        ],
      );

      return { ok: true, orderId };
    });
  }

  assignOrder(parameters = {}) {
    return this.withTransaction(() => {
      const actor = this.requireUser(parameters?.p_token);
      if (!["admin", "sales", "driver"].includes(actor.role)) {
        throw createHttpError(400, "Permission denied");
      }

      const orderId = normalizeId(parameters?.p_order_id, "Order not found.");
      const order = this.selectOrderById(orderId);
      if (!order) {
        throw createHttpError(400, "Order not found.");
      }
      if (order.status !== "active") {
        throw createHttpError(400, "Only active entries can be reassigned.");
      }

      const driverUserId = normalizeOptionalText(parameters?.p_driver_user_id);
      const allowDuplicate = actor.role === "admin" && Boolean(parameters?.p_allow_duplicate);

      if (actor.role === "driver") {
        if (order.driver_user_id !== actor.id) {
          throw createHttpError(400, "Drivers can only transfer their own assigned entries.");
        }
        if (!driverUserId) {
          throw createHttpError(400, "Drivers must transfer the entry to another driver.");
        }
        if (driverUserId === actor.id) {
          throw createHttpError(400, "Choose another driver for the transfer.");
        }
      }

      let driver = null;
      if (driverUserId) {
        driver = this.selectActiveDriver(driverUserId);
        if (!driver) {
          throw createHttpError(400, "Driver not found.");
        }
        this.assertOrderAssignmentAllowed({
          actor,
          orderId,
          driverUserId: driver.id,
          quoteNumber: order.inhouse_order_number,
          locationId: order.location_id,
          scheduledFor: order.scheduled_for,
          allowDuplicate,
        });
      }

      const nextPriority = (
        order.driver_user_id
        && !driverUserId
        && Number(order.carry_over_count || 0) > 0
        && order.priority === "high"
      )
        ? "medium"
        : order.priority;

      this.run(
        `
          update orders
          set driver_user_id = ?,
              priority = ?,
              updated_at = ?
          where id = ?
        `,
        [driver ? driver.id : null, nextPriority, nowIso(), orderId],
      );

      return { ok: true };
    });
  }

  setOrderPriority(parameters = {}) {
    const actor = this.requireUser(parameters?.p_token);
    if (actor.role !== "admin") {
      throw createHttpError(400, "Permission denied");
    }

    const orderId = normalizeId(parameters?.p_order_id, "Order not found.");
    const priority = normalizePriority(parameters?.p_priority);
    const order = this.selectOrderById(orderId);
    if (!order) {
      throw createHttpError(400, "Order not found.");
    }
    if (order.status !== "active") {
      throw createHttpError(400, "Only active entries can change priority.");
    }
    if (!ORDER_PRIORITIES.has(priority)) {
      throw createHttpError(400, "Choose a valid priority.");
    }

    this.run(
      `
        update orders
        set priority = ?,
            updated_at = ?
        where id = ?
      `,
      [priority, nowIso(), orderId],
    );

    return { ok: true, priority };
  }

  setOrderFlag(parameters = {}) {
    const actor = this.requireUser(parameters?.p_token);
    if (!["admin", "driver"].includes(actor.role)) {
      throw createHttpError(400, "Permission denied");
    }

    const orderId = normalizeId(parameters?.p_order_id, "Order not found.");
    const order = this.selectOrderById(orderId);
    if (!order) {
      throw createHttpError(400, "Order not found.");
    }
    if (order.status !== "active") {
      throw createHttpError(400, "Only active entries can be flagged.");
    }
    if (actor.role === "driver" && order.driver_user_id !== actor.id) {
      throw createHttpError(400, "Drivers can only flag their own assigned orders.");
    }

    const rawFlagType = normalizeOptionalText(parameters?.p_flag_type).toLowerCase();
    const flagType = rawFlagType || null;
    if (flagType && !ORDER_FLAG_TYPES.has(flagType)) {
      throw createHttpError(400, "Choose a valid follow-up reason.");
    }

    const note = normalizeOptionalText(parameters?.p_note);
    const timestamp = flagType ? nowIso() : null;
    this.run(
      `
        update orders
        set driver_flag_type = ?,
            driver_flag_note = ?,
            driver_flagged_at = ?,
            driver_flagged_by_user_id = ?,
            updated_at = ?
        where id = ?
      `,
      [flagType, flagType ? note : "", timestamp, flagType ? actor.id : null, nowIso(), orderId],
    );

    return { ok: true, flagType: flagType || "" };
  }

  pickUpOrder(parameters = {}) {
    const actor = this.requireUser(parameters?.p_token);
    if (!["admin", "driver"].includes(actor.role)) {
      throw createHttpError(400, "Permission denied");
    }

    const orderId = normalizeId(parameters?.p_order_id, "Order not found.");
    const order = this.selectOrderById(orderId);
    if (!order) {
      throw createHttpError(400, "Order not found.");
    }
    if (actor.role === "driver" && order.driver_user_id !== actor.id) {
      throw createHttpError(400, "Drivers can only pick up their own assigned orders.");
    }
    if (order.status !== "active") {
      throw createHttpError(400, "Only active entries can be marked as picked up.");
    }
    if (order.picked_up_at) {
      return { ok: true };
    }

    this.run(
      `
        update orders
        set picked_up_at = ?,
            picked_up_by_user_id = ?,
            updated_at = ?
        where id = ?
          and status = 'active'
          and picked_up_at is null
      `,
      [nowIso(), actor.id, nowIso(), orderId],
    );

    return { ok: true, pickedUpBy: actor.id };
  }

  completeOrder(parameters = {}) {
    const actor = this.requireUser(parameters?.p_token);
    if (!["admin", "driver"].includes(actor.role)) {
      throw createHttpError(400, "Permission denied");
    }

    const orderId = normalizeId(parameters?.p_order_id, "Order not found.");
    const order = this.selectOrderById(orderId);
    if (!order) {
      throw createHttpError(400, "Order not found.");
    }
    if (actor.role === "driver" && order.driver_user_id !== actor.id) {
      throw createHttpError(400, "Drivers can only complete their own assigned orders.");
    }
    if (order.status === "completed") {
      return { ok: true };
    }

    const completionType = normalizeCompletionType(parameters?.p_completion_type) || "office";
    if (!ORDER_COMPLETION_TYPES.has(completionType)) {
      throw createHttpError(400, "Choose a valid completion action.");
    }
    if (completionType === "factory" && !Boolean(order.move_to_factory)) {
      throw createHttpError(400, "Only factory-transfer entries can be marked as dropped at the factory.");
    }
    if (!order.picked_up_at) {
      throw createHttpError(400, "Mark the entry as picked up before dropping it off.");
    }

    const now = nowIso();
    this.run(
      `
        update orders
        set status = 'completed',
            completed_at = ?,
            completion_type = ?,
            completed_by_user_id = ?,
            driver_flag_type = null,
            driver_flag_note = '',
            driver_flagged_at = null,
            driver_flagged_by_user_id = null,
            updated_at = ?
        where id = ?
          and status <> 'completed'
      `,
      [now, completionType, actor.id, now, orderId],
    );

    return { ok: true, completionType };
  }

  deleteOrder(parameters = {}) {
    return this.withTransaction(() => {
      const actor = this.requireUser(parameters?.p_token, ["admin"]);
      const orderId = normalizeId(parameters?.p_order_id, "Order not found.");
      const order = this.selectOrderRowForLog(orderId);
      if (!order) {
        throw createHttpError(400, "Order not found.");
      }

      this.run(
        `
          insert into order_delete_log (
            id,
            order_id,
            order_number,
            reference,
            quote_number,
            sales_order_number,
            invoice_number,
            po_number,
            entry_type,
            priority,
            delivery_address,
            delivery_location_id,
            delivery_location_name,
            delivery_location_address,
            branding,
            stock_description,
            notes,
            move_to_factory,
            factory_destination_location_id,
            factory_destination_name,
            factory_destination_address,
            location_id,
            location_name,
            location_address,
            driver_user_id,
            driver_name,
            created_by_user_id,
            created_by_name,
            created_at,
            scheduled_for,
            original_scheduled_for,
            carry_over_count,
            status,
            deleted_by_user_id,
            deleted_by_name,
            deleted_by_role,
            deleted_at,
            notification_attempts,
            last_notification_error,
            notification_sent_at
          )
          values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, '', null)
        `,
        [
          randomId(),
          order.id,
          order.orderNumber,
          order.reference,
          order.quoteNumber,
          order.salesOrderNumber,
          order.invoiceNumber,
          order.poNumber,
          order.entryType,
          order.priority,
          order.deliveryAddress,
          order.deliveryLocationId,
          order.deliveryLocationName,
          order.deliveryLocationAddress,
          order.branding,
          order.stockDescription,
          order.notes,
          order.moveToFactory ? 1 : 0,
          order.factoryDestinationLocationId,
          order.factoryDestinationName,
          order.factoryDestinationAddress,
          order.locationId,
          order.locationName,
          order.locationAddress,
          order.driverUserId,
          order.driverName,
          order.createdByUserId,
          order.createdByName,
          order.createdAt,
          order.scheduledFor,
          order.originalScheduledFor,
          Number(order.carryOverCount || 0),
          order.status,
          actor.id,
          actor.name || "",
          actor.role || "admin",
          nowIso(),
        ],
      );

      this.run("delete from orders where id = ?", [orderId]);
      return { ok: true, deletedBy: actor.id };
    });
  }

  selectSnapshotUsers(actor) {
    if (actor.role === "admin") {
      const rows = this.all("select * from app_users order by lower(name), created_at");
      return rows.map((row) => this.buildUserJson(row));
    }
    if (["sales", "logistics", "driver"].includes(actor.role)) {
      const rows = this.all(
        `
          select *
          from app_users
          where role = 'driver'
            and active = 1
          order by lower(name), created_at
        `,
      );
      return rows.map((row) => this.buildUserJson(row));
    }
    return [];
  }

  selectSnapshotSuppliers(actor) {
    if (["admin", "sales"].includes(actor.role)) {
      const rows = this.all("select * from suppliers order by lower(name), created_at");
      return rows.map((row) => this.buildSupplierJson(row));
    }
    if (actor.role !== "driver") {
      return [];
    }
    const rows = this.all(
      `
        select distinct s.*
        from suppliers s
        where exists (
          select 1
          from locations l
          join orders o on o.location_id = l.id
          where l.supplier_id = s.id
            and o.driver_user_id = ?
        )
        order by lower(s.name), s.created_at
      `,
      [actor.id],
    );
    return rows.map((row) => this.buildSupplierJson(row));
  }

  selectSnapshotLocations(actor) {
    if (["admin", "sales"].includes(actor.role)) {
      const rows = this.all("select * from locations order by lower(name), created_at");
      return rows.map((row) => this.buildLocationJson(row));
    }
    if (actor.role !== "driver") {
      return [];
    }
    const rows = this.all(
      `
        select distinct l.*
        from locations l
        join orders o on o.location_id = l.id
        where o.driver_user_id = ?
        order by lower(l.name), l.created_at
      `,
      [actor.id],
    );
    return rows.map((row) => this.buildLocationJson(row));
  }

  selectSnapshotOrders(actor) {
    if (["admin", "sales"].includes(actor.role)) {
      return this.selectOrderRows();
    }
    if (actor.role !== "driver") {
      return [];
    }
    return this.selectOrderRows("where o.driver_user_id = ?", [actor.id]);
  }

  selectSnapshotStockItems(actor) {
    if (!["admin", "logistics", "sales"].includes(actor.role)) {
      return [];
    }
    const rows = this.all(
      `
        select
          s.*,
          coalesce((
            select sum(case when m.movement_type = 'in' then m.quantity else -m.quantity end)
            from stock_movements m
            where m.stock_item_id = s.id
          ), 0) as on_hand_quantity
        from stock_items s
        order by lower(s.name), s.created_at
      `,
    );
    return rows.map((row) => this.buildStockItemJson(row));
  }

  selectSnapshotStockMovements(actor) {
    if (!["admin", "logistics", "sales"].includes(actor.role)) {
      return [];
    }
    const rows = this.all(
      `
        select
          m.id,
          m.stock_item_id as stockItemId,
          s.name as itemName,
          s.sku,
          s.quote_number as quoteNumber,
          s.invoice_number as invoiceNumber,
          s.sales_order_number as salesOrderNumber,
          s.po_number as poNumber,
          s.unit,
          m.movement_type as movementType,
          m.quantity,
          m.supplier_name as supplierName,
          m.driver_user_id as driverUserId,
          coalesce(d.name, '') as driverName,
          m.notes,
          m.created_by_user_id as createdByUserId,
          c.name as createdByName,
          m.created_at as createdAt
        from stock_movements m
        join stock_items s on s.id = m.stock_item_id
        left join app_users d on d.id = m.driver_user_id
        join app_users c on c.id = m.created_by_user_id
        order by m.created_at desc
      `,
    );
    return rows.map((row) => this.buildStockMovementJson(row));
  }

  selectSnapshotArtworkRequests(actor) {
    if (!["admin", "logistics"].includes(actor.role)) {
      return [];
    }
    const rows = this.all(
      `
        select
          r.id,
          r.stock_item_id as stockItemId,
          s.name as itemName,
          s.sku,
          s.quote_number as quoteNumber,
          s.invoice_number as invoiceNumber,
          s.sales_order_number as salesOrderNumber,
          s.po_number as poNumber,
          r.requested_quantity as requestedQuantity,
          r.notes,
          r.sent_to as sentTo,
          r.requested_by_user_id as requestedByUserId,
          u.name as requestedByName,
          r.sent_at as sentAt
        from artwork_requests r
        join stock_items s on s.id = r.stock_item_id
        join app_users u on u.id = r.requested_by_user_id
        order by r.sent_at desc
      `,
    );
    return rows.map((row) => this.buildArtworkRequestJson(row));
  }

  selectOrderRows(whereSql = "", parameters = []) {
    const rows = this.all(
      `
        select
          o.id,
          'ORD-' || o.order_number as reference,
          o.order_number as orderNumber,
          o.entry_type as entryType,
          o.move_to_factory as moveToFactory,
          o.factory_destination_location_id as factoryDestinationLocationId,
          coalesce(f.name, '') as factoryDestinationName,
          coalesce(f.address, '') as factoryDestinationAddress,
          o.factory_order_number as factoryOrderNumber,
          o.inhouse_order_number as inhouseOrderNumber,
          o.factory_order_number as salesOrderNumber,
          o.inhouse_order_number as quoteNumber,
          o.invoice_number as invoiceNumber,
          o.po_number as poNumber,
          o.delivery_address as deliveryAddress,
          o.delivery_location_id as deliveryLocationId,
          coalesce(dl.name, '') as deliveryLocationName,
          coalesce(dl.address, '') as deliveryLocationAddress,
          dl.lat as deliveryLocationLat,
          dl.lng as deliveryLocationLng,
          o.branding,
          o.stock_description as stockDescription,
          o.customer_name as customerName,
          o.driver_user_id as driverUserId,
          coalesce(d.name, '') as driverName,
          o.location_id as locationId,
          l.name as locationName,
          l.address as locationAddress,
          l.location_type as locationType,
          l.contact_person as locationContactPerson,
          l.contact_number as locationContactNumber,
          o.priority,
          o.notes,
          o.driver_flag_type as driverFlagType,
          o.driver_flag_note as driverFlagNote,
          o.driver_flagged_at as driverFlaggedAt,
          o.driver_flagged_by_user_id as driverFlaggedByUserId,
          coalesce(g.name, '') as driverFlaggedByName,
          o.picked_up_at as pickedUpAt,
          o.picked_up_by_user_id as pickedUpByUserId,
          coalesce(i.name, '') as pickedUpByName,
          o.completion_type as completionType,
          o.completed_by_user_id as completedByUserId,
          coalesce(h.name, '') as completedByName,
          o.status,
          o.scheduled_for as scheduledFor,
          o.original_scheduled_for as originalScheduledFor,
          o.carry_over_count as carryOverCount,
          o.created_by_user_id as createdByUserId,
          c.name as createdByName,
          c.role as createdByRole,
          o.created_at as createdAt,
          o.completed_at as completedAt
        from orders o
        left join app_users d on d.id = o.driver_user_id
        join locations l on l.id = o.location_id
        left join locations f on f.id = o.factory_destination_location_id
        left join locations dl on dl.id = o.delivery_location_id
        left join app_users g on g.id = o.driver_flagged_by_user_id
        left join app_users h on h.id = o.completed_by_user_id
        left join app_users i on i.id = o.picked_up_by_user_id
        join app_users c on c.id = o.created_by_user_id
        ${whereSql}
        order by
          case when o.status = 'active' then 0 else 1 end,
          o.created_at desc,
          o.order_number desc
      `,
      parameters,
    );
    return rows.map((row) => this.buildOrderJson(row));
  }

  selectOrderRowForLog(orderId) {
    return this.selectOrderRows("where o.id = ?", [orderId])[0] || null;
  }

  buildUserJson(row) {
    if (!row) {
      return null;
    }
    return {
      id: row.id,
      name: row.name,
      role: row.role,
      active: Boolean(row.active),
      phone: row.phone || "",
      vehicle: row.vehicle || "",
      lastKnownLat: row.last_known_lat ?? null,
      lastKnownLng: row.last_known_lng ?? null,
      lastKnownRecordedAt: row.last_known_recorded_at || null,
      createdAt: row.created_at || null,
    };
  }

  buildSupplierJson(row) {
    return {
      id: row.id,
      name: row.name,
      contactPerson: row.contact_person || "",
      contactNumber: row.contact_number || "",
      factory: Boolean(row.factory),
      createdAt: row.created_at || null,
    };
  }

  buildLocationJson(row) {
    return {
      id: row.id,
      supplierId: row.supplier_id || "",
      locationType: row.location_type || "supplier",
      name: row.name || "",
      address: row.address || "",
      lat: row.lat ?? null,
      lng: row.lng ?? null,
      contactPerson: row.contact_person || "",
      contactNumber: row.contact_number || "",
      notes: row.notes || "",
      createdAt: row.created_at || null,
    };
  }

  buildOrderJson(row) {
    return {
      id: row.id,
      reference: row.reference || "",
      orderNumber: row.orderNumber ?? null,
      entryType: row.entryType || "delivery",
      moveToFactory: Boolean(row.moveToFactory),
      factoryDestinationLocationId: row.factoryDestinationLocationId || null,
      factoryDestinationName: row.factoryDestinationName || "",
      factoryDestinationAddress: row.factoryDestinationAddress || "",
      factoryOrderNumber: row.factoryOrderNumber || "",
      inhouseOrderNumber: row.inhouseOrderNumber || "",
      salesOrderNumber: row.salesOrderNumber || "",
      quoteNumber: row.quoteNumber || "",
      invoiceNumber: row.invoiceNumber || "",
      poNumber: row.poNumber || "",
      deliveryAddress: row.deliveryAddress || "",
      deliveryLocationId: row.deliveryLocationId || "",
      deliveryLocationName: row.deliveryLocationName || "",
      deliveryLocationAddress: row.deliveryLocationAddress || "",
      deliveryLocationLat: row.deliveryLocationLat ?? null,
      deliveryLocationLng: row.deliveryLocationLng ?? null,
      branding: row.branding || "",
      stockDescription: row.stockDescription || "",
      customerName: row.customerName || "",
      driverUserId: row.driverUserId || "",
      driverName: row.driverName || "",
      locationId: row.locationId || "",
      locationName: row.locationName || "",
      locationAddress: row.locationAddress || "",
      locationType: row.locationType || "",
      locationContactPerson: row.locationContactPerson || "",
      locationContactNumber: row.locationContactNumber || "",
      priority: row.priority || "medium",
      notes: row.notes || "",
      driverFlagType: row.driverFlagType || "",
      driverFlagNote: row.driverFlagNote || "",
      driverFlaggedAt: row.driverFlaggedAt || null,
      driverFlaggedByUserId: row.driverFlaggedByUserId || "",
      driverFlaggedByName: row.driverFlaggedByName || "",
      pickedUpAt: row.pickedUpAt || null,
      pickedUpByUserId: row.pickedUpByUserId || "",
      pickedUpByName: row.pickedUpByName || "",
      completionType: row.completionType || "",
      completedByUserId: row.completedByUserId || "",
      completedByName: row.completedByName || "",
      status: row.status || "active",
      scheduledFor: row.scheduledFor || "",
      originalScheduledFor: row.originalScheduledFor || "",
      carryOverCount: Number(row.carryOverCount || 0),
      createdByUserId: row.createdByUserId || "",
      createdByName: row.createdByName || "",
      createdByRole: row.createdByRole || "",
      createdAt: row.createdAt || null,
      completedAt: row.completedAt || null,
    };
  }

  buildDeleteLogRow(row) {
    return {
      id: row.id,
      orderId: row.orderId || "",
      orderNumber: row.orderNumber ?? null,
      reference: row.reference || "",
      quoteNumber: row.quoteNumber || "",
      salesOrderNumber: row.salesOrderNumber || "",
      invoiceNumber: row.invoiceNumber || "",
      poNumber: row.poNumber || "",
      entryType: row.entryType || "delivery",
      priority: row.priority || "medium",
      deliveryAddress: row.deliveryAddress || "",
      deliveryLocationId: row.deliveryLocationId || "",
      deliveryLocationName: row.deliveryLocationName || "",
      deliveryLocationAddress: row.deliveryLocationAddress || "",
      branding: row.branding || "",
      stockDescription: row.stockDescription || "",
      notes: row.notes || "",
      moveToFactory: Boolean(row.moveToFactory),
      factoryDestinationLocationId: row.factoryDestinationLocationId || "",
      factoryDestinationName: row.factoryDestinationName || "",
      factoryDestinationAddress: row.factoryDestinationAddress || "",
      locationId: row.locationId || "",
      locationName: row.locationName || "",
      locationAddress: row.locationAddress || "",
      driverUserId: row.driverUserId || "",
      driverName: row.driverName || "",
      createdByUserId: row.createdByUserId || "",
      createdByName: row.createdByName || "",
      createdAt: row.createdAt || null,
      scheduledFor: row.scheduledFor || "",
      originalScheduledFor: row.originalScheduledFor || "",
      carryOverCount: Number(row.carryOverCount || 0),
      status: row.status || "active",
      deletedByUserId: row.deletedByUserId || "",
      deletedByName: row.deletedByName || "",
      deletedByRole: row.deletedByRole || "",
      deletedAt: row.deletedAt || null,
      notificationAttempts: Number(row.notificationAttempts || 0),
      lastNotificationError: row.lastNotificationError || "",
      notificationSentAt: row.notificationSentAt || null,
    };
  }

  buildStockItemJson(row) {
    return {
      id: row.id,
      name: row.name || "",
      sku: row.sku || "",
      quoteNumber: row.quote_number || row.quoteNumber || "",
      invoiceNumber: row.invoice_number || row.invoiceNumber || "",
      salesOrderNumber: row.sales_order_number || row.salesOrderNumber || "",
      poNumber: row.po_number || row.poNumber || "",
      unit: row.unit || "units",
      notes: row.notes || "",
      createdSource: row.created_source || row.createdSource || "manual",
      onHandQuantity: Number(row.on_hand_quantity ?? row.onHandQuantity ?? 0),
      createdAt: row.created_at || row.createdAt || null,
      updatedAt: row.updated_at || row.updatedAt || null,
    };
  }

  buildStockMovementJson(row) {
    return {
      id: row.id,
      stockItemId: row.stockItemId || row.stock_item_id || "",
      itemName: row.itemName || "",
      sku: row.sku || "",
      quoteNumber: row.quoteNumber || "",
      invoiceNumber: row.invoiceNumber || "",
      salesOrderNumber: row.salesOrderNumber || "",
      poNumber: row.poNumber || "",
      unit: row.unit || "units",
      movementType: row.movementType || row.movement_type || "in",
      quantity: Number(row.quantity || 0),
      supplierName: row.supplierName || "",
      driverUserId: row.driverUserId || "",
      driverName: row.driverName || "",
      notes: row.notes || "",
      createdByUserId: row.createdByUserId || "",
      createdByName: row.createdByName || "",
      createdAt: row.createdAt || row.created_at || null,
    };
  }

  buildArtworkRequestJson(row) {
    return {
      id: row.id,
      stockItemId: row.stockItemId || "",
      itemName: row.itemName || "",
      sku: row.sku || "",
      quoteNumber: row.quoteNumber || "",
      invoiceNumber: row.invoiceNumber || "",
      salesOrderNumber: row.salesOrderNumber || "",
      poNumber: row.poNumber || "",
      requestedQuantity: Number(row.requestedQuantity || 0),
      notes: row.notes || "",
      sentTo: row.sentTo || "",
      requestedByUserId: row.requestedByUserId || "",
      requestedByName: row.requestedByName || "",
      sentAt: row.sentAt || null,
    };
  }

  selectUserById(userId) {
    return this.get("select * from app_users where id = ? limit 1", [userId]);
  }

  selectLocationById(locationId) {
    return this.get("select * from locations where id = ? limit 1", [locationId]);
  }

  selectStockItemById(stockItemId) {
    return this.get("select * from stock_items where id = ? limit 1", [stockItemId]);
  }

  selectStockMovementById(movementId) {
    return this.get("select * from stock_movements where id = ? limit 1", [movementId]);
  }

  selectOrderById(orderId) {
    return this.get("select * from orders where id = ? limit 1", [orderId]);
  }

  selectActiveDriver(userId) {
    const cleanId = normalizeOptionalText(userId);
    if (!cleanId) {
      return null;
    }
    return this.get(
      `
        select *
        from app_users
        where id = ?
          and role = 'driver'
          and active = 1
        limit 1
      `,
      [cleanId],
    );
  }

  findStockItemByReference(name, quoteNumber, invoiceNumber, salesOrderNumber, poNumber) {
    return this.get(
      `
        select *
        from stock_items
        where lower(trim(name)) = lower(trim(?))
          and lower(trim(quote_number)) = lower(trim(?))
          and lower(trim(invoice_number)) = lower(trim(?))
          and lower(trim(sales_order_number)) = lower(trim(?))
          and lower(trim(po_number)) = lower(trim(?))
        limit 1
      `,
      [name, quoteNumber, invoiceNumber, salesOrderNumber, poNumber],
    );
  }

  insertOrderStockItem({ actorId, name, quoteNumber, invoiceNumber, salesOrderNumber, poNumber, notes }) {
    const now = nowIso();
    const id = randomId();
    this.run(
      `
        insert into stock_items (
          id,
          name,
          sku,
          quote_number,
          invoice_number,
          sales_order_number,
          po_number,
          unit,
          notes,
          created_source,
          created_by_user_id,
          created_at,
          updated_at
        )
        values (?, ?, '', ?, ?, ?, ?, 'units', ?, 'order', ?, ?, ?)
      `,
      [id, name, quoteNumber, invoiceNumber, salesOrderNumber, poNumber, notes || "", actorId, now, now],
    );
    return { id };
  }

  assertOrderAssignmentAllowed({ actor, orderId, driverUserId, quoteNumber, locationId, scheduledFor, allowDuplicate }) {
    const duplicate = this.get(
      `
        select id
        from orders
        where id <> ?
          and lower(trim(inhouse_order_number)) = lower(trim(?))
          and location_id = ?
          and status = 'active'
        limit 1
      `,
      [orderId || "", quoteNumber, locationId],
    );
    if (duplicate) {
      throw createHttpError(400, "Duplicate blocked. This inhouse order number already has an active entry for that pickup location.");
    }

    if (!driverUserId) {
      return;
    }

    const completedStop = this.get(
      `
        select id
        from orders
        where id <> ?
          and driver_user_id = ?
          and location_id = ?
          and status = 'completed'
          and scheduled_for = ?
        limit 1
      `,
      [orderId || "", driverUserId, locationId, scheduledFor],
    );
    if (completedStop && !(actor.role === "admin" && allowDuplicate)) {
      throw createHttpError(400, "Completed stop blocked. This driver has already completed that pickup location on the scheduled date. Admin authorization is required to send them back.");
    }
  }

  getStockOnHand(stockItemId) {
    const row = this.get(
      `
        select coalesce(sum(case when movement_type = 'in' then quantity else -quantity end), 0) as on_hand
        from stock_movements
        where stock_item_id = ?
      `,
      [stockItemId],
    );
    return Number(row?.on_hand || 0);
  }

  issueSession(userId) {
    const now = Date.now();
    const createdAt = new Date(now).toISOString();
    const expiresAt = new Date(now + SESSION_TTL_MS).toISOString();
    const token = randomId();

    this.run("delete from app_sessions where user_id = ? and expires_at <= ?", [userId, createdAt]);
    this.run(
      `
        insert into app_sessions (
          token,
          user_id,
          created_at,
          last_seen_at,
          expires_at
        )
        values (?, ?, ?, ?, ?)
      `,
      [token, userId, createdAt, createdAt, expiresAt],
    );

    return token;
  }

  requireUser(token, roles = null) {
    const cleanToken = normalizeOptionalText(token);
    if (!cleanToken) {
      throw createHttpError(403, "Invalid session");
    }

    const now = Date.now();
    const nowText = new Date(now).toISOString();
    const user = this.get(
      `
        select u.*
        from app_sessions s
        join app_users u on u.id = s.user_id
        where s.token = ?
          and s.expires_at > ?
          and u.active = 1
        order by s.created_at desc
        limit 1
      `,
      [cleanToken, nowText],
    );

    if (!user) {
      throw createHttpError(403, "Invalid session");
    }
    if (Array.isArray(roles) && roles.length && !roles.includes(user.role)) {
      throw createHttpError(403, "Permission denied");
    }

    this.refreshSessionExpiry(cleanToken, now);
    return user;
  }

  refreshSessionExpiry(token, now = Date.now()) {
    const cleanToken = normalizeOptionalText(token);
    if (!cleanToken) {
      return;
    }

    const seenAt = new Date(now).toISOString();
    const expiresAt = new Date(now + SESSION_TTL_MS).toISOString();
    this.run(
      "update app_sessions set last_seen_at = ?, expires_at = ? where token = ?",
      [seenAt, expiresAt, cleanToken],
    );
  }

  ensureOfficeLocation(actorId) {
    const existing = this.get(
      `
        select id
        from locations
        where supplier_id is null
          and lower(trim(name)) = 'office'
        order by created_at asc
        limit 1
      `,
    );

    if (existing?.id) {
      return existing.id;
    }

    const now = nowIso();
    const locationId = randomId();
    this.run(
      `
        insert into locations (
          id,
          supplier_id,
          location_type,
          name,
          address,
          lat,
          lng,
          contact_person,
          contact_number,
          notes,
          created_by,
          created_at,
          updated_at
        )
        values (?, null, 'supplier', 'Office', 'Giftwrap Office', null, null, '', '', 'System pickup location for office-origin entries.', ?, ?, ?)
      `,
      [locationId, actorId, now, now],
    );
    return locationId;
  }

  rollForwardOpenOrders(today) {
    const carriedRows = this.all(
      `
        select *
        from orders
        where status = 'active'
          and driver_user_id is not null
          and scheduled_for < ?
        order by created_at desc, order_number desc
      `,
      [today],
    );

    if (!carriedRows.length) {
      return {
        today,
        updatedOrders: 0,
        carriedOrders: [],
      };
    }

    const carriedOrders = carriedRows.map((row) => {
      const nextCarryOverCount = Number(row.carry_over_count || 0) + Math.max(daysBetween(today, row.scheduled_for), 1);
      this.run(
        `
          update orders
          set scheduled_for = ?,
              priority = 'high',
              carry_over_count = ?,
              updated_at = ?
          where id = ?
        `,
        [today, nextCarryOverCount, nowIso(), row.id],
      );

      const updated = this.selectOrderRows("where o.id = ?", [row.id])[0];
      return {
        ...updated,
        previousScheduledFor: row.scheduled_for,
        scheduledFor: today,
        priority: "high",
        carryOverCount: nextCarryOverCount,
      };
    });

    carriedOrders.sort((left, right) => (
      String(left.driverName || "").localeCompare(String(right.driverName || ""), "en-ZA", { sensitivity: "base" })
      || String(left.locationName || "").localeCompare(String(right.locationName || ""), "en-ZA", { sensitivity: "base" })
      || Number(left.orderNumber || 0) - Number(right.orderNumber || 0)
    ));

    return {
      today,
      updatedOrders: carriedOrders.length,
      carriedOrders,
    };
  }

  nextOrderNumber() {
    const row = this.get("select coalesce(max(order_number), 0) + 1 as nextValue from orders");
    return Number(row?.nextValue || 1);
  }

  withTransaction(callback) {
    this.db.exec("BEGIN IMMEDIATE");
    try {
      const value = callback();
      this.db.exec("COMMIT");
      return value;
    } catch (error) {
      try {
        this.db.exec("ROLLBACK");
      } catch (rollbackError) {
        // Ignore rollback failures and preserve the original error.
      }
      throw error;
    }
  }

  prepare(sql) {
    if (!this.statementCache.has(sql)) {
      this.statementCache.set(sql, this.db.prepare(sql));
    }
    return this.statementCache.get(sql);
  }

  get(sql, parameters = []) {
    return this.prepare(sql).get(...parameters) || null;
  }

  all(sql, parameters = []) {
    return this.prepare(sql).all(...parameters) || [];
  }

  run(sql, parameters = []) {
    return this.prepare(sql).run(...parameters);
  }
}

function createHttpError(statusCode, message) {
  const error = new Error(String(message || "").trim() || "Request failed.");
  error.statusCode = statusCode;
  return error;
}

function normalizeErrorMessage(error) {
  if (!error) {
    return "Unknown error.";
  }
  if (typeof error === "string") {
    return error;
  }
  return String(error.message || error.details || error.hint || error.code || "Unknown error.");
}

function randomId() {
  return crypto.randomUUID();
}

function nowIso() {
  return new Date().toISOString();
}

function todayLocal() {
  return formatDateInTimeZone(new Date(), TIME_ZONE);
}

function formatDateInTimeZone(date, timeZone) {
  const formatter = new Intl.DateTimeFormat("en-CA", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  });

  const parts = formatter.formatToParts(date);
  const values = {};
  parts.forEach((part) => {
    if (part.type !== "literal") {
      values[part.type] = part.value;
    }
  });

  return `${values.year}-${values.month}-${values.day}`;
}

function getWeekStart(dateString) {
  const date = parseDateOnly(dateString);
  const utcDay = date.getUTCDay() || 7;
  return addDays(dateString, 1 - utcDay);
}

function addDays(dateString, days) {
  const date = parseDateOnly(dateString);
  date.setUTCDate(date.getUTCDate() + Number(days || 0));
  return formatDateOnly(date);
}

function compareDateOnly(left, right) {
  if (left === right) {
    return 0;
  }
  return left < right ? -1 : 1;
}

function daysBetween(laterDate, earlierDate) {
  const later = parseDateOnly(laterDate);
  const earlier = parseDateOnly(earlierDate);
  return Math.round((later.getTime() - earlier.getTime()) / (24 * 60 * 60 * 1000));
}

function parseDateOnly(value) {
  const normalized = normalizeOptionalDate(value);
  if (!normalized) {
    throw new Error("Invalid date.");
  }
  return new Date(`${normalized}T00:00:00.000Z`);
}

function formatDateOnly(date) {
  const year = date.getUTCFullYear();
  const month = String(date.getUTCMonth() + 1).padStart(2, "0");
  const day = String(date.getUTCDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function hashPassword(password) {
  const salt = crypto.randomBytes(16).toString("hex");
  const digest = crypto.scryptSync(password, salt, 64).toString("hex");
  return `scrypt:${salt}:${digest}`;
}

function passwordMatches(password, hash) {
  const stored = String(hash || "");
  if (/^\$2[aby]\$/.test(stored)) {
    return Boolean(bcrypt && bcrypt.compareSync(password, stored));
  }
  const parts = stored.split(":");
  if (parts.length !== 3 || parts[0] !== "scrypt") {
    return false;
  }
  const [, salt, expectedHex] = parts;
  const expected = Buffer.from(expectedHex, "hex");
  const actual = crypto.scryptSync(password, salt, expected.length);
  return expected.length === actual.length && crypto.timingSafeEqual(expected, actual);
}

function normalizeRequiredText(value, message) {
  const normalized = normalizeOptionalText(value);
  if (!normalized) {
    throw createHttpError(400, message);
  }
  return normalized;
}

function normalizeInhouseOrderNumber(value) {
  const normalized = normalizeRequiredText(value, "Inhouse order number is required.");
  const upperValue = normalized.toUpperCase();
  if (!INHOUSE_ORDER_PREFIXES.some((prefix) => upperValue.startsWith(prefix))) {
    throw createHttpError(400, `Inhouse order number must start with one of: ${INHOUSE_ORDER_PREFIX_LABEL}.`);
  }
  return normalized;
}

function normalizeOptionalText(value) {
  return String(value ?? "").trim();
}

function normalizeOptionalNumber(value) {
  if (value === null || value === undefined || String(value).trim() === "") {
    return null;
  }
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : Number.NaN;
}

function toFiniteNumber(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : NaN;
}

function normalizeOptionalDate(value) {
  const normalized = normalizeOptionalText(value);
  if (!normalized) {
    return "";
  }
  if (!/^\d{4}-\d{2}-\d{2}$/.test(normalized)) {
    throw createHttpError(400, "Invalid date.");
  }
  const date = new Date(`${normalized}T00:00:00.000Z`);
  if (Number.isNaN(date.getTime()) || formatDateOnly(date) !== normalized) {
    throw createHttpError(400, "Invalid date.");
  }
  return normalized;
}

function normalizeRole(value) {
  return normalizeOptionalText(value).toLowerCase();
}

function normalizeEmailDestinationText(value, label = "Email recipients") {
  const entries = splitEmailRecipientList(value);
  entries.forEach((entry) => {
    if (!looksLikeEmailAddress(entry)) {
      throw createHttpError(400, `${label} must contain valid email addresses.`);
    }
  });
  return entries.join(", ");
}

function splitEmailRecipientList(value) {
  if (Array.isArray(value)) {
    return value
      .map((entry) => normalizeOptionalText(entry))
      .filter(Boolean);
  }

  return String(value || "")
    .split(/[,\n;]+/)
    .map((entry) => normalizeOptionalText(entry))
    .filter(Boolean);
}

function looksLikeEmailAddress(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || "").trim());
}

function normalizeStoredBoolean(value, fallback = false) {
  const text = String(value || "").trim().toLowerCase();
  if (!text) {
    return fallback;
  }
  if (["true", "1", "yes", "on"].includes(text)) {
    return true;
  }
  if (["false", "0", "no", "off"].includes(text)) {
    return false;
  }
  return fallback;
}

function normalizeLocationType(value) {
  return normalizeOptionalText(value).toLowerCase();
}

function normalizeEntryType(value) {
  return normalizeOptionalText(value).toLowerCase();
}

function normalizePriority(value) {
  return normalizeOptionalText(value).toLowerCase() || "medium";
}

function normalizeCompletionType(value) {
  return normalizeOptionalText(value).toLowerCase();
}

function normalizeMovementType(value) {
  return normalizeOptionalText(value).toLowerCase();
}

function normalizeId(value, message) {
  const normalized = normalizeOptionalText(value);
  if (!normalized) {
    throw createHttpError(400, message);
  }
  return normalized;
}

function normalizeIdList(values) {
  if (!Array.isArray(values)) {
    return [];
  }
  return values.map((value) => normalizeOptionalText(value)).filter(Boolean);
}

function normalizeNameKey(value) {
  return normalizeOptionalText(value).toLowerCase();
}

function normalizeStockItemNames(value, fallback) {
  const source = Array.isArray(value) ? value : [fallback];
  const seen = new Set();
  const names = [];
  source.forEach((entry) => {
    const normalized = normalizeOptionalText(entry);
    const key = normalized.toLowerCase();
    if (!normalized || seen.has(key)) {
      return;
    }
    seen.add(key);
    names.push(normalized);
  });
  if (!names.length && fallback) {
    names.push(normalizeOptionalText(fallback));
  }
  return names.filter(Boolean);
}

function buildPlaceholders(count) {
  return Array.from({ length: count }, () => "?").join(", ");
}

function normalizeImportData(rawData) {
  const data = rawData && typeof rawData === "object" ? rawData : {};
  return {
    app_users: Array.isArray(data.app_users) ? data.app_users : [],
    suppliers: Array.isArray(data.suppliers) ? data.suppliers : [],
    locations: Array.isArray(data.locations) ? data.locations : [],
    orders: Array.isArray(data.orders) ? data.orders : [],
    order_delete_log: Array.isArray(data.order_delete_log) ? data.order_delete_log : [],
    stock_items: Array.isArray(data.stock_items) ? data.stock_items : [],
    stock_movements: Array.isArray(data.stock_movements) ? data.stock_movements : [],
    artwork_requests: Array.isArray(data.artwork_requests) ? data.artwork_requests : [],
    app_sessions: Array.isArray(data.app_sessions) ? data.app_sessions : [],
    app_settings: Array.isArray(data.app_settings) ? data.app_settings : [],
  };
}

function countImportRows(data) {
  return Object.fromEntries(
    Object.entries(normalizeImportData(data)).map(([tableName, rows]) => [tableName, rows.length]),
  );
}

function requireImportedText(value, message) {
  const normalized = importedNullableText(value);
  if (!normalized) {
    throw createHttpError(400, message);
  }
  return normalized;
}

function importedText(value, fallback = "") {
  if (value === null || value === undefined) {
    return fallback;
  }
  if (value instanceof Date) {
    return value.toISOString();
  }
  return String(value);
}

function importedNullableText(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }
  if (value instanceof Date) {
    return value.toISOString();
  }
  return String(value);
}

function importedTimestamp(value) {
  const normalized = importedNullableTimestamp(value);
  if (!normalized) {
    throw createHttpError(400, "Imported timestamp is required.");
  }
  return normalized;
}

function importedNullableTimestamp(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }
  if (value instanceof Date) {
    return value.toISOString();
  }

  const candidate = String(value);
  const parsed = new Date(candidate);
  if (!Number.isNaN(parsed.getTime()) && /[tT]/.test(candidate)) {
    return parsed.toISOString();
  }
  return candidate;
}

function importedDate(value) {
  const normalized = importedNullableDate(value);
  if (!normalized) {
    throw createHttpError(400, "Imported date is required.");
  }
  return normalized;
}

function importedNullableDate(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }
  if (value instanceof Date) {
    return formatDateOnly(value);
  }
  return normalizeOptionalDate(String(value));
}

function importedInteger(value, fallback = null) {
  if (value === null || value === undefined || value === "") {
    if (fallback === null) {
      throw createHttpError(400, "Imported integer is required.");
    }
    return fallback;
  }

  const number = Number(value);
  if (!Number.isInteger(number)) {
    throw createHttpError(400, "Imported integer is invalid.");
  }
  return number;
}

function importedNullableInteger(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }
  return importedInteger(value);
}

function importedNullableNumber(value) {
  if (value === null || value === undefined || value === "") {
    return null;
  }
  const number = Number(value);
  if (!Number.isFinite(number)) {
    throw createHttpError(400, "Imported number is invalid.");
  }
  return number;
}

function importedBoolean(value) {
  if (value === true || value === 1 || value === "1" || value === "true" || value === "t") {
    return 1;
  }
  return 0;
}

module.exports = {
  createLocalDatabase,
};
