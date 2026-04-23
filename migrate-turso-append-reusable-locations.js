const fs = require("fs");
const path = require("path");
const { createClient } = require("@libsql/client");
const { loadProjectEnv } = require("./load-project-env");

const ROOT_DIR = __dirname;
const TURSO_DATABASE_URL_ENV_KEYS = ["TURSO_DATABASE_URL", "LIBSQL_DATABASE_URL"];
const TURSO_AUTH_TOKEN_ENV_KEYS = ["TURSO_AUTH_TOKEN", "LIBSQL_AUTH_TOKEN"];
const TURSO_STATE_TABLE_NAME = "logistics_center_state";
const TURSO_STATE_ID = "route-ledger";

loadProjectEnv(ROOT_DIR);

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) {
    printUsage();
    return;
  }

  if (args["env-file"]) {
    loadProjectEnv(ROOT_DIR, {
      filenames: [String(args["env-file"])],
    });
  }

  if (args.url) {
    process.env.TURSO_DATABASE_URL = String(args.url).trim();
  }
  if (args.token) {
    process.env.TURSO_AUTH_TOKEN = String(args.token).trim();
  }

  const databaseUrl = getFirstConfiguredEnvironmentValue(TURSO_DATABASE_URL_ENV_KEYS);
  const authToken = getFirstConfiguredEnvironmentValue(TURSO_AUTH_TOKEN_ENV_KEYS);
  if (!databaseUrl) {
    printUsage();
    throw new Error("Set TURSO_DATABASE_URL or pass --url before running the migration.");
  }
  if (isRemoteTursoUrl(databaseUrl.value) && !authToken) {
    printUsage();
    throw new Error("Set TURSO_AUTH_TOKEN or pass --token for a Turso Cloud database.");
  }

  const client = createClient({
    url: databaseUrl.value,
    authToken: authToken?.value || undefined,
  });

  try {
    const row = await readStateRow(client);
    const backupPath = resolveBackupPath(args["backup-file"]);
    writeBackup(backupPath, {
      exportedAt: new Date().toISOString(),
      target: redactTursoUrl(databaseUrl.value),
      row,
    });

    const originalData = parseStateData(row.data);
    const migration = appendReusableLocationFields(originalData);
    const counts = safeParseJson(row.counts, {});
    const shouldApply = Boolean(args.apply);
    const updatedAt = new Date().toISOString();

    if (shouldApply && migration.changed) {
      await client.execute({
        sql: `
          update ${TURSO_STATE_TABLE_NAME}
          set data = ?,
              counts = ?,
              source = ?,
              updated_at = ?
          where id = ?
        `,
        args: [
          JSON.stringify(migration.data),
          JSON.stringify(counts),
          "append reusable delivery location schema",
          updatedAt,
          TURSO_STATE_ID,
        ],
      });
    }

    console.log(JSON.stringify({
      ok: true,
      applied: shouldApply && migration.changed,
      dryRun: !shouldApply,
      changed: migration.changed,
      target: redactTursoUrl(databaseUrl.value),
      backupPath,
      counts,
      changes: migration.changes,
      nextStep: !shouldApply
        ? "Review the backup and rerun with --apply to write the additive change."
        : "Deploy/restart the app so it reads the appended Turso state.",
    }, null, 2));
  } finally {
    if (typeof client.close === "function") {
      client.close();
    }
  }
}

async function readStateRow(client) {
  const result = await client.execute({
    sql: `
      select id, data, counts, source, updated_at
      from ${TURSO_STATE_TABLE_NAME}
      where id = ?
      limit 1
    `,
    args: [TURSO_STATE_ID],
  });

  const row = Array.isArray(result.rows) ? result.rows[0] : null;
  if (!row?.data) {
    throw new Error(`No ${TURSO_STATE_ID} row was found in ${TURSO_STATE_TABLE_NAME}. Refusing to create or seed live data.`);
  }

  return {
    id: String(row.id || TURSO_STATE_ID),
    data: String(row.data || ""),
    counts: String(row.counts || "{}"),
    source: String(row.source || ""),
    updated_at: String(row.updated_at || ""),
  };
}

function appendReusableLocationFields(data) {
  const nextData = cloneJson(data);
  const changes = {
    ordersDeliveryLocationIdAdded: 0,
    orderDeleteLogDeliveryLocationIdAdded: 0,
    orderDeleteLogDeliveryLocationNameAdded: 0,
    orderDeleteLogDeliveryLocationAddressAdded: 0,
  };

  if (Array.isArray(nextData.orders)) {
    nextData.orders = nextData.orders.map((row) => {
      const nextRow = { ...row };
      if (!Object.prototype.hasOwnProperty.call(nextRow, "delivery_location_id")) {
        nextRow.delivery_location_id = null;
        changes.ordersDeliveryLocationIdAdded += 1;
      }
      return nextRow;
    });
  }

  if (Array.isArray(nextData.order_delete_log)) {
    nextData.order_delete_log = nextData.order_delete_log.map((row) => {
      const nextRow = { ...row };
      if (!Object.prototype.hasOwnProperty.call(nextRow, "delivery_location_id")) {
        nextRow.delivery_location_id = null;
        changes.orderDeleteLogDeliveryLocationIdAdded += 1;
      }
      if (!Object.prototype.hasOwnProperty.call(nextRow, "delivery_location_name")) {
        nextRow.delivery_location_name = "";
        changes.orderDeleteLogDeliveryLocationNameAdded += 1;
      }
      if (!Object.prototype.hasOwnProperty.call(nextRow, "delivery_location_address")) {
        nextRow.delivery_location_address = "";
        changes.orderDeleteLogDeliveryLocationAddressAdded += 1;
      }
      return nextRow;
    });
  }

  const changed = Object.values(changes).some((count) => count > 0);
  return {
    changed,
    changes,
    data: nextData,
  };
}

function parseStateData(rawData) {
  const data = safeParseJson(rawData, null);
  if (!data || typeof data !== "object" || Array.isArray(data)) {
    throw new Error("The Turso state row does not contain a valid exported data object.");
  }
  return data;
}

function cloneJson(value) {
  return JSON.parse(JSON.stringify(value));
}

function safeParseJson(rawValue, fallbackValue) {
  try {
    return JSON.parse(String(rawValue || ""));
  } catch (error) {
    return fallbackValue;
  }
}

function resolveBackupPath(value) {
  const rawValue = String(value || "").trim();
  if (rawValue) {
    return path.resolve(ROOT_DIR, rawValue);
  }

  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  return path.join(ROOT_DIR, "data", `turso-state-before-reusable-locations-${stamp}.json`);
}

function writeBackup(backupPath, payload) {
  fs.mkdirSync(path.dirname(backupPath), { recursive: true });
  fs.writeFileSync(backupPath, JSON.stringify(payload, null, 2), "utf8");
}

function getFirstConfiguredEnvironmentValue(keys) {
  for (const key of keys) {
    const value = String(process.env[key] || "").trim();
    if (value) {
      return { key, value };
    }
  }
  return null;
}

function isRemoteTursoUrl(value) {
  return /^(libsql|https|http):\/\//i.test(String(value || "").trim());
}

function redactTursoUrl(value) {
  const raw = String(value || "").trim();
  if (!raw) {
    return "";
  }
  return raw.replace(/\/\/([^:@/]+):([^@/]+)@/, "//***:***@");
}

function parseArgs(argv) {
  const args = {};
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === "--help" || arg === "-h") {
      args.help = true;
      continue;
    }
    if (arg === "--apply") {
      args.apply = true;
      continue;
    }
    if (!arg.startsWith("--")) {
      continue;
    }

    const key = arg.slice(2);
    const next = argv[index + 1];
    if (!next || next.startsWith("--")) {
      args[key] = true;
      continue;
    }
    args[key] = next;
    index += 1;
  }
  return args;
}

function printUsage() {
  console.log(`
Usage:
  node migrate-turso-append-reusable-locations.js
  node migrate-turso-append-reusable-locations.js --env-file .env.local
  node migrate-turso-append-reusable-locations.js --apply

Environment:
  TURSO_DATABASE_URL   Required target Turso database URL.
  TURSO_AUTH_TOKEN     Required for Turso Cloud targets.

Options:
  --apply              Write the additive schema fields back to Turso. Without this, the script is a dry run.
  --url <value>        Target Turso database URL. Prefer env vars for production.
  --token <value>      Target Turso auth token. Prefer env vars or .env.local.
  --env-file <path>    Load an extra env file before migrating.
  --backup-file <path> Write the pre-migration state backup to a custom path.

This script does not call replaceAllData and does not copy local data into Turso.
It only updates the existing logistics_center_state row after writing a backup.
`.trim());
}

main().catch((error) => {
  console.error(error?.message || error);
  process.exitCode = 1;
});
