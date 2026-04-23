const { loadProjectEnv } = require("./load-project-env");
const { createLocalDatabase } = require("./local-database");

const ROOT_DIR = __dirname;
const TURSO_DATABASE_URL_ENV_KEYS = ["TURSO_DATABASE_URL", "LIBSQL_DATABASE_URL"];
const TURSO_AUTH_TOKEN_ENV_KEYS = ["TURSO_AUTH_TOKEN", "LIBSQL_AUTH_TOKEN"];
const TURSO_ENV_KEYS = [...TURSO_DATABASE_URL_ENV_KEYS, ...TURSO_AUTH_TOKEN_ENV_KEYS];

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

  const tursoEnv = captureTursoEnvironment();
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

  const localSnapshot = await exportLocalSnapshot();
  if (args["dry-run"]) {
    console.log(JSON.stringify({
      ok: true,
      dryRun: true,
      target: redactTursoUrl(databaseUrl.value),
      localCounts: localSnapshot.counts,
    }, null, 2));
    return;
  }

  restoreTursoEnvironment(tursoEnv);
  const tursoDatabase = createLocalDatabase(ROOT_DIR);
  const status = tursoDatabase.getStatus();
  if (!status.configured) {
    throw new Error(status.reason || "Turso database is not available.");
  }
  if (status.storage !== "turso-cloud") {
    throw new Error("TURSO_DATABASE_URL is not active; refusing to overwrite a non-Turso database.");
  }

  try {
    const result = await tursoDatabase.replaceAllData(localSnapshot.data, {
      source: String(args.source || "manual local-to-turso migration").trim() || "manual local-to-turso migration",
    });
    console.log(JSON.stringify({
      ok: true,
      target: redactTursoUrl(databaseUrl.value),
      storage: status.storage,
      tursoConnected: true,
      localCounts: localSnapshot.counts,
      cloudCounts: tursoDatabase.getTableCounts(),
      importedCounts: result.localCounts,
    }, null, 2));
  } finally {
    await tursoDatabase.close();
  }
}

async function exportLocalSnapshot() {
  const tursoEnv = captureTursoEnvironment();
  clearTursoEnvironment();

  const localDatabase = createLocalDatabase(ROOT_DIR);
  const status = localDatabase.getStatus();
  if (!status.configured) {
    restoreTursoEnvironment(tursoEnv);
    throw new Error(status.reason || "Local SQLite database is not available.");
  }

  try {
    return {
      data: localDatabase.exportAllData(),
      counts: localDatabase.getTableCounts(),
      status,
    };
  } finally {
    await localDatabase.close();
    restoreTursoEnvironment(tursoEnv);
  }
}

function captureTursoEnvironment() {
  const captured = {};
  for (const key of TURSO_ENV_KEYS) {
    captured[key] = process.env[key];
  }
  return captured;
}

function clearTursoEnvironment() {
  for (const key of TURSO_ENV_KEYS) {
    delete process.env[key];
  }
}

function restoreTursoEnvironment(captured) {
  for (const key of TURSO_ENV_KEYS) {
    if (captured[key] === undefined) {
      delete process.env[key];
    } else {
      process.env[key] = captured[key];
    }
  }
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
    if (arg === "--dry-run") {
      args["dry-run"] = true;
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
  node migrate-local-to-turso.js
  node migrate-local-to-turso.js --env-file .env.local
  node migrate-local-to-turso.js --dry-run

Environment:
  TURSO_DATABASE_URL   Required target Turso database URL.
  TURSO_AUTH_TOKEN     Required for Turso Cloud targets.

Options:
  --url <value>        Target Turso database URL. Prefer env vars for production.
  --token <value>      Target Turso auth token. Prefer env vars or .env.local.
  --source <label>     Source label stored with the Turso snapshot.
  --env-file <path>    Load an extra env file before migrating.
  --dry-run            Show local counts without writing to Turso.
`.trim());
}

main().catch((error) => {
  console.error(error?.message || error);
  process.exitCode = 1;
});
