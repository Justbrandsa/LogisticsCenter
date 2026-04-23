const fs = require("fs");
const path = require("path");

function loadProjectEnv(rootDir = process.cwd(), options = {}) {
  const filenames = Array.isArray(options.filenames) && options.filenames.length
    ? options.filenames
    : [".env.local", ".env"];
  const override = Boolean(options.override);
  const loaded = [];

  for (const filename of filenames) {
    const filePath = path.isAbsolute(filename)
      ? filename
      : path.join(rootDir, filename);
    if (!fs.existsSync(filePath)) {
      continue;
    }

    const contents = fs.readFileSync(filePath, "utf8");
    for (const entry of parseEnvContents(contents)) {
      if (!override && process.env[entry.key] !== undefined) {
        continue;
      }
      process.env[entry.key] = entry.value;
      loaded.push(entry.key);
    }
  }

  return loaded;
}

function parseEnvContents(contents) {
  const entries = [];
  const lines = String(contents || "").split(/\r?\n/);

  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (!line || line.startsWith("#")) {
      continue;
    }

    const normalizedLine = line.startsWith("export ")
      ? line.slice("export ".length).trim()
      : line;
    const separatorIndex = normalizedLine.indexOf("=");
    if (separatorIndex <= 0) {
      continue;
    }

    const key = normalizedLine.slice(0, separatorIndex).trim();
    if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(key)) {
      continue;
    }

    const value = normalizeEnvValue(normalizedLine.slice(separatorIndex + 1).trim());
    entries.push({ key, value });
  }

  return entries;
}

function normalizeEnvValue(value) {
  if (value.length >= 2) {
    const quote = value[0];
    if ((quote === "\"" || quote === "'") && value[value.length - 1] === quote) {
      const unquoted = value.slice(1, -1);
      return quote === "\""
        ? unquoted
          .replace(/\\n/g, "\n")
          .replace(/\\r/g, "\r")
          .replace(/\\t/g, "\t")
          .replace(/\\"/g, "\"")
          .replace(/\\\\/g, "\\")
        : unquoted;
    }
  }

  const commentIndex = value.search(/\s+#/);
  return commentIndex >= 0 ? value.slice(0, commentIndex).trimEnd() : value;
}

module.exports = {
  loadProjectEnv,
  parseEnvContents,
};
