const SESSION_KEY = "route-ledger-session-token-v2";
const FLASH_TIMEOUT_MS = 3200;
const TIME_ZONE = "Africa/Johannesburg";
const API_ROOT = "/api";
const PAGE_SIZES = {
  globalEntries: 12,
  driverLists: 3,
  completedEntries: 10,
};
const HUB = {
  label: "Johannesburg Dispatch Hub",
  lat: -26.2041,
  lng: 28.0473,
};

const appEl = document.getElementById("app");
const authMetaEl = document.getElementById("auth-meta");
const userActionsEl = document.getElementById("user-actions");
const pageNavEl = document.getElementById("page-nav");

const state = {
  booting: true,
  busy: false,
  missingConfig: false,
  missingConfigReason: "",
  mailConfigured: false,
  mailConfigReason: "",
  mailTo: "admin3@giftwrap.co.za",
  currentPage: "",
  pagination: {
    globalEntries: 1,
    driverLists: 1,
    completedEntries: 1,
  },
  needsBootstrap: false,
  publicState: {
    today: "",
    weekStart: "",
    weekEnd: "",
    hasUsers: false,
  },
  snapshot: {
    user: null,
    users: [],
    suppliers: [],
    locations: [],
    orders: [],
  },
};

let sessionToken = loadSessionToken();
let flash = null;
let flashTimer = null;

document.addEventListener("submit", handleSubmit);
document.addEventListener("click", handleClick);
document.addEventListener("change", handleChange);
window.addEventListener("hashchange", handleHashChange);

void boot();

async function boot() {
  try {
    const status = await fetchServerStatus();
    state.mailConfigured = Boolean(status.mailConfigured);
    state.mailConfigReason = status.mailReason || "";
    state.mailTo = status.mailTo || state.mailTo;

    if (!status.configured) {
      state.booting = false;
      state.busy = false;
      state.missingConfig = true;
      state.missingConfigReason = status.reason || "";
      state.snapshot = createEmptySnapshot();
      render();
      return;
    }

    state.missingConfig = false;
    state.missingConfigReason = "";

    if (sessionToken) {
      await refreshSnapshot();
      return;
    }

    await refreshPublicState();
  } catch (error) {
    showFlash(normalizeError(error), "error");
    state.booting = false;
    state.busy = false;
    render();
  }
}

async function fetchServerStatus() {
  const payload = await requestJson(`${API_ROOT}/status`);
  return {
    configured: Boolean(payload?.configured),
    reason: payload?.reason || "",
    mailConfigured: Boolean(payload?.mailConfigured),
    mailReason: payload?.mailReason || "",
    mailTo: payload?.mailTo || "",
  };
}

function loadSessionToken() {
  return window.localStorage.getItem(SESSION_KEY) || "";
}

function saveSessionToken(token) {
  sessionToken = token || "";
  if (sessionToken) {
    window.localStorage.setItem(SESSION_KEY, sessionToken);
  } else {
    window.localStorage.removeItem(SESSION_KEY);
  }
}

async function refreshPublicState() {
  state.booting = true;
  render();

  const data = await callRpc("get_login_state");
  state.publicState = normalizePublicState(data);
  state.needsBootstrap = !state.publicState.hasUsers;
  state.snapshot = createEmptySnapshot();
  state.currentPage = "";
  state.booting = false;
  state.busy = false;
  render();
}

async function refreshSnapshot() {
  state.booting = true;
  render();

  try {
    const data = await callRpc("get_app_snapshot", { p_token: sessionToken });
    state.snapshot = normalizeSnapshot(data);
    state.publicState = normalizePublicState(data);
    state.needsBootstrap = false;
    syncCurrentPage();
    state.booting = false;
    state.busy = false;
    render();

    if (state.snapshot.user && state.snapshot.user.role === "driver") {
      drawDriverRoute(state.snapshot.user.id);
    }
  } catch (error) {
    const message = normalizeError(error);
    if (message.toLowerCase().includes("session")) {
      saveSessionToken("");
      await refreshPublicState();
      showFlash("Your session expired. Please sign in again.", "error");
      return;
    }

    throw error;
  }
}

async function callRpc(functionName, parameters = {}) {
  const payload = await requestJson(`${API_ROOT}/rpc/${encodeURIComponent(functionName)}`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ parameters }),
  });
  return payload?.data;
}

async function requestJson(url, options = {}) {
  const response = await fetch(url, options);
  let payload = null;

  try {
    payload = await response.json();
  } catch (error) {
    payload = null;
  }

  if (!response.ok) {
    throw new Error(payload?.error || `Request failed with status ${response.status}.`);
  }

  return payload;
}

function createEmptySnapshot() {
  return {
    user: null,
    users: [],
    suppliers: [],
    locations: [],
    orders: [],
  };
}

function normalizePublicState(data) {
  return {
    today: data?.today || "",
    weekStart: data?.weekStart || "",
    weekEnd: data?.weekEnd || "",
    hasUsers: Boolean(data?.hasUsers),
  };
}

function normalizeSnapshot(data) {
  return {
    user: data?.user || null,
    users: Array.isArray(data?.users) ? data.users : [],
    suppliers: Array.isArray(data?.suppliers) ? data.suppliers : [],
    locations: Array.isArray(data?.locations) ? data.locations : [],
    orders: Array.isArray(data?.orders) ? data.orders : [],
  };
}

function normalizeError(error) {
  if (!error) {
    return "Something went wrong.";
  }

  if (typeof error === "string") {
    return error;
  }

  const candidate = error.message || error.details || error.hint;
  return candidate || "Something went wrong.";
}

async function handleSubmit(event) {
  const form = event.target;
  if (!(form instanceof HTMLFormElement)) {
    return;
  }

  event.preventDefault();
  const formData = new FormData(form);
  const formId = form.dataset.form;
  const currentUser = state.snapshot.user;

  if (formId === "bootstrap-admin") {
    await handleBootstrap(formData);
    return;
  }

  if (formId === "login") {
    await handleLogin(formData);
    return;
  }

  if (!currentUser) {
    return;
  }

  if (formId === "add-account" && currentUser.role === "admin") {
    await createAccount(formData);
  }

  if (formId === "add-supplier" && currentUser.role === "admin") {
    await createSupplier(formData);
  }

  if (formId === "add-location" && currentUser.role === "admin") {
    await createLocation(formData);
  }

  if (formId === "add-order" && (currentUser.role === "admin" || currentUser.role === "sales")) {
    await createOrder(formData, currentUser);
  }
}

async function handleClick(event) {
  const button = event.target.closest("[data-action]");
  if (!button) {
    return;
  }

  const action = button.dataset.action;
  const currentUser = state.snapshot.user;

  if (action === "logout") {
    await logout();
    return;
  }

  if (action === "change-page") {
    changePage(button.dataset.pageKey, Number(button.dataset.page));
    return;
  }

  if (action === "navigate-page") {
    setCurrentPage(button.dataset.pageId || "");
    return;
  }

  if (!currentUser) {
    return;
  }

  if (action === "export-global-csv" && (currentUser.role === "admin" || currentUser.role === "sales")) {
    exportOrdersCsv();
    return;
  }

  if (action === "email-global-csv" && (currentUser.role === "admin" || currentUser.role === "sales")) {
    await emailOrdersCsv();
    return;
  }

  if (action === "email-test" && (currentUser.role === "admin" || currentUser.role === "sales")) {
    await sendTestEmail();
    return;
  }

  if (action === "toggle-user" && currentUser.role === "admin") {
    await runMutation(
      "toggle_user_active",
      { p_token: sessionToken, p_user_id: button.dataset.userId },
      "Account status updated.",
    );
  }

  if (action === "delete-user" && currentUser.role === "admin") {
    await runMutation(
      "delete_user_account",
      { p_token: sessionToken, p_user_id: button.dataset.userId },
      "Account deleted.",
    );
  }

  if (action === "delete-supplier" && currentUser.role === "admin") {
    await runMutation(
      "delete_supplier",
      { p_token: sessionToken, p_supplier_id: button.dataset.supplierId },
      "Supplier deleted.",
    );
  }

  if (action === "delete-location" && currentUser.role === "admin") {
    await runMutation(
      "delete_location",
      { p_token: sessionToken, p_location_id: button.dataset.locationId },
      "Location deleted.",
    );
  }

  if (action === "complete-order" && (currentUser.role === "admin" || currentUser.role === "driver")) {
    await runMutation(
      "complete_order",
      { p_token: sessionToken, p_order_id: button.dataset.orderId },
      "Order marked as completed.",
    );
  }

  if (action === "delete-order" && currentUser.role === "admin") {
    await runMutation(
      "delete_order",
      { p_token: sessionToken, p_order_id: button.dataset.orderId },
      "Order deleted.",
    );
  }
}

function handleChange(event) {
  const target = event.target;
  if (!(target instanceof HTMLElement)) {
    return;
  }

  if (target.matches("[data-driver-role-select]")) {
    const driverFields = document.getElementById("driver-role-fields");
    if (driverFields) {
      driverFields.classList.toggle("hidden", target.value !== "driver");
    }
  }
}

async function handleBootstrap(formData) {
  const name = String(formData.get("name") || "").trim();
  const password = String(formData.get("password") || "").trim();
  const confirmPassword = String(formData.get("confirmPassword") || "").trim();

  if (!name || !password) {
    showFlash("Name and password are required.", "error");
    render();
    return;
  }

  if (password !== confirmPassword) {
    showFlash("Passwords do not match.", "error");
    render();
    return;
  }

  state.busy = true;
  render();

  try {
    const data = await callRpc("bootstrap_admin", {
      p_name: name,
      p_password: password,
    });
    saveSessionToken(data?.token || "");
    showFlash(`Admin account created for ${name}.`, "success");
    await refreshSnapshot();
  } catch (error) {
    state.busy = false;
    showFlash(normalizeError(error), "error");
    render();
  }
}

async function handleLogin(formData) {
  const name = String(formData.get("name") || "").trim();
  const password = String(formData.get("password") || "").trim();

  if (!name || !password) {
    showFlash("Name and password are required.", "error");
    render();
    return;
  }

  state.busy = true;
  render();

  try {
    const data = await callRpc("login_user", {
      p_name: name,
      p_password: password,
    });
    saveSessionToken(data?.token || "");
    showFlash(`Signed in as ${name}.`, "success");
    await refreshSnapshot();
  } catch (error) {
    state.busy = false;
    showFlash(normalizeError(error), "error");
    render();
  }
}

async function logout() {
  state.busy = true;
  render();

  try {
    if (sessionToken) {
      await callRpc("logout_user", { p_token: sessionToken });
    }
  } catch (error) {
    showFlash(normalizeError(error), "error");
  } finally {
    saveSessionToken("");
    await refreshPublicState();
    showFlash("Signed out.", "success");
  }
}

async function runMutation(functionName, parameters, successMessage) {
  state.busy = true;
  render();

  try {
    await callRpc(functionName, parameters);
    showFlash(successMessage, "success");
    await refreshSnapshot();
  } catch (error) {
    state.busy = false;
    showFlash(normalizeError(error), "error");
    render();
  }
}

async function createAccount(formData) {
  const role = String(formData.get("role") || "").trim();
  const name = String(formData.get("name") || "").trim();
  const password = String(formData.get("password") || "").trim();
  const phone = String(formData.get("phone") || "").trim();

  if (!name || !password || !role) {
    showFlash("Name, password, and role are required.", "error");
    render();
    return;
  }

  await runMutation(
    "create_user_account",
    {
      p_token: sessionToken,
      p_name: name,
      p_password: password,
      p_role: role,
      p_phone: phone,
    },
    `${capitalize(role)} account created for ${name}.`,
  );
}

async function createSupplier(formData) {
  const name = String(formData.get("name") || "").trim();
  if (!name) {
    showFlash("Supplier name is required.", "error");
    render();
    return;
  }

  await runMutation(
    "create_supplier",
    {
      p_token: sessionToken,
      p_name: name,
    },
    `Supplier added: ${name}.`,
  );
}

async function createLocation(formData) {
  const supplierId = String(formData.get("supplierId") || "").trim();
  const name = String(formData.get("name") || "").trim();
  const address = String(formData.get("address") || "").trim();
  const lat = Number(formData.get("lat"));
  const lng = Number(formData.get("lng"));
  const notes = String(formData.get("notes") || "").trim();

  if (!supplierId || !name || !address || Number.isNaN(lat) || Number.isNaN(lng)) {
    showFlash("Supplier, location name, address, and coordinates are required.", "error");
    render();
    return;
  }

  await runMutation(
    "create_location",
    {
      p_token: sessionToken,
      p_supplier_id: supplierId,
      p_name: name,
      p_address: address,
      p_lat: lat,
      p_lng: lng,
      p_notes: notes,
    },
    `Location added: ${name}.`,
  );
}

async function createOrder(formData, currentUser) {
  const driverUserId = String(formData.get("driverUserId") || "").trim();
  const locationId = String(formData.get("locationId") || "").trim();
  const entryType = String(formData.get("entryType") || "delivery").trim();
  const factoryOrderNumber = String(formData.get("factoryOrderNumber") || "").trim();
  const inhouseOrderNumber = String(formData.get("inhouseOrderNumber") || "").trim();
  const notice = String(formData.get("notice") || "").trim();
  const allowDuplicate = formData.get("allowDuplicate") === "on";

  if (!driverUserId || !locationId || !entryType || !factoryOrderNumber || !inhouseOrderNumber) {
    showFlash("Driver, pickup location, entry type, factory order number, and in-house order number are required.", "error");
    render();
    return;
  }

  await runMutation(
    "create_order",
    {
      p_token: sessionToken,
      p_driver_user_id: driverUserId,
      p_location_id: locationId,
      p_entry_type: entryType,
      p_factory_order_number: factoryOrderNumber,
      p_inhouse_order_number: inhouseOrderNumber,
      p_notice: notice,
      p_allow_duplicate: currentUser.role === "admin" ? allowDuplicate : false,
    },
    "Entry added to the driver list.",
  );
}

function render() {
  renderHeader();
  renderPageNavigation();

  if (state.missingConfig) {
    appEl.innerHTML = renderSetupScreen();
    return;
  }

  if (state.booting) {
    appEl.innerHTML = renderLoadingScreen();
    return;
  }

  if (!state.snapshot.user) {
    appEl.innerHTML = state.needsBootstrap ? renderBootstrapScreen() : renderLoginScreen();
    return;
  }

  if (state.snapshot.user.role === "admin") {
    appEl.innerHTML = renderAdminScreen();
    return;
  }

  if (state.snapshot.user.role === "sales") {
    appEl.innerHTML = renderSalesScreen();
    return;
  }

  appEl.innerHTML = renderDriverScreen();
  drawDriverRoute(state.snapshot.user.id);
}

function renderHeader() {
  const currentUser = state.snapshot.user;
  const liveDate = state.publicState.today
    ? formatDateOnly(state.publicState.today)
    : "Waiting for database connection";

  if (!currentUser) {
    authMetaEl.innerHTML = `
      <strong>Neon-backed workflow</strong><br>
      <span class="muted">Live date: ${escapeHtml(liveDate)}</span>
    `;
    userActionsEl.innerHTML = "";
    return;
  }

  authMetaEl.innerHTML = `
    <strong>${escapeHtml(currentUser.name)}</strong><br>
    <span class="muted">${capitalize(currentUser.role)} account</span><br>
    <span class="muted">${state.snapshot.orders.length} visible entr${state.snapshot.orders.length === 1 ? "y" : "ies"}</span>
  `;
  userActionsEl.innerHTML = `
    <button class="button button-ghost" data-action="logout"${state.busy ? " disabled" : ""}>Logout</button>
  `;
}

function renderPageNavigation() {
  if (!pageNavEl) {
    return;
  }

  const currentUser = state.snapshot.user;
  if (!currentUser) {
    pageNavEl.innerHTML = "";
    pageNavEl.classList.remove("page-nav-active");
    return;
  }

  const items = getNavigationItems(currentUser.role);
  syncCurrentPage();
  pageNavEl.classList.add("page-nav-active");
  pageNavEl.innerHTML = items
    .map(
      (item) => `
        <button
          class="page-nav-link${state.currentPage === item.id ? " is-active" : ""}"
          data-action="navigate-page"
          data-page-id="${item.id}"
          ${state.busy ? "disabled" : ""}
        >
          ${escapeHtml(item.label)}
        </button>
      `,
    )
    .join("");
}

function getNavigationItems(role) {
  if (role === "admin") {
    return [
      { id: "dashboard", label: "Dashboard" },
      { id: "entries", label: "Global List" },
      { id: "network", label: "Network" },
      { id: "users", label: "Users" },
      { id: "drivers", label: "Driver Lists" },
    ];
  }

  if (role === "sales") {
    return [
      { id: "dashboard", label: "Dashboard" },
      { id: "entries", label: "Global List" },
      { id: "drivers", label: "Driver Lists" },
    ];
  }

  return [
    { id: "route", label: "Route" },
    { id: "completed", label: "Completed" },
  ];
}

function syncCurrentPage() {
  const currentUser = state.snapshot.user;
  if (!currentUser) {
    state.currentPage = "";
    return;
  }

  const items = getNavigationItems(currentUser.role);
  const pageIds = items.map((item) => item.id);
  const hashPage = getHashPageId();
  const nextPage = pageIds.includes(hashPage)
    ? hashPage
    : pageIds.includes(state.currentPage)
      ? state.currentPage
      : items[0].id;

  state.currentPage = nextPage;
  const targetHash = `#${nextPage}`;
  if (window.location.hash !== targetHash) {
    window.history.replaceState(null, "", `${window.location.pathname}${targetHash}`);
  }
}

function setCurrentPage(pageId) {
  const currentUser = state.snapshot.user;
  if (!currentUser) {
    return;
  }

  const items = getNavigationItems(currentUser.role);
  if (!items.some((item) => item.id === pageId)) {
    return;
  }

  state.currentPage = pageId;
  window.history.replaceState(null, "", `${window.location.pathname}#${pageId}`);
  render();
}

function handleHashChange() {
  const currentUser = state.snapshot.user;
  if (!currentUser) {
    return;
  }

  const items = getNavigationItems(currentUser.role);
  const hashPage = getHashPageId();
  if (!items.some((item) => item.id === hashPage)) {
    return;
  }

  state.currentPage = hashPage;
  render();
}

function getHashPageId() {
  return (window.location.hash || "").replace(/^#/, "").trim();
}

function renderSetupScreen() {
  return `
    <section class="login-wrap">
      <div class="login-card">
        <article class="login-intro">
          <p class="eyebrow">Neon setup</p>
          <h2>Database connection is not configured yet</h2>
          <p class="muted">
            This version reads and writes everything through Neon by way of the local Node server. Add your
            connection string, then run the SQL setup file before signing in.
          </p>
          <div class="credential-list">
            <div class="credential-item">
              <strong>1. Configure the local server</strong>
              <div>Create <code>neon-config.js</code> from <code>neon-config.example.js</code> or set <code>DATABASE_URL</code>.</div>
            </div>
            <div class="credential-item">
              <strong>2. Create the database objects</strong>
              <div>Run <code>neon.sql</code> against your Neon database.</div>
            </div>
            <div class="credential-item">
              <strong>3. Configure email delivery</strong>
              <div>Create <code>mail-config.js</code> from <code>mail-config.example.js</code> or set the Microsoft Graph <code>MAIL_*</code> environment variables.</div>
            </div>
          </div>
        </article>
        <article class="login-form">
          <p class="eyebrow">Current status</p>
          <h2>Connection required</h2>
          ${renderFlash()}
          ${state.missingConfigReason ? `<p class="muted">${escapeHtml(state.missingConfigReason)}</p>` : ""}
          <p class="muted">
            After you add the server connection details, refresh this page and the app will switch to either first-admin
            setup or name-based sign-in.
          </p>
        </article>
      </div>
    </section>
  `;
}

function renderLoadingScreen() {
  return `
    <section class="login-wrap">
      <div class="login-card">
        <article class="login-intro">
          <p class="eyebrow">Route Ledger</p>
          <h2>Loading the live database workspace</h2>
          <p class="muted">
            Pulling live entries, driver lists, and account data from Neon.
          </p>
        </article>
        <article class="login-form">
          <p class="eyebrow">Please wait</p>
          <h2>${state.busy ? "Working..." : "Booting..."}</h2>
          ${renderFlash()}
        </article>
      </div>
    </section>
  `;
}

function renderBootstrapScreen() {
  return `
    <section class="login-wrap">
      <div class="login-card">
        <article class="login-intro">
          <p class="eyebrow">First-time setup</p>
          <h2>Create the first admin account</h2>
          <p class="muted">
            No users exist in the database yet. The first account created here becomes the initial admin and from then
            on all new sales and driver accounts are managed inside the app.
          </p>
          <div class="credential-list">
            <div class="credential-item">
              <strong>Name-based login</strong>
              <div>The name you enter here becomes the login name, so keep it unique.</div>
            </div>
            <div class="credential-item">
              <strong>Persistent records</strong>
              <div>Entries stay in the database until you export or manage them, so the first admin can work from one shared live list.</div>
            </div>
          </div>
        </article>
        <article class="login-form">
          <p class="eyebrow">Create admin</p>
          <h2>Initial account</h2>
          ${renderFlash()}
          <form data-form="bootstrap-admin">
            <label>
              Name
              <input name="name" type="text" placeholder="Alicia Admin" required>
            </label>
            <label>
              Password
              <input name="password" type="password" required>
            </label>
            <label>
              Confirm password
              <input name="confirmPassword" type="password" required>
            </label>
            <button type="submit" class="button button-primary"${state.busy ? " disabled" : ""}>
              ${state.busy ? "Creating..." : "Create admin"}
            </button>
          </form>
        </article>
      </div>
    </section>
  `;
}

function renderLoginScreen() {
  return `
    <section class="login-wrap">
      <div class="login-card">
        <article class="login-intro">
          <p class="eyebrow">Live database</p>
          <h2>Sign in with your name</h2>
          <p class="muted">
            User records, locations, and driver entries are all pulled from Neon. Driver-separated lists and the
            global entries register read from the same live data.
          </p>
          <div class="credential-list">
            <div class="credential-item">
              <strong>Login field changed</strong>
              <div>Users now sign in with their unique name and password instead of email.</div>
            </div>
            <div class="credential-item">
              <strong>Live date</strong>
              <div>${escapeHtml(formatDateOnly(state.publicState.today) || "Waiting for sync")}</div>
            </div>
          </div>
        </article>
        <article class="login-form">
          <p class="eyebrow">Sign in</p>
          <h2>Access your workspace</h2>
          ${renderFlash()}
          <form data-form="login">
            <label>
              Name
              <input name="name" type="text" placeholder="Daniel Dube" required>
            </label>
            <label>
              Password
              <input name="password" type="password" required>
            </label>
            <button type="submit" class="button button-primary"${state.busy ? " disabled" : ""}>
              ${state.busy ? "Signing in..." : "Sign in"}
            </button>
          </form>
        </article>
      </div>
    </section>
  `;
}

function renderAdminScreen() {
  return `
    <section class="screen-grid">
      <aside class="sidebar">
        <div class="sidebar-card">
          <p class="eyebrow">Admin scope</p>
          <h3>Full control</h3>
          <p class="muted">
            Manage users, suppliers, locations, driver entries, and CSV delivery from one live database.
          </p>
          <div class="chip-row">
            <span class="chip chip-role-admin">Admin</span>
            <span class="chip">${state.snapshot.suppliers.length} suppliers</span>
            <span class="chip">${state.snapshot.locations.length} locations</span>
          </div>
        </div>
        <div class="sidebar-card">
          <h3>Email delivery</h3>
          <p class="muted">
            ${state.mailConfigured ? `CSV email is ready for ${escapeHtml(state.mailTo)}.` : escapeHtml(state.mailConfigReason || "Email delivery is not configured yet.")}
          </p>
          <p class="muted">Use the Global List page to download the CSV, send it, or run a test email.</p>
        </div>
      </aside>
      <div class="content">
        ${renderAdminPageContent()}
      </div>
    </section>
  `;
}

function renderSalesScreen() {
  return `
    <section class="screen-grid">
      <aside class="sidebar">
        <div class="sidebar-card">
          <p class="eyebrow">Sales scope</p>
          <h3>Live entry control</h3>
          <p class="muted">
            You can view driver lists, create new entries, and email or download the global CSV, but duplicate active
            orders are still blocked for sales users.
          </p>
          <div class="chip-row">
            <span class="chip chip-role-sales">Sales</span>
            <span class="chip">${countOrdersCreatedByCurrentUser()} created by you</span>
          </div>
        </div>
        <div class="sidebar-card">
          <h3>Email delivery</h3>
          <p class="muted">
            ${state.mailConfigured ? `CSV email is ready for ${escapeHtml(state.mailTo)}.` : escapeHtml(state.mailConfigReason || "Email delivery is not configured yet.")}
          </p>
          <p class="muted">Sales users can test email delivery, but they still cannot bypass the duplicate active-order rule.</p>
        </div>
      </aside>
      <div class="content">
        ${renderSalesPageContent()}
      </div>
    </section>
  `;
}

function renderDriverScreen() {
  const plan = getRoutePlan(state.snapshot.user.id);

  return `
    <section class="screen-grid">
      <aside class="sidebar">
        <div class="sidebar-card">
          <p class="eyebrow">Driver access</p>
          <h3>${escapeHtml(state.snapshot.user.name)}</h3>
          <p class="muted">${escapeHtml(state.snapshot.user.phone || "No phone assigned")}</p>
          <div class="chip-row">
            <span class="chip chip-role-driver">Driver</span>
            <span class="chip">${plan.totalOrders} entries</span>
            <span class="chip">${plan.stops.length} stops</span>
          </div>
        </div>
        <div class="sidebar-card">
          <h3>Route mode</h3>
          <p class="muted">
            Stop order is calculated from your active assigned pickup locations. Completed entries stay on the Completed
            page.
          </p>
        </div>
      </aside>
      <div class="content">
        ${renderDriverPageContent()}
      </div>
    </section>
  `;
}

function renderAdminPageContent() {
  const activeOrders = getActiveOrders();
  const completedOrders = getCompletedOrders();

  if (state.currentPage === "entries") {
    return `
      <section class="hero-card">
        <p class="eyebrow">Global List</p>
        <h2>Create work and distribute the global list</h2>
        <p>Use this page to add live work, download the CSV, or mail the CSV straight to the admin inbox.</p>
      </section>
      ${renderFlash()}
      <section class="panel">
        <p class="eyebrow">Global List</p>
        <h3 class="panel-title">Create a new driver entry</h3>
        <p class="panel-subtitle">Admins can override the duplicate-order rule when there is a real exception.</p>
        ${renderEntryForm(state.snapshot.user, true)}
      </section>
      ${renderGlobalOrdersSection("admin")}
    `;
  }

  if (state.currentPage === "network") {
    return `
      <section class="hero-card">
        <p class="eyebrow">Network</p>
        <h2>Suppliers and pickup locations</h2>
        <p>Maintain the supplier directory and the physical pickup points used by the route lists.</p>
      </section>
      ${renderFlash()}
      ${renderSupplierNetworkSection()}
    `;
  }

  if (state.currentPage === "users") {
    return `
      <section class="hero-card">
        <p class="eyebrow">Users</p>
        <h2>Team accounts and permissions</h2>
        <p>Create admin, sales, and driver accounts, then manage their status from one place.</p>
      </section>
      ${renderFlash()}
      <section class="panel-grid">
        ${renderAccountPanel()}
      </section>
      ${renderUsersSection()}
    `;
  }

  if (state.currentPage === "drivers") {
    return `
      <section class="hero-card">
        <p class="eyebrow">Driver lists</p>
        <h2>Separated route views by driver</h2>
        <p>Each driver page groups active entries by pickup location and keeps the optimized route visible.</p>
      </section>
      ${renderFlash()}
      ${renderDriverListOverview("admin")}
    `;
  }

  return `
    <section class="hero-card">
      <p class="eyebrow">Dashboard</p>
      <h2>Dispatch control with driver lists and one global register</h2>
      <p>Create entries once, track them globally, and still keep each driver route grouped by pickup location.</p>
    </section>
    ${renderFlash()}
    <section class="metrics">
      ${renderMetric("Open entries", activeOrders.length)}
      ${renderMetric("Drivers", getDriverUsers().length)}
      ${renderMetric("Completed entries", completedOrders.length)}
    </section>
    <section class="panel-grid">
      ${renderPageSummaryCard("Global List", "Add new work and send, test, or download the CSV register.", "entries")}
      ${renderPageSummaryCard("Network", "Maintain suppliers and pickup locations.", "network")}
      ${renderPageSummaryCard("Users", "Manage admin, sales, and driver accounts.", "users")}
      ${renderPageSummaryCard("Driver lists", "Review active work separated by driver.", "drivers")}
    </section>
  `;
}

function renderSalesPageContent() {
  const ordersCreated = countOrdersCreatedByCurrentUser();
  const activeOrders = getActiveOrders();

  if (state.currentPage === "entries") {
    return `
      <section class="hero-card">
        <p class="eyebrow">Global List</p>
        <h2>Create work and manage the global list</h2>
        <p>Sales can add work, download the CSV, or email the latest global list to the admin inbox.</p>
      </section>
      ${renderFlash()}
      <section class="panel">
        <p class="eyebrow">Global List</p>
        <h3 class="panel-title">Append work to a driver list</h3>
        <p class="panel-subtitle">The duplicate-order rule is enforced by the database.</p>
        ${renderEntryForm(state.snapshot.user, false)}
      </section>
      ${renderGlobalOrdersSection("sales")}
    `;
  }

  if (state.currentPage === "drivers") {
    return `
      <section class="hero-card">
        <p class="eyebrow">Driver lists</p>
        <h2>Review the live driver-separated queues</h2>
        <p>Each driver list groups active entries by pickup location and keeps the route sequence visible.</p>
      </section>
      ${renderFlash()}
      ${renderDriverListOverview("sales")}
    `;
  }

  return `
    <section class="hero-card">
      <p class="eyebrow">Dashboard</p>
      <h2>View driver coverage and manage live entry flow</h2>
      <p>Driver lists and the global entries register stay in sync so dispatch and reporting reflect the same data.</p>
    </section>
    ${renderFlash()}
    <section class="metrics">
      ${renderMetric("Drivers visible", getDriverUsers().length)}
      ${renderMetric("Entries created", ordersCreated)}
      ${renderMetric("Open entries", activeOrders.length)}
    </section>
    <section class="panel-grid">
      ${renderPageSummaryCard("Global List", "Create new work and email, test, or download the CSV register.", "entries")}
      ${renderPageSummaryCard("Driver lists", "Review active work separated by driver.", "drivers")}
    </section>
  `;
}

function renderDriverPageContent() {
  const currentUser = state.snapshot.user;
  const plan = getRoutePlan(currentUser.id);
  const completedOrders = getCompletedOrders(currentUser.id);

  if (state.currentPage === "completed") {
    return `
      <section class="hero-card">
        <p class="eyebrow">Completed</p>
        <h2>Your completed entries</h2>
        <p>Completed work is kept in the database and listed here separately from the live route.</p>
      </section>
      ${renderFlash()}
      <section class="metrics">
        ${renderMetric("Completed entries", completedOrders.length)}
        ${renderMetric("Active stops", plan.stops.length)}
        ${renderMetric("Estimated km", plan.totalKm.toFixed(1))}
      </section>
      ${renderCompletedOrders(currentUser.id)}
    `;
  }

  return `
    <section class="hero-card">
      <p class="eyebrow">Route</p>
      <h2>Optimized run sheet for your live active entries</h2>
      <p>You only see the entries assigned to your name. Completing an entry removes it from the live route sequence.</p>
    </section>
    ${renderFlash()}
    <section class="metrics">
      ${renderMetric("Active stops", plan.stops.length)}
      ${renderMetric("Estimated km", plan.totalKm.toFixed(1))}
      ${renderMetric("Completed entries", completedOrders.length)}
    </section>
    <section class="route-canvas-card">
      <p class="eyebrow">Route sketch</p>
      <h3 class="panel-title">Dispatch hub to optimized stop sequence</h3>
      <div class="route-canvas-wrap">
        <canvas id="route-canvas" width="900" height="340"></canvas>
      </div>
    </section>
    <section class="driver-grid">
      ${
        plan.stops.length
          ? plan.stops.map((stop, index) => renderStopCard(stop, index, "driver")).join("")
          : `
            <div class="empty-state">
              No active entries are assigned to you right now. When work is loaded for you, it will appear here.
            </div>
          `
      }
    </section>
  `;
}

function renderPageSummaryCard(title, description, pageId) {
  return `
    <article class="panel summary-panel">
      <p class="eyebrow">${escapeHtml(title)}</p>
      <h3 class="panel-title">${escapeHtml(title)}</h3>
      <p class="panel-subtitle">${escapeHtml(description)}</p>
      <button class="button button-secondary" data-action="navigate-page" data-page-id="${pageId}"${state.busy ? " disabled" : ""}>
        Open ${escapeHtml(title)}
      </button>
    </article>
  `;
}

function renderAccountPanel() {
  return `
    <article class="panel">
      <p class="eyebrow">Accounts</p>
      <h3 class="panel-title">Add team member</h3>
      <p class="panel-subtitle">Create admin, sales, or driver accounts with name-based login.</p>
      <form data-form="add-account">
        <div class="form-grid">
          <label>
            Name
            <input name="name" type="text" required>
          </label>
          <label>
            Role
            <select name="role" data-driver-role-select required>
              <option value="sales">Sales</option>
              <option value="driver">Driver</option>
              <option value="admin">Admin</option>
            </select>
          </label>
        </div>
        <div class="form-grid">
          <label>
            Password
            <input name="password" type="password" required>
          </label>
          <label>
            Note
            <input type="text" value="Name is also the login field" readonly>
          </label>
        </div>
        <div id="driver-role-fields" class="hidden">
          <div class="form-grid">
            <label>
              Driver phone
              <input name="phone" type="text" placeholder="071 555 0100">
            </label>
          </div>
        </div>
        <button type="submit" class="button button-primary"${state.busy ? " disabled" : ""}>
          Create account
        </button>
      </form>
    </article>
  `;
}

function renderSupplierNetworkSection() {
  return `
    <section class="table-card">
      <p class="eyebrow">Supplier network</p>
      <h3 class="panel-title">Suppliers and locations</h3>
      <p class="panel-subtitle">Both parts of the delivery network are managed together here.</p>
      <div class="panel-grid">
        <article class="panel">
          <p class="eyebrow">Suppliers</p>
          <h3 class="panel-title">Add supplier</h3>
          <form data-form="add-supplier">
            <label>
              Supplier name
              <input name="name" type="text" required>
            </label>
            <button type="submit" class="button button-primary"${state.busy ? " disabled" : ""}>
              Add supplier
            </button>
          </form>
          <div class="table-scroll">
            <table>
              <thead>
                <tr>
                  <th>Supplier</th>
                  <th>Locations</th>
                  <th>Action</th>
                </tr>
              </thead>
              <tbody>
                ${
                  state.snapshot.suppliers.length
                    ? state.snapshot.suppliers.map((supplier) => renderSupplierRow(supplier)).join("")
                    : `
                      <tr>
                        <td colspan="3">No suppliers added yet.</td>
                      </tr>
                    `
                }
              </tbody>
            </table>
          </div>
        </article>
        <article class="panel">
          <p class="eyebrow">Locations</p>
          <h3 class="panel-title">Add location</h3>
          <form data-form="add-location">
            <div class="form-grid">
              <label>
                Supplier
                <select name="supplierId" required>
                  ${renderSupplierOptions()}
                </select>
              </label>
              <label>
                Location name
                <input name="name" type="text" required>
              </label>
            </div>
            <label>
              Address
              <input name="address" type="text" required>
            </label>
            <div class="form-grid">
              <label>
                Latitude
                <input name="lat" type="number" step="0.0001" required>
              </label>
              <label>
                Longitude
                <input name="lng" type="number" step="0.0001" required>
              </label>
            </div>
            <label>
              Notes
              <textarea name="notes" placeholder="Dock, gate, or access notes"></textarea>
            </label>
            <button type="submit" class="button button-primary"${state.busy ? " disabled" : ""}>
              Add location
            </button>
          </form>
          <div class="table-scroll">
            <table>
              <thead>
                <tr>
                  <th>Location</th>
                  <th>Supplier</th>
                  <th>Address</th>
                  <th>Action</th>
                </tr>
              </thead>
              <tbody>
                ${
                  state.snapshot.locations.length
                    ? state.snapshot.locations.map((location) => renderLocationRow(location)).join("")
                    : `
                      <tr>
                        <td colspan="4">No locations added yet.</td>
                      </tr>
                    `
                }
              </tbody>
            </table>
          </div>
        </article>
      </div>
    </section>
  `;
}

function renderUsersSection() {
  return `
    <section class="table-card">
      <p class="eyebrow">Users</p>
      <h3 class="panel-title">Sales and driver account management</h3>
      <div class="table-scroll">
        <table>
          <thead>
            <tr>
              <th>Name</th>
              <th>Role</th>
              <th>Status</th>
              <th>Details</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
            ${
              state.snapshot.users.length
                ? state.snapshot.users.map((user) => renderUserRow(user)).join("")
                : `
                  <tr>
                    <td colspan="5">No users available.</td>
                  </tr>
                `
            }
          </tbody>
        </table>
      </div>
    </section>
  `;
}

function countOrdersCreatedByCurrentUser() {
  if (!state.snapshot.user) {
    return 0;
  }

  return state.snapshot.orders.filter((order) => order.createdByUserId === state.snapshot.user.id).length;
}

function renderEntryForm(currentUser, allowDuplicateOverride) {
  return `
    <form data-form="add-order">
      <div class="form-grid">
        <label>
          Driver
          <select name="driverUserId" required>
            ${renderDriverOptions()}
          </select>
        </label>
        <label>
          Pickup location
          <select name="locationId" required>
            ${renderLocationOptions()}
          </select>
        </label>
      </div>
      <div class="form-grid">
        <label>
          Collection or delivery
          <select name="entryType" required>
            <option value="collection">Collection</option>
            <option value="delivery" selected>Delivery</option>
          </select>
        </label>
        <label>
          Created by
          <input class="readonly-field" type="text" value="${escapeHtml(currentUser.name)}" readonly>
        </label>
      </div>
      <div class="form-grid">
        <label>
          Factory order number
          <input name="factoryOrderNumber" type="text" required>
        </label>
        <label>
          In-house order number
          <input name="inhouseOrderNumber" type="text" required>
        </label>
      </div>
      <label>
        Notice
        <textarea name="notice" placeholder="Special instruction, handover note, or anything dispatch should keep on the order"></textarea>
      </label>
      ${
        allowDuplicateOverride
          ? `
            <label class="inline-check">
              <input type="checkbox" name="allowDuplicate">
              Allow a duplicate active order for this driver
            </label>
            <p class="field-note">
              This only bypasses the block when the same driver already has another active entry with the same order numbers.
            </p>
          `
          : ""
      }
      <button type="submit" class="button button-primary"${state.busy ? " disabled" : ""}>
        Create entry
      </button>
    </form>
  `;
}

function renderGlobalOrdersSection(viewerRole) {
  const sortedOrders = [...state.snapshot.orders].sort(orderDisplaySort);
  const page = getPaginationData(sortedOrders, "globalEntries", PAGE_SIZES.globalEntries);
  const canExport = viewerRole === "admin" || viewerRole === "sales";

  return `
    <section class="table-card">
      <div class="table-toolbar">
        <div>
          <p class="eyebrow">Global List</p>
          <h3 class="panel-title">All visible entries</h3>
          <p class="panel-subtitle">This table combines every visible entry across the driver-separated lists.</p>
        </div>
        ${
          canExport
            ? `
              <div class="action-row">
                <button class="button button-secondary" data-action="export-global-csv"${state.busy ? " disabled" : ""}>
                  Download CSV
                </button>
                <button
                  class="button button-secondary"
                  data-action="email-test"
                  ${state.busy || !state.mailConfigured ? "disabled" : ""}
                >
                  Test Email
                </button>
                <button
                  class="button button-primary"
                  data-action="email-global-csv"
                  ${state.busy || !state.mailConfigured ? "disabled" : ""}
                >
                  Email CSV
                </button>
              </div>
            `
            : ""
        }
      </div>
      ${
        canExport && !state.mailConfigured
          ? `<p class="field-note">${escapeHtml(state.mailConfigReason || "Email delivery is not configured yet.")}</p>`
          : ""
      }
      <div class="table-scroll">
        <table>
          <thead>
            <tr>
              <th>Reference</th>
              <th>In-house</th>
              <th>Factory</th>
              <th>Type</th>
              <th>Driver</th>
              <th>Pickup location</th>
              <th>Created by</th>
              <th>Status</th>
              <th>Notice</th>
              <th>Created</th>
            </tr>
          </thead>
          <tbody>
            ${
              page.items.length
                ? page.items.map((order) => renderGlobalOrderRow(order)).join("")
                : `
                  <tr>
                    <td colspan="10">No entries available yet.</td>
                  </tr>
                `
            }
          </tbody>
        </table>
      </div>
      ${renderPaginationControls("globalEntries", page)}
    </section>
  `;
}

function renderDriverListOverview(viewerRole) {
  const drivers = getDriverUsers();
  const page = getPaginationData(drivers, "driverLists", PAGE_SIZES.driverLists);
  const content = page.items
    .map((driver) => {
      const plan = getRoutePlan(driver.id);
      const duplicateCount = countDuplicateOrders(driver.id);
      return `
        <article class="panel">
          <div class="stop-header">
            <div>
              <p class="eyebrow">${escapeHtml(driver.phone || "Driver list")}</p>
              <h3 class="panel-title">${escapeHtml(driver.name)}</h3>
            </div>
            <div class="chip-row">
              <span class="chip">${plan.totalOrders} entries</span>
              <span class="chip">${plan.totalKm.toFixed(1)} km</span>
              ${
                duplicateCount
                  ? `<span class="chip chip-warning">${duplicateCount} duplicate order${duplicateCount === 1 ? "" : "s"}</span>`
                  : ""
              }
            </div>
          </div>
          <p class="panel-subtitle">Active pickup route sequence.</p>
          ${
            plan.stops.length
              ? plan.stops.map((stop, index) => renderStopCard(stop, index, viewerRole)).join("")
              : '<div class="empty-state">No active work on this driver list.</div>'
          }
        </article>
      `;
    })
    .join("");

  return `
    <section class="table-card">
      <div class="table-toolbar">
        <div>
          <p class="eyebrow">Driver lists</p>
          <h3 class="panel-title">Separated driver views</h3>
          <p class="panel-subtitle">Each page groups active entries by driver and pickup location.</p>
        </div>
      </div>
      <div class="panel-grid">
        ${content || '<div class="empty-state">No driver accounts have been created yet.</div>'}
      </div>
      ${renderPaginationControls("driverLists", page)}
    </section>
  `;
}

function renderCompletedOrders(driverUserId) {
  const completed = getCompletedOrders(driverUserId).sort(orderDisplaySort);
  const page = getPaginationData(completed, "completedEntries", PAGE_SIZES.completedEntries);

  return `
    <section class="table-card">
      <p class="eyebrow">Completed</p>
      <h3 class="panel-title">Finished entries</h3>
      <div class="table-scroll">
        <table>
          <thead>
            <tr>
              <th>Reference</th>
              <th>In-house</th>
              <th>Factory</th>
              <th>Type</th>
              <th>Completed</th>
            </tr>
          </thead>
          <tbody>
            ${
              page.items.length
                ? page.items
                    .map(
                      (order) => `
                        <tr>
                          <td>${escapeHtml(order.reference)}</td>
                          <td>${escapeHtml(order.inhouseOrderNumber || "")}</td>
                          <td>${escapeHtml(order.factoryOrderNumber || "")}</td>
                          <td>${renderTypeChip(order.entryType)}</td>
                          <td>${escapeHtml(formatDateTime(order.completedAt) || "Not completed")}</td>
                        </tr>
                      `,
                    )
                    .join("")
                : `
                  <tr>
                    <td colspan="5">No completed entries yet.</td>
                  </tr>
                `
            }
          </tbody>
        </table>
      </div>
      ${renderPaginationControls("completedEntries", page)}
    </section>
  `;
}

function renderGlobalOrderRow(order) {
  return `
    <tr>
      <td>${escapeHtml(order.reference)}</td>
      <td>${escapeHtml(order.inhouseOrderNumber || "")}</td>
      <td>${escapeHtml(order.factoryOrderNumber || "")}</td>
      <td>${renderTypeChip(order.entryType)}</td>
      <td>${escapeHtml(order.driverName || "Unknown")}</td>
      <td>
        <strong>${escapeHtml(order.locationName || "Unknown")}</strong><br>
        <span class="muted">${escapeHtml(order.locationAddress || "")}</span>
      </td>
      <td>${escapeHtml(order.createdByName || "Unknown")}</td>
      <td>${renderStatusChip(order.status)}</td>
      <td>${renderOrderNotice(order, "None")}</td>
      <td>${escapeHtml(formatDateTime(order.createdAt))}</td>
    </tr>
  `;
}

function renderMetric(label, value) {
  return `
    <article class="metric-card">
      <p class="metric-label">${escapeHtml(label)}</p>
      <p class="metric-value">${escapeHtml(String(value))}</p>
    </article>
  `;
}

function renderUserRow(user) {
  const detailParts = [];
  if (user.phone) {
    detailParts.push(escapeHtml(user.phone));
  }

  return `
    <tr>
      <td>${escapeHtml(user.name)}</td>
      <td><span class="chip chip-role-${user.role}">${capitalize(user.role)}</span></td>
      <td>${user.active ? '<span class="chip chip-success">Active</span>' : '<span class="chip chip-warning">Inactive</span>'}</td>
      <td>${detailParts.length ? detailParts.join("<br>") : "n/a"}</td>
      <td>
        <div class="action-row">
          <button class="button button-secondary" data-action="toggle-user" data-user-id="${user.id}"${state.busy ? " disabled" : ""}>
            ${user.active ? "Disable" : "Enable"}
          </button>
          <button class="button button-danger" data-action="delete-user" data-user-id="${user.id}"${state.busy ? " disabled" : ""}>
            Delete
          </button>
        </div>
      </td>
    </tr>
  `;
}

function renderSupplierRow(supplier) {
  const locationCount = state.snapshot.locations.filter((location) => location.supplierId === supplier.id).length;
  return `
    <tr>
      <td>${escapeHtml(supplier.name)}</td>
      <td>${locationCount}</td>
      <td>
        <button class="button button-danger" data-action="delete-supplier" data-supplier-id="${supplier.id}"${state.busy ? " disabled" : ""}>
          Delete
        </button>
      </td>
    </tr>
  `;
}

function renderLocationRow(location) {
  return `
    <tr>
      <td>
        <strong>${escapeHtml(location.name)}</strong><br>
        <span class="muted">${escapeHtml(location.notes || "No notes")}</span>
      </td>
      <td>${escapeHtml(location.supplierName || "Unknown")}</td>
      <td>${escapeHtml(location.address)}</td>
      <td>
        <button class="button button-danger" data-action="delete-location" data-location-id="${location.id}"${state.busy ? " disabled" : ""}>
          Delete
        </button>
      </td>
    </tr>
  `;
}

function renderStopCard(stop, index, viewerRole) {
  const allowComplete = viewerRole === "admin" || viewerRole === "driver";
  const allowDelete = viewerRole === "admin";

  return `
    <article class="stop-card">
      <div class="stop-header">
        <div>
          <p class="eyebrow">Stop ${index + 1}</p>
          <h4 class="stop-title">${escapeHtml(stop.location.name)}</h4>
          <p class="stop-address">${escapeHtml(stop.location.address)}</p>
        </div>
        <div class="chip-row">
          <span class="chip">${stop.legKm.toFixed(1)} km leg</span>
          <span class="chip">${stop.orders.length} entr${stop.orders.length === 1 ? "y" : "ies"}</span>
        </div>
      </div>
      <div class="stop-orders">
        ${stop.orders
          .map(
            (order) => `
              <div class="order-card">
                <strong>${escapeHtml(order.reference)}</strong>
                <div class="order-meta">
                  <span>In-house ${escapeHtml(order.inhouseOrderNumber || "")}</span>
                  <span>Factory ${escapeHtml(order.factoryOrderNumber || "")}</span>
                </div>
                <div class="chip-row">
                  ${renderTypeChip(order.entryType)}
                  <span class="chip">Created by ${escapeHtml(order.createdByName)}</span>
                  ${renderStatusChip(order.status)}
                </div>
                ${renderOrderNotice(order)}
                <p>${escapeHtml(order.locationName || stop.location.name)}<br>${escapeHtml(order.locationAddress || stop.location.address)}</p>
                ${
                  allowComplete || allowDelete
                    ? `
                      <div class="action-row">
                        ${
                          allowComplete
                            ? `
                              <button class="button button-secondary" data-action="complete-order" data-order-id="${order.id}"${state.busy ? " disabled" : ""}>
                                Complete
                              </button>
                            `
                            : ""
                        }
                        ${
                          allowDelete
                            ? `
                              <button class="button button-danger" data-action="delete-order" data-order-id="${order.id}"${state.busy ? " disabled" : ""}>
                                Delete
                              </button>
                            `
                            : ""
                        }
                      </div>
                    `
                    : ""
                }
              </div>
            `,
          )
          .join("")}
      </div>
    </article>
  `;
}

function renderTypeChip(entryType) {
  return `<span class="chip chip-type">${capitalize(entryType || "entry")}</span>`;
}

function renderStatusChip(status) {
  const statusClass = status === "completed" ? "chip-success" : "chip-type";
  return `<span class="chip ${statusClass}">${capitalize(status || "active")}</span>`;
}

function renderFlash() {
  if (!flash) {
    return "";
  }

  return `<div class="flash flash-${flash.type}">${escapeHtml(flash.message)}</div>`;
}

function renderSupplierOptions() {
  if (!state.snapshot.suppliers.length) {
    return '<option value="">Create a supplier first</option>';
  }

  return state.snapshot.suppliers
    .map((supplier) => `<option value="${supplier.id}">${escapeHtml(supplier.name)}</option>`)
    .join("");
}

function renderDriverOptions() {
  const drivers = getDriverUsers().filter((user) => user.active);
  if (!drivers.length) {
    return '<option value="">Create a driver first</option>';
  }

  return drivers
    .map((driver) => `<option value="${driver.id}">${escapeHtml(driver.name)}</option>`)
    .join("");
}

function renderLocationOptions() {
  if (!state.snapshot.locations.length) {
    return '<option value="">Create a location first</option>';
  }

  return state.snapshot.locations
    .map((location) => {
      const supplierName = location.supplierName ? ` - ${escapeHtml(location.supplierName)}` : "";
      return `<option value="${location.id}">${escapeHtml(location.name)}${supplierName}</option>`;
    })
    .join("");
}

function getUsersByRole(role) {
  return state.snapshot.users.filter((user) => user.role === role);
}

function getDriverUsers() {
  if (state.snapshot.user && state.snapshot.user.role === "driver") {
    return [state.snapshot.user];
  }
  return state.snapshot.users.filter((user) => user.role === "driver");
}

function getLocation(locationId) {
  return state.snapshot.locations.find((location) => location.id === locationId) || null;
}

function getOrdersForDriver(driverUserId) {
  return state.snapshot.orders.filter((order) => order.driverUserId === driverUserId);
}

function getActiveOrders() {
  return state.snapshot.orders.filter((order) => order.status === "active");
}

function getCompletedOrders(driverUserId = "") {
  return state.snapshot.orders.filter((order) => {
    if (order.status !== "completed") {
      return false;
    }

    if (!driverUserId) {
      return true;
    }

    return order.driverUserId === driverUserId;
  });
}

function countActiveStops() {
  const activeStopKeys = new Set(
    getActiveOrders().map((order) => `${order.driverUserId}:${order.locationId}`),
  );
  return activeStopKeys.size;
}

function countDuplicateOrders(driverUserId) {
  const counts = {};
  getActiveOrders()
    .filter((order) => order.driverUserId === driverUserId)
    .forEach((order) => {
      const key = `${String(order.factoryOrderNumber || "").toLowerCase()}::${String(order.inhouseOrderNumber || "").toLowerCase()}`;
      counts[key] = (counts[key] || 0) + 1;
    });

  return Object.values(counts).filter((count) => count > 1).length;
}

function getRoutePlan(driverUserId) {
  const activeOrders = getOrdersForDriver(driverUserId)
    .filter((order) => order.status === "active")
    .sort(orderDisplaySort);

  const grouped = new Map();

  activeOrders.forEach((order) => {
    const location = getLocation(order.locationId);
    if (!location) {
      return;
    }

    if (!grouped.has(location.id)) {
      grouped.set(location.id, {
        id: location.id,
        location,
        orders: [],
        lat: Number(location.lat),
        lng: Number(location.lng),
      });
    }

    grouped.get(location.id).orders.push(order);
  });

  const orderedStops = optimizeRoute(Array.from(grouped.values()), {
    lat: HUB.lat,
    lng: HUB.lng,
  });

  const enrichedStops = [];
  let currentPoint = { lat: HUB.lat, lng: HUB.lng };
  orderedStops.forEach((stop) => {
    enrichedStops.push({
      ...stop,
      legKm: haversineKm(currentPoint, stop),
    });
    currentPoint = stop;
  });

  return {
    totalOrders: activeOrders.length,
    totalKm: totalRouteDistance(enrichedStops, { lat: HUB.lat, lng: HUB.lng }),
    stops: enrichedStops,
  };
}

function optimizeRoute(stops, start) {
  if (stops.length <= 1) {
    return stops;
  }

  const unvisited = [...stops];
  const ordered = [];
  let current = { lat: start.lat, lng: start.lng };

  while (unvisited.length) {
    let bestIndex = 0;
    let bestDistance = Number.POSITIVE_INFINITY;

    unvisited.forEach((candidate, index) => {
      const distance = haversineKm(current, candidate);
      if (distance < bestDistance) {
        bestDistance = distance;
        bestIndex = index;
      }
    });

    const [nextStop] = unvisited.splice(bestIndex, 1);
    ordered.push(nextStop);
    current = nextStop;
  }

  return twoOpt(ordered, start);
}

function twoOpt(route, start) {
  if (route.length < 4) {
    return route;
  }

  let best = [...route];
  let improved = true;

  while (improved) {
    improved = false;
    for (let i = 0; i < best.length - 1; i += 1) {
      for (let k = i + 1; k < best.length; k += 1) {
        const candidate = best
          .slice(0, i)
          .concat(best.slice(i, k + 1).reverse(), best.slice(k + 1));

        if (totalRouteDistance(candidate, start) + 0.01 < totalRouteDistance(best, start)) {
          best = candidate;
          improved = true;
        }
      }
    }
  }

  return best;
}

function totalRouteDistance(route, start) {
  if (!route.length) {
    return 0;
  }

  let total = 0;
  let current = start;
  route.forEach((stop) => {
    total += haversineKm(current, stop);
    current = stop;
  });
  return total;
}

function haversineKm(pointA, pointB) {
  const toRadians = (value) => (value * Math.PI) / 180;
  const earthRadiusKm = 6371;
  const dLat = toRadians(pointB.lat - pointA.lat);
  const dLng = toRadians(pointB.lng - pointA.lng);
  const lat1 = toRadians(pointA.lat);
  const lat2 = toRadians(pointB.lat);

  const a =
    Math.sin(dLat / 2) ** 2 +
    Math.cos(lat1) * Math.cos(lat2) * Math.sin(dLng / 2) ** 2;
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return earthRadiusKm * c;
}

function drawDriverRoute(driverUserId) {
  const canvas = document.getElementById("route-canvas");
  if (!(canvas instanceof HTMLCanvasElement)) {
    return;
  }

  const context = canvas.getContext("2d");
  if (!context) {
    return;
  }

  const plan = getRoutePlan(driverUserId);
  const points = [
    { label: HUB.label, lat: HUB.lat, lng: HUB.lng, isHub: true },
    ...plan.stops.map((stop, index) => ({
      label: `${index + 1}. ${stop.location.name}`,
      lat: Number(stop.location.lat),
      lng: Number(stop.location.lng),
      isHub: false,
    })),
  ];

  context.clearRect(0, 0, canvas.width, canvas.height);
  context.fillStyle = "#f6efe3";
  context.fillRect(0, 0, canvas.width, canvas.height);

  if (points.length === 1) {
    context.fillStyle = "#5b665e";
    context.font = "16px Trebuchet MS";
    context.fillText("No active route to draw yet.", 32, 40);
    return;
  }

  const padding = 48;
  const latitudes = points.map((point) => point.lat);
  const longitudes = points.map((point) => point.lng);
  const minLat = Math.min(...latitudes);
  const maxLat = Math.max(...latitudes);
  const minLng = Math.min(...longitudes);
  const maxLng = Math.max(...longitudes);

  const scale = (value, min, max, size) => {
    if (Math.abs(max - min) < 0.0001) {
      return size / 2;
    }
    return ((value - min) / (max - min)) * size;
  };

  const projected = points.map((point) => ({
    ...point,
    x: padding + scale(point.lng, minLng, maxLng, canvas.width - padding * 2),
    y:
      padding +
      (canvas.height - padding * 2 - scale(point.lat, minLat, maxLat, canvas.height - padding * 2)),
  }));

  context.strokeStyle = "rgba(21, 94, 82, 0.7)";
  context.lineWidth = 4;
  context.beginPath();
  projected.forEach((point, index) => {
    if (index === 0) {
      context.moveTo(point.x, point.y);
    } else {
      context.lineTo(point.x, point.y);
    }
  });
  context.stroke();

  projected.forEach((point, index) => {
    context.beginPath();
    context.fillStyle = point.isHub ? "#d9703d" : "#155e52";
    context.arc(point.x, point.y, point.isHub ? 11 : 9, 0, Math.PI * 2);
    context.fill();

    context.fillStyle = "#fff";
    context.font = "bold 12px Trebuchet MS";
    context.fillText(point.isHub ? "H" : String(index), point.x - 4, point.y + 4);

    context.fillStyle = "#16221d";
    context.font = "13px Trebuchet MS";
    context.fillText(point.label, point.x + 14, point.y - 10);
  });
}

function orderDisplaySort(left, right) {
  if (left.status !== right.status) {
    return left.status === "active" ? -1 : 1;
  }

  const leftCompleted = left.completedAt || "";
  const rightCompleted = right.completedAt || "";
  if (leftCompleted && rightCompleted && leftCompleted !== rightCompleted) {
    return rightCompleted.localeCompare(leftCompleted);
  }

  return right.createdAt.localeCompare(left.createdAt);
}

function getPaginationData(items, pageKey, pageSize) {
  const totalItems = items.length;
  const totalPages = Math.max(1, Math.ceil(totalItems / pageSize));
  const currentPage = clampPage(state.pagination[pageKey] || 1, totalPages);
  const startIndex = (currentPage - 1) * pageSize;

  if (state.pagination[pageKey] !== currentPage) {
    state.pagination[pageKey] = currentPage;
  }

  return {
    items: items.slice(startIndex, startIndex + pageSize),
    currentPage,
    totalItems,
    totalPages,
    startItem: totalItems ? startIndex + 1 : 0,
    endItem: Math.min(startIndex + pageSize, totalItems),
  };
}

function renderPaginationControls(pageKey, page) {
  if (page.totalPages <= 1) {
    return "";
  }

  return `
    <div class="pagination">
      <span class="pagination-summary">
        Showing ${page.startItem}-${page.endItem} of ${page.totalItems}
      </span>
      <div class="action-row">
        <button
          class="button button-ghost"
          data-action="change-page"
          data-page-key="${pageKey}"
          data-page="${page.currentPage - 1}"
          ${page.currentPage <= 1 || state.busy ? "disabled" : ""}
        >
          Previous
        </button>
        <span class="chip">Page ${page.currentPage} of ${page.totalPages}</span>
        <button
          class="button button-ghost"
          data-action="change-page"
          data-page-key="${pageKey}"
          data-page="${page.currentPage + 1}"
          ${page.currentPage >= page.totalPages || state.busy ? "disabled" : ""}
        >
          Next
        </button>
      </div>
    </div>
  `;
}

function clampPage(page, totalPages) {
  if (!Number.isFinite(page)) {
    return 1;
  }

  return Math.min(Math.max(Math.trunc(page), 1), totalPages);
}

function changePage(pageKey, page) {
  if (!Object.prototype.hasOwnProperty.call(state.pagination, pageKey)) {
    return;
  }

  state.pagination[pageKey] = Math.max(1, Math.trunc(page) || 1);
  render();
}

function exportOrdersCsv() {
  const sortedOrders = [...state.snapshot.orders].sort(orderDisplaySort);
  if (!sortedOrders.length) {
    showFlash("There are no entries to export yet.", "error");
    return;
  }

  const rows = sortedOrders.map((order) => [
    order.reference,
    order.driverName,
    order.locationName,
    order.entryType,
    order.factoryOrderNumber,
    order.inhouseOrderNumber,
    order.createdByName,
    order.status,
    getOrderNoticeText(order),
    formatDateTime(order.createdAt),
    formatDateTime(order.completedAt),
  ]);
  const csv = [
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
    ...rows,
  ]
    .map((columns) => columns.map(escapeCsvValue).join(","))
    .join("\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  const dateStamp = new Date().toISOString().slice(0, 10);

  anchor.href = url;
  anchor.download = `route-ledger-${dateStamp}.csv`;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
  showFlash("CSV export downloaded.", "success");
}

async function emailOrdersCsv() {
  if (!state.mailConfigured) {
    showFlash(state.mailConfigReason || "Email delivery is not configured yet.", "error");
    return;
  }

  state.busy = true;
  render();

  try {
    const payload = await requestJson(`${API_ROOT}/export/email`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ token: sessionToken }),
    });

    showFlash(`CSV emailed to ${payload?.sentTo || state.mailTo}.`, "success");
    state.busy = false;
    render();
  } catch (error) {
    state.busy = false;
    showFlash(normalizeError(error), "error");
    render();
  }
}

async function sendTestEmail() {
  if (!state.mailConfigured) {
    showFlash(state.mailConfigReason || "Email delivery is not configured yet.", "error");
    return;
  }

  state.busy = true;
  render();

  try {
    const payload = await requestJson(`${API_ROOT}/export/email/test`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ token: sessionToken }),
    });

    showFlash(`Test email sent to ${payload?.sentTo || state.mailTo}.`, "success");
    state.busy = false;
    render();
  } catch (error) {
    state.busy = false;
    showFlash(normalizeError(error), "error");
    render();
  }
}

function renderOrderNotice(order, emptyLabel = "") {
  const lines = getOrderNoticeLines(order);
  if (!lines.length) {
    return emptyLabel ? `<span class="muted">${escapeHtml(emptyLabel)}</span>` : "";
  }

  return `
    <div class="order-notice">
      ${lines.map((line) => `<span class="order-notice-line">${escapeHtml(line)}</span>`).join("")}
    </div>
  `;
}

function getOrderNoticeLines(order) {
  const lines = [];
  const notice = String(order?.notes || "").trim();
  const rolloverNotice = getRolloverNoticeText(order);

  if (notice) {
    lines.push(notice);
  }

  if (rolloverNotice) {
    lines.push(rolloverNotice);
  }

  return lines;
}

function getOrderNoticeText(order) {
  return getOrderNoticeLines(order).join(" | ");
}

function getRolloverNoticeText(order) {
  const carryOverCount = Number(order?.carryOverCount || 0);
  if (!carryOverCount) {
    return "";
  }

  const scheduledFor = formatDateOnly(order?.scheduledFor);
  const originalScheduledFor = formatDateOnly(order?.originalScheduledFor);
  const dayLabel = carryOverCount === 1 ? "day" : "days";

  if (scheduledFor && originalScheduledFor) {
    return `Rolled to ${scheduledFor} from ${originalScheduledFor} (${carryOverCount} ${dayLabel}).`;
  }

  return `Rolled to the next day (${carryOverCount} ${dayLabel}).`;
}

function escapeCsvValue(value) {
  const text = String(value || "");
  return `"${text.replaceAll('"', '""')}"`;
}

function showFlash(message, type) {
  flash = { message, type };
  if (flashTimer) {
    window.clearTimeout(flashTimer);
  }

  render();

  flashTimer = window.setTimeout(() => {
    flash = null;
    render();
  }, FLASH_TIMEOUT_MS);
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function capitalize(value) {
  if (!value) {
    return "";
  }
  return value.charAt(0).toUpperCase() + value.slice(1);
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
