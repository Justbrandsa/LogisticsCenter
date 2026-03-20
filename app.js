const SESSION_KEY = "route-ledger-session-token-v2";
const FLASH_TIMEOUT_MS = 3200;
const TIME_ZONE = "Africa/Johannesburg";
const API_ROOT = "/api";
const PAGE_SIZES = {
  globalEntries: 12,
  assignments: 12,
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
  artworkTo: "artwork3@giftwrap.co.za",
  currentPage: "",
  editingSupplierId: "",
  pagination: {
    globalEntries: 1,
    assignments: 1,
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
    stockItems: [],
    stockMovements: [],
    artworkRequests: [],
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
    state.artworkTo = status.artworkTo || state.artworkTo;

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
    artworkTo: payload?.artworkTo || "",
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
  state.editingSupplierId = "";
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
    if (state.editingSupplierId && !state.snapshot.suppliers.some((supplier) => supplier.id === state.editingSupplierId)) {
      state.editingSupplierId = "";
    }
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
    stockItems: [],
    stockMovements: [],
    artworkRequests: [],
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
    stockItems: Array.isArray(data?.stockItems) ? data.stockItems : [],
    stockMovements: Array.isArray(data?.stockMovements) ? data.stockMovements : [],
    artworkRequests: Array.isArray(data?.artworkRequests) ? data.artworkRequests : [],
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

  if (formId === "add-stock-item" && (currentUser.role === "admin" || currentUser.role === "logistics")) {
    await createStockItem(formData);
  }

  if (formId === "add-stock-movement" && (currentUser.role === "admin" || currentUser.role === "logistics")) {
    await recordStockMovement(formData);
  }

  if (formId === "request-artwork" && (currentUser.role === "admin" || currentUser.role === "logistics")) {
    await requestArtwork(formData);
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

  if (action === "edit-supplier" && currentUser.role === "admin") {
    state.editingSupplierId = String(button.dataset.supplierId || "");
    render();
    return;
  }

  if (action === "cancel-edit-supplier" && currentUser.role === "admin") {
    state.editingSupplierId = "";
    render();
    return;
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

  if (action === "save-order-assignment" && (currentUser.role === "admin" || currentUser.role === "sales")) {
    await saveOrderAssignment(button, currentUser);
    return;
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

  const orderForm = target.closest('form[data-form="add-order"]');
  if (
    orderForm instanceof HTMLFormElement
    && (target.matches('[name="entryType"]') || target.matches('[name="locationId"]'))
  ) {
    syncMoveToFactoryField(orderForm);
  }

  const stockMovementForm = target.closest('form[data-form="add-stock-movement"]');
  if (stockMovementForm instanceof HTMLFormElement && target.matches('[name="movementType"]')) {
    syncStockMovementFields(stockMovementForm);
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
    return true;
  } catch (error) {
    state.busy = false;
    showFlash(normalizeError(error), "error");
    render();
    return false;
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
  const supplierId = String(formData.get("supplierId") || "").trim();
  const name = String(formData.get("name") || "").trim();
  const contactPerson = String(formData.get("contactPerson") || "").trim();
  const contactNumber = String(formData.get("contactNumber") || "").trim();
  const factory = formData.get("factory") === "on";

  if (!name || !contactNumber) {
    showFlash("Supplier name and contact number are required.", "error");
    render();
    return;
  }

  const ok = await runMutation(
    supplierId ? "update_supplier" : "create_supplier",
    supplierId
      ? {
          p_token: sessionToken,
          p_supplier_id: supplierId,
          p_name: name,
          p_contact_person: contactPerson,
          p_contact_number: contactNumber,
          p_factory: factory,
        }
      : {
          p_token: sessionToken,
          p_name: name,
          p_contact_person: contactPerson,
          p_contact_number: contactNumber,
          p_factory: factory,
        },
    supplierId ? `Supplier updated: ${name}.` : `Supplier added: ${name}.`,
  );

  if (ok) {
    state.editingSupplierId = "";
    render();
  }
}

async function createLocation(formData) {
  const locationType = String(formData.get("locationType") || "").trim();
  const name = String(formData.get("name") || "").trim();
  const address = String(formData.get("address") || "").trim();
  const lat = Number(formData.get("lat"));
  const lng = Number(formData.get("lng"));
  const contactPerson = String(formData.get("contactPerson") || "").trim();
  const contactNumber = String(formData.get("contactNumber") || "").trim();

  if (!locationType || !name || !address || !contactNumber || Number.isNaN(lat) || Number.isNaN(lng)) {
    showFlash("Location name, type, address, coordinates, and contact number are required.", "error");
    render();
    return;
  }

  await runMutation(
    "create_location",
    {
      p_token: sessionToken,
      p_location_type: locationType,
      p_name: name,
      p_address: address,
      p_lat: lat,
      p_lng: lng,
      p_contact_person: contactPerson,
      p_contact_number: contactNumber,
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
  const moveToFactory = formData.get("moveToFactory") === "on";
  const allowDuplicate = formData.get("allowDuplicate") === "on";

  if (!locationId || !entryType || !factoryOrderNumber || !inhouseOrderNumber) {
    showFlash("Pickup location, entry type, factory order number, and in-house order number are required.", "error");
    render();
    return;
  }

  await runMutation(
    "create_order",
    {
      p_token: sessionToken,
      p_driver_user_id: driverUserId || null,
      p_location_id: locationId,
      p_entry_type: entryType,
      p_factory_order_number: factoryOrderNumber,
      p_inhouse_order_number: inhouseOrderNumber,
      p_notice: notice,
      p_move_to_factory: moveToFactory,
      p_allow_duplicate: currentUser.role === "admin" ? allowDuplicate : false,
    },
    driverUserId ? "Entry added to the driver list." : "Entry created in the unassigned queue.",
  );
}

async function saveOrderAssignment(button, currentUser) {
  const orderId = String(button.dataset.orderId || "").trim();
  if (!orderId) {
    return;
  }

  const driverField = document.querySelector(`[data-assignment-driver][data-order-id="${orderId}"]`);
  if (!(driverField instanceof HTMLSelectElement)) {
    return;
  }

  const allowDuplicateField = document.querySelector(`[data-assignment-allow-duplicate][data-order-id="${orderId}"]`);
  const driverUserId = String(driverField.value || "").trim();
  const selectedLabel = driverField.selectedOptions[0]?.textContent?.trim() || "Unassigned";

  await runMutation(
    "assign_order",
    {
      p_token: sessionToken,
      p_order_id: orderId,
      p_driver_user_id: driverUserId || null,
      p_allow_duplicate:
        currentUser.role === "admin" && allowDuplicateField instanceof HTMLInputElement
          ? allowDuplicateField.checked
          : false,
    },
    driverUserId
      ? `Entry assigned to ${selectedLabel}.`
      : "Entry moved to the unassigned queue.",
  );
}

async function createStockItem(formData) {
  const name = String(formData.get("name") || "").trim();
  const sku = String(formData.get("sku") || "").trim();
  const unit = String(formData.get("unit") || "").trim();
  const notes = String(formData.get("notes") || "").trim();

  if (!name) {
    showFlash("Stock item name is required.", "error");
    render();
    return;
  }

  await runMutation(
    "create_stock_item",
    {
      p_token: sessionToken,
      p_name: name,
      p_sku: sku,
      p_unit: unit || "units",
      p_notes: notes,
    },
    `Stock item added: ${name}.`,
  );
}

async function recordStockMovement(formData) {
  const stockItemId = String(formData.get("stockItemId") || "").trim();
  const movementType = String(formData.get("movementType") || "in").trim();
  const quantity = Number(formData.get("quantity"));
  const supplierName = String(formData.get("supplierName") || "").trim();
  const driverUserId = String(formData.get("driverUserId") || "").trim();
  const notes = String(formData.get("notes") || "").trim();

  if (!stockItemId || !movementType || !Number.isInteger(quantity) || quantity <= 0) {
    showFlash("Stock item, movement type, and a valid quantity are required.", "error");
    render();
    return;
  }

  if (movementType === "in" && !supplierName) {
    showFlash("Supplier is required for stock coming in.", "error");
    render();
    return;
  }

  if (movementType === "out" && !driverUserId) {
    showFlash("Driver is required for stock going out.", "error");
    render();
    return;
  }

  await runMutation(
    "record_stock_movement",
    {
      p_token: sessionToken,
      p_stock_item_id: stockItemId,
      p_movement_type: movementType,
      p_quantity: quantity,
      p_supplier_name: supplierName,
      p_driver_user_id: driverUserId || null,
      p_notes: notes,
    },
    movementType === "in" ? "Stock receipt recorded." : "Stock issue recorded.",
  );
}

async function requestArtwork(formData) {
  const stockItemId = String(formData.get("stockItemId") || "").trim();
  const requestedQuantity = Number(formData.get("requestedQuantity"));
  const notes = String(formData.get("notes") || "").trim();

  if (!stockItemId || !Number.isInteger(requestedQuantity) || requestedQuantity <= 0) {
    showFlash("Stock item and requested quantity are required.", "error");
    render();
    return;
  }

  state.busy = true;
  render();

  try {
    const payload = await requestJson(`${API_ROOT}/artwork/request`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        token: sessionToken,
        stockItemId,
        requestedQuantity,
        notes,
      }),
    });

    showFlash(`Artwork request emailed to ${payload?.sentTo || state.artworkTo}.`, "success");
    await refreshSnapshot();
  } catch (error) {
    state.busy = false;
    showFlash(normalizeError(error), "error");
    render();
  }
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
    syncPostRenderUi();
    return;
  }

  if (state.snapshot.user.role === "sales") {
    appEl.innerHTML = renderSalesScreen();
    syncPostRenderUi();
    return;
  }

  if (state.snapshot.user.role === "logistics") {
    appEl.innerHTML = renderLogisticsScreen();
    syncPostRenderUi();
    return;
  }

  appEl.innerHTML = renderDriverScreen();
  syncPostRenderUi();
  drawDriverRoute(state.snapshot.user.id);
}

function syncPostRenderUi() {
  const orderForm = document.querySelector('form[data-form="add-order"]');
  if (orderForm instanceof HTMLFormElement) {
    syncMoveToFactoryField(orderForm);
  }

  const stockMovementForm = document.querySelector('form[data-form="add-stock-movement"]');
  if (stockMovementForm instanceof HTMLFormElement) {
    syncStockMovementFields(stockMovementForm);
  }
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

  const summaryLine = currentUser.role === "logistics"
    ? `${state.snapshot.stockMovements.length} logged stock movement${state.snapshot.stockMovements.length === 1 ? "" : "s"}`
    : `${state.snapshot.orders.length} visible entr${state.snapshot.orders.length === 1 ? "y" : "ies"}`;

  authMetaEl.innerHTML = `
    <strong>${escapeHtml(currentUser.name)}</strong><br>
    <span class="muted">${capitalize(currentUser.role)} account</span><br>
    <span class="muted">${escapeHtml(summaryLine)}</span>
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
      { id: "assignments", label: "Assignments" },
      { id: "stock", label: "Stock" },
      { id: "network", label: "Network" },
      { id: "users", label: "Users" },
      { id: "drivers", label: "Driver Lists" },
    ];
  }

  if (role === "sales") {
    return [
      { id: "dashboard", label: "Dashboard" },
      { id: "entries", label: "Global List" },
      { id: "assignments", label: "Assignments" },
      { id: "drivers", label: "Driver Lists" },
    ];
  }

  if (role === "logistics") {
    return [
      { id: "dashboard", label: "Dashboard" },
      { id: "stock", label: "Stock" },
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
            Manage users, pickup locations, stock tracking, driver entries, and CSV delivery from one live database.
          </p>
          <div class="chip-row">
            <span class="chip chip-role-admin">Admin</span>
            <span class="chip">${state.snapshot.locations.length} locations</span>
            <span class="chip">${state.snapshot.stockItems.length} stock items</span>
          </div>
        </div>
        <div class="sidebar-card">
          <h3>Email delivery</h3>
          <p class="muted">
            ${
              state.mailConfigured
                ? `CSV email is ready for ${escapeHtml(state.mailTo)} and artwork requests go to ${escapeHtml(state.artworkTo)}.`
                : escapeHtml(state.mailConfigReason || "Email delivery is not configured yet.")
            }
          </p>
          <p class="muted">Use the Global List page for CSV delivery and the Stock page for artwork requests.</p>
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

function renderLogisticsScreen() {
  return `
    <section class="screen-grid">
      <aside class="sidebar">
        <div class="sidebar-card">
          <p class="eyebrow">Logistics scope</p>
          <h3>Stock control</h3>
          <p class="muted">
            Track incoming and outgoing stock, keep a live on-hand view, and request artwork from the artwork department.
          </p>
          <div class="chip-row">
            <span class="chip chip-role-logistics">Logistics</span>
            <span class="chip">${state.snapshot.stockItems.length} stock items</span>
            <span class="chip">${getStockOnHandTotal()} on hand</span>
          </div>
        </div>
        <div class="sidebar-card">
          <h3>Artwork requests</h3>
          <p class="muted">
            ${
              state.mailConfigured
                ? `Requests will email ${escapeHtml(state.artworkTo)} through the configured mail account.`
                : escapeHtml(state.mailConfigReason || "Email delivery is not configured yet.")
            }
          </p>
          <p class="muted">Use the Stock page to log movements and send artwork requests after the stock is queued.</p>
        </div>
      </aside>
      <div class="content">
        ${renderLogisticsPageContent()}
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
  const unassignedOrders = getUnassignedActiveOrders();
  const completedOrders = getCompletedOrders();

  if (state.currentPage === "entries") {
    return `
      <section class="hero-card">
        <p class="eyebrow">Global List</p>
        <h2>Create work and distribute the global list</h2>
        <p>Use this page to add live work with an optional driver assignment, then download or email the shared CSV.</p>
      </section>
      ${renderFlash()}
      <section class="panel">
        <p class="eyebrow">Global List</p>
        <h3 class="panel-title">Create a new entry</h3>
        <p class="panel-subtitle">Leave the driver unassigned to queue work for later dispatch. Admins can still override the duplicate rule when assigning to a driver.</p>
        ${renderEntryForm(state.snapshot.user, true)}
      </section>
      ${renderGlobalOrdersSection("admin")}
    `;
  }

  if (state.currentPage === "assignments") {
    return `
      <section class="hero-card">
        <p class="eyebrow">Assignments</p>
        <h2>Assign queued work to drivers</h2>
        <p>Review active entries, keep new work unassigned when needed, and move it onto driver lists once dispatch is ready.</p>
      </section>
      ${renderFlash()}
      <section class="metrics">
        ${renderMetric("Unassigned", unassignedOrders.length)}
        ${renderMetric("Assigned", activeOrders.length - unassignedOrders.length)}
        ${renderMetric("Active drivers", getActiveDriverUsers().length)}
      </section>
      ${renderAssignmentManager("admin")}
    `;
  }

  if (state.currentPage === "stock") {
    return renderStockWorkspace({
      viewerRole: "admin",
      title: "Stock tracking for logistics",
      subtitle: "Record stock coming in, stock going out with drivers, and artwork requests from one live workspace.",
    });
  }

  if (state.currentPage === "network") {
    return `
      <section class="hero-card">
        <p class="eyebrow">Network</p>
        <h2>Pickup locations</h2>
        <p>Maintain the physical pickup points used by the route lists.</p>
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
        <p>Create admin, sales, driver, and logistics accounts, then manage their status from one place.</p>
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
      ${renderMetric("Unassigned", unassignedOrders.length)}
      ${renderMetric("Drivers", getDriverUsers().length)}
      ${renderMetric("Completed entries", completedOrders.length)}
    </section>
    <section class="panel-grid">
      ${renderPageSummaryCard("Global List", "Add new work and send, test, or download the CSV register.", "entries")}
      ${renderPageSummaryCard("Assignments", "Allocate queued work to drivers and reassign active entries.", "assignments")}
      ${renderPageSummaryCard("Stock", "Track stock in/out and send artwork requests.", "stock")}
      ${renderPageSummaryCard("Network", "Maintain pickup locations.", "network")}
      ${renderPageSummaryCard("Users", "Manage admin, sales, driver, and logistics accounts.", "users")}
      ${renderPageSummaryCard("Driver lists", "Review active work separated by driver.", "drivers")}
    </section>
  `;
}

function renderSalesPageContent() {
  const ordersCreated = countOrdersCreatedByCurrentUser();
  const activeOrders = getActiveOrders();
  const unassignedOrders = getUnassignedActiveOrders();

  if (state.currentPage === "entries") {
    return `
      <section class="hero-card">
        <p class="eyebrow">Global List</p>
        <h2>Create work and manage the global list</h2>
        <p>Sales can add work with an optional driver assignment, then download or email the latest global list.</p>
      </section>
      ${renderFlash()}
      <section class="panel">
        <p class="eyebrow">Global List</p>
        <h3 class="panel-title">Create a new entry</h3>
        <p class="panel-subtitle">Leave the driver unassigned to hold the work for later dispatch. Duplicate protection still applies when a driver is selected.</p>
        ${renderEntryForm(state.snapshot.user, false)}
      </section>
      ${renderGlobalOrdersSection("sales")}
    `;
  }

  if (state.currentPage === "assignments") {
    return `
      <section class="hero-card">
        <p class="eyebrow">Assignments</p>
        <h2>Allocate active work to drivers</h2>
        <p>Use this page to work through the unassigned queue and reassign active entries when the route plan changes.</p>
      </section>
      ${renderFlash()}
      <section class="metrics">
        ${renderMetric("Unassigned", unassignedOrders.length)}
        ${renderMetric("Assigned", activeOrders.length - unassignedOrders.length)}
        ${renderMetric("Active drivers", getActiveDriverUsers().length)}
      </section>
      ${renderAssignmentManager("sales")}
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
      ${renderMetric("Unassigned", unassignedOrders.length)}
    </section>
    <section class="panel-grid">
      ${renderPageSummaryCard("Global List", "Create new work and email, test, or download the CSV register.", "entries")}
      ${renderPageSummaryCard("Assignments", "Assign queued work to drivers and rebalance active entries.", "assignments")}
      ${renderPageSummaryCard("Driver lists", "Review active work separated by driver.", "drivers")}
    </section>
  `;
}

function renderLogisticsPageContent() {
  const stockItems = state.snapshot.stockItems;
  const stockMovements = state.snapshot.stockMovements;
  const artworkRequests = state.snapshot.artworkRequests;

  if (state.currentPage === "stock") {
    return renderStockWorkspace({
      viewerRole: "logistics",
      title: "Logistics stock workspace",
      subtitle: "Manage stock in, stock out, and artwork requests without leaving the logistics workflow.",
    });
  }

  return `
    <section class="hero-card">
      <p class="eyebrow">Dashboard</p>
      <h2>Live stock tracking for logistics</h2>
      <p>Keep a running ledger of stock receipts and issues, then request artwork from the artwork department when new work is needed.</p>
    </section>
    ${renderFlash()}
    <section class="metrics">
      ${renderMetric("Stock items", stockItems.length)}
      ${renderMetric("On hand", getStockOnHandTotal())}
      ${renderMetric("Movements", stockMovements.length)}
      ${renderMetric("Artwork requests", artworkRequests.length)}
    </section>
    <section class="panel-grid">
      ${renderPageSummaryCard("Stock", "Log stock in/out and send artwork requests.", "stock")}
    </section>
  `;
}

function renderStockWorkspace({ viewerRole, title, subtitle }) {
  const stockMovements = state.snapshot.stockMovements;
  const inboundCount = stockMovements.filter((movement) => movement.movementType === "in").length;
  const outboundCount = stockMovements.filter((movement) => movement.movementType === "out").length;

  return `
    <section class="hero-card">
      <p class="eyebrow">Stock</p>
      <h2>${escapeHtml(title)}</h2>
      <p>${escapeHtml(subtitle)}</p>
    </section>
    ${renderFlash()}
    <section class="metrics">
      ${renderMetric("Stock items", state.snapshot.stockItems.length)}
      ${renderMetric("On hand", getStockOnHandTotal())}
      ${renderMetric("Stock in", inboundCount)}
      ${renderMetric("Stock out", outboundCount)}
      ${renderMetric("Artwork requests", state.snapshot.artworkRequests.length)}
    </section>
    <section class="panel-grid">
      ${renderStockItemPanel()}
      ${renderStockMovementPanel()}
    </section>
    ${renderArtworkRequestPanel(viewerRole)}
    ${renderStockItemsSection()}
    ${renderStockMovementsSection()}
    ${renderArtworkRequestsSection()}
  `;
}

function renderStockItemPanel() {
  return `
    <article class="panel">
      <p class="eyebrow">Stock master</p>
      <h3 class="panel-title">Add stock item</h3>
      <p class="panel-subtitle">Create the item once, then log every stock movement against it.</p>
      <form data-form="add-stock-item">
        <label>
          Item name
          <input name="name" type="text" required>
        </label>
        <div class="form-grid">
          <label>
            Stock code
            <input name="sku" type="text" placeholder="SKU-001">
          </label>
          <label>
            Unit
            <input name="unit" type="text" value="units">
          </label>
        </div>
        <label>
          Notes
          <textarea name="notes" placeholder="Material, finish, size, or any internal note"></textarea>
        </label>
        <button type="submit" class="button button-primary"${state.busy ? " disabled" : ""}>
          Add stock item
        </button>
      </form>
    </article>
  `;
}

function renderStockMovementPanel() {
  return `
    <article class="panel">
      <p class="eyebrow">Stock ledger</p>
      <h3 class="panel-title">Log stock in or out</h3>
      <p class="panel-subtitle">Every stock movement is timestamped and linked to the staff member who recorded it.</p>
      <form data-form="add-stock-movement">
        <div class="form-grid">
          <label>
            Stock item
            <select name="stockItemId" required>
              ${renderStockItemOptions()}
            </select>
          </label>
          <label>
            Movement type
            <select name="movementType" required>
              <option value="in">Stock in</option>
              <option value="out">Stock out</option>
            </select>
          </label>
        </div>
        <div class="form-grid">
          <label>
            Quantity
            <input name="quantity" type="number" min="1" step="1" required>
          </label>
          <label data-stock-supplier-field>
            Supplier
            <input name="supplierName" type="text" placeholder="Supplier name">
          </label>
          <label data-stock-driver-field class="hidden">
            Driver
            <select name="driverUserId">
              ${renderDriverOptions("", true)}
            </select>
          </label>
        </div>
        <p class="field-note" data-stock-movement-note>
          Record which supplier the stock arrived from.
        </p>
        <label>
          Notes
          <textarea name="notes" placeholder="Batch note, dispatch note, or anything the team should keep on record"></textarea>
        </label>
        <button type="submit" class="button button-primary"${state.busy ? " disabled" : ""}>
          Save movement
        </button>
      </form>
    </article>
  `;
}

function renderArtworkRequestPanel(viewerRole) {
  return `
    <article class="panel">
      <p class="eyebrow">Artwork</p>
      <h3 class="panel-title">Email artwork request</h3>
      <p class="panel-subtitle">
        Send a request to ${escapeHtml(state.artworkTo)} so the artwork department can prepare what logistics needs.
      </p>
      <form data-form="request-artwork">
        <div class="form-grid">
          <label>
            Stock item
            <select name="stockItemId" required>
              ${renderStockItemOptions()}
            </select>
          </label>
          <label>
            Requested quantity
            <input name="requestedQuantity" type="number" min="1" step="1" required>
          </label>
        </div>
        <label>
          Notes
          <textarea name="notes" placeholder="Artwork size, finish, due date, or any design instruction"></textarea>
        </label>
        <button type="submit" class="button button-secondary"${state.busy || !state.mailConfigured ? " disabled" : ""}>
          Send artwork request
        </button>
        ${
          !state.mailConfigured
            ? `<p class="field-note">${escapeHtml(state.mailConfigReason || "Email delivery is not configured yet.")}</p>`
            : ""
        }
      </form>
    </article>
  `;
}

function renderStockItemsSection() {
  return `
    <section class="table-card">
      <p class="eyebrow">Current stock</p>
      <h3 class="panel-title">On-hand summary</h3>
      <div class="table-scroll">
        <table>
          <thead>
            <tr>
              <th>Item</th>
              <th>Stock code</th>
              <th>Unit</th>
              <th>On hand</th>
              <th>Notes</th>
              <th>Updated</th>
            </tr>
          </thead>
          <tbody>
            ${
              state.snapshot.stockItems.length
                ? state.snapshot.stockItems.map((item) => renderStockItemRow(item)).join("")
                : `
                  <tr>
                    <td colspan="6">No stock items added yet.</td>
                  </tr>
                `
            }
          </tbody>
        </table>
      </div>
    </section>
  `;
}

function renderStockMovementsSection() {
  return `
    <section class="table-card">
      <p class="eyebrow">Stock history</p>
      <h3 class="panel-title">Movement ledger</h3>
      <div class="table-scroll">
        <table>
          <thead>
            <tr>
              <th>Logged</th>
              <th>Item</th>
              <th>Type</th>
              <th>Quantity</th>
              <th>Supplier / driver</th>
              <th>Notes</th>
              <th>Recorded by</th>
            </tr>
          </thead>
          <tbody>
            ${
              state.snapshot.stockMovements.length
                ? state.snapshot.stockMovements.map((movement) => renderStockMovementRow(movement)).join("")
                : `
                  <tr>
                    <td colspan="7">No stock movements logged yet.</td>
                  </tr>
                `
            }
          </tbody>
        </table>
      </div>
    </section>
  `;
}

function renderArtworkRequestsSection() {
  return `
    <section class="table-card">
      <p class="eyebrow">Artwork requests</p>
      <h3 class="panel-title">Email request log</h3>
      <div class="table-scroll">
        <table>
          <thead>
            <tr>
              <th>Sent</th>
              <th>Item</th>
              <th>Quantity</th>
              <th>Requested by</th>
              <th>Sent to</th>
              <th>Notes</th>
            </tr>
          </thead>
          <tbody>
            ${
              state.snapshot.artworkRequests.length
                ? state.snapshot.artworkRequests.map((request) => renderArtworkRequestRow(request)).join("")
                : `
                  <tr>
                    <td colspan="6">No artwork requests sent yet.</td>
                  </tr>
                `
            }
          </tbody>
        </table>
      </div>
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
      <p class="panel-subtitle">Create admin, sales, driver, or logistics accounts with name-based login.</p>
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
              <option value="logistics">Logistics</option>
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
      <p class="eyebrow">Location network</p>
      <h3 class="panel-title">Pickup locations</h3>
      <p class="panel-subtitle">Each location now stores its own type and contact details.</p>
      <article class="panel">
        <p class="eyebrow">Locations</p>
        <h3 class="panel-title">Add location</h3>
        <form data-form="add-location">
          <div class="form-grid">
            <label>
              Location Name
              <input name="name" type="text" required>
            </label>
            <label>
              Supplier or factory
              <select name="locationType" required>
                <option value="supplier">Supplier</option>
                <option value="factory">Factory</option>
                <option value="both">Both</option>
              </select>
            </label>
          </div>
          <label>
            Physical address
            <input name="address" type="text" required>
          </label>
          <div class="form-grid">
            <label>
              Latitude
              <input name="lat" type="number" step="0.000001" required>
            </label>
            <label>
              Longitude
              <input name="lng" type="number" step="0.000001" required>
            </label>
          </div>
          <div class="form-grid">
            <label>
              Contact person
              <input name="contactPerson" type="text">
            </label>
            <label>
              Contact number
              <input name="contactNumber" type="tel" required>
            </label>
          </div>
          <button type="submit" class="button button-primary"${state.busy ? " disabled" : ""}>
            Add location
          </button>
        </form>
        <div class="table-scroll">
          <table>
            <thead>
              <tr>
                <th>Location</th>
                <th>Supplier or factory</th>
                <th>Physical address</th>
                <th>Contact</th>
                <th>Action</th>
              </tr>
            </thead>
            <tbody>
              ${
                state.snapshot.locations.length
                  ? state.snapshot.locations.map((location) => renderLocationRow(location)).join("")
                  : `
                    <tr>
                      <td colspan="5">No locations added yet.</td>
                    </tr>
                  `
              }
            </tbody>
          </table>
        </div>
      </article>
    </section>
  `;
}

function renderUsersSection() {
  return `
    <section class="table-card">
      <p class="eyebrow">Users</p>
      <h3 class="panel-title">Team account management</h3>
      <p class="panel-subtitle">Admin, sales, driver, and logistics accounts all appear in this list.</p>
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
          <select name="driverUserId">
            ${renderDriverOptions("", true)}
          </select>
        </label>
        <label>
          Pickup location
          <select name="locationId" required>
            ${renderLocationOptions()}
          </select>
        </label>
      </div>
      <p class="field-note">
        Leave the driver as Unassigned if dispatch will allocate it later.
      </p>
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
      <label class="inline-check">
        <input type="checkbox" name="moveToFactory" disabled>
        Move collected stock to a factory
      </label>
      <p class="field-note" data-move-to-factory-note>
        Switch the entry type to Collection to enable factory transfer.
      </p>
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

function renderAssignmentManager(viewerRole) {
  const activeOrders = [...getActiveOrders()].sort(orderAssignmentSort);
  const page = getPaginationData(activeOrders, "assignments", PAGE_SIZES.assignments);

  return `
    <section class="table-card">
      <div class="table-toolbar">
        <div>
          <p class="eyebrow">Assignments</p>
          <h3 class="panel-title">Active entry allocation</h3>
          <p class="panel-subtitle">Unassigned entries appear first. Move work onto a driver list when the route is ready, or return it to the queue.</p>
        </div>
      </div>
      <div class="table-scroll">
        <table>
          <thead>
            <tr>
              <th>Entry</th>
              <th>Current driver</th>
              <th>Pickup location</th>
              <th>Created by</th>
              <th>Assign to</th>
              <th>Action</th>
            </tr>
          </thead>
          <tbody>
            ${
              page.items.length
                ? page.items.map((order) => renderAssignmentRow(order, viewerRole)).join("")
                : `
                  <tr>
                    <td colspan="6">No active entries available for assignment.</td>
                  </tr>
                `
            }
          </tbody>
        </table>
      </div>
      ${renderPaginationControls("assignments", page)}
    </section>
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
              <th>Move to factory</th>
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
                    <td colspan="11">No entries available yet.</td>
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
      <td>${renderMoveToFactoryValue(order)}</td>
      <td>${renderDriverAssignmentValue(order)}</td>
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

function renderAssignmentRow(order, viewerRole) {
  return `
    <tr>
      <td>
        <strong>${escapeHtml(order.reference)}</strong><br>
        <span class="muted">In-house ${escapeHtml(order.inhouseOrderNumber || "")}</span><br>
        <span class="muted">Factory ${escapeHtml(order.factoryOrderNumber || "")}</span>
        <div class="chip-row">
          ${renderTypeChip(order.entryType)}
          ${order.moveToFactory ? '<span class="chip chip-warning">Factory move</span>' : ""}
        </div>
        ${renderOrderNotice(order)}
      </td>
      <td>${renderDriverAssignmentValue(order)}</td>
      <td>
        <strong>${escapeHtml(order.locationName || "Unknown")}</strong><br>
        <span class="muted">${escapeHtml(order.locationAddress || "")}</span>
      </td>
      <td>
        ${escapeHtml(order.createdByName || "Unknown")}<br>
        <span class="muted">${escapeHtml(formatDateTime(order.createdAt))}</span>
      </td>
      <td>
        <div class="assignment-control">
          <select data-assignment-driver data-order-id="${order.id}">
            ${renderDriverOptions(order.driverUserId || "", true)}
          </select>
          ${
            viewerRole === "admin"
              ? `
                <label class="inline-check assignment-inline">
                  <input type="checkbox" data-assignment-allow-duplicate data-order-id="${order.id}">
                  Allow duplicate
                </label>
              `
              : ""
          }
        </div>
      </td>
      <td>
        <button class="button button-primary" data-action="save-order-assignment" data-order-id="${order.id}"${state.busy ? " disabled" : ""}>
          ${order.driverUserId ? "Save" : "Assign"}
        </button>
      </td>
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
      <td>${escapeHtml(supplier.contactPerson || "Not set")}</td>
      <td>${escapeHtml(supplier.contactNumber || "Not set")}</td>
      <td>${supplier.factory ? "Yes" : "No"}</td>
      <td>${locationCount}</td>
      <td>
        <div class="action-row">
          <button class="button button-secondary" data-action="edit-supplier" data-supplier-id="${supplier.id}"${state.busy ? " disabled" : ""}>
            Edit
          </button>
          <button class="button button-danger" data-action="delete-supplier" data-supplier-id="${supplier.id}"${state.busy ? " disabled" : ""}>
            Delete
          </button>
        </div>
      </td>
    </tr>
  `;
}

function renderLocationRow(location) {
  return `
    <tr>
      <td>
        <strong>${escapeHtml(location.name)}</strong>
      </td>
      <td>${escapeHtml(capitalize(location.locationType || ""))}</td>
      <td>${escapeHtml(location.address)}</td>
      <td>${escapeHtml(location.contactPerson || "Not set")}<br><span class="muted">${escapeHtml(location.contactNumber || "Not set")}</span></td>
      <td>
        <button class="button button-danger" data-action="delete-location" data-location-id="${location.id}"${state.busy ? " disabled" : ""}>
          Delete
        </button>
      </td>
    </tr>
  `;
}

function renderStockItemRow(item) {
  return `
    <tr>
      <td>${escapeHtml(item.name)}</td>
      <td>${escapeHtml(item.sku || "Not set")}</td>
      <td>${escapeHtml(item.unit || "units")}</td>
      <td>${escapeHtml(String(item.onHandQuantity || 0))}</td>
      <td>${escapeHtml(item.notes || "None")}</td>
      <td>${escapeHtml(formatDateTime(item.updatedAt || item.createdAt) || "Not updated")}</td>
    </tr>
  `;
}

function renderStockMovementRow(movement) {
  return `
    <tr>
      <td>${escapeHtml(formatDateTime(movement.createdAt) || "")}</td>
      <td>
        <strong>${escapeHtml(movement.itemName || "Unknown")}</strong><br>
        <span class="muted">${escapeHtml(movement.sku || "No stock code")}</span>
      </td>
      <td>${movement.movementType === "in" ? '<span class="chip chip-success">Stock in</span>' : '<span class="chip chip-warning">Stock out</span>'}</td>
      <td>${escapeHtml(String(movement.quantity || 0))} ${escapeHtml(movement.unit || "units")}</td>
      <td>${escapeHtml(getStockMovementPartyLabel(movement))}</td>
      <td>${escapeHtml(movement.notes || "None")}</td>
      <td>${escapeHtml(movement.createdByName || "Unknown")}</td>
    </tr>
  `;
}

function renderArtworkRequestRow(request) {
  return `
    <tr>
      <td>${escapeHtml(formatDateTime(request.sentAt) || "")}</td>
      <td>
        <strong>${escapeHtml(request.itemName || "Unknown")}</strong><br>
        <span class="muted">${escapeHtml(request.sku || "No stock code")}</span>
      </td>
      <td>${escapeHtml(String(request.requestedQuantity || 0))}</td>
      <td>${escapeHtml(request.requestedByName || "Unknown")}</td>
      <td>${escapeHtml(request.sentTo || state.artworkTo)}</td>
      <td>${escapeHtml(request.notes || "None")}</td>
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
                  ${order.moveToFactory ? '<span class="chip chip-warning">Factory move</span>' : ""}
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

function renderMoveToFactoryValue(order) {
  const label = getMoveToFactoryLabel(order);
  if (!label) {
    return '<span class="muted">No</span>';
  }

  return escapeHtml(label);
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

function renderDriverOptions(selectedDriverId = "", includeUnassigned = false) {
  const drivers = getActiveDriverUsers();
  const options = [];

  if (includeUnassigned) {
    options.push(`<option value=""${selectedDriverId ? "" : " selected"}>Unassigned</option>`);
  }

  if (!drivers.length) {
    return options.length ? options.join("") : '<option value="">Create a driver first</option>';
  }

  return options
    .concat(
      drivers.map(
        (driver) => `<option value="${driver.id}"${driver.id === selectedDriverId ? " selected" : ""}>${escapeHtml(driver.name)}</option>`,
      ),
    )
    .join("");
}

function renderStockItemOptions(selectedStockItemId = "") {
  if (!state.snapshot.stockItems.length) {
    return '<option value="">Create a stock item first</option>';
  }

  return state.snapshot.stockItems
    .map((item) => {
      const label = item.sku ? `${item.name} (${item.sku})` : item.name;
      return `<option value="${item.id}"${item.id === selectedStockItemId ? " selected" : ""}>${escapeHtml(label)}</option>`;
    })
    .join("");
}

function renderLocationOptions() {
  if (!state.snapshot.locations.length) {
    return '<option value="">Create a location first</option>';
  }

  return state.snapshot.locations
    .map((location) => {
      const typeLabel = location.locationType ? ` - ${escapeHtml(capitalize(location.locationType))}` : "";
      return `<option value="${location.id}">${escapeHtml(location.name)}${typeLabel}</option>`;
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

function getActiveDriverUsers() {
  return getDriverUsers().filter((user) => user.active);
}

function getLocation(locationId) {
  return state.snapshot.locations.find((location) => location.id === locationId) || null;
}

function getSupplier(supplierId) {
  return state.snapshot.suppliers.find((supplier) => supplier.id === supplierId) || null;
}

function getEditingSupplier() {
  return state.editingSupplierId ? getSupplier(state.editingSupplierId) : null;
}

function getOrdersForDriver(driverUserId) {
  return state.snapshot.orders.filter((order) => order.driverUserId === driverUserId);
}

function getActiveOrders() {
  return state.snapshot.orders.filter((order) => order.status === "active");
}

function getUnassignedActiveOrders() {
  return getActiveOrders().filter((order) => !order.driverUserId);
}

function getStockOnHandTotal() {
  return state.snapshot.stockItems.reduce((sum, item) => sum + Number(item.onHandQuantity || 0), 0);
}

function getStockMovementPartyLabel(movement) {
  if (movement?.movementType === "in") {
    return movement?.supplierName || "Supplier not set";
  }

  return movement?.driverName || "Driver not set";
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
    getActiveOrders()
      .filter((order) => order.driverUserId)
      .map((order) => `${order.driverUserId}:${order.locationId}`),
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

function orderAssignmentSort(left, right) {
  const leftAssigned = Boolean(left.driverUserId);
  const rightAssigned = Boolean(right.driverUserId);
  if (leftAssigned !== rightAssigned) {
    return leftAssigned ? 1 : -1;
  }

  return orderDisplaySort(left, right);
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
    getDriverDisplayName(order),
    order.locationName,
    order.entryType,
    getMoveToFactoryLabel(order) || "No",
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
      "Move to factory",
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
  const moveToFactory = getMoveToFactoryText(order);
  const rolloverNotice = getRolloverNoticeText(order);

  if (notice) {
    lines.push(notice);
  }

  if (moveToFactory) {
    lines.push(moveToFactory);
  }

  if (rolloverNotice) {
    lines.push(rolloverNotice);
  }

  return lines;
}

function getOrderNoticeText(order) {
  return getOrderNoticeLines(order).join(" | ");
}

function getMoveToFactoryLabel(order) {
  return order?.moveToFactory ? "Yes" : "";
}

function getMoveToFactoryText(order) {
  const label = getMoveToFactoryLabel(order);
  if (!label) {
    return "";
  }

  return "Move collected stock to a factory.";
}

function getDriverDisplayName(order) {
  return String(order?.driverName || "").trim() || "Unassigned";
}

function renderDriverAssignmentValue(order) {
  if (order?.driverUserId) {
    return escapeHtml(getDriverDisplayName(order));
  }

  return '<span class="chip chip-warning">Unassigned</span>';
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

function syncMoveToFactoryField(form) {
  const entryTypeField = form.querySelector('[name="entryType"]');
  const moveToFactoryField = form.querySelector('[name="moveToFactory"]');
  const noteEl = form.querySelector("[data-move-to-factory-note]");

  if (
    !(entryTypeField instanceof HTMLSelectElement)
    || !(moveToFactoryField instanceof HTMLInputElement)
    || !(noteEl instanceof HTMLElement)
  ) {
    return;
  }

  const isCollection = entryTypeField.value === "collection";
  moveToFactoryField.disabled = !isCollection;
  if (!isCollection) {
    moveToFactoryField.checked = false;
  }

  if (!isCollection) {
    noteEl.textContent = "Switch the entry type to Collection to enable factory transfer.";
    return;
  }

  noteEl.textContent = "Check this when the collected stock must be moved to a factory.";
}

function syncStockMovementFields(form) {
  const movementTypeField = form.querySelector('[name="movementType"]');
  const supplierField = form.querySelector("[data-stock-supplier-field]");
  const driverField = form.querySelector("[data-stock-driver-field]");
  const noteEl = form.querySelector("[data-stock-movement-note]");

  if (
    !(movementTypeField instanceof HTMLSelectElement)
    || !(supplierField instanceof HTMLElement)
    || !(driverField instanceof HTMLElement)
    || !(noteEl instanceof HTMLElement)
  ) {
    return;
  }

  const supplierInput = supplierField.querySelector('[name="supplierName"]');
  const driverSelect = driverField.querySelector('[name="driverUserId"]');
  const isInbound = movementTypeField.value === "in";

  supplierField.classList.toggle("hidden", !isInbound);
  driverField.classList.toggle("hidden", isInbound);

  if (supplierInput instanceof HTMLInputElement) {
    supplierInput.required = isInbound;
    if (!isInbound) {
      supplierInput.value = "";
    }
  }

  if (driverSelect instanceof HTMLSelectElement) {
    driverSelect.required = !isInbound;
    if (isInbound) {
      driverSelect.value = "";
    }
  }

  noteEl.textContent = isInbound
    ? "Record which supplier the stock arrived from."
    : "Record which driver the stock went out with.";
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
