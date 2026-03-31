const SESSION_KEY = "route-ledger-session-token-v2";
const FLASH_TIMEOUT_MS = 3200;
const TIME_ZONE = "Africa/Johannesburg";
const API_ROOT = "/api";
const STOCK_QR_TYPE = "route-ledger-stock";
const STOCK_QR_VERSION = 1;
const STOCK_SCANNER_FORMATS = ["qr_code", "code_128"];
const STOCK_LABEL_SKU_MAX_LENGTH = 10;
const STOCK_LABEL_CODE_LENGTH = 8;
const STOCK_LABEL_ALPHABET = "0123456789ABCDEFGHJKMNPQRSTVWXYZ";
const STOCK_RECENT_ACTIVITY_DAYS = 7;
const UUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
const ORDER_FLAG_LABELS = Object.freeze({
  not_collected: "Not collected",
  not_ready: "Not yet ready",
});
const DEFAULT_ORDER_PRIORITY = "medium";
const PRIORITY_STOP_VALUE = "high";
const ORDER_PRIORITY_LEVELS = Object.freeze({
  high: 0,
  medium: 1,
  low: 2,
});
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
  source: "hub",
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
  editingUserId: "",
  editingSupplierId: "",
  editingLocationId: "",
  editingStockItemId: "",
  editingStockMovementId: "",
  stockMovementSelectedItemId: "",
  stockQrItemId: "",
  stockQrSvg: "",
  stockQrBusy: false,
  stockQrSharing: false,
  stockScannerOpen: false,
  stockScannerStatus: "",
  stockScannerPendingItemId: "",
  stockScannerPendingCode: "",
  stockSearchQuery: "",
  stockMovementsSectionOpen: false,
  stockArtworkPanelOpen: false,
  stockArtworkRequestsSectionOpen: false,
  stockOpenItemCards: {},
  stockScannerZoomSupported: false,
  stockScannerZoomValue: 1,
  stockScannerZoomMin: 1,
  stockScannerZoomMax: 1,
  stockScannerZoomStep: 0.1,
  assignmentDriverFilter: "",
  flaggingOrderId: "",
  driverRouteOrigin: null,
  driverRouteOriginLoading: false,
  driverRouteOriginAttempted: false,
  driverRouteOriginError: "",
  driverRouteOriginUserId: "",
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
let stockScannerStream = null;
let stockScannerAnimationFrame = 0;
let stockScannerStarting = false;
let stockScannerStopRequested = false;
let stockScannerDetector = null;
let stockScannerVideoTrack = null;
let routeMap = null;
let routeMapContainer = null;
let routeMapLayers = null;

document.addEventListener("submit", handleSubmit);
document.addEventListener("click", handleClick);
document.addEventListener("change", handleChange);
document.addEventListener("input", handleInput);
document.addEventListener("keydown", handleKeyDown);
window.addEventListener("hashchange", handleHashChange);
window.addEventListener("beforeunload", () => {
  void stopStockScanner();
});

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
  state.editingUserId = "";
  state.editingSupplierId = "";
  state.editingLocationId = "";
  state.editingStockItemId = "";
  state.editingStockMovementId = "";
  state.stockMovementSelectedItemId = "";
  state.stockQrItemId = "";
  state.stockQrSvg = "";
  state.stockQrBusy = false;
  state.stockQrSharing = false;
  state.stockScannerOpen = false;
  state.stockScannerStatus = "";
  state.stockScannerPendingItemId = "";
  state.stockScannerPendingCode = "";
  state.assignmentDriverFilter = "";
  state.flaggingOrderId = "";
  resetDriverRouteOrigin();
  void stopStockScanner();
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
    if (!state.snapshot.user || state.snapshot.user.role !== "driver") {
      resetDriverRouteOrigin();
    } else if (state.driverRouteOriginUserId && state.driverRouteOriginUserId !== state.snapshot.user.id) {
      resetDriverRouteOrigin();
    }
    if (state.editingUserId && !state.snapshot.users.some((user) => user.id === state.editingUserId)) {
      state.editingUserId = "";
    }
    if (state.editingSupplierId && !state.snapshot.suppliers.some((supplier) => supplier.id === state.editingSupplierId)) {
      state.editingSupplierId = "";
    }
    if (state.editingLocationId && !state.snapshot.locations.some((location) => location.id === state.editingLocationId)) {
      state.editingLocationId = "";
    }
    if (state.editingStockItemId && !state.snapshot.stockItems.some((item) => item.id === state.editingStockItemId)) {
      state.editingStockItemId = "";
    }
    if (state.editingStockMovementId && !state.snapshot.stockMovements.some((movement) => movement.id === state.editingStockMovementId)) {
      state.editingStockMovementId = "";
    }
    if (state.stockMovementSelectedItemId && !state.snapshot.stockItems.some((item) => item.id === state.stockMovementSelectedItemId)) {
      state.stockMovementSelectedItemId = "";
    }
    if (state.stockQrItemId && !state.snapshot.stockItems.some((item) => item.id === state.stockQrItemId)) {
      state.stockQrItemId = "";
      state.stockQrSvg = "";
      state.stockQrBusy = false;
      state.stockQrSharing = false;
    }
    if (state.stockScannerPendingItemId && !state.snapshot.stockItems.some((item) => item.id === state.stockScannerPendingItemId)) {
      state.stockScannerPendingItemId = "";
      state.stockScannerPendingCode = "";
    }
    if (state.flaggingOrderId) {
      const flaggedOrder = state.snapshot.orders.find((order) => order.id === state.flaggingOrderId && order.status === "active");
      if (!flaggedOrder) {
        state.flaggingOrderId = "";
      }
    }
    if (
      state.assignmentDriverFilter
      && state.assignmentDriverFilter !== "unassigned"
      && !state.snapshot.users.some((user) => user.id === state.assignmentDriverFilter && user.role === "driver")
    ) {
      state.assignmentDriverFilter = "";
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

async function requestText(url, options = {}) {
  const response = await fetch(url, options);
  const payload = await response.text();

  if (!response.ok) {
    const contentType = String(response.headers.get("Content-Type") || "").toLowerCase();
    if (contentType.includes("application/json")) {
      try {
        const parsed = JSON.parse(payload);
        throw new Error(parsed?.error || `Request failed with status ${response.status}.`);
      } catch (error) {
        throw new Error(normalizeError(error) || `Request failed with status ${response.status}.`);
      }
    }

    throw new Error(payload || `Request failed with status ${response.status}.`);
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

function resetDriverRouteOrigin() {
  state.driverRouteOrigin = null;
  state.driverRouteOriginLoading = false;
  state.driverRouteOriginAttempted = false;
  state.driverRouteOriginError = "";
  state.driverRouteOriginUserId = "";
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

  if (formId === "flag-order" && (currentUser.role === "admin" || currentUser.role === "driver")) {
    await saveOrderFlag(formData);
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

  if (action === "edit-user" && currentUser.role === "admin") {
    state.editingUserId = String(button.dataset.userId || "");
    render();
    return;
  }

  if (action === "cancel-edit-user" && currentUser.role === "admin") {
    state.editingUserId = "";
    render();
    return;
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
    const supplierName = String(button.dataset.supplierName || "this supplier").trim();
    const confirmed = window.confirm(`Delete supplier "${supplierName}"? This cannot be undone.`);
    if (!confirmed) {
      return;
    }

    await runMutation(
      "delete_supplier",
      { p_token: sessionToken, p_supplier_id: button.dataset.supplierId },
      "Supplier deleted.",
    );
  }

  if (action === "delete-location" && currentUser.role === "admin") {
    const locationName = String(button.dataset.locationName || "this location").trim();
    const confirmed = window.confirm(`Delete location "${locationName}"? This cannot be undone.`);
    if (!confirmed) {
      return;
    }

    await runMutation(
      "delete_location",
      { p_token: sessionToken, p_location_id: button.dataset.locationId },
      "Location deleted.",
    );
  }

  if (action === "edit-location" && currentUser.role === "admin") {
    state.editingLocationId = String(button.dataset.locationId || "");
    render();
    return;
  }

  if (action === "cancel-edit-location" && currentUser.role === "admin") {
    state.editingLocationId = "";
    render();
    return;
  }

  if (action === "edit-stock-movement" && (currentUser.role === "admin" || currentUser.role === "logistics")) {
    const stockMovementId = String(button.dataset.stockMovementId || "");
    if (state.editingStockMovementId === stockMovementId) {
      cancelStockMovementEdit();
      return;
    }

    const movement = getStockMovement(stockMovementId);
    if (!movement) {
      showFlash("That stock movement could not be found.", "error");
      render();
      return;
    }

    state.editingStockMovementId = stockMovementId;
    state.stockMovementSelectedItemId = movement.stockItemId || "";
    state.stockMovementsSectionOpen = true;
    render();
    focusStockMovementForm();
    return;
  }

  if (action === "cancel-edit-stock-movement" && (currentUser.role === "admin" || currentUser.role === "logistics")) {
    cancelStockMovementEdit();
    return;
  }

  if (action === "edit-stock-item" && currentUser.role === "admin") {
    const stockItemId = String(button.dataset.stockItemId || "");
    if (state.editingStockItemId === stockItemId) {
      cancelStockItemEdit();
      return;
    }

    state.editingStockItemId = stockItemId;
    render();
    focusStockItemForm();
    return;
  }

  if (action === "cancel-edit-stock-item" && currentUser.role === "admin") {
    cancelStockItemEdit();
    return;
  }

  if (action === "delete-stock-item" && currentUser.role === "admin") {
    await runMutation(
      "delete_stock_item",
      { p_token: sessionToken, p_stock_item_id: button.dataset.stockItemId },
      "Stock item deleted.",
    );
    return;
  }

  if (action === "open-stock-qr" && (currentUser.role === "admin" || currentUser.role === "logistics")) {
    await openStockQrPreview(String(button.dataset.stockItemId || ""));
    return;
  }

  if (action === "close-stock-qr" && (currentUser.role === "admin" || currentUser.role === "logistics")) {
    state.stockQrItemId = "";
    state.stockQrSvg = "";
    state.stockQrBusy = false;
    state.stockQrSharing = false;
    render();
    return;
  }

  if (action === "print-stock-qr" && (currentUser.role === "admin" || currentUser.role === "logistics")) {
    printStockQrLabel(String(button.dataset.stockItemId || ""));
    return;
  }

  if (action === "share-stock-qr" && (currentUser.role === "admin" || currentUser.role === "logistics")) {
    await shareStockQrLabel(String(button.dataset.stockItemId || ""));
    return;
  }

  if (action === "copy-stock-qr" && (currentUser.role === "admin" || currentUser.role === "logistics")) {
    await copyStockQrValue(String(button.dataset.stockItemId || ""));
    return;
  }

  if (action === "open-stock-scanner" && (currentUser.role === "admin" || currentUser.role === "logistics")) {
    state.stockScannerOpen = true;
    state.stockScannerStatus = getStockScannerHint();
    state.stockScannerPendingItemId = "";
    state.stockScannerPendingCode = "";
    await stopStockScanner();
    render();
    return;
  }

  if (action === "clear-stock-search" && (currentUser.role === "admin" || currentUser.role === "logistics")) {
    state.stockSearchQuery = "";
    render();
    return;
  }

  if (action === "toggle-stock-section" && (currentUser.role === "admin" || currentUser.role === "logistics")) {
    const section = String(button.dataset.stockSection || "");
    if (section === "movements") {
      state.stockMovementsSectionOpen = !state.stockMovementsSectionOpen;
    } else if (section === "artwork-form") {
      state.stockArtworkPanelOpen = !state.stockArtworkPanelOpen;
    } else if (section === "artwork-log") {
      state.stockArtworkRequestsSectionOpen = !state.stockArtworkRequestsSectionOpen;
    }
    render();
    return;
  }

  if (action === "toggle-stock-item-card" && (currentUser.role === "admin" || currentUser.role === "logistics")) {
    const stockItemId = String(button.dataset.stockItemId || "");
    if (stockItemId) {
      state.stockOpenItemCards = {
        ...state.stockOpenItemCards,
        [stockItemId]: !state.stockOpenItemCards[stockItemId],
      };
      render();
    }
    return;
  }

  if (action === "start-stock-camera" && (currentUser.role === "admin" || currentUser.role === "logistics")) {
    await startStockScanner();
    return;
  }

  if (action === "close-stock-scanner" && (currentUser.role === "admin" || currentUser.role === "logistics")) {
    state.stockScannerOpen = false;
    state.stockScannerStatus = "";
    state.stockScannerPendingItemId = "";
    state.stockScannerPendingCode = "";
    await stopStockScanner();
    render();
    return;
  }

  if (action === "confirm-stock-scan" && (currentUser.role === "admin" || currentUser.role === "logistics")) {
    await confirmStockScanSelection();
    return;
  }

  if (action === "retry-stock-scan" && (currentUser.role === "admin" || currentUser.role === "logistics")) {
    clearPendingStockScan("Ready to scan another stock QR code.");
    render();
    return;
  }

  if (action === "apply-stock-manual-qr" && (currentUser.role === "admin" || currentUser.role === "logistics")) {
    const scannerPanel = button.closest(".scanner-panel");
    const input = scannerPanel instanceof HTMLElement
      ? scannerPanel.querySelector("[data-stock-manual-value]")
      : null;
    const rawValue = input instanceof HTMLInputElement ? input.value.trim() : "";

    if (!rawValue) {
      showFlash("Enter the QR value printed on the label.", "error");
      render();
      return;
    }

    try {
      await applyStockScanResult(rawValue);
    } catch (error) {
      showFlash(normalizeStockScannerError(error), "error");
      state.stockScannerStatus = normalizeStockScannerError(error);
      render();
    }
    return;
  }

  if (action === "clear-stock-selection" && (currentUser.role === "admin" || currentUser.role === "logistics")) {
    state.stockMovementSelectedItemId = "";
    render();
    return;
  }

  if (action === "save-order-assignment" && (currentUser.role === "admin" || currentUser.role === "sales")) {
    await saveOrderAssignment(button, currentUser);
    return;
  }

  if (action === "toggle-order-priority" && currentUser.role === "admin") {
    await toggleOrderPriority(String(button.dataset.orderId || "").trim());
    return;
  }

  if (action === "toggle-order-flag" && (currentUser.role === "admin" || currentUser.role === "driver")) {
    const orderId = String(button.dataset.orderId || "").trim();
    const order = getOrder(orderId);
    if (!order || order.status !== "active") {
      return;
    }

    state.flaggingOrderId = state.flaggingOrderId === orderId ? "" : orderId;
    render();
    return;
  }

  if (action === "cancel-order-flag" && (currentUser.role === "admin" || currentUser.role === "driver")) {
    state.flaggingOrderId = "";
    render();
    return;
  }

  if (action === "clear-order-flag" && (currentUser.role === "admin" || currentUser.role === "driver")) {
    await clearOrderFlag(String(button.dataset.orderId || "").trim());
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
    const orderReference = String(button.dataset.orderReference || "this entry").trim();
    const confirmed = window.confirm(`Delete ${orderReference}? This cannot be undone.`);
    if (!confirmed) {
      return;
    }

    await runMutation(
      "delete_order",
      { p_token: sessionToken, p_order_id: button.dataset.orderId },
      "Order deleted.",
    );
    return;
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

  if (target.matches("[data-assignment-filter]") && target instanceof HTMLSelectElement) {
    state.assignmentDriverFilter = target.value;
    state.pagination.assignments = 1;
    render();
    return;
  }

  const orderForm = target.closest('form[data-form="add-order"]');
  if (
    orderForm instanceof HTMLFormElement
    && (target.matches('[name="entryType"]') || target.matches('[name="locationId"]') || target.matches('[name="moveToFactory"]'))
  ) {
    syncMoveToFactoryField(orderForm);
  }

  const stockMovementForm = target.closest('form[data-form="add-stock-movement"]');
  if (stockMovementForm instanceof HTMLFormElement && target.matches('[name="movementType"]')) {
    syncStockMovementFields(stockMovementForm);
  }

  if (stockMovementForm instanceof HTMLFormElement && target.matches('[name="stockItemId"]')) {
    state.stockMovementSelectedItemId = target instanceof HTMLSelectElement ? target.value : "";
  }

  if (target.matches("[data-stock-scan-upload]") && target instanceof HTMLInputElement && target.files?.[0]) {
    void handleStockScanUpload(target.files[0], target);
  }
}

function handleInput(event) {
  const target = event.target;
  if (!(target instanceof HTMLElement)) {
    return;
  }

  if (target.matches("[data-stock-search]") && target instanceof HTMLInputElement) {
    state.stockSearchQuery = target.value;
    render();
    return;
  }

  if (target.matches("[data-stock-scanner-zoom]") && target instanceof HTMLInputElement) {
    const zoomValue = Number(target.value);
    if (Number.isFinite(zoomValue)) {
      void setStockScannerZoom(zoomValue);
    }
  }
}

function handleKeyDown(event) {
  if (event.key !== "Escape" || state.busy) {
    return;
  }

  if (state.flaggingOrderId) {
    event.preventDefault();
    state.flaggingOrderId = "";
    render();
    return;
  }

  if (state.currentPage !== "stock") {
    return;
  }

  if (state.editingStockMovementId) {
    event.preventDefault();
    cancelStockMovementEdit();
    return;
  }

  if (state.editingStockItemId) {
    event.preventDefault();
    cancelStockItemEdit();
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
  const userId = String(formData.get("userId") || "").trim();
  const role = String(formData.get("role") || "").trim();
  const name = String(formData.get("name") || "").trim();
  const password = String(formData.get("password") || "").trim();
  const phone = String(formData.get("phone") || "").trim();

  if (!name || !role || (!userId && !password)) {
    showFlash(userId ? "Name and role are required." : "Name, password, and role are required.", "error");
    render();
    return;
  }

  const ok = await runMutation(
    userId ? "update_user_account" : "create_user_account",
    userId
      ? {
          p_token: sessionToken,
          p_user_id: userId,
          p_name: name,
          p_role: role,
          p_phone: phone,
          p_password: password || null,
        }
      : {
          p_token: sessionToken,
          p_name: name,
          p_password: password,
          p_role: role,
          p_phone: phone,
        },
    userId ? `Account updated: ${name}.` : `${capitalize(role)} account created for ${name}.`,
  );

  if (ok) {
    state.editingUserId = "";
    render();
  }
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
  const locationId = String(formData.get("locationId") || "").trim();
  const locationType = String(formData.get("locationType") || "").trim();
  const name = String(formData.get("name") || "").trim();
  const address = String(formData.get("address") || "").trim();
  const lat = parseOptionalNumber(formData.get("lat"));
  const lng = parseOptionalNumber(formData.get("lng"));
  const contactPerson = String(formData.get("contactPerson") || "").trim();
  const contactNumber = String(formData.get("contactNumber") || "").trim();

  if (!locationType || !name || !address) {
    showFlash("Location name, type, and address are required.", "error");
    render();
    return;
  }

  if (Number.isNaN(lat) || Number.isNaN(lng)) {
    showFlash("Latitude and longitude must be valid numbers.", "error");
    render();
    return;
  }

  if ((lat === null) !== (lng === null)) {
    showFlash("Enter both latitude and longitude, or leave both blank.", "error");
    render();
    return;
  }

  const ok = await runMutation(
    locationId ? "update_location" : "create_location",
    locationId
      ? {
          p_token: sessionToken,
          p_location_id: locationId,
          p_location_type: locationType,
          p_name: name,
          p_address: address,
          p_lat: lat,
          p_lng: lng,
          p_contact_person: contactPerson,
          p_contact_number: contactNumber,
        }
      : {
          p_token: sessionToken,
          p_location_type: locationType,
          p_name: name,
          p_address: address,
          p_lat: lat,
          p_lng: lng,
          p_contact_person: contactPerson,
          p_contact_number: contactNumber,
        },
    locationId ? `Location updated: ${name}.` : `Location added: ${name}.`,
  );

  if (ok) {
    state.editingLocationId = "";
    render();
  }
}

async function createOrder(formData, currentUser) {
  const driverUserId = String(formData.get("driverUserId") || "").trim();
  const locationId = String(formData.get("locationId") || "").trim();
  const entryType = String(formData.get("entryType") || "delivery").trim();
  const quoteNumber = String(formData.get("quoteNumber") || "").trim();
  const salesOrderNumber = String(formData.get("salesOrderNumber") || "").trim();
  const invoiceNumber = String(formData.get("invoiceNumber") || "").trim();
  const poNumber = String(formData.get("poNumber") || "").trim();
  const notice = String(formData.get("notice") || "").trim();
  const moveToFactory = formData.get("moveToFactory") === "on";
  const factoryDestinationLocationId = String(formData.get("factoryDestinationLocationId") || "").trim();
  const allowDuplicate = formData.get("allowDuplicate") === "on";

  if (!locationId || !entryType || !quoteNumber) {
    showFlash("Pickup location, entry type, and quote number are required.", "error");
    render();
    return;
  }

  if (moveToFactory && !factoryDestinationLocationId) {
    showFlash("Select which factory the collected stock should go to.", "error");
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
      p_quote_number: quoteNumber,
      p_sales_order_number: salesOrderNumber,
      p_invoice_number: invoiceNumber,
      p_po_number: poNumber,
      p_notice: notice,
      p_move_to_factory: moveToFactory,
      p_factory_destination_location_id: moveToFactory ? factoryDestinationLocationId : null,
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

async function toggleOrderPriority(orderId) {
  if (!orderId) {
    return;
  }

  const order = getOrder(orderId);
  if (!order || order.status !== "active") {
    return;
  }

  const nextPriority = isPriorityOrder(order) ? DEFAULT_ORDER_PRIORITY : PRIORITY_STOP_VALUE;

  await runMutation(
    "set_order_priority",
    {
      p_token: sessionToken,
      p_order_id: orderId,
      p_priority: nextPriority,
    },
    nextPriority === PRIORITY_STOP_VALUE
      ? `${order.reference || "Entry"} marked as a priority stop.`
      : `Priority cleared for ${order.reference || "the entry"}.`,
  );
}

async function saveOrderFlag(formData) {
  const orderId = String(formData.get("orderId") || "").trim();
  const flagType = String(formData.get("flagType") || "").trim();
  const note = String(formData.get("flagNote") || "").trim();
  const order = getOrder(orderId);

  if (!orderId || !flagType) {
    showFlash("Choose why this entry is being flagged.", "error");
    render();
    return;
  }

  const ok = await runMutation(
    "set_order_flag",
    {
      p_token: sessionToken,
      p_order_id: orderId,
      p_flag_type: flagType,
      p_note: note || null,
    },
    `Follow-up saved for ${order?.reference || "the entry"}.`,
  );

  if (ok) {
    state.flaggingOrderId = "";
    render();
  }
}

async function clearOrderFlag(orderId) {
  if (!orderId) {
    return;
  }

  const order = getOrder(orderId);
  if (!order) {
    return;
  }

  const ok = await runMutation(
    "set_order_flag",
    {
      p_token: sessionToken,
      p_order_id: orderId,
      p_flag_type: null,
      p_note: null,
    },
    `Follow-up cleared for ${order.reference || "the entry"}.`,
  );

  if (ok) {
    if (state.flaggingOrderId === orderId) {
      state.flaggingOrderId = "";
    }
    render();
  }
}

async function createStockItem(formData) {
  const stockItemId = String(formData.get("stockItemId") || "").trim();
  const currentUser = state.snapshot.user;
  const name = String(formData.get("name") || "").trim();
  const sku = String(formData.get("sku") || "").trim();
  const quoteNumber = String(formData.get("quoteNumber") || "").trim();
  const invoiceNumber = String(formData.get("invoiceNumber") || "").trim();
  const salesOrderNumber = String(formData.get("salesOrderNumber") || "").trim();
  const poNumber = String(formData.get("poNumber") || "").trim();
  const unit = String(formData.get("unit") || "").trim();
  const notes = String(formData.get("notes") || "").trim();

  if (stockItemId && currentUser?.role !== "admin") {
    showFlash("Only admins can edit stock items.", "error");
    render();
    return;
  }

  if (!name) {
    showFlash("Stock description is required.", "error");
    render();
    return;
  }

  if (!quoteNumber) {
    showFlash("Quote number is required for QR-managed stock.", "error");
    render();
    return;
  }

  const ok = await runMutation(
    stockItemId ? "update_stock_item" : "create_stock_item",
    stockItemId
      ? {
          p_token: sessionToken,
          p_stock_item_id: stockItemId,
          p_name: name,
          p_sku: sku,
          p_quote_number: quoteNumber,
          p_invoice_number: invoiceNumber,
          p_sales_order_number: salesOrderNumber,
          p_po_number: poNumber,
          p_unit: unit || "units",
          p_notes: notes,
        }
      : {
          p_token: sessionToken,
          p_name: name,
          p_sku: sku,
          p_quote_number: quoteNumber,
          p_invoice_number: invoiceNumber,
          p_sales_order_number: salesOrderNumber,
          p_po_number: poNumber,
          p_unit: unit || "units",
          p_notes: notes,
        },
    stockItemId ? `Stock item updated: ${name}.` : `Stock item added: ${name}.`,
  );

  if (ok && stockItemId) {
    state.editingStockItemId = "";
    render();
  }
}

async function recordStockMovement(formData) {
  const stockMovementId = String(formData.get("stockMovementId") || "").trim();
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

  const ok = await runMutation(
    stockMovementId ? "update_stock_movement" : "record_stock_movement",
    {
      p_token: sessionToken,
      ...(stockMovementId ? { p_stock_movement_id: stockMovementId } : {}),
      p_stock_item_id: stockItemId,
      p_movement_type: movementType,
      p_quantity: quantity,
      p_supplier_name: supplierName,
      p_driver_user_id: driverUserId || null,
      p_notes: notes,
    },
    stockMovementId
      ? (movementType === "in" ? "Stock receipt updated." : "Stock issue updated.")
      : (movementType === "in" ? "Stock receipt recorded." : "Stock issue recorded."),
  );

  if (ok && stockMovementId) {
    state.editingStockMovementId = "";
    render();
  }
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
  destroyDriverRouteMap();
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
  ensureDriverRouteOrigin();
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

  void syncStockScannerUi();
}

function focusStockItemForm() {
  window.requestAnimationFrame(() => {
    const form = document.querySelector('form[data-form="add-stock-item"]');
    if (!(form instanceof HTMLFormElement)) {
      return;
    }

    form.scrollIntoView({ behavior: "smooth", block: "start" });
    const firstField = form.querySelector('[name="name"]');
    if (firstField instanceof HTMLElement) {
      firstField.focus();
    }
  });
}

function focusStockMovementForm() {
  window.requestAnimationFrame(() => {
    const form = document.querySelector('form[data-form="add-stock-movement"]');
    if (!(form instanceof HTMLFormElement)) {
      return;
    }

    form.scrollIntoView({ behavior: "smooth", block: "start" });
    const firstField = form.querySelector('[name="stockItemId"]');
    if (firstField instanceof HTMLElement) {
      firstField.focus();
    }
  });
}

function cancelStockItemEdit() {
  if (!state.editingStockItemId) {
    return;
  }

  state.editingStockItemId = "";
  render();
}

function cancelStockMovementEdit() {
  if (!state.editingStockMovementId && !state.stockMovementSelectedItemId) {
    return;
  }

  state.editingStockMovementId = "";
  state.stockMovementSelectedItemId = "";
  render();
}

async function syncStockScannerUi() {
  if (!state.stockScannerOpen || state.currentPage !== "stock") {
    await stopStockScanner();
    return;
  }

  const videoEl = document.querySelector("[data-stock-scanner-video]");
  if (!(videoEl instanceof HTMLVideoElement)) {
    await stopStockScanner();
  }
}

function getStockItemById(stockItemId) {
  return state.snapshot.stockItems.find((item) => item.id === stockItemId) || null;
}

function getEditingStockItem() {
  return state.editingStockItemId ? getStockItemById(state.editingStockItemId) : null;
}

function getStockMovement(stockMovementId) {
  return state.snapshot.stockMovements.find((movement) => movement.id === stockMovementId) || null;
}

function getEditingStockMovement() {
  return state.editingStockMovementId ? getStockMovement(state.editingStockMovementId) : null;
}

function isRecentStockMovement(value, days = STOCK_RECENT_ACTIVITY_DAYS) {
  const timestamp = Date.parse(String(value || ""));
  if (!Number.isFinite(timestamp)) {
    return false;
  }

  return (Date.now() - timestamp) <= (days * 24 * 60 * 60 * 1000);
}

function getRecentStockMovements(movementType, limit = 6) {
  const recentMovements = [];
  const seenStockItemIds = new Set();

  state.snapshot.stockMovements.forEach((movement) => {
    if (movement.movementType !== movementType) {
      return;
    }

    if (!isRecentStockMovement(movement.createdAt)) {
      return;
    }

    if (seenStockItemIds.has(movement.stockItemId)) {
      return;
    }

    seenStockItemIds.add(movement.stockItemId);
    recentMovements.push(movement);
  });

  return recentMovements.slice(0, limit);
}

function getLatestStockMovementByItemId() {
  const latestMovementByItemId = new Map();

  state.snapshot.stockMovements.forEach((movement) => {
    if (!latestMovementByItemId.has(movement.stockItemId)) {
      latestMovementByItemId.set(movement.stockItemId, movement);
    }
  });

  return latestMovementByItemId;
}

function normalizeStockSearchValue(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "");
}

function matchesStockSearch(record, query = state.stockSearchQuery) {
  const needle = normalizeStockSearchValue(query);
  if (!needle) {
    return true;
  }

  return [
    record?.name,
    record?.sku,
    record?.quoteNumber,
    record?.salesOrderNumber,
    record?.invoiceNumber,
    record?.poNumber,
  ].some((value) => normalizeStockSearchValue(value).includes(needle));
}

function getFilteredStockItems() {
  return state.snapshot.stockItems.filter((item) => matchesStockSearch(item));
}

function getStockItemSearchSummary(filteredCount, totalCount) {
  const query = String(state.stockSearchQuery || "").trim();
  if (!query) {
    return "Search by quote, sales order, invoice, PO, stock code, or item name.";
  }

  if (!filteredCount) {
    return `No stock items match "${query}".`;
  }

  return `Showing ${filteredCount} of ${totalCount} stock items for "${query}".`;
}

function renderStockDisclosure({ tag = "section", containerClass = "table-card", sectionKey, eyebrow, title, subtitle = "", summary = "", open = false, body }) {
  return `
    <${tag} class="${containerClass} stock-disclosure${open ? " is-open" : ""}">
      <div class="stock-disclosure-header">
        <div class="stock-disclosure-copy">
          <p class="eyebrow">${escapeHtml(eyebrow)}</p>
          <h3 class="panel-title">${escapeHtml(title)}</h3>
          ${subtitle ? `<p class="panel-subtitle">${escapeHtml(subtitle)}</p>` : ""}
          ${summary ? `<p class="stock-disclosure-summary">${escapeHtml(summary)}</p>` : ""}
        </div>
        <button
          type="button"
          class="button button-ghost stock-disclosure-toggle"
          data-action="toggle-stock-section"
          data-stock-section="${escapeHtml(sectionKey)}"
          ${state.busy ? " disabled" : ""}
        >
          ${open ? "Collapse" : "Expand"}
        </button>
      </div>
      ${open ? `<div class="stock-disclosure-body">${body}</div>` : ""}
    </${tag}>
  `;
}

function isLocalHost() {
  const host = String(window.location.hostname || "").trim().toLowerCase();
  return host === "localhost" || host === "127.0.0.1";
}

function supportsStockQrDetection() {
  return typeof window.BarcodeDetector === "function";
}

function canUseLiveStockScanner() {
  return supportsStockQrDetection() && Boolean(navigator.mediaDevices?.getUserMedia) && (window.isSecureContext || isLocalHost());
}

async function getStockScannerDetector() {
  if (!supportsStockQrDetection()) {
    throw new Error("This browser does not support QR detection.");
  }

  if (stockScannerDetector) {
    return stockScannerDetector;
  }

  let formats = STOCK_SCANNER_FORMATS.slice();

  if (typeof window.BarcodeDetector.getSupportedFormats === "function") {
    try {
      const supportedFormats = await window.BarcodeDetector.getSupportedFormats();
      if (Array.isArray(supportedFormats)) {
        if (!supportedFormats.includes("qr_code")) {
          throw new Error("This browser does not support QR scanning.");
        }
        formats = STOCK_SCANNER_FORMATS.filter((format) => supportedFormats.includes(format));
      }
    } catch (error) {
      if (String(error?.message || "").toLowerCase().includes("qr scanning")) {
        throw error;
      }
      formats = STOCK_SCANNER_FORMATS.slice();
    }
  }

  if (!formats.length) {
    throw new Error("This browser does not support stock QR scanning.");
  }

  stockScannerDetector = new window.BarcodeDetector({ formats });
  return stockScannerDetector;
}

function getStockScannerHint() {
  if (!supportsStockQrDetection()) {
    return "This browser does not support QR detection. Use the printed QR value for stock selection.";
  }

  if (canUseLiveStockScanner()) {
    return "Start the camera to scan a stock QR code, upload a QR image, or enter the printed QR value.";
  }

  if (navigator.mediaDevices?.getUserMedia && !window.isSecureContext && !isLocalHost()) {
    return "Live camera scan needs HTTPS or localhost. Upload a QR image here, or enter the printed QR value.";
  }

  return "Upload a QR image here, or enter the printed QR value.";
}

function resetStockScannerCameraState() {
  stockScannerVideoTrack = null;
  state.stockScannerZoomSupported = false;
  state.stockScannerZoomValue = 1;
  state.stockScannerZoomMin = 1;
  state.stockScannerZoomMax = 1;
  state.stockScannerZoomStep = 0.1;
}

function getActiveStockScannerStatus() {
  if (state.stockScannerZoomSupported) {
    return "Scanning for a stock QR code. Move slowly closer until it looks sharp, or use Zoom.";
  }

  return "Scanning for a stock QR code. Move slowly closer until it looks sharp.";
}

async function configureStockScannerTrack(track) {
  resetStockScannerCameraState();

  if (!track) {
    return;
  }

  stockScannerVideoTrack = track;

  let capabilities = null;
  let settings = null;

  try {
    capabilities = typeof track.getCapabilities === "function" ? track.getCapabilities() : null;
  } catch (error) {
    capabilities = null;
  }

  try {
    settings = typeof track.getSettings === "function" ? track.getSettings() : null;
  } catch (error) {
    settings = null;
  }

  const focusModes = Array.isArray(capabilities?.focusMode) ? capabilities.focusMode : [];
  if (focusModes.length) {
    const preferredFocusMode = focusModes.includes("continuous")
      ? "continuous"
      : (focusModes.includes("single-shot") ? "single-shot" : "");

    if (preferredFocusMode) {
      try {
        await track.applyConstraints({ advanced: [{ focusMode: preferredFocusMode }] });
      } catch (error) {
        // Ignore devices that expose focus metadata but reject browser focus controls.
      }
    }
  }

  const zoomCapability = capabilities?.zoom;
  const zoomMin = Number(zoomCapability?.min);
  const zoomMax = Number(zoomCapability?.max);
  const zoomStep = Number(zoomCapability?.step);
  const currentZoom = Number(settings?.zoom);

  if (Number.isFinite(zoomMin) && Number.isFinite(zoomMax) && zoomMax > zoomMin) {
    state.stockScannerZoomSupported = true;
    state.stockScannerZoomMin = Math.round(zoomMin * 100) / 100;
    state.stockScannerZoomMax = Math.round(zoomMax * 100) / 100;
    state.stockScannerZoomStep = Number.isFinite(zoomStep) && zoomStep > 0
      ? Math.max(Math.round(zoomStep * 100) / 100, 0.01)
      : Math.max(Math.round(((zoomMax - zoomMin) / 20) * 100) / 100, 0.01);
    state.stockScannerZoomValue = Math.min(
      Math.max(Number.isFinite(currentZoom) ? currentZoom : zoomMin, zoomMin),
      zoomMax,
    );
  }
}

async function bindStockScannerVideo(stream) {
  const videoEl = document.querySelector("[data-stock-scanner-video]");
  if (!(videoEl instanceof HTMLVideoElement)) {
    return null;
  }

  videoEl.muted = true;
  videoEl.srcObject = stream;
  videoEl.setAttribute("playsinline", "true");
  await videoEl.play();
  return videoEl;
}

async function setStockScannerZoom(zoomValue) {
  if (!stockScannerVideoTrack || !state.stockScannerZoomSupported || !Number.isFinite(zoomValue)) {
    return;
  }

  const nextZoom = Math.min(Math.max(zoomValue, state.stockScannerZoomMin), state.stockScannerZoomMax);

  try {
    await stockScannerVideoTrack.applyConstraints({ advanced: [{ zoom: nextZoom }] });
    state.stockScannerZoomValue = nextZoom;
  } catch (error) {
    // Ignore devices that expose zoom but do not allow live updates from the browser.
  }
}

async function startStockScanner() {
  if (!state.stockScannerOpen) {
    return;
  }

  clearPendingStockScan();

  if (!supportsStockQrDetection()) {
    state.stockScannerStatus = "This browser does not support QR detection.";
    render();
    return;
  }

  if (!navigator.mediaDevices?.getUserMedia) {
    state.stockScannerStatus = "This browser cannot open the device camera.";
    render();
    return;
  }

  if (!window.isSecureContext && !isLocalHost()) {
    state.stockScannerStatus = "Live camera scan needs HTTPS or localhost. Upload a QR image instead.";
    render();
    return;
  }

  if (stockScannerStream || stockScannerStarting) {
    return;
  }

  stockScannerStarting = true;
  stockScannerStopRequested = false;

  try {
    const detector = await getStockScannerDetector();
    const stream = await navigator.mediaDevices.getUserMedia({
      audio: false,
      video: {
        facingMode: { ideal: "environment" },
        width: { ideal: 1920 },
        height: { ideal: 1080 },
      },
    });

    if (stockScannerStopRequested) {
      stream.getTracks().forEach((track) => track.stop());
      return;
    }

    stockScannerStream = stream;
    await configureStockScannerTrack(stream.getVideoTracks()[0] || null);
    state.stockScannerStatus = getActiveStockScannerStatus();
    render();

    const videoEl = await bindStockScannerVideo(stream);
    if (!(videoEl instanceof HTMLVideoElement)) {
      await stopStockScanner();
      return;
    }

    scanStockVideoFrame(videoEl, detector);
  } catch (error) {
    await stopStockScanner();
    state.stockScannerStatus = normalizeStockScannerError(error);
    render();
  } finally {
    stockScannerStarting = false;
  }
}

function scanStockVideoFrame(videoEl, detector = stockScannerDetector) {
  if (!stockScannerStream || stockScannerStopRequested || !(videoEl instanceof HTMLVideoElement)) {
    return;
  }

  const runDetection = async () => {
    if (!stockScannerStream || stockScannerStopRequested) {
      return;
    }

    try {
      const detections = await detector.detect(videoEl);
      const rawValue = detections.find((entry) => typeof entry?.rawValue === "string" && entry.rawValue.trim())?.rawValue?.trim();

      if (rawValue) {
        await applyStockScanResult(rawValue);
        return;
      }
    } catch (error) {
      if (!isIgnorableStockScannerError(error)) {
        state.stockScannerStatus = normalizeStockScannerError(error);
        await stopStockScanner();
        render();
        return;
      }
    }

    stockScannerAnimationFrame = window.requestAnimationFrame(() => {
      void runDetection();
    });
  };

  void runDetection();
}

async function stopStockScanner() {
  stockScannerStopRequested = true;

  if (stockScannerAnimationFrame) {
    window.cancelAnimationFrame(stockScannerAnimationFrame);
    stockScannerAnimationFrame = 0;
  }

  if (stockScannerStream) {
    stockScannerStream.getTracks().forEach((track) => track.stop());
    stockScannerStream = null;
  }

  resetStockScannerCameraState();

  const videoEl = document.querySelector("[data-stock-scanner-video]");
  if (videoEl instanceof HTMLVideoElement) {
    videoEl.pause();
    videoEl.srcObject = null;
  }
}

async function handleStockScanUpload(file, inputEl) {
  if (!supportsStockQrDetection()) {
    showFlash("This browser does not support QR detection from images.", "error");
    render();
    return;
  }

  try {
    state.stockScannerStatus = "Reading QR image.";
    render();

    const image = await loadImageFromFile(file);
    const detector = await getStockScannerDetector();
    const detections = await detector.detect(image);
    const rawValue = detections.find((entry) => typeof entry?.rawValue === "string" && entry.rawValue.trim())?.rawValue?.trim();

    if (!rawValue) {
      throw new Error("No stock QR code was found in that image.");
    }

    await applyStockScanResult(rawValue);
  } catch (error) {
    showFlash(normalizeStockScannerError(error), "error");
    state.stockScannerStatus = normalizeStockScannerError(error);
    render();
  } finally {
    if (inputEl instanceof HTMLInputElement) {
      inputEl.value = "";
    }
  }
}

function loadImageFromFile(file) {
  return new Promise((resolve, reject) => {
    const objectUrl = URL.createObjectURL(file);
    const image = new Image();
    image.onload = () => {
      URL.revokeObjectURL(objectUrl);
      resolve(image);
    };
    image.onerror = () => {
      URL.revokeObjectURL(objectUrl);
      reject(new Error("That image could not be read."));
    };
    image.src = objectUrl;
  });
}

async function applyStockScanResult(rawValue) {
  const parsed = parseStockQrPayload(rawValue);
  if (!parsed?.stockItemId) {
    throw new Error("That QR code is not a Route Ledger stock label.");
  }

  const item = getStockItemById(parsed.stockItemId);
  if (!item) {
    throw new Error("That QR code does not match a stock item in this snapshot.");
  }

  await stopStockScanner();
  state.stockScannerPendingItemId = item.id;
  state.stockScannerPendingCode = buildStockQrPayload(item);
  state.stockScannerStatus = `Scanned ${buildStockQrPayload(item)}. Confirm to select ${item.name}.`;
  render();
}

async function confirmStockScanSelection() {
  const item = getStockItemById(state.stockScannerPendingItemId);
  if (!item) {
    clearPendingStockScan(getStockScannerHint());
    showFlash("Scan confirmation expired. Scan the label again.", "error");
    render();
    return;
  }

  state.stockMovementSelectedItemId = item.id;
  state.stockScannerOpen = false;
  state.stockScannerStatus = "";
  clearPendingStockScan();
  await stopStockScanner();
  showFlash(`Stock item selected: ${item.name}.`, "success");
  render();
}

function clearPendingStockScan(status = "") {
  state.stockScannerPendingItemId = "";
  state.stockScannerPendingCode = "";
  if (status) {
    state.stockScannerStatus = status;
  }
}

function buildStockQrPayload(item) {
  const preferredSku = normalizeStockLabelText(item?.sku || "");
  if (preferredSku && preferredSku.length <= STOCK_LABEL_SKU_MAX_LENGTH && /^[A-Z0-9-]+$/.test(preferredSku)) {
    return preferredSku;
  }

  const uuidHex = getStockItemUuidHex(item);
  if (uuidHex) {
    return encodeStockLabelCode(BigInt(`0x${uuidHex.slice(-10)}`), STOCK_LABEL_CODE_LENGTH);
  }

  return normalizeStockLabelText(item?.quoteNumber || item?.id || "STOCK");
}

function parseStockQrPayload(rawValue) {
  const value = String(rawValue || "").trim();
  if (!value) {
    return null;
  }

  if (UUID_PATTERN.test(value)) {
    return { stockItemId: value.toLowerCase() };
  }

  const normalizedValue = normalizeStockLabelText(value);
  const matchedItem = state.snapshot.stockItems.find((item) => (
    buildStockQrPayload(item) === normalizedValue
      || buildLegacyStockLabelCode(item) === normalizedValue
  ));
  if (matchedItem) {
    return { stockItemId: matchedItem.id };
  }

  try {
    const payload = JSON.parse(value);
    if (
      payload
      && payload.type === STOCK_QR_TYPE
      && Number(payload.version) === STOCK_QR_VERSION
      && UUID_PATTERN.test(String(payload.stockItemId || ""))
    ) {
      return {
        ...payload,
        stockItemId: String(payload.stockItemId || "").toLowerCase(),
      };
    }
  } catch (error) {
    return null;
  }

  return null;
}

function normalizeStockLabelText(value) {
  return String(value || "").trim().toUpperCase();
}

function getStockItemUuidHex(item) {
  const uuidHex = String(item?.id || "").replace(/-/g, "").trim().toLowerCase();
  return uuidHex.length === 32 ? uuidHex : "";
}

function encodeStockLabelCode(value, minLength = 1) {
  let current = BigInt(value || 0);
  let result = "";

  do {
    const digit = Number(current % 32n);
    result = `${STOCK_LABEL_ALPHABET[digit]}${result}`;
    current /= 32n;
  } while (current > 0n);

  return result.padStart(minLength, "0");
}

function buildLegacyStockLabelCode(item) {
  const uuidHex = getStockItemUuidHex(item);
  if (!uuidHex) {
    return "";
  }

  return encodeStockLabelCode(BigInt(`0x${uuidHex.slice(-16)}`), 13);
}

function normalizeStockScannerError(error) {
  const message = normalizeError(error);
  const lowerMessage = message.toLowerCase();

  if (lowerMessage.includes("qr scanning")) {
    return "This browser does not support QR scanning.";
  }

  if (lowerMessage.includes("permission") || lowerMessage.includes("denied")) {
    return "Camera access was blocked. Allow camera access, or upload a QR image instead.";
  }

  if (lowerMessage.includes("not found")) {
    return "No camera was found. Upload a QR image instead.";
  }

  return message;
}

function isIgnorableStockScannerError(error) {
  const message = String(error?.message || "").toLowerCase();
  return message.includes("source is unavailable") || message.includes("service unavailable");
}

async function openStockQrPreview(stockItemId) {
  const item = getStockItemById(stockItemId);
  if (!item) {
    showFlash("Stock item not found.", "error");
    render();
    return;
  }

  state.stockQrItemId = item.id;
  state.stockQrSvg = "";
  state.stockQrBusy = true;
  state.stockQrSharing = false;
  render();

  try {
    state.stockQrSvg = await requestText(`${API_ROOT}/qr/svg`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        text: buildStockQrPayload(item),
      }),
    });
  } catch (error) {
    state.stockQrItemId = "";
    state.stockQrSvg = "";
    state.stockQrSharing = false;
    showFlash(normalizeError(error), "error");
  } finally {
    state.stockQrBusy = false;
    render();
  }
}

function printStockQrLabel(stockItemId) {
  const item = getStockItemById(stockItemId);
  if (!item || !state.stockQrSvg || state.stockQrItemId !== stockItemId) {
    showFlash("Open the QR preview before printing.", "error");
    render();
    return;
  }

  const printWindow = window.open("", "_blank", "noopener,noreferrer");
  if (!printWindow) {
    showFlash("The browser blocked the print window.", "error");
    render();
    return;
  }

  const labelCode = buildStockQrPayload(item);
  const labelCaption = getStockLabelCaption(item);

  printWindow.document.write(`
    <!DOCTYPE html>
    <html lang="en">
      <head>
        <meta charset="UTF-8">
        <title>${escapeHtml(item.name)}</title>
        <style>
          @page { size: 40mm 14mm; margin: 0; }
          html, body { width: 40mm; height: 14mm; margin: 0; padding: 0; }
          body { font-family: Arial, sans-serif; color: #173c34; }
          .label {
            width: 40mm;
            height: 14mm;
            box-sizing: border-box;
            padding: 0.8mm 1mm;
            border: 0.2mm solid #cddcd7;
            overflow: hidden;
            display: grid;
            grid-template-columns: 11mm minmax(0, 1fr);
            gap: 0.9mm;
            align-items: center;
          }
          .qr {
            width: 11mm;
            height: 11mm;
            display: flex;
            align-items: center;
            justify-content: center;
          }
          .qr svg { width: 100%; height: 100%; display: block; }
          .details {
            min-width: 0;
            display: grid;
            gap: 0.45mm;
          }
          .caption {
            margin: 0;
            font-size: 1.95mm;
            line-height: 1.05;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
          }
          .code {
            margin: 0;
            font-size: 2.1mm;
            font-weight: 700;
            line-height: 1;
            letter-spacing: 0.16mm;
          }
        </style>
      </head>
      <body>
        <div class="label">
          <div class="qr">${state.stockQrSvg}</div>
          <div class="details">
            <p class="caption">${escapeHtml(labelCaption)}</p>
            <p class="code">${escapeHtml(labelCode)}</p>
          </div>
        </div>
        <script>window.print();</script>
      </body>
    </html>
  `);
  printWindow.document.close();
}

async function shareStockQrLabel(stockItemId) {
  const item = getStockItemById(stockItemId);
  if (!item || !state.stockQrSvg || state.stockQrItemId !== stockItemId) {
    showFlash("Open the QR preview before sharing.", "error");
    render();
    return;
  }

  if (!supportsStockLabelShare()) {
    showFlash("This browser cannot share the label file.", "error");
    render();
    return;
  }

  state.stockQrSharing = true;
  render();

  let file;
  try {
    file = await createStockLabelFile(item, state.stockQrSvg);
  } catch (error) {
    state.stockQrSharing = false;
    showFlash(normalizeError(error), "error");
    render();
    return;
  }

  if (typeof navigator.canShare === "function" && !navigator.canShare({ files: [file] })) {
    state.stockQrSharing = false;
    showFlash("This browser cannot share the label file.", "error");
    render();
    return;
  }

  try {
    await navigator.share({
      title: `${item.name} QR label`,
      text: getReferenceLines(item).join(" | ") || `QR label for ${item.name}`,
      files: [file],
    });
    showFlash("PNG QR label shared. Choose Labelnize to import it.", "success");
  } catch (error) {
    if (String(error?.name || "") !== "AbortError") {
      showFlash(normalizeError(error), "error");
    }
  } finally {
    state.stockQrSharing = false;
    render();
  }
}

function supportsStockLabelShare() {
  return typeof navigator.share === "function" && typeof File === "function";
}

async function createStockLabelFile(item, svg) {
  const nameParts = [
    sanitizeFileStem(item?.name || ""),
    sanitizeFileStem(buildStockQrPayload(item)),
  ].filter(Boolean);
  const fileName = `${nameParts.join("-") || "stock-label"}-qr-label.png`;
  const blob = await renderStockLabelPng(item, svg);
  return new File([blob], fileName, { type: "image/png" });
}

function sanitizeFileStem(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

async function copyStockQrValue(stockItemId) {
  const item = getStockItemById(stockItemId);
  if (!item) {
    showFlash("Stock item not found.", "error");
    render();
    return;
  }

  try {
    await copyText(buildStockQrPayload(item));
    showFlash("QR value copied.", "success");
  } catch (error) {
    showFlash(normalizeError(error), "error");
  } finally {
    render();
  }
}

async function renderStockLabelPng(item, svg) {
  const width = 1200;
  const height = 420;
  const padding = 40;
  const qrSize = 300;
  const labelCode = buildStockQrPayload(item);
  const caption = getStockLabelCaption(item);
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d");
  if (!context) {
    throw new Error("This browser cannot prepare the label image.");
  }

  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, width, height);
  context.strokeStyle = "#cddcd7";
  context.lineWidth = 6;
  context.strokeRect(6, 6, width - 12, height - 12);

  context.fillStyle = "#173c34";
  context.font = "600 38px Arial";
  context.textAlign = "left";
  context.textBaseline = "middle";
  context.fillText(trimCanvasText(context, caption, width - ((padding * 3) + qrSize)), qrSize + (padding * 2), 110);

  const qrImage = await loadSvgImage(svg);
  context.drawImage(qrImage, padding, 60, qrSize, qrSize);

  context.font = "700 58px Arial";
  context.textAlign = "left";
  context.fillText(labelCode, qrSize + (padding * 2), 210);

  const blob = await canvasToBlob(canvas, "image/png");
  return blob;
}

function trimCanvasText(context, text, maxWidth) {
  let value = String(text || "").trim();
  if (context.measureText(value).width <= maxWidth) {
    return value;
  }

  while (value && context.measureText(`${value}...`).width > maxWidth) {
    value = value.slice(0, -1).trimEnd();
  }

  return `${value || ""}...`;
}

function getStockLabelCaption(item) {
  return String(item?.sku || item?.quoteNumber || item?.name || "Stock item").trim();
}

function loadSvgImage(svg) {
  return new Promise((resolve, reject) => {
    const objectUrl = URL.createObjectURL(new Blob([svg], { type: "image/svg+xml;charset=utf-8" }));
    const image = new Image();
    image.onload = () => {
      URL.revokeObjectURL(objectUrl);
      resolve(image);
    };
    image.onerror = () => {
      URL.revokeObjectURL(objectUrl);
      reject(new Error("The QR image could not be prepared."));
    };
    image.src = objectUrl;
  });
}

function canvasToBlob(canvas, type) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (blob) {
        resolve(blob);
        return;
      }

      reject(new Error("The label image could not be prepared."));
    }, type);
  });
}

async function copyText(text) {
  if (navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(text);
    return;
  }

  const input = document.createElement("textarea");
  input.value = String(text || "");
  input.setAttribute("readonly", "true");
  input.style.position = "fixed";
  input.style.opacity = "0";
  document.body.appendChild(input);
  input.focus();
  input.select();

  const copied = document.execCommand("copy");
  document.body.removeChild(input);

  if (!copied) {
    throw new Error("This browser could not copy the QR value.");
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
            Manage users, pickup locations, stock records, driver entries, and email delivery from one live database.
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
            Track incoming and outgoing stock, keep a live on-hand view, and send artwork requests from the same workspace.
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
          <p class="muted">Use the Stock page to log movements, review history, and send artwork requests once stock is queued.</p>
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
      title: "Stock control and cleanup",
      subtitle: "Record stock in and out, send artwork requests, and remove unused stock items before they build history.",
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
      ${renderPageSummaryCard("Stock", "Track stock, send artwork requests, and remove unused items before history exists.", "stock")}
      ${renderPageSummaryCard("Network", "Maintain pickup locations with optional contact details.", "network")}
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
      <p>Keep a running ledger of stock receipts and issues, review stock history, and request artwork when new work is needed.</p>
    </section>
    ${renderFlash()}
    <section class="metrics">
      ${renderMetric("Stock items", stockItems.length)}
      ${renderMetric("On hand", getStockOnHandTotal())}
      ${renderMetric("Movements", stockMovements.length)}
      ${renderMetric("Artwork requests", artworkRequests.length)}
    </section>
    <section class="panel-grid">
      ${renderPageSummaryCard("Stock", "Log stock in/out, review history, and send artwork requests.", "stock")}
    </section>
  `;
}

function renderStockWorkspace({ viewerRole, title, subtitle }) {
  const stockMovements = state.snapshot.stockMovements;
  const recentInbound = getRecentStockMovements("in");
  const recentOutbound = getRecentStockMovements("out");

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
      ${renderMetric("Movements", stockMovements.length)}
      ${renderMetric("Recent arrivals", recentInbound.length)}
      ${renderMetric("Recent shipments", recentOutbound.length)}
      ${renderMetric("Artwork requests", state.snapshot.artworkRequests.length)}
    </section>
    ${renderRecentStockActivitySection({ recentInbound, recentOutbound })}
    <section class="panel-grid">
      ${renderStockItemPanel(viewerRole)}
      ${renderStockMovementPanel()}
    </section>
    ${renderStockQrPreviewPanel()}
    ${renderStockItemsSection(viewerRole)}
    ${renderStockMovementsSection(viewerRole)}
    ${renderArtworkRequestPanel(viewerRole)}
    ${renderArtworkRequestsSection()}
  `;
}

function renderStockItemPanel(viewerRole) {
  const editingItem = viewerRole === "admin" ? getEditingStockItem() : null;
  const isEditing = Boolean(editingItem);

  return `
    <article class="panel">
      <p class="eyebrow">Stock master</p>
      <h3 class="panel-title">${isEditing ? "Edit stock item" : "Add stock item"}</h3>
      <p class="panel-subtitle">
        ${
          isEditing
            ? "Update the stock master record. Existing movement history stays linked to this item."
            : "Create the item once with its order references, then log movements and artwork requests against it."
        }
      </p>
      ${
        isEditing
          ? `
            <div class="stock-edit-banner">
              <span class="chip">Editing stock item</span>
              <span>Last updated ${escapeHtml(formatDateTime(editingItem.updatedAt || editingItem.createdAt) || "recently")}.</span>
            </div>
          `
          : ""
      }
      <form data-form="add-stock-item">
        <input name="stockItemId" type="hidden" value="${escapeHtml(editingItem?.id || "")}">
        <label>
          Stock description
          <input name="name" type="text" value="${escapeHtml(editingItem?.name || "")}" required>
        </label>
        <div class="form-grid">
          <label>
            Quote number
            <input name="quoteNumber" type="text" value="${escapeHtml(editingItem?.quoteNumber || "")}" required>
          </label>
          <label>
            Sales order number
            <input name="salesOrderNumber" type="text" value="${escapeHtml(editingItem?.salesOrderNumber || "")}" placeholder="Optional">
          </label>
        </div>
        <div class="form-grid">
          <label>
            Invoice number
            <input name="invoiceNumber" type="text" value="${escapeHtml(editingItem?.invoiceNumber || "")}" placeholder="Optional">
          </label>
          <label>
            PO number
            <input name="poNumber" type="text" value="${escapeHtml(editingItem?.poNumber || "")}" placeholder="Optional">
          </label>
        </div>
        <div class="form-grid">
          <label>
            Stock code
            <input name="sku" type="text" value="${escapeHtml(editingItem?.sku || "")}" placeholder="SKU-001">
          </label>
          <label>
            Unit
            <input name="unit" type="text" value="${escapeHtml(editingItem?.unit || "units")}">
          </label>
        </div>
        <label>
          Notes
          <textarea name="notes" placeholder="Material, finish, size, or any internal note">${escapeHtml(editingItem?.notes || "")}</textarea>
        </label>
        <div class="action-row">
          <button type="submit" class="button button-primary"${state.busy ? " disabled" : ""}>
            ${isEditing ? "Save stock item" : "Add stock item"}
          </button>
          ${
            isEditing
              ? `
                <button type="button" class="button button-ghost" data-action="cancel-edit-stock-item"${state.busy ? " disabled" : ""}>
                  Cancel
                </button>
              `
              : ""
          }
        </div>
      </form>
    </article>
  `;
}

function renderStockMovementPanel() {
  const editingMovement = getEditingStockMovement();
  const isEditing = Boolean(editingMovement);
  const movementType = editingMovement?.movementType === "out" ? "out" : "in";
  const isInbound = movementType === "in";
  const selectedStockItemId = state.stockMovementSelectedItemId || editingMovement?.stockItemId || "";

  return `
    <article class="panel">
      <p class="eyebrow">Stock ledger</p>
      <h3 class="panel-title">${isEditing ? "Edit stock movement" : "Log stock in or out"}</h3>
      <p class="panel-subtitle">
        ${
          isEditing
            ? "Update the selected movement. The original logged time and recorded-by details stay unchanged."
            : "Every stock movement is timestamped and linked to the staff member who recorded it."
        }
      </p>
      ${
        isEditing
          ? `
            <div class="stock-edit-banner">
              <span class="chip ${movementType === "in" ? "chip-success" : "chip-warning"}">
                ${movementType === "in" ? "Editing receipt" : "Editing issue"}
              </span>
              <span>
                Logged ${escapeHtml(formatDateTime(editingMovement.createdAt) || "")}
                by ${escapeHtml(editingMovement.createdByName || "Unknown")}.
              </span>
            </div>
          `
          : ""
      }
      <div class="action-row">
        <button type="button" class="button button-secondary" data-action="open-stock-scanner"${state.busy ? " disabled" : ""}>
          Scan QR
        </button>
        ${
          isEditing
            ? `
              <button type="button" class="button button-ghost" data-action="cancel-edit-stock-movement"${state.busy ? " disabled" : ""}>
                Cancel edit
              </button>
            `
            : `
              <button type="button" class="button button-ghost" data-action="clear-stock-selection"${state.busy || !state.stockMovementSelectedItemId ? " disabled" : ""}>
                Clear selection
              </button>
            `
        }
      </div>
      ${renderStockScannerPanel()}
      <form data-form="add-stock-movement">
        <input name="stockMovementId" type="hidden" value="${escapeHtml(editingMovement?.id || "")}">
        <div class="form-grid">
          <label>
            Stock item
            <select name="stockItemId" required>
              ${renderStockItemOptions(selectedStockItemId)}
            </select>
          </label>
          <label>
            Movement type
            <select name="movementType" required>
              <option value="in"${movementType === "in" ? " selected" : ""}>Stock in</option>
              <option value="out"${movementType === "out" ? " selected" : ""}>Stock out</option>
            </select>
          </label>
        </div>
        <div class="form-grid">
          <label>
            Quantity
            <input name="quantity" type="number" min="1" step="1" value="${escapeHtml(String(editingMovement?.quantity || ""))}" required>
          </label>
          <label data-stock-supplier-field${isInbound ? "" : ' class="hidden"'}>
            Supplier
            <input name="supplierName" type="text" value="${escapeHtml(editingMovement?.supplierName || "")}" placeholder="Supplier name">
          </label>
          <label data-stock-driver-field${isInbound ? ' class="hidden"' : ""}>
            Driver
            <select name="driverUserId">
              ${renderDriverOptions(editingMovement?.driverUserId || "", true)}
            </select>
          </label>
        </div>
        <p class="field-note" data-stock-movement-note>
          ${isInbound ? "Record which supplier the stock arrived from." : "Record which driver the stock went out with."}
        </p>
        <label>
          Notes
          <textarea name="notes" placeholder="Batch note, dispatch note, or anything the team should keep on record">${escapeHtml(editingMovement?.notes || "")}</textarea>
        </label>
        <button type="submit" class="button button-primary"${state.busy ? " disabled" : ""}>
          ${isEditing ? "Save changes" : "Save movement"}
        </button>
      </form>
    </article>
  `;
}

function renderArtworkRequestPanel(viewerRole) {
  return renderStockDisclosure({
    tag: "article",
    containerClass: "panel",
    sectionKey: "artwork-form",
    eyebrow: "Artwork",
    title: "Email artwork request",
    subtitle: `Send a request to ${state.artworkTo} so the artwork department can prepare what logistics needs.`,
    summary: "Collapsed until you need to send a request.",
    open: state.stockArtworkPanelOpen,
    body: `
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
    `,
  });
}

function renderStockItemsSection(viewerRole) {
  const allowDelete = viewerRole === "admin";
  const filteredItems = getFilteredStockItems();
  const latestMovementByItemId = getLatestStockMovementByItemId();
  const totalItems = state.snapshot.stockItems.length;

  return `
    <section class="table-card">
      <div class="table-toolbar stock-table-toolbar">
        <div class="stock-section-copy">
          <p class="eyebrow">Stock register</p>
          <h3 class="panel-title">On-hand summary</h3>
          <p class="panel-subtitle">
            ${
              allowDelete
                ? "Admins can delete stock items at any time. Deleting an item also removes its movement and artwork history."
                : "Each item stays linked to its movement and artwork history for traceability."
            }
          </p>
          <p class="stock-results-note">${escapeHtml(getStockItemSearchSummary(filteredItems.length, totalItems))}</p>
        </div>
        <label class="stock-search">
          Search order numbers
          <div class="stock-search-controls">
            <input
              type="search"
              value="${escapeHtml(state.stockSearchQuery)}"
              placeholder="Quote, sales order, invoice, PO, stock code"
              autocapitalize="characters"
              autocomplete="off"
              spellcheck="false"
              data-stock-search
            >
            ${
              state.stockSearchQuery
                ? `<button type="button" class="button button-ghost" data-action="clear-stock-search"${state.busy ? " disabled" : ""}>Clear</button>`
                : ""
            }
          </div>
        </label>
      </div>
      <div class="table-scroll stock-items-table">
        <table class="responsive-stack">
          <thead>
            <tr>
              <th>Item</th>
              <th>References</th>
              <th>Stock code</th>
              <th>Unit</th>
              <th>On hand</th>
              <th>Notes</th>
              <th>Updated</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
            ${
              filteredItems.length
                ? filteredItems.map((item) => renderStockItemRow(item, viewerRole, latestMovementByItemId)).join("")
                : `
                  <tr>
                    <td colspan="8">${totalItems ? "No stock items match this search." : "No stock items added yet."}</td>
                  </tr>
                `
            }
          </tbody>
        </table>
      </div>
      <div class="stock-item-mobile-list">
        ${
          filteredItems.length
            ? filteredItems.map((item) => renderStockItemCard(item, viewerRole, latestMovementByItemId)).join("")
            : `<div class="empty-state">${escapeHtml(totalItems ? "No stock items match this search." : "No stock items added yet.")}</div>`
        }
      </div>
    </section>
  `;
}

function renderRecentStockActivitySection({ recentInbound, recentOutbound }) {
  return `
    <section class="table-card">
      <div class="table-toolbar">
        <div class="stock-section-copy">
          <p class="eyebrow">Recent activity</p>
          <h3 class="panel-title">Recently arrived and shipped stock</h3>
          <p class="panel-subtitle">
            Latest stock activity from the last ${escapeHtml(String(STOCK_RECENT_ACTIVITY_DAYS))} days, grouped by item.
          </p>
        </div>
      </div>
      <div class="recent-stock-grid">
        ${renderRecentStockActivityColumn({
          title: "Recently arrived",
          subtitle: "Latest stock-in activity per item.",
          emptyLabel: `No stock arrived in the last ${STOCK_RECENT_ACTIVITY_DAYS} days.`,
          movements: recentInbound,
        })}
        ${renderRecentStockActivityColumn({
          title: "Recently shipped",
          subtitle: "Latest stock-out activity per item.",
          emptyLabel: `No stock shipped in the last ${STOCK_RECENT_ACTIVITY_DAYS} days.`,
          movements: recentOutbound,
        })}
      </div>
    </section>
  `;
}

function renderRecentStockActivityColumn({ title, subtitle, emptyLabel, movements }) {
  const isInbound = title.toLowerCase().includes("arrived");

  return `
    <article class="recent-stock-column">
      <div class="recent-stock-column-head">
        <div>
          <h4 class="recent-stock-column-title">${escapeHtml(title)}</h4>
          <p class="recent-stock-column-note">${escapeHtml(subtitle)}</p>
        </div>
        <span class="chip ${isInbound ? "chip-success" : "chip-warning"}">${escapeHtml(String(movements.length))}</span>
      </div>
      ${
        movements.length
          ? `
            <div class="recent-stock-list">
              ${movements.map((movement) => renderRecentStockActivityEntry(movement)).join("")}
            </div>
          `
          : `<p class="stock-results-note">${escapeHtml(emptyLabel)}</p>`
      }
    </article>
  `;
}

function renderRecentStockActivityEntry(movement) {
  const item = getStockItemById(movement.stockItemId);
  const onHandLabel = item
    ? `${Number(item.onHandQuantity || 0)} ${item.unit || movement.unit || "units"} on hand now`
    : "";
  const referenceSummary = getReferenceLines(movement).join(" | ") || "No order references";

  return `
    <article class="recent-stock-entry">
      <div class="recent-stock-entry-head">
        <div class="recent-stock-entry-copy">
          <h5 class="recent-stock-entry-title">${escapeHtml(movement.itemName || "Unknown")}</h5>
          <p class="recent-stock-entry-summary">${escapeHtml(referenceSummary)}</p>
        </div>
        <span class="chip ${movement.movementType === "in" ? "chip-success" : "chip-warning"}">
          ${escapeHtml(String(movement.quantity || 0))} ${escapeHtml(movement.unit || "units")}
        </span>
      </div>
      <div class="recent-stock-meta">
        <span>${escapeHtml(formatDateTime(movement.createdAt) || "")}</span>
        <span>${escapeHtml(getStockMovementPartyLabel(movement))}</span>
        ${onHandLabel ? `<span>${escapeHtml(onHandLabel)}</span>` : ""}
      </div>
    </article>
  `;
}

function renderStockMovementsSection(viewerRole) {
  return renderStockDisclosure({
    sectionKey: "movements",
    eyebrow: "Stock history",
    title: "Movement ledger",
    subtitle: "Use Edit to correct the item, movement type, quantity, party, or notes for an existing stock entry.",
    summary: `${state.snapshot.stockMovements.length} logged stock movement${state.snapshot.stockMovements.length === 1 ? "" : "s"}.`,
    open: state.stockMovementsSectionOpen,
    body: `
      <div class="table-scroll">
        <table class="responsive-stack">
          <thead>
            <tr>
              <th>Logged</th>
              <th>Item</th>
              <th>Type</th>
              <th>Quantity</th>
              <th>Supplier / driver</th>
              <th>Notes</th>
              <th>Recorded by</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
            ${
              state.snapshot.stockMovements.length
                ? state.snapshot.stockMovements.map((movement) => renderStockMovementRow(movement, viewerRole)).join("")
                : `
                  <tr>
                    <td colspan="8">No stock movements logged yet.</td>
                  </tr>
                `
            }
          </tbody>
        </table>
      </div>
    `,
  });
}

function renderStockQrPreviewPanel() {
  const item = getStockItemById(state.stockQrItemId);
  if (!item) {
    return "";
  }

  return `
    <section class="table-card qr-preview-card">
      <div class="table-toolbar">
        <div>
          <p class="eyebrow">QR Label</p>
          <h3 class="panel-title">${escapeHtml(item.name)}</h3>
          <p class="panel-subtitle">${escapeHtml(getStockLabelCaption(item))}</p>
        </div>
        <div class="action-row">
          <button class="button button-secondary" data-action="print-stock-qr" data-stock-item-id="${item.id}"${state.stockQrBusy || state.stockQrSharing || !state.stockQrSvg ? " disabled" : ""}>
            Print label
          </button>
          <button class="button button-secondary" data-action="share-stock-qr" data-stock-item-id="${item.id}"${state.stockQrBusy || state.stockQrSharing || !state.stockQrSvg || !supportsStockLabelShare() ? " disabled" : ""}>
            ${state.stockQrSharing ? "Sharing..." : "Share QR"}
          </button>
          <button class="button button-secondary" data-action="copy-stock-qr" data-stock-item-id="${item.id}"${state.stockQrBusy || state.stockQrSharing ? " disabled" : ""}>
            Copy QR value
          </button>
          <button class="button button-ghost" data-action="close-stock-qr"${state.stockQrBusy || state.stockQrSharing ? " disabled" : ""}>
            Close
          </button>
        </div>
      </div>
      ${
        state.stockQrBusy
          ? '<p class="field-note">Generating QR label...</p>'
          : `
            <div class="qr-preview-body">
              <div class="qr-preview-code">${state.stockQrSvg}</div>
              <div class="qr-preview-meta">
                <span>QR value: ${escapeHtml(buildStockQrPayload(item))}</span>
                <span>Format: QR code</span>
                ${getReferenceLines(item).map((line) => `<span>${escapeHtml(line)}</span>`).join("")}
                <span>Stock code: ${escapeHtml(item.sku || "Not set")}</span>
              </div>
            </div>
          `
      }
    </section>
  `;
}

function renderStockScannerPanel() {
  if (!state.stockScannerOpen) {
    return "";
  }

  const pendingItem = getStockItemById(state.stockScannerPendingItemId);

  return `
    <section class="scanner-panel">
      <p class="field-note">${escapeHtml(state.stockScannerStatus || getStockScannerHint())}</p>
      ${
        pendingItem
          ? `
            <div class="scan-confirm-card">
              <p class="eyebrow">Scan found</p>
              <h4 class="panel-title">${escapeHtml(pendingItem.name)}</h4>
              <div class="scan-confirm-meta">
                <span>QR value: ${escapeHtml(state.stockScannerPendingCode || buildStockQrPayload(pendingItem))}</span>
                <span>Stock code: ${escapeHtml(pendingItem.sku || "Not set")}</span>
                <span>${escapeHtml(getReferenceLines(pendingItem).join(" | ") || "No order references set.")}</span>
              </div>
              <div class="action-row">
                <button type="button" class="button button-primary" data-action="confirm-stock-scan"${state.busy ? " disabled" : ""}>
                  Use this item
                </button>
                <button type="button" class="button button-secondary" data-action="retry-stock-scan"${state.busy ? " disabled" : ""}>
                  Scan again
                </button>
                <button type="button" class="button button-ghost" data-action="close-stock-scanner"${state.busy ? " disabled" : ""}>
                  Close scanner
                </button>
              </div>
            </div>
          `
          : `
            <div class="scanner-grid">
              <div class="scanner-video-wrap">
                <video data-stock-scanner-video class="scanner-video" muted autoplay playsinline></video>
              </div>
              <div class="scanner-actions">
                <button type="button" class="button button-primary" data-action="start-stock-camera"${state.busy || !canUseLiveStockScanner() ? " disabled" : ""}>
                  Start camera
                </button>
                ${
                  state.stockScannerZoomSupported
                    ? `
                      <label class="scanner-zoom">
                        Zoom
                        <input
                          type="range"
                          min="${escapeHtml(String(state.stockScannerZoomMin))}"
                          max="${escapeHtml(String(state.stockScannerZoomMax))}"
                          step="${escapeHtml(String(state.stockScannerZoomStep))}"
                          value="${escapeHtml(String(state.stockScannerZoomValue))}"
                          data-stock-scanner-zoom
                        >
                        <span class="field-note">Move closer first. Use zoom only if the QR still looks soft.</span>
                      </label>
                    `
                    : ""
                }
                <label class="scanner-upload">
                  Upload QR image
                  <input type="file" accept="image/*" capture="environment" data-stock-scan-upload${supportsStockQrDetection() ? "" : " disabled"}>
                </label>
                <div class="scanner-manual">
                  <label>
                    Enter QR value
                    <input type="text" data-stock-manual-value placeholder="Type the code printed on the label" autocapitalize="characters" autocomplete="off" spellcheck="false">
                  </label>
                  <button type="button" class="button button-secondary" data-action="apply-stock-manual-qr"${state.busy ? " disabled" : ""}>
                    Use QR value
                  </button>
                </div>
                <button type="button" class="button button-ghost" data-action="close-stock-scanner"${state.busy ? " disabled" : ""}>
                  Close scanner
                </button>
              </div>
            </div>
          `
      }
    </section>
  `;
}

function renderArtworkRequestsSection() {
  return renderStockDisclosure({
    sectionKey: "artwork-log",
    eyebrow: "Artwork requests",
    title: "Email request log",
    summary: `${state.snapshot.artworkRequests.length} artwork request${state.snapshot.artworkRequests.length === 1 ? "" : "s"} sent.`,
    open: state.stockArtworkRequestsSectionOpen,
    body: `
      <div class="table-scroll">
        <table class="responsive-stack">
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
    `,
  });
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
      ${renderMetric("Priority stops", plan.priorityStopCount)}
      ${renderMetric("Estimated km", plan.totalKm.toFixed(1))}
      ${renderMetric("Completed entries", completedOrders.length)}
    </section>
    <section class="route-canvas-card">
      <p class="eyebrow">Route map</p>
      <h3 class="panel-title">${escapeHtml(getDriverRouteHeading(currentUser.id))}</h3>
      <div class="route-canvas-wrap">
        <div id="route-map" class="route-map" aria-label="${escapeHtml(getDriverRouteAriaLabel(currentUser.id))}"></div>
        <canvas id="route-canvas" class="route-canvas-fallback hidden" width="900" height="340"></canvas>
      </div>
      <p id="route-map-status" class="route-map-status hidden"></p>
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
  const editingUser = getEditingUser();
  const isEditing = Boolean(editingUser);
  const selectedRole = editingUser?.role || "sales";
  const selectedPhone = editingUser?.phone || "";

  return `
    <article class="panel">
      <p class="eyebrow">Accounts</p>
      <h3 class="panel-title">${isEditing ? "Edit team member" : "Add team member"}</h3>
      <p class="panel-subtitle">
        ${
          isEditing
            ? "Update the selected account details. Leave the password blank to keep the current password."
            : "Create admin, sales, driver, or logistics accounts with name-based login."
        }
      </p>
      <form data-form="add-account">
        <input name="userId" type="hidden" value="${escapeHtml(editingUser?.id || "")}">
        <div class="form-grid">
          <label>
            Name
            <input name="name" type="text" value="${escapeHtml(editingUser?.name || "")}" required>
          </label>
          <label>
            Role
            <select name="role" data-driver-role-select required>
              <option value="sales"${selectedRole === "sales" ? " selected" : ""}>Sales</option>
              <option value="driver"${selectedRole === "driver" ? " selected" : ""}>Driver</option>
              <option value="logistics"${selectedRole === "logistics" ? " selected" : ""}>Logistics</option>
              <option value="admin"${selectedRole === "admin" ? " selected" : ""}>Admin</option>
            </select>
          </label>
        </div>
        <div class="form-grid">
          <label>
            ${isEditing ? "Password reset" : "Password"}
            <input name="password" type="password"${isEditing ? "" : " required"}>
          </label>
          <label>
            Note
            <input type="text" value="${escapeHtml(isEditing ? "Leave password blank to keep it unchanged" : "Name is also the login field")}" readonly>
          </label>
        </div>
        <div id="driver-role-fields"${selectedRole === "driver" ? "" : ' class="hidden"'}>
          <div class="form-grid">
            <label>
              Driver phone
              <input name="phone" type="text" value="${escapeHtml(selectedPhone)}" placeholder="071 555 0100">
            </label>
          </div>
        </div>
        <div class="action-row">
          <button type="submit" class="button button-primary"${state.busy ? " disabled" : ""}>
            ${isEditing ? "Save account" : "Create account"}
          </button>
          ${
            isEditing
              ? `
                <button type="button" class="button button-ghost" data-action="cancel-edit-user"${state.busy ? " disabled" : ""}>
                  Cancel
                </button>
              `
              : ""
          }
        </div>
      </form>
    </article>
  `;
}

function renderSupplierNetworkSection() {
  const editingLocation = getEditingLocation();
  const isEditingLocation = Boolean(editingLocation);

  return `
    <section class="table-card">
      <p class="eyebrow">Location network</p>
      <h3 class="panel-title">Pickup locations</h3>
      <p class="panel-subtitle">Each location stores its own type, address, optional coordinates, and optional contact details.</p>
      <article class="panel">
        <p class="eyebrow">Locations</p>
        <h3 class="panel-title">${isEditingLocation ? "Edit location" : "Add location"}</h3>
        <p class="panel-subtitle">
          ${
            isEditingLocation
              ? "Update the selected location details. Contact details and coordinates can be left blank when unavailable."
              : "Contact details and coordinates can be left blank when they are not available."
          }
        </p>
        <form data-form="add-location">
          <input name="locationId" type="hidden" value="${escapeHtml(editingLocation?.id || "")}">
          <div class="form-grid">
            <label>
              Location name
              <input name="name" type="text" value="${escapeHtml(editingLocation?.name || "")}" required>
            </label>
            <label>
              Location type
              <select name="locationType" required>
                <option value="supplier"${editingLocation?.locationType === "supplier" ? " selected" : ""}>Supplier</option>
                <option value="factory"${editingLocation?.locationType === "factory" ? " selected" : ""}>Factory</option>
                <option value="both"${editingLocation?.locationType === "both" ? " selected" : ""}>Both</option>
              </select>
            </label>
          </div>
          <label>
            Physical address
            <input name="address" type="text" value="${escapeHtml(editingLocation?.address || "")}" required>
          </label>
          <div class="form-grid">
            <label>
              Latitude
              <input name="lat" type="number" step="0.000001" value="${escapeHtml(editingLocation?.lat ?? "")}" placeholder="Optional">
            </label>
            <label>
              Longitude
              <input name="lng" type="number" step="0.000001" value="${escapeHtml(editingLocation?.lng ?? "")}" placeholder="Optional">
            </label>
          </div>
          <div class="form-grid">
            <label>
              Contact person
              <input name="contactPerson" type="text" value="${escapeHtml(editingLocation?.contactPerson || "")}" placeholder="Optional">
            </label>
            <label>
              Contact number
              <input name="contactNumber" type="tel" value="${escapeHtml(editingLocation?.contactNumber || "")}" placeholder="Optional">
            </label>
          </div>
          <div class="action-row">
            <button type="submit" class="button button-primary"${state.busy ? " disabled" : ""}>
              ${isEditingLocation ? "Save location" : "Add location"}
            </button>
            ${
              isEditingLocation
                ? `
                  <button type="button" class="button button-ghost" data-action="cancel-edit-location"${state.busy ? " disabled" : ""}>
                    Cancel
                  </button>
                `
                : ""
            }
          </div>
        </form>
        <div class="table-scroll">
          <table class="responsive-stack">
            <thead>
              <tr>
                <th>Location</th>
                <th>Location type</th>
                <th>Physical address</th>
                <th>Contact</th>
                <th>Actions</th>
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
        <table class="responsive-stack">
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
          Quote number
          <input name="quoteNumber" type="text" required>
        </label>
        <label>
          Sales order number
          <input name="salesOrderNumber" type="text" placeholder="Optional">
        </label>
      </div>
      <div class="form-grid">
        <label>
          Invoice number
          <input name="invoiceNumber" type="text" placeholder="Optional">
        </label>
        <label>
          PO number
          <input name="poNumber" type="text" placeholder="Optional">
        </label>
      </div>
      <label class="inline-check">
        <input type="checkbox" name="moveToFactory" disabled>
        Move collected stock to a factory
      </label>
      <label class="hidden" data-move-to-factory-destination>
        Destination factory
        <select name="factoryDestinationLocationId">
          ${renderFactoryLocationOptions()}
        </select>
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
              This only bypasses the block when the same driver already has another active entry with the same quote number.
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
  const filteredOrders = getFilteredAssignmentOrders();
  const page = getPaginationData(filteredOrders, "assignments", PAGE_SIZES.assignments);

  return `
    <section class="table-card">
      <div class="table-toolbar">
        <div class="stock-section-copy">
          <p class="eyebrow">Assignments</p>
          <h3 class="panel-title">Active entry allocation</h3>
          <p class="panel-subtitle">Unassigned entries appear first. Move work onto a driver list when the route is ready, or return it to the queue.</p>
          <p class="stock-results-note">${escapeHtml(getAssignmentFilterSummary(filteredOrders.length))}</p>
        </div>
        <label class="assignment-filter">
          Filter by driver
          <select data-assignment-filter>
            ${renderAssignmentFilterOptions(state.assignmentDriverFilter)}
          </select>
        </label>
      </div>
      <div class="table-scroll">
        <table class="responsive-stack">
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
                    <td colspan="6">No active entries match this driver filter.</td>
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
  const canDelete = viewerRole === "admin";

  return `
    <section class="table-card">
      <div class="table-toolbar">
        <div>
          <p class="eyebrow">Global List</p>
          <h3 class="panel-title">All visible entries</h3>
          <p class="panel-subtitle">
            ${
              canDelete
                ? "This table combines every visible entry across the driver-separated lists. Admins can remove entries here if needed."
                : "This table combines every visible entry across the driver-separated lists."
            }
          </p>
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
        <table class="responsive-stack">
          <thead>
            <tr>
              <th>Reference</th>
              <th>Order references</th>
              <th>Type</th>
              <th>Move to factory</th>
              <th>Driver</th>
              <th>Pickup location</th>
              <th>Created by</th>
              <th>Status</th>
              <th>Notice</th>
              <th>Created</th>
              ${canDelete ? "<th>Actions</th>" : ""}
            </tr>
          </thead>
          <tbody>
            ${
              page.items.length
                ? page.items.map((order) => renderGlobalOrderRow(order, viewerRole)).join("")
                : `
                  <tr>
                    <td colspan="${canDelete ? "11" : "10"}">No entries available yet.</td>
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
                plan.priorityStopCount
                  ? `<span class="chip chip-priority-high">${plan.priorityStopCount} priority stop${plan.priorityStopCount === 1 ? "" : "s"}</span>`
                  : ""
              }
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
        <table class="responsive-stack">
          <thead>
            <tr>
              <th>Reference</th>
              <th>Order references</th>
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
                          <td data-label="Reference">${escapeHtml(order.reference)}</td>
                          <td data-label="Order references">${renderReferenceSummary(order)}</td>
                          <td data-label="Type">${renderTypeChip(order.entryType)}</td>
                          <td data-label="Completed">${escapeHtml(formatDateTime(order.completedAt) || "Not completed")}</td>
                        </tr>
                      `,
                    )
                    .join("")
                : `
                  <tr>
                    <td colspan="4">No completed entries yet.</td>
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

function renderGlobalOrderRow(order, viewerRole) {
  const canDelete = viewerRole === "admin";

  return `
    <tr>
      <td data-label="Reference">${escapeHtml(order.reference)}</td>
      <td data-label="Order references">${renderReferenceSummary(order)}</td>
      <td data-label="Type">
        <div class="chip-row">
          ${renderTypeChip(order.entryType)}
          ${renderOrderPriorityChip(order)}
        </div>
      </td>
      <td data-label="Move to factory">${renderMoveToFactoryValue(order)}</td>
      <td data-label="Driver">${renderDriverAssignmentValue(order)}</td>
      <td data-label="Pickup location">
        <strong>${escapeHtml(order.locationName || "Unknown")}</strong><br>
        <span class="muted">${escapeHtml(order.locationAddress || "")}</span>
      </td>
      <td data-label="Created by">${escapeHtml(order.createdByName || "Unknown")}</td>
      <td data-label="Status">${renderStatusChip(order.status)}</td>
      <td data-label="Notice">${renderOrderNotice(order, "None")}</td>
      <td data-label="Created">${escapeHtml(formatDateTime(order.createdAt))}</td>
      ${
        canDelete
          ? `
            <td data-label="Actions">
              <div class="action-row">
                <button
                  class="button button-danger"
                  data-action="delete-order"
                  data-order-id="${order.id}"
                  data-order-reference="${escapeHtml(order.reference)}"
                  ${state.busy ? " disabled" : ""}
                >
                  Delete
                </button>
              </div>
            </td>
          `
          : ""
      }
    </tr>
  `;
}

function renderAssignmentRow(order, viewerRole) {
  return `
    <tr>
      <td data-label="Entry">
        <strong>${escapeHtml(order.reference)}</strong><br>
        ${renderReferenceSummary(order)}
        <div class="chip-row">
          ${renderTypeChip(order.entryType)}
          ${renderOrderPriorityChip(order)}
          ${order.moveToFactory ? '<span class="chip chip-warning">Factory move</span>' : ""}
          ${renderOrderFlagChip(order)}
        </div>
        ${renderOrderNotice(order)}
      </td>
      <td data-label="Current driver">${renderDriverAssignmentValue(order)}</td>
      <td data-label="Pickup location">
        <strong>${escapeHtml(order.locationName || "Unknown")}</strong><br>
        <span class="muted">${escapeHtml(order.locationAddress || "")}</span>
      </td>
      <td data-label="Created by">
        ${escapeHtml(order.createdByName || "Unknown")}<br>
        <span class="muted">${escapeHtml(formatDateTime(order.createdAt))}</span>
      </td>
      <td data-label="Assign to">
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
      <td data-label="Action">
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
      <td data-label="Name">${escapeHtml(user.name)}</td>
      <td data-label="Role"><span class="chip chip-role-${user.role}">${capitalize(user.role)}</span></td>
      <td data-label="Status">${user.active ? '<span class="chip chip-success">Active</span>' : '<span class="chip chip-warning">Inactive</span>'}</td>
      <td data-label="Details">${detailParts.length ? detailParts.join("<br>") : "n/a"}</td>
      <td data-label="Actions">
        <div class="action-row">
          <button class="button button-secondary" data-action="edit-user" data-user-id="${user.id}"${state.busy ? " disabled" : ""}>
            Edit
          </button>
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
      <td data-label="Name">${escapeHtml(supplier.name)}</td>
      <td data-label="Contact person">${escapeHtml(supplier.contactPerson || "Not set")}</td>
      <td data-label="Contact number">${escapeHtml(supplier.contactNumber || "Not set")}</td>
      <td data-label="Factory">${supplier.factory ? "Yes" : "No"}</td>
      <td data-label="Locations">${locationCount}</td>
      <td data-label="Actions">
        <div class="action-row">
          <button class="button button-secondary" data-action="edit-supplier" data-supplier-id="${supplier.id}"${state.busy ? " disabled" : ""}>
            Edit
          </button>
          <button class="button button-danger" data-action="delete-supplier" data-supplier-id="${supplier.id}" data-supplier-name="${escapeHtml(supplier.name)}"${state.busy ? " disabled" : ""}>
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
      <td data-label="Location">
        <strong>${escapeHtml(location.name)}</strong>
      </td>
      <td data-label="Location type">${escapeHtml(capitalize(location.locationType || ""))}</td>
      <td data-label="Physical address">${escapeHtml(location.address)}</td>
      <td data-label="Contact">${escapeHtml(location.contactPerson || "Not set")}<br><span class="muted">${escapeHtml(location.contactNumber || "Not set")}</span></td>
      <td data-label="Actions">
        <div class="action-row">
          <button class="button button-secondary" data-action="edit-location" data-location-id="${location.id}"${state.busy ? " disabled" : ""}>
            Edit
          </button>
          <button class="button button-danger" data-action="delete-location" data-location-id="${location.id}" data-location-name="${escapeHtml(location.name)}"${state.busy ? " disabled" : ""}>
            Delete
          </button>
        </div>
      </td>
    </tr>
  `;
}

function getReferenceLines(record) {
  const quoteNumber = String(record?.quoteNumber || record?.inhouseOrderNumber || "").trim();
  const salesOrderNumber = String(record?.salesOrderNumber || record?.factoryOrderNumber || "").trim();
  const invoiceNumber = String(record?.invoiceNumber || "").trim();
  const poNumber = String(record?.poNumber || "").trim();
  const lines = [];

  if (quoteNumber) {
    lines.push(`Quote ${quoteNumber}`);
  }

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

function renderReferenceSummary(record, emptyLabel = "No order references") {
  const lines = getReferenceLines(record);
  if (!lines.length) {
    return `<span class="muted">${escapeHtml(emptyLabel)}</span>`;
  }

  return lines.map((line) => `<span class="muted">${escapeHtml(line)}</span>`).join("<br>");
}

function renderRecentStockItemBadge(movement) {
  if (!movement || !isRecentStockMovement(movement.createdAt)) {
    return "";
  }

  return `
    <div class="chip-row stock-item-activity">
      <span class="chip ${movement.movementType === "in" ? "chip-success" : "chip-warning"}">
        ${movement.movementType === "in" ? "Recent arrival" : "Recent shipment"}
      </span>
    </div>
  `;
}

function renderStockItemRow(item, viewerRole, latestMovementByItemId = getLatestStockMovementByItemId()) {
  const allowDelete = viewerRole === "admin";
  const isEditing = state.editingStockItemId === item.id;
  const activityBadge = renderRecentStockItemBadge(latestMovementByItemId.get(item.id));
  const actions = [];

  if (allowDelete) {
    actions.push(
      `<button class="button button-secondary" data-action="edit-stock-item" data-stock-item-id="${item.id}"${state.busy || isEditing ? " disabled" : ""}>${isEditing ? "Editing" : "Edit"}</button>`,
    );
  }

  actions.push(
    `<button class="button button-secondary" data-action="open-stock-qr" data-stock-item-id="${item.id}"${state.busy ? " disabled" : ""}>QR</button>`,
  );

  if (allowDelete) {
    actions.push(
      `<button class="button button-danger" data-action="delete-stock-item" data-stock-item-id="${item.id}"${state.busy ? " disabled" : ""}>Delete</button>`,
    );
  }

  return `
    <tr>
      <td data-label="Item">
        <strong>${escapeHtml(item.name)}</strong>
        ${activityBadge}
      </td>
      <td data-label="References">${renderReferenceSummary(item)}</td>
      <td data-label="Stock code">${escapeHtml(item.sku || "Not set")}</td>
      <td data-label="Unit">${escapeHtml(item.unit || "units")}</td>
      <td data-label="On hand">${escapeHtml(String(item.onHandQuantity || 0))}</td>
      <td data-label="Notes">${escapeHtml(item.notes || "None")}</td>
      <td data-label="Updated">${escapeHtml(formatDateTime(item.updatedAt || item.createdAt) || "Not updated")}</td>
      <td data-label="Actions">
        <div class="action-row">
          ${actions.join("")}
        </div>
      </td>
    </tr>
  `;
}

function renderStockItemCard(item, viewerRole, latestMovementByItemId = getLatestStockMovementByItemId()) {
  const allowDelete = viewerRole === "admin";
  const isOpen = Boolean(state.stockOpenItemCards[item.id]);
  const isEditing = state.editingStockItemId === item.id;
  const referenceSummary = getReferenceLines(item).join(" | ") || "No order references set.";
  const activityBadge = renderRecentStockItemBadge(latestMovementByItemId.get(item.id));
  const actions = [];

  if (allowDelete) {
    actions.push(
      `<button class="button button-secondary" data-action="edit-stock-item" data-stock-item-id="${item.id}"${state.busy || isEditing ? " disabled" : ""}>${isEditing ? "Editing" : "Edit"}</button>`,
    );
  }

  actions.push(
    `<button class="button button-secondary" data-action="open-stock-qr" data-stock-item-id="${item.id}"${state.busy ? " disabled" : ""}>QR</button>`,
  );

  if (allowDelete) {
    actions.push(
      `<button class="button button-danger" data-action="delete-stock-item" data-stock-item-id="${item.id}"${state.busy ? " disabled" : ""}>Delete</button>`,
    );
  }

  return `
    <article class="stock-item-mobile-card${isOpen ? " is-open" : ""}">
      <div class="stock-item-mobile-head">
        <div class="stock-item-mobile-copy">
          <h4 class="stock-item-mobile-title">${escapeHtml(item.name)}</h4>
          <p class="stock-item-mobile-summary">${escapeHtml(referenceSummary)}</p>
          ${activityBadge}
        </div>
        <div class="stock-item-mobile-side">
          <span class="stock-item-mobile-onhand">${escapeHtml(String(item.onHandQuantity || 0))} ${escapeHtml(item.unit || "units")}</span>
          <button
            type="button"
            class="button button-ghost stock-item-mobile-toggle"
            data-action="toggle-stock-item-card"
            data-stock-item-id="${item.id}"
            ${state.busy ? " disabled" : ""}
          >
            ${isOpen ? "Hide details" : "Show details"}
          </button>
        </div>
      </div>
      ${
        isOpen
          ? `
            <div class="stock-item-mobile-body">
              <div class="stock-item-mobile-grid">
                <div class="stock-item-mobile-field">
                  <span class="stock-item-mobile-label">Stock code</span>
                  <strong>${escapeHtml(item.sku || "Not set")}</strong>
                </div>
                <div class="stock-item-mobile-field">
                  <span class="stock-item-mobile-label">Updated</span>
                  <strong>${escapeHtml(formatDateTime(item.updatedAt || item.createdAt) || "Not updated")}</strong>
                </div>
                <div class="stock-item-mobile-field">
                  <span class="stock-item-mobile-label">References</span>
                  <div>${renderReferenceSummary(item)}</div>
                </div>
                <div class="stock-item-mobile-field">
                  <span class="stock-item-mobile-label">Notes</span>
                  <strong>${escapeHtml(item.notes || "None")}</strong>
                </div>
              </div>
              <div class="action-row">
                ${actions.join("")}
              </div>
            </div>
          `
          : ""
      }
    </article>
  `;
}

function renderStockMovementRow(movement, viewerRole) {
  const canEdit = viewerRole === "admin" || viewerRole === "logistics";
  const isEditing = state.editingStockMovementId === movement.id;

  return `
    <tr>
      <td data-label="Logged">${escapeHtml(formatDateTime(movement.createdAt) || "")}</td>
      <td data-label="Item">
        <strong>${escapeHtml(movement.itemName || "Unknown")}</strong><br>
        <span class="muted">${escapeHtml(movement.sku || "No stock code")}</span><br>
        ${renderReferenceSummary(movement)}
      </td>
      <td data-label="Type">${movement.movementType === "in" ? '<span class="chip chip-success">Stock in</span>' : '<span class="chip chip-warning">Stock out</span>'}</td>
      <td data-label="Quantity">${escapeHtml(String(movement.quantity || 0))} ${escapeHtml(movement.unit || "units")}</td>
      <td data-label="Supplier / driver">${escapeHtml(getStockMovementPartyLabel(movement))}</td>
      <td data-label="Notes">${escapeHtml(movement.notes || "None")}</td>
      <td data-label="Recorded by">${escapeHtml(movement.createdByName || "Unknown")}</td>
      <td data-label="Actions">
        ${
          canEdit
            ? `
              <div class="action-row">
                <button
                  class="button button-secondary"
                  data-action="edit-stock-movement"
                  data-stock-movement-id="${movement.id}"
                  ${state.busy || isEditing ? " disabled" : ""}
                >
                  ${isEditing ? "Editing" : "Edit"}
                </button>
              </div>
            `
            : '<span class="muted">View only</span>'
        }
      </td>
    </tr>
  `;
}

function renderArtworkRequestRow(request) {
  return `
    <tr>
      <td data-label="Sent">${escapeHtml(formatDateTime(request.sentAt) || "")}</td>
      <td data-label="Item">
        <strong>${escapeHtml(request.itemName || "Unknown")}</strong><br>
        <span class="muted">${escapeHtml(request.sku || "No stock code")}</span><br>
        ${renderReferenceSummary(request)}
      </td>
      <td data-label="Quantity">${escapeHtml(String(request.requestedQuantity || 0))}</td>
      <td data-label="Requested by">${escapeHtml(request.requestedByName || "Unknown")}</td>
      <td data-label="Sent to">${escapeHtml(request.sentTo || state.artworkTo)}</td>
      <td data-label="Notes">${escapeHtml(request.notes || "None")}</td>
    </tr>
  `;
}

function renderStopCard(stop, index, viewerRole) {
  const allowComplete = viewerRole === "admin" || viewerRole === "driver";
  const allowDelete = viewerRole === "admin";
  const allowFlag = viewerRole === "admin" || viewerRole === "driver";
  const allowPriority = viewerRole === "admin";
  const allowNavigate = viewerRole === "admin" || viewerRole === "driver";
  const navigationUrl = allowNavigate ? getGoogleMapsNavigateUrl(stop.location) : "";
  const legLabel = stop.hasCoordinates && stop.legKm !== null
    ? `${stop.legKm.toFixed(1)} km leg`
    : "Coordinates pending";

  return `
    <article class="stop-card${stop.isPriority ? " stop-card-priority" : ""}">
      <div class="stop-header">
        <div>
          <p class="eyebrow">Stop ${index + 1}</p>
          <h4 class="stop-title">${escapeHtml(stop.location.name)}</h4>
          <p class="stop-address">${escapeHtml(stop.location.address)}</p>
        </div>
        <div class="chip-row">
          ${stop.isPriority ? '<span class="chip chip-priority-high">Priority stop</span>' : ""}
          <span class="chip">${legLabel}</span>
          <span class="chip">${stop.orders.length} entr${stop.orders.length === 1 ? "y" : "ies"}</span>
        </div>
      </div>
      ${
        navigationUrl
          ? `
            <div class="action-row stop-actions">
              <a
                class="button button-primary"
                href="${escapeHtml(navigationUrl)}"
                target="_blank"
                rel="noreferrer noopener"
              >
                Navigate
              </a>
            </div>
          `
          : ""
      }
      <div class="stop-orders">
        ${stop.orders
          .map((order) => {
            const isFlagging = state.flaggingOrderId === order.id;
            const canFlag = allowFlag && order.status === "active";
            const isPriority = isPriorityOrder(order);

            return `
              <div class="order-card${isPriority ? " order-card-priority" : ""}">
                <strong>${escapeHtml(order.reference)}</strong>
                <div class="order-meta">
                  ${getReferenceLines(order).map((line) => `<span>${escapeHtml(line)}</span>`).join("")}
                </div>
                <div class="chip-row">
                  ${renderTypeChip(order.entryType)}
                  ${renderOrderPriorityChip(order)}
                  ${order.moveToFactory ? '<span class="chip chip-warning">Factory move</span>' : ""}
                  ${renderOrderFlagChip(order)}
                  <span class="chip">Created by ${escapeHtml(order.createdByName)}</span>
                  ${renderStatusChip(order.status)}
                </div>
                ${renderOrderNotice(order)}
                <p>${escapeHtml(order.locationName || stop.location.name)}<br>${escapeHtml(order.locationAddress || stop.location.address)}</p>
                ${
                  allowComplete || allowDelete || canFlag || allowPriority
                    ? `
                      <div class="action-row">
                        ${
                          allowPriority
                            ? `
                              <button
                                class="button ${isPriority ? "button-secondary" : "button-primary"}"
                                data-action="toggle-order-priority"
                                data-order-id="${order.id}"
                                ${state.busy ? " disabled" : ""}
                              >
                                ${isPriority ? "Clear priority" : "Make priority"}
                              </button>
                            `
                            : ""
                        }
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
                          canFlag
                            ? `
                              <button
                                class="button button-ghost"
                                data-action="toggle-order-flag"
                                data-order-id="${order.id}"
                                ${state.busy ? " disabled" : ""}
                              >
                                ${isFlagging ? "Cancel flag" : order.driverFlagType ? "Update flag" : "Flag issue"}
                              </button>
                            `
                            : ""
                        }
                        ${
                          allowDelete
                            ? `
                              <button class="button button-danger" data-action="delete-order" data-order-id="${order.id}" data-order-reference="${escapeHtml(order.reference)}"${state.busy ? " disabled" : ""}>
                                Delete
                              </button>
                            `
                            : ""
                        }
                      </div>
                    `
                    : ""
                }
                ${isFlagging ? renderOrderFlagForm(order) : ""}
              </div>
            `;
          })
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

function getGoogleMapsNavigateUrl(location) {
  const destination = getStopNavigationDestination(location);
  if (!destination) {
    return "";
  }

  const params = new URLSearchParams({
    api: "1",
    destination,
    travelmode: "driving",
    dir_action: "navigate",
  });

  return `https://www.google.com/maps/dir/?${params.toString()}`;
}

function getStopNavigationDestination(location) {
  const coordinates = getCoordinates(location);
  if (coordinates) {
    return `${coordinates.lat},${coordinates.lng}`;
  }

  const name = String(location?.name || "").trim();
  const address = String(location?.address || "").trim();
  return [name, address].filter(Boolean).join(", ");
}

function getOrderPriority(order) {
  const priority = String(order?.priority || DEFAULT_ORDER_PRIORITY).trim().toLowerCase();
  return Object.prototype.hasOwnProperty.call(ORDER_PRIORITY_LEVELS, priority) ? priority : DEFAULT_ORDER_PRIORITY;
}

function getOrderPriorityRank(order) {
  return ORDER_PRIORITY_LEVELS[getOrderPriority(order)] ?? ORDER_PRIORITY_LEVELS[DEFAULT_ORDER_PRIORITY];
}

function isPriorityOrder(order) {
  return getOrderPriority(order) === PRIORITY_STOP_VALUE;
}

function renderOrderPriorityChip(order) {
  return isPriorityOrder(order) ? '<span class="chip chip-priority-high">Priority stop</span>' : "";
}

function renderOrderFlagChip(order) {
  const label = getOrderFlagLabel(order);
  if (!label) {
    return "";
  }

  return `<span class="chip chip-warning">Follow-up: ${escapeHtml(label)}</span>`;
}

function renderOrderFlagForm(order) {
  const hasFlag = Boolean(getOrderFlagLabel(order));
  const savedAt = formatDateTime(order?.driverFlaggedAt);
  const savedBy = String(order?.driverFlaggedByName || "").trim();
  const detailParts = [];

  if (hasFlag) {
    detailParts.push(`Current flag: ${getOrderFlagLabel(order)}.`);
    if (savedAt || savedBy) {
      detailParts.push(`Last updated ${[savedAt, savedBy ? `by ${savedBy}` : ""].filter(Boolean).join(" ")}.`);
    }
  } else {
    detailParts.push("Choose what happened at this stop and add any note the office should see.");
  }

  return `
    <form class="order-flag-form" data-form="flag-order">
      <input name="orderId" type="hidden" value="${escapeHtml(order.id)}">
      <label>
        Follow-up reason
        <select name="flagType" required>
          ${renderOrderFlagTypeOptions(order?.driverFlagType || "")}
        </select>
      </label>
      <label>
        Driver note
        <textarea name="flagNote" placeholder="What happened at this stop?">${escapeHtml(order?.driverFlagNote || "")}</textarea>
      </label>
      <p class="field-note">${escapeHtml(detailParts.join(" "))}</p>
      <div class="action-row">
        <button type="submit" class="button button-primary"${state.busy ? " disabled" : ""}>
          ${hasFlag ? "Save follow-up" : "Log follow-up"}
        </button>
        <button type="button" class="button button-ghost" data-action="cancel-order-flag"${state.busy ? " disabled" : ""}>
          Cancel
        </button>
        ${
          hasFlag
            ? `
              <button
                type="button"
                class="button button-secondary"
                data-action="clear-order-flag"
                data-order-id="${order.id}"
                ${state.busy ? " disabled" : ""}
              >
                Clear flag
              </button>
            `
            : ""
        }
      </div>
    </form>
  `;
}

function renderOrderFlagTypeOptions(selectedFlagType = "") {
  return Object.entries(ORDER_FLAG_LABELS)
    .map(
      ([value, label]) => `<option value="${value}"${value === selectedFlagType ? " selected" : ""}>${escapeHtml(label)}</option>`,
    )
    .join("");
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

function getFactoryLocations() {
  return state.snapshot.locations.filter((location) => ["factory", "both"].includes(location.locationType));
}

function renderDriverOptions(selectedDriverId = "", includeUnassigned = false) {
  const drivers = getActiveDriverUsers();
  const selectedDriver = selectedDriverId ? getUser(selectedDriverId) : null;
  const options = [];

  if (includeUnassigned) {
    options.push(`<option value=""${selectedDriverId ? "" : " selected"}>Unassigned</option>`);
  }

  if (
    selectedDriver
    && selectedDriver.role === "driver"
    && !drivers.some((driver) => driver.id === selectedDriver.id)
  ) {
    options.push(
      `<option value="${selectedDriver.id}" selected>${escapeHtml(selectedDriver.name)} (inactive)</option>`,
    );
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

function getAssignmentDriverFilterOptions() {
  const counts = new Map();
  let unassignedCount = 0;

  getActiveOrders().forEach((order) => {
    if (order.driverUserId) {
      counts.set(order.driverUserId, (counts.get(order.driverUserId) || 0) + 1);
    } else {
      unassignedCount += 1;
    }
  });

  const drivers = getDriverUsers()
    .slice()
    .sort((left, right) => String(left.name || "").localeCompare(String(right.name || ""), "en-ZA"));

  return {
    totalCount: getActiveOrders().length,
    unassignedCount,
    options: drivers.map((driver) => ({
      id: driver.id,
      name: driver.active ? driver.name : `${driver.name} (inactive)`,
      count: counts.get(driver.id) || 0,
    })),
  };
}

function renderAssignmentFilterOptions(selectedDriverId = "") {
  const { totalCount, unassignedCount, options } = getAssignmentDriverFilterOptions();

  return [
    `<option value=""${selectedDriverId ? "" : " selected"}>All drivers (${totalCount})</option>`,
    `<option value="unassigned"${selectedDriverId === "unassigned" ? " selected" : ""}>Unassigned (${unassignedCount})</option>`,
    ...options.map(
      (driver) => `<option value="${driver.id}"${driver.id === selectedDriverId ? " selected" : ""}>${escapeHtml(driver.name)} (${driver.count})</option>`,
    ),
  ].join("");
}

function renderFactoryLocationOptions(selectedLocationId = "") {
  const factories = getFactoryLocations();
  if (!factories.length) {
    return '<option value="">Create a factory location first</option>';
  }

  return [
    `<option value=""${selectedLocationId ? "" : " selected"}>Select a factory</option>`,
    ...factories.map((location) => {
      const typeLabel = location.locationType === "both" ? " (Both)" : "";
      return `<option value="${location.id}"${location.id === selectedLocationId ? " selected" : ""}>${escapeHtml(location.name)}${typeLabel}</option>`;
    }),
  ].join("");
}

function renderStockItemOptions(selectedStockItemId = "") {
  if (!state.snapshot.stockItems.length) {
    return '<option value="">Create a stock item first</option>';
  }

  return state.snapshot.stockItems
    .map((item) => {
      const labelParts = [item.name];
      const quoteNumber = String(item.quoteNumber || "").trim();
      if (quoteNumber) {
        labelParts.push(`Quote ${quoteNumber}`);
      }
      if (item.sku) {
        labelParts.push(item.sku);
      }
      const label = labelParts.join(" | ");
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

function getUser(userId) {
  return state.snapshot.users.find((user) => user.id === userId) || null;
}

function getEditingUser() {
  return state.editingUserId ? getUser(state.editingUserId) : null;
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

function getOrder(orderId) {
  return state.snapshot.orders.find((order) => order.id === orderId) || null;
}

function getSupplier(supplierId) {
  return state.snapshot.suppliers.find((supplier) => supplier.id === supplierId) || null;
}

function getEditingSupplier() {
  return state.editingSupplierId ? getSupplier(state.editingSupplierId) : null;
}

function getEditingLocation() {
  return state.editingLocationId ? getLocation(state.editingLocationId) : null;
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

function getFilteredAssignmentOrders() {
  const filterValue = String(state.assignmentDriverFilter || "").trim();
  const activeOrders = getActiveOrders();

  if (!filterValue) {
    return activeOrders.sort(orderAssignmentSort);
  }

  if (filterValue === "unassigned") {
    return activeOrders.filter((order) => !order.driverUserId).sort(orderAssignmentSort);
  }

  return activeOrders.filter((order) => order.driverUserId === filterValue).sort(orderAssignmentSort);
}

function getAssignmentFilterSummary(filteredCount) {
  const filterValue = String(state.assignmentDriverFilter || "").trim();

  if (!filterValue) {
    return `Showing ${filteredCount} active entr${filteredCount === 1 ? "y" : "ies"}. Unassigned work still appears first.`;
  }

  if (filterValue === "unassigned") {
    return filteredCount
      ? `Showing ${filteredCount} unassigned active entr${filteredCount === 1 ? "y" : "ies"}.`
      : "No unassigned active entries right now.";
  }

  const driver = getUser(filterValue);
  const driverName = driver?.name || "that driver";
  return filteredCount
    ? `Showing ${filteredCount} active entr${filteredCount === 1 ? "y" : "ies"} for ${driverName}.`
    : `No active entries are currently assigned to ${driverName}.`;
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

function getRouteOrigin(driverUserId) {
  const currentUser = state.snapshot.user;
  if (
    currentUser
    && currentUser.role === "driver"
    && currentUser.id === driverUserId
    && state.driverRouteOrigin
  ) {
    return state.driverRouteOrigin;
  }

  return HUB;
}

function getDriverRouteHeading(driverUserId) {
  return getRouteOrigin(driverUserId).source === "driver"
    ? "Your location to optimized stop sequence"
    : "Dispatch hub to optimized stop sequence";
}

function getDriverRouteAriaLabel(driverUserId) {
  return getRouteOrigin(driverUserId).source === "driver"
    ? "Map of your current location and optimized stop sequence"
    : "Map of the dispatch hub and optimized stop sequence";
}

function ensureDriverRouteOrigin() {
  const currentUser = state.snapshot.user;
  if (
    !currentUser
    || currentUser.role !== "driver"
    || state.driverRouteOrigin
    || state.driverRouteOriginLoading
    || state.driverRouteOriginAttempted
  ) {
    return;
  }

  if (!navigator.geolocation) {
    state.driverRouteOriginAttempted = true;
    state.driverRouteOriginError = "Driver location is unavailable in this browser, so Johannesburg dispatch is still being used as the route start.";
    return;
  }

  state.driverRouteOriginLoading = true;
  state.driverRouteOriginUserId = currentUser.id;

  navigator.geolocation.getCurrentPosition(
    (position) => {
      const activeUser = state.snapshot.user;
      if (!activeUser || activeUser.role !== "driver" || activeUser.id !== currentUser.id) {
        return;
      }

      state.driverRouteOrigin = {
        label: `${activeUser.name || "Driver"} location`,
        lat: position.coords.latitude,
        lng: position.coords.longitude,
        source: "driver",
      };
      state.driverRouteOriginLoading = false;
      state.driverRouteOriginAttempted = true;
      state.driverRouteOriginError = "";
      state.driverRouteOriginUserId = activeUser.id;
      render();
    },
    (error) => {
      const activeUser = state.snapshot.user;
      if (!activeUser || activeUser.role !== "driver" || activeUser.id !== currentUser.id) {
        return;
      }

      state.driverRouteOrigin = null;
      state.driverRouteOriginLoading = false;
      state.driverRouteOriginAttempted = true;
      state.driverRouteOriginError = getDriverRouteOriginError(error);
      state.driverRouteOriginUserId = activeUser.id;
      render();
    },
    {
      enableHighAccuracy: true,
      timeout: 10000,
      maximumAge: 300000,
    },
  );
}

function getDriverRouteOriginError(error) {
  if (error && typeof error === "object" && "code" in error) {
    if (error.code === 1) {
      return "Location access is off, so Johannesburg dispatch is still being used as the route start.";
    }

    if (error.code === 2) {
      return "Your location could not be read, so Johannesburg dispatch is still being used as the route start.";
    }

    if (error.code === 3) {
      return "Location lookup timed out, so Johannesburg dispatch is still being used as the route start.";
    }
  }

  return "Driver location is unavailable right now, so Johannesburg dispatch is still being used as the route start.";
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
      const key = String(order.quoteNumber || order.inhouseOrderNumber || "").toLowerCase();
      counts[key] = (counts[key] || 0) + 1;
    });

  return Object.values(counts).filter((count) => count > 1).length;
}

function getRoutePlan(driverUserId) {
  const activeOrders = getOrdersForDriver(driverUserId)
    .filter((order) => order.status === "active")
    .sort(orderRouteSort);
  const routeOrigin = getRouteOrigin(driverUserId);

  const grouped = new Map();

  activeOrders.forEach((order) => {
    const location = getLocation(order.locationId);
    if (!location) {
      return;
    }

    const coordinates = getCoordinates(location);

    if (!grouped.has(location.id)) {
      grouped.set(location.id, {
        id: location.id,
        location,
        orders: [],
        lat: coordinates?.lat ?? null,
        lng: coordinates?.lng ?? null,
        hasCoordinates: Boolean(coordinates),
        priorityRank: getOrderPriorityRank(order),
        isPriority: isPriorityOrder(order),
      });
    }

    const stop = grouped.get(location.id);
    stop.orders.push(order);
    stop.priorityRank = Math.min(stop.priorityRank, getOrderPriorityRank(order));
    stop.isPriority = stop.isPriority || isPriorityOrder(order);
  });

  const stops = Array.from(grouped.values()).map((stop) => ({
    ...stop,
    orders: [...stop.orders].sort(orderRouteSort),
  }));
  const priorityRouteableStops = stops.filter((stop) => stop.hasCoordinates && stop.isPriority);
  const standardRouteableStops = stops.filter((stop) => stop.hasCoordinates && !stop.isPriority);
  const priorityOrderedStops = optimizeRoute(priorityRouteableStops, routeOrigin);
  const standardRouteStart = priorityOrderedStops.length
    ? priorityOrderedStops[priorityOrderedStops.length - 1]
    : routeOrigin;
  const standardOrderedStops = optimizeRoute(standardRouteableStops, standardRouteStart);
  const routeableStops = priorityOrderedStops.concat(standardOrderedStops);
  const unroutedStops = stops
    .filter((stop) => !stop.hasCoordinates)
    .sort(stopDisplaySort);

  const enrichedStops = [];
  let currentPoint = routeOrigin;
  routeableStops.forEach((stop) => {
    enrichedStops.push({
      ...stop,
      legKm: haversineKm(currentPoint, stop),
    });
    currentPoint = stop;
  });
  unroutedStops.forEach((stop) => {
    enrichedStops.push({
      ...stop,
      legKm: null,
    });
  });

  return {
    origin: routeOrigin,
    totalOrders: activeOrders.length,
    totalKm: totalRouteDistance(routeableStops, routeOrigin),
    priorityStopCount: stops.filter((stop) => stop.isPriority).length,
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
  const plan = getRoutePlan(driverUserId);
  const routeableStops = plan.stops.filter((stop) => stop.hasCoordinates);
  if (drawDriverRouteMap(plan, routeableStops)) {
    return;
  }

  drawDriverRouteFallback(plan, routeableStops);
}

function drawDriverRouteMap(plan, routeableStops) {
  const Leaflet = window.L;
  const mapEl = document.getElementById("route-map");
  const canvas = document.getElementById("route-canvas");
  const routeOrigin = plan.origin || HUB;

  if (!Leaflet || !(mapEl instanceof HTMLElement)) {
    return false;
  }

  mapEl.classList.remove("hidden");
  if (canvas instanceof HTMLCanvasElement) {
    canvas.classList.add("hidden");
  }

  if (routeMap && routeMapContainer !== mapEl) {
    destroyDriverRouteMap();
  }

  if (!routeMap) {
    routeMap = Leaflet.map(mapEl, {
      zoomControl: true,
      scrollWheelZoom: false,
    });
    routeMapContainer = mapEl;
    Leaflet.tileLayer("https://tile.openstreetmap.org/{z}/{x}/{y}.png", {
      maxZoom: 19,
      attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',
    }).addTo(routeMap);
    routeMapLayers = Leaflet.layerGroup().addTo(routeMap);
  } else if (!routeMapLayers) {
    routeMapLayers = Leaflet.layerGroup().addTo(routeMap);
    routeMapContainer = mapEl;
  }

  routeMapLayers.clearLayers();

  const bounds = [[routeOrigin.lat, routeOrigin.lng]];
  const routePoints = [[routeOrigin.lat, routeOrigin.lng]];

  if (routeableStops.length) {
    routeableStops.forEach((stop) => {
      routePoints.push([stop.lat, stop.lng]);
      bounds.push([stop.lat, stop.lng]);
    });

    Leaflet.polyline(routePoints, {
      color: "#155e52",
      weight: 4,
      opacity: 0.82,
    }).addTo(routeMapLayers);
  }

  Leaflet.marker([routeOrigin.lat, routeOrigin.lng], {
    icon: buildDriverRouteMarkerIcon("H", true),
  })
    .addTo(routeMapLayers)
    .bindPopup(buildDriverRouteOriginPopup(plan));

  routeableStops.forEach((stop, index) => {
    Leaflet.marker([stop.lat, stop.lng], {
      icon: buildDriverRouteMarkerIcon(String(index + 1), false),
    })
      .addTo(routeMapLayers)
      .bindPopup(buildDriverRouteStopPopup(stop, index));
  });

  if (bounds.length > 1) {
    routeMap.fitBounds(bounds, { padding: [36, 36] });
  } else {
    routeMap.setView([routeOrigin.lat, routeOrigin.lng], routeOrigin.source === "driver" ? 11 : 9);
  }

  setDriverRouteStatus(getDriverRouteStatus(plan, routeableStops));
  window.requestAnimationFrame(() => {
    routeMap?.invalidateSize();
  });

  return true;
}

function drawDriverRouteFallback(plan, routeableStops) {
  const canvas = document.getElementById("route-canvas");
  const mapEl = document.getElementById("route-map");
  const routeOrigin = plan.origin || HUB;
  const statusParts = [];
  if (!window.L) {
    statusParts.push("Live map unavailable right now, so the route sketch fallback is shown instead.");
  }
  statusParts.push(getDriverRouteStatus(plan, routeableStops));
  const fallbackStatus = statusParts.filter(Boolean).join(" ");

  if (mapEl instanceof HTMLElement) {
    mapEl.classList.add("hidden");
  }
  if (!(canvas instanceof HTMLCanvasElement)) {
    setDriverRouteStatus(fallbackStatus);
    return;
  }

  canvas.classList.remove("hidden");
  const context = canvas.getContext("2d");
  if (!context) {
    setDriverRouteStatus(fallbackStatus);
    return;
  }

  const points = [
    { label: routeOrigin.label, lat: routeOrigin.lat, lng: routeOrigin.lng, isHub: true },
    ...routeableStops.map((stop, index) => ({
      label: `${index + 1}. ${stop.location.name}`,
      lat: stop.lat,
      lng: stop.lng,
      isHub: false,
    })),
  ];

  context.clearRect(0, 0, canvas.width, canvas.height);
  context.fillStyle = "#f6efe3";
  context.fillRect(0, 0, canvas.width, canvas.height);

  if (points.length === 1) {
    context.fillStyle = "#5b665e";
    context.font = "16px Trebuchet MS";
    context.fillText("No mapped route to draw yet.", 32, 40);
    setDriverRouteStatus(fallbackStatus);
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

  if (routeableStops.length < plan.stops.length) {
    context.fillStyle = "#5b665e";
    context.font = "14px Trebuchet MS";
    context.fillText("Stops without coordinates are excluded from the map.", 32, canvas.height - 24);
  }

  setDriverRouteStatus(fallbackStatus);
}

function destroyDriverRouteMap() {
  if (routeMap) {
    routeMap.remove();
  }
  routeMap = null;
  routeMapContainer = null;
  routeMapLayers = null;
}

function setDriverRouteStatus(message) {
  const statusEl = document.getElementById("route-map-status");
  if (!(statusEl instanceof HTMLElement)) {
    return;
  }

  statusEl.textContent = message || "";
  statusEl.classList.toggle("hidden", !message);
}

function getDriverRouteStatus(plan, routeableStops) {
  const originMessage = getDriverRouteOriginStatus(plan.origin);
  const priorityMessage = plan.priorityStopCount
    ? `${plan.priorityStopCount} priority stop${plan.priorityStopCount === 1 ? " is" : "s are"} highlighted first when coordinates are available.`
    : "";
  if (!plan.stops.length) {
    return [originMessage, priorityMessage, "No active entries are assigned to you right now."].filter(Boolean).join(" ");
  }

  const missingCoordinatesCount = plan.stops.length - routeableStops.length;
  if (!routeableStops.length) {
    return [originMessage, priorityMessage, "These stops still need coordinates before they can appear on the live map."].filter(Boolean).join(" ");
  }

  if (missingCoordinatesCount) {
    return [
      originMessage,
      priorityMessage,
      `${missingCoordinatesCount} stop${missingCoordinatesCount === 1 ? "" : "s"} without coordinates ${missingCoordinatesCount === 1 ? "is" : "are"} excluded from the mapped route.`,
    ].filter(Boolean).join(" ");
  }

  return [
    originMessage,
    priorityMessage,
    `Numbered markers follow the optimized stop order from the ${plan.origin?.source === "driver" ? "driver location" : "dispatch hub"}.`,
  ].filter(Boolean).join(" ");
}

function buildDriverRouteMarkerIcon(label, isHub) {
  const Leaflet = window.L;
  return Leaflet.divIcon({
    className: "route-map-marker-shell",
    html: `<span class="route-map-marker ${isHub ? "route-map-marker-hub" : "route-map-marker-stop"}">${escapeHtml(label)}</span>`,
    iconSize: [40, 40],
    iconAnchor: [20, 20],
    popupAnchor: [0, -18],
  });
}

function getDriverRouteOriginStatus(routeOrigin) {
  if (routeOrigin?.source === "driver") {
    return "Route starts from the driver's current location.";
  }

  return state.driverRouteOriginError || "";
}

function buildDriverRouteOriginPopup(plan) {
  const routeOrigin = plan.origin || HUB;
  const stopLabel = `${plan.stops.length} stop${plan.stops.length === 1 ? "" : "s"}`;
  const totalKmLabel = plan.totalKm ? `${plan.totalKm.toFixed(1)} km planned` : "Route starts here";
  return `
    <div class="route-map-popup">
      <p class="eyebrow">${escapeHtml(routeOrigin.source === "driver" ? "Driver location" : "Dispatch hub")}</p>
      <h4>${escapeHtml(routeOrigin.label)}</h4>
      <p>${escapeHtml(stopLabel)}</p>
      <p class="muted">${escapeHtml(totalKmLabel)}</p>
    </div>
  `;
}

function buildDriverRouteStopPopup(stop, index) {
  const entryLabel = `${stop.orders.length} entr${stop.orders.length === 1 ? "y" : "ies"}`;
  const references = stop.orders
    .map((order) => order.reference)
    .filter(Boolean)
    .slice(0, 3)
    .join(", ");
  const extraReferences = stop.orders.length > 3 ? ` +${stop.orders.length - 3} more` : "";
  const legLabel = stop.legKm != null ? `${stop.legKm.toFixed(1)} km from the previous stop` : "Coordinates pending";

  return `
    <div class="route-map-popup">
      <p class="eyebrow">Stop ${escapeHtml(String(index + 1))}</p>
      <h4>${escapeHtml(stop.location.name)}</h4>
      <p>${escapeHtml(stop.location.address || "Address not set")}</p>
      <p>${escapeHtml(entryLabel)}</p>
      ${references ? `<p class="muted">Refs: ${escapeHtml(references)}${escapeHtml(extraReferences)}</p>` : ""}
      <p class="muted">${escapeHtml(legLabel)}</p>
    </div>
  `;
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

function orderRouteSort(left, right) {
  const priorityCompare = getOrderPriorityRank(left) - getOrderPriorityRank(right);
  if (priorityCompare) {
    return priorityCompare;
  }

  return orderDisplaySort(left, right);
}

function orderAssignmentSort(left, right) {
  const leftAssigned = Boolean(left.driverUserId);
  const rightAssigned = Boolean(right.driverUserId);
  if (leftAssigned !== rightAssigned) {
    return leftAssigned ? 1 : -1;
  }

  return orderDisplaySort(left, right);
}

function stopDisplaySort(left, right) {
  const priorityCompare = (left.priorityRank ?? ORDER_PRIORITY_LEVELS[DEFAULT_ORDER_PRIORITY])
    - (right.priorityRank ?? ORDER_PRIORITY_LEVELS[DEFAULT_ORDER_PRIORITY]);
  if (priorityCompare) {
    return priorityCompare;
  }

  const leftName = String(left?.location?.name || "").toLowerCase();
  const rightName = String(right?.location?.name || "").toLowerCase();
  return leftName.localeCompare(rightName);
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
    order.quoteNumber || order.inhouseOrderNumber,
    order.salesOrderNumber || order.factoryOrderNumber,
    order.invoiceNumber || "",
    order.poNumber || "",
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
  const driverFlag = getOrderFlagNoticeText(order);
  const moveToFactory = getMoveToFactoryText(order);
  const rolloverNotice = getRolloverNoticeText(order);

  if (driverFlag) {
    lines.push(driverFlag);
  }

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

function getOrderFlagLabel(order) {
  return getOrderFlagTypeLabel(order?.driverFlagType);
}

function getOrderFlagTypeLabel(flagType) {
  return ORDER_FLAG_LABELS[String(flagType || "").trim()] || "";
}

function getOrderFlagNoticeText(order) {
  const label = getOrderFlagLabel(order);
  if (!label) {
    return "";
  }

  const note = String(order?.driverFlagNote || "").trim();
  const flaggedAt = formatDateTime(order?.driverFlaggedAt);
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
  const destinationField = form.querySelector("[data-move-to-factory-destination]");
  const destinationSelect = form.querySelector('[name="factoryDestinationLocationId"]');
  const noteEl = form.querySelector("[data-move-to-factory-note]");

  if (
    !(entryTypeField instanceof HTMLSelectElement)
    || !(moveToFactoryField instanceof HTMLInputElement)
    || !(destinationField instanceof HTMLElement)
    || !(destinationSelect instanceof HTMLSelectElement)
    || !(noteEl instanceof HTMLElement)
  ) {
    return;
  }

  const isCollection = entryTypeField.value === "collection";
  const hasFactoryOptions = Array.from(destinationSelect.options).some((option) => option.value);
  moveToFactoryField.disabled = !isCollection;
  if (!isCollection) {
    moveToFactoryField.checked = false;
    destinationSelect.value = "";
  }

  const showDestination = isCollection && moveToFactoryField.checked;
  destinationField.classList.toggle("hidden", !showDestination);
  destinationSelect.required = showDestination;

  if (!showDestination) {
    noteEl.textContent = isCollection
      ? "Check this when the collected stock must be moved to a factory."
      : "Switch the entry type to Collection to enable factory transfer.";
    return;
  }

  if (!isCollection) {
    noteEl.textContent = "Switch the entry type to Collection to enable factory transfer.";
    return;
  }

  noteEl.textContent = hasFactoryOptions
    ? "Select which factory the collected stock should go to."
    : "No factory destinations are available yet. Add a factory or Both location first.";
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

function parseOptionalNumber(value) {
  const text = String(value ?? "").trim();
  if (!text) {
    return null;
  }

  const parsed = Number(text);
  return Number.isFinite(parsed) ? parsed : Number.NaN;
}

function getCoordinates(value) {
  const lat = Number(value?.lat);
  const lng = Number(value?.lng);

  if (!Number.isFinite(lat) || !Number.isFinite(lng)) {
    return null;
  }

  return { lat, lng };
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
