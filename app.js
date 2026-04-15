const SESSION_KEY = "route-ledger-session-token-v2";
const SNAPSHOT_REFRESH_LOG_KEY = "route-ledger-last-refresh-v1";
const THEME_KEY = "route-ledger-theme-v1";
const FLASH_TIMEOUT_MS = 3200;
const TIME_ZONE = "Africa/Johannesburg";
const API_ROOT = "/api";
const APP_NAME = "Logictics Centre";
const STOCK_QR_TYPE = "route-ledger-stock";
const STOCK_QR_VERSION = 1;
const STOCK_SCANNER_FORMATS = ["qr_code", "code_128"];
const STOCK_LABEL_SKU_MAX_LENGTH = 10;
const STOCK_LABEL_CODE_LENGTH = 8;
const STOCK_LABEL_ALPHABET = "0123456789ABCDEFGHJKMNPQRSTVWXYZ";
const STOCK_RECENT_ACTIVITY_HOURS = 24;
const UUID_PATTERN = /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i;
const ORDER_FLAG_LABELS = Object.freeze({
  not_collected: "Not collected",
  not_ready: "Not yet ready",
});
const ORDER_COMPLETION_LABELS = Object.freeze({
  office: "Dropped at office",
  factory: "Dropped at factory",
});
const DEFAULT_ORDER_PRIORITY = "medium";
const PRIORITY_STOP_VALUE = "high";
const ORDER_PRIORITY_LEVELS = Object.freeze({
  high: 0,
  medium: 1,
  low: 2,
});
const AUTO_REFRESH_INTERVAL_MS = 60000;
const PAGE_SIZES = {
  globalEntries: 12,
  assignments: 12,
  driverLists: 3,
  completedEntries: 10,
};
const GUIDE_PAGE_SUMMARIES = Object.freeze({
  dashboard: "Start here for the live summary of the workspace so you can see what needs attention before moving into detailed pages.",
  entries: "Use this page to create work, search the grouped global register, and handle CSV or email actions for the live list.",
  assignments: "Use this page to move active work onto drivers, rebalance it between drivers, or return it to the queue without losing context.",
  stock: "Use this page to review stock activity, QR tools, movement history, and artwork requests according to your permissions.",
  network: "Use this page to maintain the supplier, pickup, and factory location records the rest of the app depends on.",
  users: "Use this page to manage team accounts, roles, passwords, and driver phone numbers in one place.",
  drivers: "Use this page to review each driver's live stop sequence, grouped work, and route context.",
  route: "Use this page to work through your assigned stops, navigate to locations, and update entry progress while you are out on route.",
  completed: "Use this page to review work that has already been finished so it stays separate from the live route.",
  guide: "Use this page when you need a role-based walkthrough, onboarding help, or a reminder of what each page is for.",
});
const ROLE_GUIDES = Object.freeze({
  admin: {
    title: "How to use the admin workspace",
    subtitle: "Admins oversee the whole live system, from creating work and dispatching routes to maintaining master data and correcting mistakes.",
    overview: "Use the admin workspace when you need full control of the operation. Admins can create and assign work, maintain suppliers and pickup locations, manage users, correct stock activity, clear rollover markers, and use higher-impact actions such as duplicate override or delete when something in the live list needs careful intervention.",
    roleFocus: [
      { label: "When to use this role", text: "Choose the admin workspace for cross-team tasks, data corrections, dispatch decisions, and setup work that affects the whole system." },
      { label: "What you can control", text: "Admins can manage users, locations, assignments, priorities, stock records, route visibility, CSV and email actions, and system cleanup tasks." },
      { label: "What needs extra care", text: "Override and delete actions change shared live data for everyone, so they should only be used after checking the current entry, stop, and stock history." },
    ],
    startingChecks: [
      "Start on Dashboard to review open work, unassigned jobs, active drivers, and completed volume before making dispatch decisions.",
      "Check the live date and last refresh message in the header so you know the screen is working from the latest snapshot.",
      "If you plan to email the CSV or send artwork requests, confirm mail delivery is available before relying on that step.",
    ],
    pageNotes: {
      dashboard: "Use Dashboard for your first read of the day. It tells you how much live work is open, what is still unassigned, how many drivers are active, and how much has already been completed.",
      entries: "Use Global List to create new work, search the live register, review grouped stops, and send or export the shared CSV that office staff use as a working list.",
      assignments: "Use Assignments when dispatch is ready to move jobs onto drivers, rebalance active work between drivers, or return a job to the unassigned queue.",
      stock: "Use Stock to manage stock items, correct movement history, work with QR labels or scanning, and send artwork requests when a production handoff is needed.",
      network: "Use Network to keep suppliers, pickup locations, addresses, and coordinates accurate so the route and assignment pages work from clean location data.",
      users: "Use Users to create, edit, disable, or delete admin, sales, driver, and logistics accounts, including driver phone numbers and password resets.",
      drivers: "Use Driver Lists to understand what each driver is carrying, which stops are priority, and where the most recently shared driver position was recorded.",
    },
    dailyFlow: [
      "Start on Dashboard so you can spot pressure points early, especially unassigned work, driver coverage, and any unusual completed volume.",
      "Move to Global List to create new entries, review the live grouped register, and confirm that high-priority or duplicate-sensitive jobs have been captured correctly.",
      "Open Assignments once dispatch is ready and move queued work onto the right drivers without overloading a single route.",
      "Check Driver Lists after dispatch to confirm stop order, priority jobs, and the latest driver location information.",
      "Use Stock and Network during the day whenever source data, movement history, QR tooling, or location records need correction.",
    ],
    keyTasks: [
      { label: "Create entries", text: "When an entry is saved, matching stock items are also created from the stock description, so this step affects both the route list and the stock ledger." },
      { label: "Dispatch cleanly", text: "Keep jobs Unassigned until the route plan is ready, then allocate them from Assignments so driver lists stay deliberate instead of being edited repeatedly." },
      { label: "Use override carefully", text: "Admin override is there for deliberate duplicate quotes or same-day return situations, not as a shortcut around normal duplicate protection." },
      { label: "Protect history", text: "Deleting a stock item also removes its movement and artwork history, so only delete when you are certain the record should not exist." },
    ],
    tips: [
      "Search before creating a new entry, especially when a quote, customer, or pickup stop may already be active in the live list.",
      "Add coordinates to locations whenever possible so route planning and map ordering can work from real geographic points.",
      "Clear rollover markers only after the carry-over report has been checked, otherwise you may remove the signal before the team has acted on it.",
    ],
  },
  sales: {
    title: "How to use the sales workspace",
    subtitle: "Sales is built for creating work quickly, dispatching it cleanly, and keeping the live customer-facing list accurate.",
    overview: "Use the sales workspace for day-to-day entry creation and dispatch support. Sales can add work, edit active entries, assign or rebalance active entries, review live driver queues, and share the global CSV, while admin-only controls such as delete, override, and stock editing stay protected.",
    roleFocus: [
      { label: "When to use this role", text: "Choose sales for normal office operations where the goal is to create, review, assign, and share live work without changing protected system data." },
      { label: "What you can control", text: "Sales users can create entries, edit active entries, review the live list, assign work to drivers, view stock history, and send or test CSV email actions." },
      { label: "What stays restricted", text: "Sales cannot bypass admin-only duplicate protection, delete entries, edit stock records, or manage users and network data." },
    ],
    startingChecks: [
      "Check Dashboard first to understand open work, unassigned volume, and how much activity has already been loaded onto drivers.",
      "Search Global List before creating a new job so you do not accidentally duplicate a quote or stop already in progress.",
      "If you need to share the list externally, confirm the email controls or export action you plan to use before the work gets busy.",
    ],
    pageNotes: {
      dashboard: "Use Dashboard to get a quick office view of open work, your current entry load, the number of unassigned jobs, and how much is already moving through drivers.",
      entries: "Use Global List to create new entries, search existing work, check grouped stops, and send or export the live CSV that the wider team uses.",
      assignments: "Use Assignments to work through the unassigned queue and rebalance active jobs when the route plan changes during the day.",
      stock: "Use Stock as a read-only reference when you need visibility into recent arrivals, current on-hand quantities, or movement history linked to active work.",
      drivers: "Use Driver Lists to confirm what each driver currently has, how the route is grouped, and whether priority work is showing where you expect it.",
    },
    dailyFlow: [
      "Create new work from Global List as soon as it is confirmed, and leave the driver Unassigned if dispatch will decide later.",
      "Use Assignments to move through the unassigned queue in batches and keep route changes tidy instead of editing one stop at a time in different places.",
      "Open Driver Lists to confirm what each driver is carrying and whether priority jobs are showing in the right sequence.",
      "Send or export the CSV when the live list is ready to share with the rest of the team or outside stakeholders.",
    ],
    keyTasks: [
      { label: "Create work", text: "Enter the quote first, then complete the stock description, pickup location, and delivery details so the live list and stock records stay clear." },
      { label: "Assign work", text: "Use the assignment filter to focus on Unassigned items so you can dispatch quickly without losing track of work already placed on drivers." },
      { label: "Share the list", text: "Use Download CSV, Test Email, or Email CSV from Global List depending on whether you need a quick export, a delivery test, or the live file sent out." },
      { label: "Check stock", text: "Use Stock when you need extra context about recent arrivals or current on-hand quantities linked to a customer or reference." },
    ],
    tips: [
      "Sales cannot use admin override, delete entries, edit stock records, or change protected setup data such as users and locations.",
      "Duplicate checks and completed-stop protection still apply even when you assign a driver at entry creation time.",
      "If a customer calls about an existing job, search Global List first instead of creating a fresh entry for the same quote or location.",
    ],
  },
  logistics: {
    title: "How to use the logistics workspace",
    subtitle: "Logistics keeps the stock side of the operation accurate, traceable, and ready for the next handoff.",
    overview: "Use the logistics workspace when the main job is stock control rather than route dispatch. Logistics focuses on keeping item records correct, logging stock in and out, using QR tools for speed, and sending artwork requests without exposing wider admin or user-management controls.",
    roleFocus: [
      { label: "When to use this role", text: "Choose logistics when your work is centered on stock intake, stock release, QR handling, and artwork preparation rather than customer entry creation." },
      { label: "What you can control", text: "Logistics can add stock items, log movements, search history, use QR labels and scanning, and send artwork requests from the stock workspace." },
      { label: "What stays restricted", text: "Logistics cannot manage users, locations, or dispatch lists, and only admins can edit or delete an existing stock item once it has been created." },
    ],
    startingChecks: [
      "Start with the stock dashboard view to see movement volume, total on-hand position, and whether there are outstanding artwork requests to follow up.",
      "Search existing stock before creating a new item so references and descriptions stay grouped instead of drifting into duplicates.",
      "If you are using QR scanning, make sure the camera or upload method you plan to use is working before you begin logging movements.",
    ],
    pageNotes: {
      dashboard: "Use Dashboard for a quick view of stock totals, recent movement volume, and artwork request activity so you know where attention is needed first.",
      stock: "Use Stock for nearly everything in this role: creating items, logging stock in and out, searching history, opening QR labels, scanning QR codes, and sending artwork requests.",
    },
    dailyFlow: [
      "Review recent stock activity first so you can see what arrived, what shipped, and whether anything looks incomplete or unusual.",
      "Add new stock items only when they do not already exist, and make sure they are linked to the right references before movements start building up.",
      "Log stock in and stock out as the movement happens so the on-hand position stays trustworthy for the rest of the team.",
      "Use QR labels or scanning whenever possible to speed up item selection and reduce manual picking mistakes.",
      "Send artwork requests once the stock and quantity are confirmed so the next production step can start with the right details.",
    ],
    keyTasks: [
      { label: "Add stock", text: "Each new stock item should have a clear description, at least one useful reference, and a correct opening quantity so later movement history makes sense." },
      { label: "Log movements", text: "Capture supplier details for stock in and the driver or destination context for stock out so the ledger remains traceable when questions come back later." },
      { label: "Use QR tools", text: "Open QR from the stock register or use Scan QR inside the movement form when you want faster selection and fewer manual lookup errors." },
      { label: "Request artwork", text: "Send quantity and notes from the stock workspace as soon as the job is ready so production receives a clear, linked request." },
    ],
    tips: [
      "Logistics can add stock items and manage movements, but only admins can edit or delete an existing stock item after it has been created.",
      "If live camera scanning fails, switch to uploading a QR image or typing the printed QR value manually instead of delaying the movement log.",
      "If artwork or email buttons are disabled, mail delivery is not configured yet and the request cannot be sent from inside the app.",
    ],
  },
  driver: {
    title: "How to use the driver workspace",
    subtitle: "The driver workspace is designed to keep the day simple: see your route, work each stop, and keep the office updated as conditions change.",
    overview: "Use the driver workspace while you are out on the road. It shows only your active route and your completed work, so you can focus on the next stop, update the office with notes or flags, and transfer jobs when the route changes in real time.",
    roleFocus: [
      { label: "When to use this role", text: "Choose the driver workspace when you need a focused route view without office setup pages, stock controls, or admin actions getting in the way." },
      { label: "What you can control", text: "Drivers can navigate to stops, mark entries picked up, complete work, flag problems, add notes, and transfer active jobs to another driver when needed." },
      { label: "What stays simple", text: "Drivers do not create users, change setup data, or edit stock history. The goal is to keep the route screen clear and action-oriented while you are moving." },
    ],
    startingChecks: [
      "Open Route first and allow location access if you want the map to start from your real position instead of the dispatch hub fallback.",
      "Review the grouped stops before you leave so you understand where priority work sits in the route for the day.",
      "Check that the phone number and account details shown for you look correct, especially if the office may need to reach you during a transfer or issue.",
    ],
    pageNotes: {
      route: "Use Route as your main working page. It shows your stops in order, lets you open directions, and gives you the buttons you need to update each job as it moves through pickup and delivery.",
      completed: "Use Completed to review jobs you already finished without cluttering the live route, especially if the office asks you to confirm what was done earlier in the day.",
    },
    dailyFlow: [
      "Open Route at the start of the day and review the stop order before you begin moving so you understand the current plan.",
      "Use Navigate when you need directions, then work through the stops in order unless the office has asked you to change the route.",
      "Mark an item Picked up as soon as it is on the vehicle so the office can see that the job has moved from waiting to active load.",
      "Complete the item only when it has reached the right final handoff point, such as the client, office, or factory.",
      "Use Flag issue or Transfer as soon as something changes so the office sees the problem while there is still time to respond.",
    ],
    keyTasks: [
      { label: "Navigate", text: "Use the Navigate button to open the stop directly in Google Maps instead of manually retyping the address while you are on the move." },
      { label: "Progress work", text: "Mark entries Picked up before completing them so the handover history shows that the job was loaded and then finished in the correct order." },
      { label: "Flag problems", text: "Use Not collected or Not yet ready with a clear driver note so the office can follow up with the customer or supplier immediately." },
      { label: "Transfer items", text: "Transfer an active job to another active driver if the route changes and the work can be completed faster by someone else." },
    ],
    tips: [
      "Completed items leave the Route page automatically and appear on Completed, so do not look for them in the active list afterward.",
      "If location access is blocked, route planning starts from the Johannesburg Dispatch Hub instead of your current position.",
      "Priority stops are highlighted first when the route can be mapped using stored coordinates, so watch for those cards before you start driving.",
    ],
  },
});
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
const pageShellEl = document.querySelector(".page-shell");

const state = {
  booting: true,
  busy: false,
  theme: loadThemePreference(),
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
  editingOrderId: "",
  orderEditReturnPage: "",
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
  globalListSearchQuery: "",
  globalListScheduledDate: "",
  networkSearchQuery: "",
  entryFormOpen: false,
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
  transferringOrderId: "",
  driverOpenStops: {},
  lastSnapshotRefreshAt: loadLastSnapshotRefreshAt(),
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
let adminDriverMap = null;
let adminDriverMapContainer = null;
let adminDriverMapLayers = null;
let snapshotAutoRefreshTimer = 0;

applyTheme(state.theme, { persist: false });

document.addEventListener("submit", handleSubmit);
document.addEventListener("click", handleClick);
document.addEventListener("change", handleChange);
document.addEventListener("input", handleInput);
document.addEventListener("keydown", handleKeyDown);
window.addEventListener("hashchange", handleHashChange);
window.addEventListener("beforeunload", () => {
  stopSnapshotAutoRefresh();
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

function loadThemePreference() {
  const rawValue = window.localStorage.getItem(THEME_KEY);
  return rawValue === "light" ? "light" : "dark";
}

function loadLastSnapshotRefreshAt() {
  const rawValue = window.localStorage.getItem(SNAPSHOT_REFRESH_LOG_KEY);
  const parsedValue = Number(rawValue);
  return Number.isFinite(parsedValue) && parsedValue > 0 ? parsedValue : 0;
}

function applyTheme(theme, { persist = true } = {}) {
  const normalizedTheme = theme === "light" ? "light" : "dark";
  state.theme = normalizedTheme;
  document.documentElement.dataset.theme = normalizedTheme;
  if (persist) {
    window.localStorage.setItem(THEME_KEY, normalizedTheme);
  }
  return normalizedTheme;
}

function saveLastSnapshotRefreshAt(timestamp) {
  const normalizedTimestamp = Number(timestamp);
  if (Number.isFinite(normalizedTimestamp) && normalizedTimestamp > 0) {
    state.lastSnapshotRefreshAt = normalizedTimestamp;
    window.localStorage.setItem(SNAPSHOT_REFRESH_LOG_KEY, String(normalizedTimestamp));
    return;
  }

  state.lastSnapshotRefreshAt = 0;
  window.localStorage.removeItem(SNAPSHOT_REFRESH_LOG_KEY);
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
  stopSnapshotAutoRefresh();
  state.booting = true;
  state.editingUserId = "";
  state.editingSupplierId = "";
  state.editingLocationId = "";
  state.editingOrderId = "";
  state.orderEditReturnPage = "";
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
  state.entryFormOpen = false;
  state.flaggingOrderId = "";
  state.transferringOrderId = "";
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

async function refreshSnapshot(options = {}) {
  const { silent = false } = options;
  if (!silent) {
    state.booting = true;
    render();
  }

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
    if (
      state.editingOrderId
      && !state.snapshot.orders.some((order) => order.id === state.editingOrderId)
    ) {
      state.editingOrderId = "";
      state.orderEditReturnPage = "";
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
    if (state.transferringOrderId) {
      const transferringOrder = state.snapshot.orders.find((order) => order.id === state.transferringOrderId && order.status === "active");
      if (!transferringOrder) {
        state.transferringOrderId = "";
      }
    }
    if (
      state.assignmentDriverFilter
      && state.assignmentDriverFilter !== "unassigned"
      && !state.snapshot.users.some((user) => user.id === state.assignmentDriverFilter && user.role === "driver")
    ) {
      state.assignmentDriverFilter = "";
    }
    if (
      state.globalListScheduledDate
      && !state.snapshot.orders.some((order) => getOrderScheduledForValue(order) === state.globalListScheduledDate)
    ) {
      state.globalListScheduledDate = "";
    }
    if (
      state.editingOrderId
      && !["admin", "sales"].includes(String(state.snapshot.user?.role || ""))
    ) {
      state.editingOrderId = "";
      state.orderEditReturnPage = "";
    }
    state.publicState = normalizePublicState(data);
    state.needsBootstrap = false;
    syncCurrentPage();
    saveLastSnapshotRefreshAt(Date.now());
    state.booting = false;
    state.busy = false;
    render();
    scheduleSnapshotAutoRefresh();

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

    if (silent) {
      console.error("Automatic refresh failed.", error);
      scheduleSnapshotAutoRefresh();
      return;
    }

    throw error;
  }
}

function stopSnapshotAutoRefresh() {
  if (snapshotAutoRefreshTimer) {
    window.clearTimeout(snapshotAutoRefreshTimer);
    snapshotAutoRefreshTimer = 0;
  }
}

function scheduleSnapshotAutoRefresh() {
  stopSnapshotAutoRefresh();
  if (!sessionToken) {
    return;
  }

  snapshotAutoRefreshTimer = window.setTimeout(() => {
    void runSnapshotAutoRefresh();
  }, AUTO_REFRESH_INTERVAL_MS);
}

function isEntryFormSessionActive() {
  return Boolean(state.entryFormOpen || state.editingOrderId);
}

async function runSnapshotAutoRefresh() {
  if (!sessionToken) {
    stopSnapshotAutoRefresh();
    return;
  }

  if (shouldDeferSnapshotAutoRefresh()) {
    scheduleSnapshotAutoRefresh();
    return;
  }

  await refreshSnapshot({ silent: true });
}

function shouldDeferSnapshotAutoRefresh() {
  if (state.busy || state.booting || state.stockScannerOpen || state.stockQrBusy || state.stockQrSharing) {
    return true;
  }

  if (isEntryFormSessionActive()) {
    return true;
  }

  if (
    state.flaggingOrderId
    || state.editingUserId
    || state.editingSupplierId
    || state.editingLocationId
    || state.editingStockItemId
    || state.editingStockMovementId
  ) {
    return true;
  }

  const activeElement = document.activeElement;
  if (
    activeElement instanceof HTMLElement
    && activeElement.matches("input, textarea, select")
    && !activeElement.hasAttribute("readonly")
    && !activeElement.hasAttribute("disabled")
  ) {
    return true;
  }

  return false;
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
  const message = candidate || "Something went wrong.";
  const lowerMessage = String(message).toLowerCase();

  if (
    lowerMessage.includes("unknown rpc function")
    || lowerMessage.includes("request failed with status 404")
  ) {
    return "The running server is missing this update. Restart node serve.js, refresh the app, then try again.";
  }

  if (
    lowerMessage.includes("function public.create_order")
    || lowerMessage.includes("function public.update_order")
    || lowerMessage.includes("function public.assign_order")
    || lowerMessage.includes("function public.record_driver_position")
    || lowerMessage.includes("column \"last_known_lat\"")
    || lowerMessage.includes("column \"last_known_lng\"")
    || lowerMessage.includes("column \"last_known_recorded_at\"")
  ) {
    return "The app code is ahead of the database. Apply the latest neon.sql update, then try again.";
  }

  return message;
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

  if (formId === "edit-order" && (currentUser.role === "admin" || currentUser.role === "sales")) {
    await updateOrder(formData, currentUser);
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

  if (action === "toggle-theme") {
    applyTheme(state.theme === "light" ? "dark" : "light");
    render();
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

  if (action === "toggle-entry-form" && (currentUser.role === "admin" || currentUser.role === "sales")) {
    const willOpen = !state.entryFormOpen;
    state.entryFormOpen = willOpen;
    if (!state.entryFormOpen) {
      state.editingOrderId = "";
      state.orderEditReturnPage = "";
    }
    render();
    if (willOpen) {
      focusOrderForm();
    } else {
      void refreshSnapshot({ silent: true });
    }
    return;
  }

  if (action === "edit-order" && (currentUser.role === "admin" || currentUser.role === "sales")) {
    openOrderEditor(String(button.dataset.orderId || "").trim());
    return;
  }

  if (action === "cancel-edit-order" && (currentUser.role === "admin" || currentUser.role === "sales")) {
    cancelOrderEdit();
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

  if (action === "email-rollover-test" && (currentUser.role === "admin" || currentUser.role === "sales")) {
    await sendRolloverTestEmail();
    return;
  }

  if (action === "clear-all-order-priorities" && currentUser.role === "admin") {
    await clearAllOrderPriorities();
    return;
  }

  if (action === "clear-order-rollovers" && currentUser.role === "admin") {
    await clearOrderRollovers();
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

  if (action === "toggle-stop-card") {
    const stopCardKey = String(button.dataset.stopCardKey || "").trim();
    if (!stopCardKey) {
      return;
    }

    state.driverOpenStops = {
      ...state.driverOpenStops,
      [stopCardKey]: !isStopCardOpen(stopCardKey),
    };
    render();
    return;
  }

  if (action === "clear-stock-search" && (currentUser.role === "admin" || currentUser.role === "logistics" || currentUser.role === "sales")) {
    state.stockSearchQuery = "";
    render();
    return;
  }

  if (action === "clear-global-list-search" && (currentUser.role === "admin" || currentUser.role === "sales")) {
    state.globalListSearchQuery = "";
    state.pagination.globalEntries = 1;
    render();
    return;
  }

  if (action === "clear-global-list-date-filter" && (currentUser.role === "admin" || currentUser.role === "sales")) {
    state.globalListScheduledDate = "";
    state.pagination.globalEntries = 1;
    render();
    return;
  }

  if (action === "set-global-list-date-filter" && (currentUser.role === "admin" || currentUser.role === "sales")) {
    state.globalListScheduledDate = String(button.dataset.scheduledFor || "").trim();
    state.pagination.globalEntries = 1;
    render();
    return;
  }

  if (action === "clear-network-search" && currentUser.role === "admin") {
    state.networkSearchQuery = "";
    render();
    return;
  }

  if (action === "toggle-stock-section" && (currentUser.role === "admin" || currentUser.role === "logistics" || currentUser.role === "sales")) {
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

  if (action === "toggle-stock-item-card" && (currentUser.role === "admin" || currentUser.role === "logistics" || currentUser.role === "sales")) {
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

  if (action === "save-location-assignment" && (currentUser.role === "admin" || currentUser.role === "sales")) {
    await saveLocationAssignment(button, currentUser);
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

    state.transferringOrderId = "";
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

  if (action === "toggle-transfer-order" && (currentUser.role === "admin" || currentUser.role === "driver")) {
    const orderId = String(button.dataset.orderId || "").trim();
    const order = getOrder(orderId);
    if (!order || order.status !== "active" || !getTransferDriverChoices(order).length) {
      return;
    }

    state.flaggingOrderId = "";
    state.transferringOrderId = state.transferringOrderId === orderId ? "" : orderId;
    render();
    return;
  }

  if (action === "cancel-transfer-order" && (currentUser.role === "admin" || currentUser.role === "driver")) {
    state.transferringOrderId = "";
    render();
    return;
  }

  if (action === "save-transfer-order" && (currentUser.role === "admin" || currentUser.role === "driver")) {
    await saveOrderTransfer(button, currentUser);
    return;
  }

  if (action === "pick-up-order" && (currentUser.role === "admin" || currentUser.role === "driver")) {
    await runMutation(
      "pick_up_order",
      {
        p_token: sessionToken,
        p_order_id: button.dataset.orderId,
      },
      "Entry marked as picked up.",
    );
    return;
  }

  if (action === "complete-order" && (currentUser.role === "admin" || currentUser.role === "driver")) {
    const completionType = String(button.dataset.completionType || "").trim() || "office";
    const order = getOrder(String(button.dataset.orderId || "").trim());
    const completionLabel = getOrderCompletionTypeLabel(completionType, order).toLowerCase() || "completed";
    await runMutation(
      "complete_order",
      {
        p_token: sessionToken,
        p_order_id: button.dataset.orderId,
        p_completion_type: completionType,
      },
      `Entry marked as ${completionLabel}.`,
    );
    return;
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

  if (target.matches("[data-global-list-date-filter]") && target instanceof HTMLSelectElement) {
    state.globalListScheduledDate = target.value;
    state.pagination.globalEntries = 1;
    render();
    return;
  }

  const orderForm = target.closest('form[data-form="add-order"], form[data-form="edit-order"]');
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

  if (target.matches("[data-global-list-search]") && target instanceof HTMLInputElement) {
    state.globalListSearchQuery = target.value;
    state.pagination.globalEntries = 1;
    render();
    return;
  }

  if (target.matches("[data-network-search]") && target instanceof HTMLInputElement) {
    state.networkSearchQuery = target.value;
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
    const data = await callRpc(functionName, parameters);
    const warning = String(data?.warning || "").trim();
    showFlash([successMessage, warning].filter(Boolean).join(" "), "success");
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
  const scheduledFor = String(formData.get("scheduledFor") || getDefaultScheduledDateValue()).trim();
  const priority = formData.get("priority") === "on" ? PRIORITY_STOP_VALUE : DEFAULT_ORDER_PRIORITY;
  const quoteNumber = String(formData.get("quoteNumber") || "").trim();
  const salesOrderNumber = String(formData.get("salesOrderNumber") || "").trim();
  const invoiceNumber = String(formData.get("invoiceNumber") || "").trim();
  const poNumber = String(formData.get("poNumber") || "").trim();
  const branding = String(formData.get("branding") || "").trim();
  const stockDescription = String(formData.get("stockDescription") || "").trim();
  const stockItemNames = getStockItemDisplayLines(stockDescription);
  const deliveryAddress = String(formData.get("deliveryAddress") || "").trim();
  const notice = String(formData.get("notice") || "").trim();
  const moveToFactory = formData.get("moveToFactory") === "on";
  const factoryDestinationLocationId = String(formData.get("factoryDestinationLocationId") || "").trim();
  const allowDuplicate = formData.get("allowDuplicate") === "on";

  if (!locationId || !entryType || !quoteNumber) {
    showFlash("Pickup location, entry type, and quote number are required.", "error");
    render();
    return;
  }

  if (!scheduledFor) {
    showFlash("Choose which date this entry should appear on the driver list.", "error");
    render();
    return;
  }

  if (!stockDescription) {
    showFlash("Stock description is required so drivers know what stock is on this entry.", "error");
    render();
    return;
  }

  if (entryType === "delivery" && !deliveryAddress) {
    showFlash("Delivery address is required for delivery entries.", "error");
    render();
    return;
  }

  if (moveToFactory && !factoryDestinationLocationId) {
    showFlash("Select which factory the collected stock should go to.", "error");
    render();
    return;
  }

  const ok = await runMutation(
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
      p_branding: branding,
      p_stock_description: stockDescription,
      p_stock_item_names: stockItemNames.length ? stockItemNames : null,
      p_delivery_address: entryType === "delivery" ? deliveryAddress : null,
      p_scheduled_for: scheduledFor,
      p_priority: priority,
      p_notice: notice,
      p_move_to_factory: moveToFactory,
      p_factory_destination_location_id: moveToFactory ? factoryDestinationLocationId : null,
      p_allow_duplicate: currentUser.role === "admin" ? allowDuplicate : false,
    },
    driverUserId ? "Entry added to the driver list." : "Entry created in the unassigned queue.",
  );

  if (ok) {
    state.editingOrderId = "";
    state.orderEditReturnPage = "";
    state.entryFormOpen = false;
    render();
  }
}

async function updateOrder(formData, currentUser) {
  const orderId = String(formData.get("orderId") || state.editingOrderId || "").trim();
  const driverUserId = String(formData.get("driverUserId") || "").trim();
  const locationId = String(formData.get("locationId") || "").trim();
  const entryType = String(formData.get("entryType") || "delivery").trim();
  const scheduledFor = String(formData.get("scheduledFor") || getDefaultScheduledDateValue()).trim();
  const priority = formData.get("priority") === "on" ? PRIORITY_STOP_VALUE : DEFAULT_ORDER_PRIORITY;
  const quoteNumber = String(formData.get("quoteNumber") || "").trim();
  const salesOrderNumber = String(formData.get("salesOrderNumber") || "").trim();
  const invoiceNumber = String(formData.get("invoiceNumber") || "").trim();
  const poNumber = String(formData.get("poNumber") || "").trim();
  const branding = String(formData.get("branding") || "").trim();
  const stockDescription = String(formData.get("stockDescription") || "").trim();
  const deliveryAddress = String(formData.get("deliveryAddress") || "").trim();
  const notice = String(formData.get("notice") || "").trim();
  const moveToFactory = formData.get("moveToFactory") === "on";
  const factoryDestinationLocationId = String(formData.get("factoryDestinationLocationId") || "").trim();
  const allowDuplicate = formData.get("allowDuplicate") === "on";

  if (!orderId) {
    showFlash("That entry could not be found for editing.", "error");
    render();
    return;
  }

  if (!locationId || !entryType || !quoteNumber) {
    showFlash("Pickup location, entry type, and quote number are required.", "error");
    render();
    return;
  }

  if (!scheduledFor) {
    showFlash("Choose which date this entry should appear on the driver list.", "error");
    render();
    return;
  }

  if (!stockDescription) {
    showFlash("Stock description is required so drivers know what stock is on this entry.", "error");
    render();
    return;
  }

  if (entryType === "delivery" && !deliveryAddress) {
    showFlash("Delivery address is required for delivery entries.", "error");
    render();
    return;
  }

  if (moveToFactory && !factoryDestinationLocationId) {
    showFlash("Select which factory the collected stock should go to.", "error");
    render();
    return;
  }

  const ok = await runMutation(
    "update_order",
    {
      p_token: sessionToken,
      p_order_id: orderId,
      p_driver_user_id: driverUserId || null,
      p_location_id: locationId,
      p_entry_type: entryType,
      p_quote_number: quoteNumber,
      p_sales_order_number: salesOrderNumber,
      p_invoice_number: invoiceNumber,
      p_po_number: poNumber,
      p_branding: branding,
      p_stock_description: stockDescription,
      p_delivery_address: entryType === "delivery" ? deliveryAddress : null,
      p_scheduled_for: scheduledFor,
      p_priority: priority,
      p_allow_duplicate: currentUser.role === "admin" ? allowDuplicate : false,
      p_notice: notice,
      p_move_to_factory: moveToFactory,
      p_factory_destination_location_id: moveToFactory ? factoryDestinationLocationId : null,
    },
    "Entry updated.",
  );

  if (!ok) {
    return;
  }

  const returnPage = state.orderEditReturnPage;
  state.editingOrderId = "";
  state.orderEditReturnPage = "";
  state.entryFormOpen = false;

  if (returnPage && returnPage !== "entries") {
    state.currentPage = returnPage;
    syncEntryFormVisibility(returnPage);
    window.history.replaceState(null, "", `${window.location.pathname}#${returnPage}`);
    render();
    return;
  }

  render();
}

function openOrderEditor(orderId) {
  const order = getOrder(orderId);
  const currentUser = state.snapshot.user;
  if (!order) {
    showFlash("That entry could not be found for editing.", "error");
    render();
    return;
  }

  if (!currentUser || !["admin", "sales"].includes(currentUser.role)) {
    showFlash("You do not have permission to edit this entry.", "error");
    render();
    return;
  }

  if (currentUser.role === "sales" && order.status !== "active") {
    showFlash("Sales can only edit active entries.", "error");
    render();
    return;
  }

  state.flaggingOrderId = "";
  state.transferringOrderId = "";
  state.editingOrderId = order.id;
  state.orderEditReturnPage = state.currentPage && state.currentPage !== "entries" ? state.currentPage : "";
  state.entryFormOpen = true;

  if (state.currentPage === "entries") {
    render();
    focusOrderForm();
    return;
  }

  setCurrentPage("entries");
  focusOrderForm();
}

function cancelOrderEdit() {
  const returnPage = state.orderEditReturnPage;
  state.editingOrderId = "";
  state.orderEditReturnPage = "";
  state.entryFormOpen = false;

  if (returnPage && returnPage !== "entries") {
    setCurrentPage(returnPage);
    void refreshSnapshot({ silent: true });
    return;
  }

  render();
  void refreshSnapshot({ silent: true });
}

async function saveLocationAssignment(button, currentUser) {
  const groupKey = String(button.dataset.groupKey || "").trim();
  if (!groupKey) {
    return;
  }

  const group = getFilteredAssignmentLocationGroups().find((entry) => entry.key === groupKey);
  if (!group || !group.orders.length) {
    return;
  }

  const container = button.closest(".stop-card");
  const driverField = container?.querySelector("[data-assignment-location-driver]");
  if (!(driverField instanceof HTMLSelectElement)) {
    return;
  }

  const allowDuplicateField = container?.querySelector("[data-assignment-location-allow-duplicate]");
  const rawDriverValue = String(driverField.value || "").trim();
  if (group.hasMixedDriverSelection && !rawDriverValue) {
    showFlash("Choose a driver or Unassigned for that pickup location first.", "error");
    render();
    return;
  }

  const driverUserId = rawDriverValue === "__unassigned__" ? "" : rawDriverValue;
  const selectedLabel = rawDriverValue === "__unassigned__"
    ? "Unassigned"
    : (driverField.selectedOptions[0]?.textContent?.trim() || "Unassigned");
  const ordersToUpdate = group.orders.filter((order) => String(order.driverUserId || "").trim() !== driverUserId);

  if (!ordersToUpdate.length) {
    showFlash(
      driverUserId
        ? `All visible entries at ${group.locationName || "that pickup location"} are already assigned to ${selectedLabel}.`
        : `All visible entries at ${group.locationName || "that pickup location"} are already unassigned.`,
      "success",
    );
    render();
    return;
  }

  state.busy = true;
  render();

  let updatedCount = 0;
  const warnings = [];

  try {
    for (const order of ordersToUpdate) {
      const data = await callRpc("assign_order", {
        p_token: sessionToken,
        p_order_id: order.id,
        p_driver_user_id: driverUserId || null,
        p_allow_duplicate:
          currentUser.role === "admin" && allowDuplicateField instanceof HTMLInputElement
            ? allowDuplicateField.checked
            : false,
      });
      updatedCount += 1;
      const warning = String(data?.warning || "").trim();
      if (warning) {
        warnings.push(warning);
      }
    }

    showFlash(
      [
        driverUserId
          ? `${updatedCount} entr${updatedCount === 1 ? "y" : "ies"} at ${group.locationName || "that pickup location"} assigned to ${selectedLabel}.`
          : `${updatedCount} entr${updatedCount === 1 ? "y" : "ies"} at ${group.locationName || "that pickup location"} moved to the unassigned queue.`,
        ...warnings,
      ].filter(Boolean).join(" "),
      "success",
    );
    await refreshSnapshot();
  } catch (error) {
    if (updatedCount) {
      showFlash(
        `${updatedCount} entr${updatedCount === 1 ? "y was" : "ies were"} updated at ${group.locationName || "that pickup location"} before the assignment stopped. ${normalizeError(error)}`,
        "error",
      );
      await refreshSnapshot();
      return;
    }

    state.busy = false;
    showFlash(normalizeError(error), "error");
    render();
  }
}

async function saveOrderTransfer(button, currentUser) {
  const orderId = String(button.dataset.orderId || "").trim();
  if (!orderId) {
    return;
  }

  const order = getOrder(orderId);
  if (!order || order.status !== "active") {
    return;
  }

  const driverField = document.querySelector(`[data-transfer-driver][data-order-id="${orderId}"]`);
  if (!(driverField instanceof HTMLSelectElement)) {
    return;
  }

  const driverUserId = String(driverField.value || "").trim();
  const selectedLabel = driverField.selectedOptions[0]?.textContent?.trim() || "the selected driver";
  if (!driverUserId) {
    showFlash("Select another driver first.", "error");
    render();
    return;
  }

  const ok = await runMutation(
    "assign_order",
    {
      p_token: sessionToken,
      p_order_id: orderId,
      p_driver_user_id: driverUserId,
      p_allow_duplicate: false,
    },
    currentUser.role === "driver"
      ? `Entry transferred to ${selectedLabel}.`
      : `Entry reassigned to ${selectedLabel}.`,
  );

  if (ok) {
    state.transferringOrderId = "";
    render();
  }
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

async function clearAllOrderPriorities() {
  const priorityOrders = getActivePriorityOrders();
  const count = priorityOrders.length;

  if (!count) {
    showFlash("There are no active priority entries to clear.", "error");
    render();
    return;
  }

  const confirmed = window.confirm(
    `Clear priority from ${count} active entr${count === 1 ? "y" : "ies"}?`,
  );
  if (!confirmed) {
    return;
  }

  await runMutation(
    "clear_all_order_priorities",
    { p_token: sessionToken },
    `Priority cleared for ${count} active entr${count === 1 ? "y" : "ies"}.`,
  );
}

async function clearOrderRollovers() {
  const rolloverOrders = getActiveCarryOverOrders();
  const count = rolloverOrders.length;

  if (!count) {
    showFlash("There are no active rollover entries to clear.", "error");
    render();
    return;
  }

  const confirmed = window.confirm(
    `Clear rollover markers from ${count} active entr${count === 1 ? "y" : "ies"}?`,
  );
  if (!confirmed) {
    return;
  }

  await runMutation(
    "clear_order_rollovers",
    { p_token: sessionToken },
    `Rollover cleared for ${count} active entr${count === 1 ? "y" : "ies"}.`,
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
  const editingItem = stockItemId ? getStockItemById(stockItemId) : null;
  const name = String(formData.get("name") || "").trim();
  const sku = String(formData.get("sku") || "").trim();
  const quoteNumber = String(formData.get("quoteNumber") || "").trim();
  const invoiceNumber = String(formData.get("invoiceNumber") || "").trim();
  const salesOrderNumber = String(formData.get("salesOrderNumber") || "").trim();
  const poNumber = String(formData.get("poNumber") || "").trim();
  const initialQuantityValue = String(formData.get("initialQuantity") || "").trim();
  const initialQuantity = Number.parseInt(initialQuantityValue, 10);
  const unitLabel = getStockUnitLabel(editingItem);
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

  if (![quoteNumber, salesOrderNumber, invoiceNumber, poNumber].some(Boolean)) {
    showFlash("Enter at least one quote, sales order, invoice, or PO number.", "error");
    render();
    return;
  }

  if (!stockItemId && (!Number.isInteger(initialQuantity) || initialQuantity <= 0)) {
    showFlash("Opening stock must be greater than zero.", "error");
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
        p_unit: unitLabel,
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
        p_unit: unitLabel,
        p_initial_quantity: initialQuantity,
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
  destroyAdminDriverLocationMap();
  renderHeader();
  renderPageNavigation();
  if (pageShellEl) {
    pageShellEl.classList.toggle("workspace-shell", Boolean(state.snapshot.user));
  }

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
  document
    .querySelectorAll('form[data-form="add-order"], form[data-form="edit-order"]')
    .forEach((form) => {
      if (form instanceof HTMLFormElement) {
        syncMoveToFactoryField(form);
      }
    });

  const stockMovementForm = document.querySelector('form[data-form="add-stock-movement"]');
  if (stockMovementForm instanceof HTMLFormElement) {
    syncStockMovementFields(stockMovementForm);
  }

  void syncStockScannerUi();

  const currentUser = state.snapshot.user;
  if (currentUser?.role === "admin" && state.currentPage === "drivers") {
    drawAdminDriverLocationMap();
  }
}

function focusOrderForm() {
  window.requestAnimationFrame(() => {
    const form = document.querySelector('form[data-form="edit-order"], form[data-form="add-order"]');
    if (!(form instanceof HTMLFormElement)) {
      return;
    }

    form.scrollIntoView({ behavior: "smooth", block: "start" });
    const firstField = form.querySelector('[name="locationId"]');
    if (firstField instanceof HTMLElement) {
      firstField.focus();
    }
  });
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

function isRecentStockMovement(value, hours = STOCK_RECENT_ACTIVITY_HOURS) {
  const timestamp = Date.parse(String(value || ""));
  if (!Number.isFinite(timestamp)) {
    return false;
  }

  return (Date.now() - timestamp) <= (hours * 60 * 60 * 1000);
}

function normalizeStockMatchValue(value) {
  return String(value || "").trim().toLowerCase();
}

function getStockActivityReferenceFields(record) {
  return {
    name: normalizeStockMatchValue(record?.name || record?.stockDescription),
    quoteNumber: normalizeStockMatchValue(record?.quoteNumber || record?.inhouseOrderNumber),
    salesOrderNumber: normalizeStockMatchValue(record?.salesOrderNumber || record?.factoryOrderNumber),
    invoiceNumber: normalizeStockMatchValue(record?.invoiceNumber),
    poNumber: normalizeStockMatchValue(record?.poNumber),
  };
}

function findMatchingStockItem(record, stockItemName = "") {
  const recordFields = {
    ...getStockActivityReferenceFields(record),
    name: normalizeStockMatchValue(stockItemName || record?.name || record?.stockDescription),
  };
  let bestMatch = null;
  let bestScore = 0;

  state.snapshot.stockItems.forEach((item) => {
    const itemFields = getStockActivityReferenceFields(item);
    const exactTupleMatch = (
      recordFields.name === itemFields.name
      && recordFields.quoteNumber === itemFields.quoteNumber
      && recordFields.salesOrderNumber === itemFields.salesOrderNumber
      && recordFields.invoiceNumber === itemFields.invoiceNumber
      && recordFields.poNumber === itemFields.poNumber
    );

    let score = exactTupleMatch ? 100 : 0;
    let referenceMatches = 0;

    ["quoteNumber", "salesOrderNumber", "invoiceNumber", "poNumber"].forEach((field) => {
      if (recordFields[field] && itemFields[field] && recordFields[field] === itemFields[field]) {
        score += 10;
        referenceMatches += 1;
      }
    });

    if (recordFields.name && itemFields.name && recordFields.name === itemFields.name) {
      score += 1;
    }

    if (!exactTupleMatch && !referenceMatches) {
      return;
    }

    if (score > bestScore) {
      bestScore = score;
      bestMatch = item;
    }
  });

  return bestMatch;
}

function getUniqueStockItemNames(value) {
  const names = getStockItemDisplayLines(value);
  const seenNames = new Set();
  const uniqueNames = [];

  names.forEach((name) => {
    const key = normalizeStockMatchValue(name);
    if (!key || seenNames.has(key)) {
      return;
    }

    seenNames.add(key);
    uniqueNames.push(name);
  });

  return uniqueNames;
}

function getRecentCreatedStockActivities() {
  return state.snapshot.stockItems
    .filter((item) => (
      isRecentStockMovement(item.createdAt)
      && String(item.createdSource || "manual").trim() !== "order"
    ))
    .map((item) => ({
      id: `stock-created-${item.id}`,
      stockItemId: item.id,
      itemName: item.name,
      sku: item.sku,
      quoteNumber: item.quoteNumber,
      invoiceNumber: item.invoiceNumber,
      salesOrderNumber: item.salesOrderNumber,
      poNumber: item.poNumber,
      unit: item.unit,
      movementType: "in",
      quantity: Number(item.onHandQuantity || 0),
      supplierName: "Stock item created",
      driverUserId: "",
      driverName: "",
      notes: item.notes || "",
      createdByUserId: "",
      createdByName: "",
      createdAt: item.createdAt,
      activityKind: "created",
    }));
}

function getRecentOfficeDropActivities() {
  return state.snapshot.orders
    .filter((order) => (
      order.status === "completed"
      && order.entryType === "collection"
      && order.completionType === "office"
      && isRecentStockMovement(order.completedAt)
    ))
    .flatMap((order) => {
      const stockItemNames = getUniqueStockItemNames(order.stockDescription);
      const fallbackNames = stockItemNames.length
        ? stockItemNames
        : [String(order.stockDescription || "").trim() || order.reference || "Unknown"];

      return fallbackNames.map((stockItemName, index) => {
        const matchedItem = findMatchingStockItem(order, stockItemName);

        return {
          id: `office-drop-${order.id}-${index}`,
          stockItemId: matchedItem?.id || `office-drop-${order.id}-${index}`,
          itemName: matchedItem?.name || stockItemName,
          sku: matchedItem?.sku || "",
          quoteNumber: matchedItem?.quoteNumber || order.quoteNumber || order.inhouseOrderNumber || "",
          invoiceNumber: matchedItem?.invoiceNumber || order.invoiceNumber || "",
          salesOrderNumber: matchedItem?.salesOrderNumber || order.salesOrderNumber || order.factoryOrderNumber || "",
          poNumber: matchedItem?.poNumber || order.poNumber || "",
          unit: matchedItem?.unit || "units",
          movementType: "in",
          quantity: 0,
          supplierName: order.locationName || "",
          driverUserId: order.driverUserId || "",
          driverName: order.driverName || order.completedByName || "",
          notes: order.notes || "",
          createdByUserId: order.completedByUserId || "",
          createdByName: order.completedByName || "",
          createdAt: order.completedAt,
          activityKind: "office_drop",
          locationName: order.locationName || "",
          locationAddress: order.locationAddress || "",
        };
      });
    });
}

function getRecentStockMovements(movementType, limit = 6) {
  const recentMovements = state.snapshot.stockMovements
    .filter((movement) => movement.movementType === movementType && isRecentStockMovement(movement.createdAt));

  if (movementType === "in") {
    recentMovements.push(...getRecentCreatedStockActivities());
    recentMovements.push(...getRecentOfficeDropActivities());
  }

  recentMovements.sort((left, right) => Date.parse(String(right.createdAt || "")) - Date.parse(String(left.createdAt || "")));

  const seenStockItemIds = new Set();
  return recentMovements.filter((movement) => {
    const dedupeKey = String(movement.stockItemId || movement.id || "").trim();
    if (!dedupeKey) {
      return true;
    }

    if (seenStockItemIds.has(dedupeKey)) {
      return false;
    }

    seenStockItemIds.add(dedupeKey);
    return true;
  }).slice(0, limit);
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
  return normalizeSearchValue(value);
}

function normalizeSearchValue(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "");
}

function getSearchTerms(query) {
  return String(query || "")
    .toLowerCase()
    .split(/[^a-z0-9]+/g)
    .map((term) => term.trim())
    .filter(Boolean);
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

function matchesGlobalOrderSearch(order, query = state.globalListSearchQuery) {
  const needle = normalizeSearchValue(query);
  if (!needle) {
    return true;
  }

  return [
    order?.reference,
    ...getReferenceLines(order),
    order?.locationName,
    order?.locationAddress,
    order?.deliveryAddress,
    getDriverDisplayName(order),
    order?.createdByName,
    order?.entryType,
    order?.status,
    order?.stockDescription,
    order?.branding,
    order?.notes,
    order?.factoryDestinationName,
    order?.driverFlagNote,
    order?.customerName,
  ].some((value) => normalizeSearchValue(value).includes(needle));
}

function matchesGlobalLocationSearch(group, query = state.globalListSearchQuery) {
  const needle = normalizeSearchValue(query);
  if (!needle) {
    return true;
  }

  return [
    group?.locationName,
    group?.locationAddress,
  ].some((value) => normalizeSearchValue(value).includes(needle));
}

function buildFilteredLocationGroup(group, orders) {
  return {
    ...group,
    orders,
    activeCount: orders.filter((order) => order.status === "active").length,
    completedCount: orders.filter((order) => order.status === "completed").length,
    priorityCount: orders.filter((order) => isPriorityOrder(order)).length,
  };
}

function getFilteredGlobalLocationGroups(groups = getGlobalLocationGroups(), query = state.globalListSearchQuery) {
  const needle = normalizeSearchValue(query);
  if (!needle) {
    return groups;
  }

  return groups
    .map((group) => {
      const locationMatches = matchesGlobalLocationSearch(group, query);
      const matchingOrders = locationMatches
        ? group.orders
        : group.orders.filter((order) => matchesGlobalOrderSearch(order, query));

      if (!matchingOrders.length) {
        return null;
      }

      return buildFilteredLocationGroup(group, matchingOrders);
    })
    .filter(Boolean);
}

function getGlobalListSearchSummary(filteredGroups, totalGroups, totalEntries) {
  const query = String(state.globalListSearchQuery || "").trim();
  const scheduledDate = String(state.globalListScheduledDate || "").trim();
  const filteredEntries = filteredGroups.reduce((sum, group) => sum + group.orders.length, 0);
  const dateLabel = formatDateOnly(scheduledDate) || scheduledDate;
  const dateSuffix = dateLabel ? ` on ${dateLabel}` : "";

  if (!query && !filteredEntries) {
    return dateLabel
      ? `No visible entries are scheduled for ${dateLabel}.`
      : "No entries available yet.";
  }

  if (!query) {
    return `${totalEntries} visible entr${totalEntries === 1 ? "y" : "ies"} across ${totalGroups} pickup location${totalGroups === 1 ? "" : "s"}${dateSuffix}.`;
  }

  if (!filteredEntries) {
    return `No entries or locations match "${query}"${dateSuffix}.`;
  }

  return `Showing ${filteredEntries} of ${totalEntries} entr${totalEntries === 1 ? "y" : "ies"} across ${filteredGroups.length} of ${totalGroups} location${totalGroups === 1 ? "" : "s"}${dateSuffix} for "${query}".`;
}

function getGlobalListEmptyStateMessage(totalEntries) {
  const query = String(state.globalListSearchQuery || "").trim();
  const scheduledDate = String(state.globalListScheduledDate || "").trim();
  const dateLabel = formatDateOnly(scheduledDate) || scheduledDate;

  if (query) {
    return `No entries or locations match "${query}"${dateLabel ? ` on ${dateLabel}` : ""}.`;
  }

  if (dateLabel && !totalEntries) {
    return `No entries are booked for ${dateLabel} yet.`;
  }

  return "No entries available yet.";
}

function getGlobalListScheduledDateOptions(orders = state.snapshot.orders) {
  const counts = new Map();
  const liveDate = getLiveScheduleDate();

  orders.forEach((order) => {
    const scheduledFor = getOrderScheduledForValue(order);
    if (!scheduledFor) {
      return;
    }

    const current = counts.get(scheduledFor) || {
      value: scheduledFor,
      count: 0,
      activeCount: 0,
      completedCount: 0,
      isLiveDate: scheduledFor === liveDate,
    };

    current.count += 1;
    if (order.status === "active") {
      current.activeCount += 1;
    }
    if (order.status === "completed") {
      current.completedCount += 1;
    }

    counts.set(scheduledFor, current);
  });

  return Array.from(counts.values()).sort((left, right) => String(left.value || "").localeCompare(String(right.value || "")));
}

function renderGlobalListDateFilterOptions(dateOptions, totalEntries) {
  return `
    <option value="">All scheduled dates${totalEntries ? ` (${totalEntries})` : ""}</option>
    ${dateOptions.map((option) => `
      <option value="${escapeHtml(option.value)}"${state.globalListScheduledDate === option.value ? " selected" : ""}>
        ${escapeHtml(`${formatDateOnly(option.value) || option.value} (${option.count})${option.isLiveDate ? " | Live date" : ""}`)}
      </option>
    `).join("")}
  `;
}

function renderGlobalListDateFilterPanel(dateOptions, totalEntries) {
  const selectedDate = String(state.globalListScheduledDate || "").trim();

  return `
    <section class="global-list-date-panel">
      <div class="global-list-date-panel-header">
        <div>
          <p class="eyebrow">Booked dates</p>
          <p class="network-search-hint">
            Choose a scheduled day to focus the global list. The live date is marked so you can see today's workload quickly.
          </p>
        </div>
        <div class="chip-row">
          <span class="chip">${dateOptions.length} booked day${dateOptions.length === 1 ? "" : "s"}</span>
          ${selectedDate
      ? `<span class="chip chip-success">Filtered to ${escapeHtml(formatDateOnly(selectedDate) || selectedDate)}</span>`
      : `<span class="chip">Showing all dates</span>`
    }
        </div>
      </div>
      <div class="global-list-booked-date-list" role="list" aria-label="Booked dates">
        <button
          type="button"
          class="button button-ghost global-list-booked-date-button${!selectedDate ? " is-selected" : ""}"
          data-action="set-global-list-date-filter"
          data-scheduled-for=""
          aria-pressed="${!selectedDate ? "true" : "false"}"
          ${state.busy ? " disabled" : ""}
        >
          <span class="global-list-booked-date-label">All scheduled dates</span>
          <span class="global-list-booked-date-meta">${totalEntries} entr${totalEntries === 1 ? "y" : "ies"}</span>
        </button>
        ${dateOptions.map((option) => `
          <button
            type="button"
            class="button button-ghost global-list-booked-date-button${selectedDate === option.value ? " is-selected" : ""}"
            data-action="set-global-list-date-filter"
            data-scheduled-for="${escapeHtml(option.value)}"
            aria-pressed="${selectedDate === option.value ? "true" : "false"}"
            ${state.busy ? " disabled" : ""}
          >
            <span class="global-list-booked-date-label">${escapeHtml(formatDateOnly(option.value) || option.value)}</span>
            <span class="global-list-booked-date-meta">
              ${option.count} entr${option.count === 1 ? "y" : "ies"}${option.isLiveDate ? " | Live date" : ""}
            </span>
          </button>
        `).join("")}
      </div>
    </section>
  `;
}

function matchesNetworkSearch(location, query = state.networkSearchQuery) {
  const searchTerms = getSearchTerms(query);
  if (!searchTerms.length) {
    return true;
  }

  const fields = [
    location?.name,
    location?.locationType,
    location?.address,
    location?.contactPerson,
    location?.contactNumber,
  ]
    .map((value) => normalizeSearchValue(value))
    .filter(Boolean);

  return searchTerms.every((term) => fields.some((value) => value.includes(term)));
}

function getFilteredNetworkLocations() {
  return state.snapshot.locations.filter((location) => matchesNetworkSearch(location));
}

function getNetworkSearchSummary(filteredCount, totalCount) {
  const query = String(state.networkSearchQuery || "").trim();
  if (!query) {
    return `${totalCount} pickup location${totalCount === 1 ? "" : "s"} in the network.`;
  }

  if (!filteredCount) {
    return `No locations match "${query}".`;
  }

  return `Showing ${filteredCount} of ${totalCount} location${totalCount === 1 ? "" : "s"} for "${query}".`;
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
    throw new Error(`That QR code is not a ${APP_NAME} stock label.`);
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

  const referenceValues = getStockReferenceValues(item);
  if (referenceValues.length) {
    return referenceValues[0];
  }

  const uuidHex = getStockItemUuidHex(item);
  if (uuidHex) {
    return encodeStockLabelCode(BigInt(`0x${uuidHex.slice(-10)}`), STOCK_LABEL_CODE_LENGTH);
  }

  return normalizeStockLabelText(item?.id || item?.name || "STOCK");
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
    getStockLookupValues(item).includes(normalizedValue)
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

function getStockReferenceValues(record) {
  const values = [
    record?.quoteNumber,
    record?.salesOrderNumber,
    record?.invoiceNumber,
    record?.poNumber,
  ]
    .map((value) => normalizeStockLabelText(value))
    .filter(Boolean);

  return values.filter((value, index) => values.indexOf(value) === index);
}

function getStockLookupValues(record) {
  const values = [];
  const sku = normalizeStockLabelText(record?.sku || "");
  if (sku) {
    values.push(sku);
  }

  getStockReferenceValues(record).forEach((value) => {
    if (!values.includes(value)) {
      values.push(value);
    }
  });

  const qrPayload = normalizeStockLabelText(buildStockQrPayload(record));
  if (qrPayload && !values.includes(qrPayload)) {
    values.push(qrPayload);
  }

  return values;
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
  const firstReference = getStockReferenceValues(item)[0] || "";
  return String(item?.sku || firstReference || item?.name || "Stock item").trim();
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
      <div class="topbar-status-strip">
        <span class="topbar-pill topbar-pill-live">Live date ${escapeHtml(liveDate)}</span>
        <span class="topbar-pill">Neon-backed workflow</span>
      </div>
    `;
    userActionsEl.innerHTML = `
      <div class="topbar-actions">
        ${renderThemeToggleButton()}
      </div>
    `;
    return;
  }

  const summaryLine = currentUser.role === "logistics"
    ? `${state.snapshot.stockMovements.length} logged stock movement${state.snapshot.stockMovements.length === 1 ? "" : "s"}`
    : `${state.snapshot.orders.length} visible entr${state.snapshot.orders.length === 1 ? "y" : "ies"}`;
  const refreshLine = getSnapshotRefreshSummary();
  const workspaceLabel = currentUser.role === "driver"
    ? "Driver route workspace"
    : `${capitalize(currentUser.role)} workspace`;
  const initials = getUserInitials(currentUser.name);

  authMetaEl.innerHTML = `
    <div class="topbar-status-strip">
      <span class="topbar-pill topbar-pill-live">Live date ${escapeHtml(liveDate)}</span>
      <span class="topbar-pill">${escapeHtml(summaryLine)}</span>
      <span class="topbar-pill">${escapeHtml(refreshLine)}</span>
    </div>
    <div class="topbar-user-card">
      <span class="topbar-user-avatar">${escapeHtml(initials)}</span>
      <span class="topbar-user-copy">
        <strong>${escapeHtml(currentUser.name)}</strong>
        <span>${escapeHtml(workspaceLabel)}</span>
      </span>
    </div>
  `;
  userActionsEl.innerHTML = `
    <div class="topbar-actions">
      ${renderThemeToggleButton()}
      <button class="button button-secondary" data-action="logout"${state.busy ? " disabled" : ""}>Logout</button>
    </div>
  `;
}

function renderThemeToggleButton() {
  const isLightTheme = state.theme === "light";
  return `
    <button
      class="button theme-toggle${isLightTheme ? " is-light" : ""}"
      type="button"
      data-action="toggle-theme"
      aria-pressed="${isLightTheme ? "true" : "false"}"
      title="Switch to ${isLightTheme ? "dark" : "light"} mode"
    >
      <span class="theme-toggle-track" aria-hidden="true">
        <span class="theme-toggle-thumb"></span>
      </span>
      <span class="theme-toggle-copy">${isLightTheme ? "Light" : "Dark"}</span>
    </button>
  `;
}

function getUserInitials(name) {
  const parts = String(name || "")
    .trim()
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2);

  if (!parts.length) {
    return "RL";
  }

  return parts.map((part) => part.charAt(0).toUpperCase()).join("");
}

function getSnapshotRefreshSummary() {
  const lastRefreshLabel = formatPreciseDateTime(state.lastSnapshotRefreshAt);
  if (isEntryFormSessionActive()) {
    if (!lastRefreshLabel) {
      return "Auto refresh is paused while the entry form is open. Save or close the form to sync again.";
    }

    return `Last refresh ${lastRefreshLabel}. Auto refresh is paused while the entry form is open.`;
  }

  if (!lastRefreshLabel) {
    return `Auto refresh every ${Math.round(AUTO_REFRESH_INTERVAL_MS / 1000)} sec. Waiting for first sync.`;
  }

  return `Last refresh ${lastRefreshLabel}. Auto refresh every ${Math.round(AUTO_REFRESH_INTERVAL_MS / 1000)} sec.`;
}

function renderPageNavigation() {
  if (!pageNavEl) {
    return;
  }

  const currentUser = state.snapshot.user;
  if (!currentUser) {
    pageNavEl.innerHTML = "";
    pageNavEl.removeAttribute("data-role");
    pageNavEl.classList.remove("page-nav-active");
    return;
  }

  const items = getNavigationItems(currentUser.role);
  const liveDate = state.publicState.today
    ? formatDateOnly(state.publicState.today)
    : "Waiting for sync";
  syncCurrentPage();
  pageNavEl.dataset.role = currentUser.role;
  pageNavEl.classList.add("page-nav-active");
  pageNavEl.innerHTML = `
    <div class="page-nav-panel">
      <div class="page-nav-header">
        <span class="page-nav-logo">RL</span>
        <span class="page-nav-header-copy">
          <strong>${escapeHtml(APP_NAME)}</strong>
          <span>${escapeHtml(capitalize(currentUser.role))} workspace</span>
        </span>
      </div>
      <div class="page-nav-section">
        <p class="page-nav-group">Navigation</p>
        <div class="page-nav-items">
          ${items
      .map(
        (item) => `
                  <button
                    class="page-nav-link${state.currentPage === item.id ? " is-active" : ""}"
                    data-action="navigate-page"
                    data-page-id="${item.id}"
                    ${state.busy ? "disabled" : ""}
                  >
                    <span class="page-nav-link-dot" aria-hidden="true"></span>
                    <span class="page-nav-link-text">${escapeHtml(item.label)}</span>
                  </button>
                `,
      )
      .join("")
    }
        </div>
      </div>
      <div class="page-nav-footer">
        <span class="page-nav-footer-label">Live date</span>
        <strong>${escapeHtml(liveDate)}</strong>
      </div>
    </div>
  `;
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
      { id: "guide", label: "Guide" },
    ];
  }

  if (role === "sales") {
    return [
      { id: "dashboard", label: "Dashboard" },
      { id: "entries", label: "Global List" },
      { id: "assignments", label: "Assignments" },
      { id: "stock", label: "Stock" },
      { id: "drivers", label: "Driver Lists" },
      { id: "guide", label: "Guide" },
    ];
  }

  if (role === "logistics") {
    return [
      { id: "dashboard", label: "Dashboard" },
      { id: "stock", label: "Stock" },
      { id: "guide", label: "Guide" },
    ];
  }

  return [
    { id: "route", label: "Route" },
    { id: "completed", label: "Completed" },
    { id: "guide", label: "Guide" },
  ];
}

function syncEntryFormVisibility(pageId = state.currentPage) {
  const currentUser = state.snapshot.user;
  const canUseEntryForm = Boolean(
    currentUser
    && (currentUser.role === "admin" || currentUser.role === "sales")
    && pageId === "entries",
  );

  if (!canUseEntryForm) {
    state.entryFormOpen = false;
    state.editingOrderId = "";
    state.orderEditReturnPage = "";
  }

  if (
    state.editingOrderId
    && !["admin", "sales"].includes(String(currentUser?.role || ""))
  ) {
    state.editingOrderId = "";
    state.orderEditReturnPage = "";
  }
}

function syncCurrentPage() {
  const currentUser = state.snapshot.user;
  if (!currentUser) {
    state.currentPage = "";
    syncEntryFormVisibility("");
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
  syncEntryFormVisibility(nextPage);
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
  syncEntryFormVisibility(pageId);
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
  syncEntryFormVisibility(hashPage);
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
          <p class="eyebrow">${escapeHtml(APP_NAME)}</p>
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
    <section class="screen-grid screen-grid-main">
      <div class="content">
        ${renderAdminPageContent()}
      </div>
    </section>
  `;
}

function renderSalesScreen() {
  return `
    <section class="screen-grid screen-grid-main">
      <div class="content">
        ${renderSalesPageContent()}
      </div>
    </section>
  `;
}

function renderLogisticsScreen() {
  return `
    <section class="screen-grid screen-grid-main">
      <div class="content">
        ${renderLogisticsPageContent()}
      </div>
    </section>
  `;
}

function renderDriverScreen() {
  return `
    <section class="screen-grid screen-grid-main">
      <div class="content">
        ${renderDriverPageContent()}
      </div>
    </section>
  `;
}

function renderRoleGuidePage(role) {
  const guide = ROLE_GUIDES[role];
  if (!guide) {
    return "";
  }

  return `
    <section class="hero-card guide-hero">
      <p class="eyebrow">${escapeHtml(capitalize(role))} guide</p>
      <h2>${escapeHtml(guide.title)}</h2>
      <p>${escapeHtml(guide.overview)}</p>
      <div class="chip-row">
        <span class="chip chip-role-${escapeHtml(role)}">${escapeHtml(capitalize(role))}</span>
        <span class="chip">${escapeHtml(String(getNavigationItems(role).length))} pages</span>
        <span class="chip">Role-based walkthrough</span>
      </div>
    </section>
    ${renderFlash()}
    <section class="panel-grid guide-grid">
      ${guide.roleFocus?.length
      ? renderGuidePanel({
        eyebrow: "Role scope",
        title: "What this role is responsible for",
        items: guide.roleFocus,
      })
      : ""
    }
      ${renderGuidePanel({
      eyebrow: "Your pages",
      title: "What each page helps you do",
      subtitle: guide.subtitle,
      items: getNavigationItems(role).map((item) => ({
        label: item.label,
        text: guide.pageNotes?.[item.id] || GUIDE_PAGE_SUMMARIES[item.id] || "Use this page for the actions available to your role.",
      })),
    })}
      ${guide.startingChecks?.length
      ? renderGuidePanel({
        eyebrow: "Start here",
        title: "What to check first",
        items: guide.startingChecks,
        ordered: true,
      })
      : ""
    }
      ${renderGuidePanel({
      eyebrow: "Daily flow",
      title: "Recommended routine",
      items: guide.dailyFlow,
      ordered: true,
    })}
      ${renderGuidePanel({
      eyebrow: "Key tasks",
      title: "What matters most in this role",
      items: guide.keyTasks,
    })}
      ${renderGuidePanel({
      eyebrow: "Keep in mind",
      title: "Important tips and limits",
      items: guide.tips,
    })}
    </section>
  `;
}

function renderGuidePanel({ eyebrow, title, subtitle = "", items = [], ordered = false }) {
  return `
    <article class="panel guide-panel">
      <p class="eyebrow">${escapeHtml(eyebrow)}</p>
      <h3 class="panel-title">${escapeHtml(title)}</h3>
      ${subtitle ? `<p class="panel-subtitle">${escapeHtml(subtitle)}</p>` : ""}
      ${renderGuideList(items, ordered)}
    </article>
  `;
}

function renderGuideList(items = [], ordered = false) {
  if (!items.length) {
    return '<p class="guide-note">No guide items are available for this role yet.</p>';
  }

  const tagName = ordered ? "ol" : "ul";
  return `
    <${tagName} class="guide-list">
      ${items.map((item) => renderGuideListItem(item)).join("")}
    </${tagName}>
  `;
}

function renderGuideListItem(item) {
  if (item && typeof item === "object" && !Array.isArray(item)) {
    const label = String(item.label || "").trim();
    const text = String(item.text || "").trim();
    if (label && text) {
      return `<li><strong>${escapeHtml(label)}:</strong> ${escapeHtml(text)}</li>`;
    }

    return `<li>${escapeHtml(label || text || "")}</li>`;
  }

  return `<li>${escapeHtml(String(item || "").trim())}</li>`;
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
      ${renderEntryComposerSection(
      state.snapshot.user,
      true,
      "Leave the driver unassigned to queue work for later dispatch. Admins can still authorize duplicate quote entries or send a driver back to a stop completed earlier today.",
    )}
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
        <p>Each driver page groups active entries by pickup location and now includes the latest recorded driver positions on a map.</p>
      </section>
      ${renderFlash()}
      ${renderDriverListOverview("admin")}
    `;
  }

  if (state.currentPage === "guide") {
    return renderRoleGuidePage("admin");
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
      ${renderEntryComposerSection(
      state.snapshot.user,
      false,
      "Leave the driver unassigned to hold the work for later dispatch. Duplicate protection and completed-stop protection still apply when a driver is selected.",
    )}
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

  if (state.currentPage === "stock") {
    return renderStockWorkspace({
      viewerRole: "sales",
      title: "Read-only stock visibility",
      subtitle: "Review live stock that has arrived and the current on-hand view without leaving the sales workspace.",
    });
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

  if (state.currentPage === "guide") {
    return renderRoleGuidePage("sales");
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
      ${renderPageSummaryCard("Stock", "Review what has arrived in stock without changing logistics records.", "stock")}
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

  if (state.currentPage === "guide") {
    return renderRoleGuidePage("logistics");
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
  const canManageStock = viewerRole === "admin" || viewerRole === "logistics";
  const canSeeArtwork = canManageStock;
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
      ${canSeeArtwork ? renderMetric("Artwork requests", state.snapshot.artworkRequests.length) : ""}
    </section>
    ${renderRecentStockActivitySection({ recentInbound, recentOutbound })}
    ${canManageStock
      ? `
          <section class="panel-grid">
            ${renderStockItemPanel(viewerRole)}
            ${renderStockMovementPanel()}
          </section>
          ${renderStockQrPreviewPanel()}
        `
      : `
          <section class="panel-grid">
            ${renderStockReadOnlyPanel()}
          </section>
        `
    }
    ${renderStockItemsSection(viewerRole)}
    ${renderStockMovementsSection(viewerRole)}
    ${canSeeArtwork ? renderArtworkRequestPanel(viewerRole) : ""}
    ${canSeeArtwork ? renderArtworkRequestsSection() : ""}
  `;
}

function renderStockReadOnlyPanel() {
  return `
    <article class="panel">
      <p class="eyebrow">Sales visibility</p>
      <h3 class="panel-title">Incoming stock overview</h3>
      <p class="panel-subtitle">
        Sales can search the live stock register, review recent arrivals, and open the movement ledger without changing logistics data.
      </p>
      <p class="field-note">Use the Recent arrivals section below to see what has come in most recently.</p>
    </article>
  `;
}

function renderStockItemPanel(viewerRole) {
  const editingItem = viewerRole === "admin" ? getEditingStockItem() : null;
  const isEditing = Boolean(editingItem);
  const onHandValue = Number(editingItem?.onHandQuantity || 0);

  return `
    <article class="panel">
      <p class="eyebrow">Stock master</p>
      <h3 class="panel-title">${isEditing ? "Edit stock item" : "Add stock item"}</h3>
      <p class="panel-subtitle">
        ${isEditing
      ? "Update the stock master record. Existing movement history stays linked to this item."
      : "Create the item once with its opening stock and order references, then log movements and artwork requests against it."
    }
      </p>
      ${isEditing
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
            <input name="quoteNumber" type="text" value="${escapeHtml(editingItem?.quoteNumber || "")}" placeholder="Optional">
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
        <p class="field-note">Enter at least one quote, sales order, invoice, or PO number for this stock item.</p>
        <div class="form-grid">
          <label>
            Stock code
            <input name="sku" type="text" value="${escapeHtml(editingItem?.sku || "")}" placeholder="SKU-001">
          </label>
          ${isEditing
      ? `
                <label>
                  On hand now
                  <input type="text" value="${escapeHtml(String(onHandValue))}" disabled>
                </label>
              `
      : `
                <label>
                  Opening stock
                  <input name="initialQuantity" type="number" min="1" step="1" inputmode="numeric" placeholder="0" required>
                </label>
              `
    }
        </div>
        <p class="field-note">
          ${isEditing
      ? "Use the stock movement form to add or remove quantity."
      : "Adding the item records this quantity as stock in right away."
    }
        </p>
        <label>
          Notes
          <textarea name="notes" placeholder="Material, finish, size, or any internal note">${escapeHtml(editingItem?.notes || "")}</textarea>
        </label>
        <div class="action-row">
          <button type="submit" class="button button-primary"${state.busy ? " disabled" : ""}>
            ${isEditing ? "Save stock item" : "Add stock item"}
          </button>
          ${isEditing
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
        ${isEditing
      ? "Update the selected movement. The original logged time and recorded-by details stay unchanged."
      : "Every stock movement is timestamped and linked to the staff member who recorded it."
    }
      </p>
      ${isEditing
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
        ${isEditing
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
        ${!state.mailConfigured
        ? `<p class="field-note">${escapeHtml(state.mailConfigReason || "Email delivery is not configured yet.")}</p>`
        : ""
      }
      </form>
    `,
  });
}

function renderStockItemsSection(viewerRole) {
  const allowDelete = viewerRole === "admin";
  const readOnlyView = viewerRole === "sales";
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
            ${readOnlyView
      ? "Sales can search the live stock register and see which items have arrived, but only admin or logistics can change stock records."
      : allowDelete
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
            ${state.stockSearchQuery
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
              <th>Notes</th>
              <th>Updated</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
            ${filteredItems.length
      ? filteredItems.map((item) => renderStockItemRow(item, viewerRole, latestMovementByItemId)).join("")
      : `
                  <tr>
                    <td colspan="7">${totalItems ? "No stock items match this search." : "No stock items added yet."}</td>
                  </tr>
                `
    }
          </tbody>
        </table>
      </div>
      <div class="stock-item-mobile-list">
        ${filteredItems.length
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
            Latest stock activity from the last ${escapeHtml(String(STOCK_RECENT_ACTIVITY_HOURS))} hours, including manually created stock items and driver office drops, grouped by item.
          </p>
        </div>
      </div>
      <div class="recent-stock-grid">
        ${renderRecentStockActivityColumn({
    title: "Recently arrived",
    subtitle: "Latest stock-in activity, manually created stock items, and driver office drops per stock line.",
    emptyLabel: `No stock arrived, was created, or was dropped at the office in the last ${STOCK_RECENT_ACTIVITY_HOURS} hours.`,
    movements: recentInbound,
  })}
        ${renderRecentStockActivityColumn({
    title: "Recently shipped",
    subtitle: "Latest stock-out activity per item.",
    emptyLabel: `No stock shipped in the last ${STOCK_RECENT_ACTIVITY_HOURS} hours.`,
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
      ${movements.length
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
  const unitLabel = getStockUnitLabel(item || movement);
  const isCreatedActivity = movement.activityKind === "created";
  const isOfficeDropActivity = movement.activityKind === "office_drop";
  const onHandLabel = item
    ? `${Number(item.onHandQuantity || 0)} ${unitLabel} on hand now`
    : "";
  const referenceSummary = getReferenceLines(movement).join(" | ") || "No order references";
  const quantityLabel = isCreatedActivity
    ? "New item"
    : isOfficeDropActivity
      ? "Office drop"
      : `${Number(movement.quantity || 0)} ${unitLabel}`;
  const partyLabel = isCreatedActivity
    ? "Stock item created"
    : isOfficeDropActivity
      ? (movement.driverName ? `Dropped at office by ${movement.driverName}` : "Dropped at office")
      : getStockMovementPartyLabel(movement);
  const locationLabel = isOfficeDropActivity
    ? [movement.locationName, movement.locationAddress].filter(Boolean).join(" | ")
    : "";

  return `
    <article class="recent-stock-entry">
      <div class="recent-stock-entry-head">
        <div class="recent-stock-entry-copy">
          <h5 class="recent-stock-entry-title">${escapeHtml(movement.itemName || "Unknown")}</h5>
          <p class="recent-stock-entry-summary">${escapeHtml(referenceSummary)}</p>
        </div>
        <span class="chip ${movement.movementType === "in" ? "chip-success" : "chip-warning"}">
          ${escapeHtml(quantityLabel)}
        </span>
      </div>
      <div class="recent-stock-meta">
        <span>${escapeHtml(formatDateTime(movement.createdAt) || "")}</span>
        <span>${escapeHtml(partyLabel)}</span>
        ${locationLabel ? `<span>${escapeHtml(locationLabel)}</span>` : ""}
        ${onHandLabel ? `<span>${escapeHtml(onHandLabel)}</span>` : ""}
      </div>
    </article>
  `;
}

function renderStockMovementsSection(viewerRole) {
  const canEdit = viewerRole === "admin" || viewerRole === "logistics";
  return renderStockDisclosure({
    sectionKey: "movements",
    eyebrow: "Stock history",
    title: "Movement ledger",
    subtitle: canEdit
      ? "Use Edit to correct the item, movement type, quantity, party, or notes for an existing stock entry."
      : "Review the stock-in and stock-out history without changing logistics records.",
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
            ${state.snapshot.stockMovements.length
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
      ${state.stockQrBusy
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
      ${pendingItem
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
                ${state.stockScannerZoomSupported
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
            ${state.snapshot.artworkRequests.length
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
  const plan = getRoutePlan(currentUser.id, { scheduledFor: getLiveScheduleDate() });
  const completedOrders = getCompletedOrders(currentUser.id);

  if (state.currentPage === "guide") {
    return renderRoleGuidePage("driver");
  }

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
      <p>You only see the entries assigned to your name for the live date. Completing an entry removes it from the live route sequence.</p>
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
      ${plan.stops.length
      ? plan.stops.map((stop, index) => renderStopCard(stop, index, "driver", currentUser.id)).join("")
      : `
            <div class="empty-state">
              No active entries are scheduled for you today. Future-dated work will appear here on the scheduled day.
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
        ${isEditing
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
          ${isEditing
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
  const filteredLocations = getFilteredNetworkLocations();
  const totalLocations = state.snapshot.locations.length;

  return `
    <section class="table-card">
      <p class="eyebrow">Location network</p>
      <h3 class="panel-title">Pickup locations</h3>
      <p class="panel-subtitle">Each location stores its own type, address, optional coordinates, and optional contact details.</p>
      <article class="panel">
        <p class="eyebrow">Locations</p>
        <h3 class="panel-title">${isEditingLocation ? "Edit location" : "Add location"}</h3>
        <p class="panel-subtitle">
          ${isEditingLocation
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
            ${isEditingLocation
      ? `
                  <button type="button" class="button button-ghost" data-action="cancel-edit-location"${state.busy ? " disabled" : ""}>
                    Cancel
                  </button>
                `
      : ""
    }
          </div>
        </form>
        <div class="table-toolbar stock-table-toolbar">
          <div class="stock-section-copy">
            <p class="stock-results-note">${escapeHtml(getNetworkSearchSummary(filteredLocations.length, totalLocations))}</p>
            <p class="network-search-hint">Search visible location fields as you type. Use multiple words to narrow the list by name, type, address, or contact.</p>
          </div>
          <label class="stock-search">
            Search locations
            <div class="stock-search-controls">
              <input
                type="search"
                value="${escapeHtml(state.networkSearchQuery)}"
                placeholder="Try: factory sandton or giftwrap office"
                autocomplete="off"
                spellcheck="false"
                data-network-search
              >
              ${state.networkSearchQuery
      ? `<button type="button" class="button button-ghost" data-action="clear-network-search"${state.busy ? " disabled" : ""}>Clear</button>`
      : ""
    }
            </div>
          </label>
        </div>
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
              ${filteredLocations.length
      ? filteredLocations.map((location) => renderLocationRow(location)).join("")
      : `
                    <tr>
                      <td colspan="5">${totalLocations ? "No locations match this search." : "No locations added yet."}</td>
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
            ${state.snapshot.users.length
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

function renderEntryComposerSection(currentUser, allowDuplicateOverride, subtitle) {
  const editingOrder = getEditingOrder();
  const isEditing = Boolean(editingOrder);
  const isOpen = state.entryFormOpen || isEditing;

  return `
    <section class="panel entry-composer">
      <div class="entry-composer-header">
        <div class="entry-composer-copy">
          <p class="eyebrow">Global List</p>
          <h3 class="panel-title">${isEditing ? "Edit entry" : "Create a new entry"}</h3>
          <p class="panel-subtitle">${escapeHtml(
    isEditing
      ? "Update the entry details, then save the correction back into the system."
      : subtitle,
  )}</p>
        </div>
        <button
          type="button"
          class="button ${isOpen ? "button-ghost" : "button-primary"}"
          data-action="toggle-entry-form"
          aria-expanded="${isOpen ? "true" : "false"}"
          ${state.busy ? "disabled" : ""}
        >
          ${isEditing ? "Close editor" : isOpen ? "Hide entry form" : "Add new entry"}
        </button>
      </div>
      ${isOpen
      ? `
              <div class="entry-composer-body">
                ${isEditing
        ? `<p class="field-note">Editing updates the saved entry details. Existing stock item records are not renamed automatically, so correct stock references separately if needed.</p>`
        : ""
      }
                ${renderEntryForm(currentUser, allowDuplicateOverride, { order: editingOrder })}
              </div>
            `
      : `<p class="entry-composer-note">Open the form only when you need to capture a new job. The live global list stays visible below for easier scanning.</p>`
    }
    </section>
  `;
}

function renderEntryForm(currentUser, allowDuplicateOverride, options = {}) {
  const editingOrder = options.order || null;
  const isEditing = Boolean(editingOrder);
  const formId = isEditing ? "edit-order" : "add-order";
  const selectedDriverId = String(editingOrder?.driverUserId || "").trim();
  const selectedLocationId = String(editingOrder?.locationId || "").trim();
  const entryType = String(editingOrder?.entryType || "delivery").trim() || "delivery";
  const scheduledFor = String(editingOrder?.scheduledFor || getDefaultScheduledDateValue()).trim();
  const createdByName = String(editingOrder?.createdByName || currentUser.name || "").trim();
  const deliveryAddress = String(editingOrder?.deliveryAddress || "").trim();
  const quoteNumber = String(editingOrder?.quoteNumber || "").trim();
  const salesOrderNumber = String(editingOrder?.salesOrderNumber || "").trim();
  const invoiceNumber = String(editingOrder?.invoiceNumber || "").trim();
  const poNumber = String(editingOrder?.poNumber || "").trim();
  const stockDescription = String(editingOrder?.stockDescription || "").trim();
  const branding = String(editingOrder?.branding || "").trim();
  const moveToFactory = Boolean(editingOrder?.moveToFactory);
  const factoryDestinationLocationId = String(editingOrder?.factoryDestinationLocationId || "").trim();
  const notice = String(editingOrder?.notes || "").trim();
  const isPriority = getOrderPriority(editingOrder) === PRIORITY_STOP_VALUE;

  return `
    <form data-form="${formId}">
      ${isEditing ? `<input type="hidden" name="orderId" value="${escapeHtml(editingOrder?.id || "")}">` : ""}
      <div class="form-grid">
        <label>
          Driver
          <select name="driverUserId">
            ${renderDriverOptions(selectedDriverId, true)}
          </select>
        </label>
        <label>
          Pickup location
          <select name="locationId" required>
            ${renderLocationOptions(selectedLocationId)}
          </select>
        </label>
      </div>
      <p class="field-note">
        Leave the driver as Unassigned if dispatch will allocate it later.
      </p>
      <div class="form-grid-3">
        <label>
          Collection or delivery
          <select name="entryType" required>
            <option value="collection"${entryType === "collection" ? " selected" : ""}>Collection</option>
            <option value="delivery"${entryType === "delivery" ? " selected" : ""}>Delivery</option>
          </select>
        </label>
        <label>
          Scheduled date
          <input name="scheduledFor" type="date" value="${escapeHtml(scheduledFor)}" required>
        </label>
        <label>
          ${isEditing ? "Originally created by" : "Created by"}
          <input class="readonly-field" type="text" value="${escapeHtml(createdByName)}" readonly>
        </label>
      </div>
      <p class="field-note">
        Use a future date to plan work ahead of time. The daily CSV and email export only include entries scheduled for the live date.
      </p>
      <label class="hidden" data-delivery-address-field>
        Delivery address
        <textarea
          name="deliveryAddress"
          placeholder="Required for delivery entries"
          data-delivery-address-input
        >${escapeHtml(deliveryAddress)}</textarea>
      </label>
      <p class="field-note hidden" data-delivery-address-note>
        Add the address the driver must deliver this order to.
      </p>
      <label class="inline-check">
        <input type="checkbox" name="priority"${isPriority ? " checked" : ""}>
        Mark this as a priority stop
      </label>
      <div class="form-grid">
        <label>
          Quote number
          <input name="quoteNumber" type="text" value="${escapeHtml(quoteNumber)}" required>
        </label>
        <label>
          Sales order number
          <input name="salesOrderNumber" type="text" value="${escapeHtml(salesOrderNumber)}" placeholder="Optional">
        </label>
      </div>
      <div class="form-grid">
        <label>
          Invoice number
          <input name="invoiceNumber" type="text" value="${escapeHtml(invoiceNumber)}" placeholder="Optional">
        </label>
        <label>
          PO number
          <input name="poNumber" type="text" value="${escapeHtml(poNumber)}" placeholder="Optional">
        </label>
      </div>
      <label>
        Stock description
        <textarea name="stockDescription" placeholder="What stock is on this entry? Put one item per line." required>${escapeHtml(stockDescription)}</textarea>
      </label>
      <p class="field-note">
        ${isEditing
      ? "This updates the entry details shown in the live route list. If the stock item records also need correction, update those separately in the Stock page."
      : "Saving this entry will also create matching stock items automatically in the Stock page. Put each item on its own line to create separate stock records."
    }
      </p>
      <label>
        Branding
        <input name="branding" type="text" value="${escapeHtml(branding)}" placeholder="Optional">
      </label>
      <label class="inline-check">
        <input type="checkbox" name="moveToFactory" disabled${moveToFactory ? " checked" : ""}>
        Move collected stock to a factory
      </label>
      <label class="hidden" data-move-to-factory-destination>
        Destination factory
        <select name="factoryDestinationLocationId">
          ${renderFactoryLocationOptions(factoryDestinationLocationId)}
        </select>
      </label>
      <p class="field-note" data-move-to-factory-note>
        Switch the entry type to Collection to enable factory transfer.
      </p>
      <label>
        Notice
        <textarea name="notice" placeholder="Special instruction, handover note, or anything dispatch should keep on the order">${escapeHtml(notice)}</textarea>
      </label>
      ${allowDuplicateOverride
      ? `
            <label class="inline-check">
              <input type="checkbox" name="allowDuplicate">
              Admin override for duplicate or return stop
            </label>
            <p class="field-note">
              This only lets admins send a driver back to a stop they already completed today, or keep two active entries with the same quote number on that driver's list.
            </p>
          `
      : ""
    }
      <div class="action-row">
        <button type="submit" class="button button-primary"${state.busy ? " disabled" : ""}>
          ${isEditing ? "Save changes" : "Create entry"}
        </button>
        ${isEditing
      ? `<button type="button" class="button button-ghost" data-action="cancel-edit-order"${state.busy ? " disabled" : ""}>Cancel</button>`
      : ""
    }
      </div>
    </form>
  `;
}

function renderAssignmentManager(viewerRole) {
  const filteredOrders = getFilteredAssignmentOrders();
  const locationGroups = getFilteredAssignmentLocationGroups(filteredOrders);
  const page = getPaginationData(locationGroups, "assignments", PAGE_SIZES.assignments);

  return `
    <section class="table-card">
      <div class="table-toolbar">
        <div class="stock-section-copy">
          <p class="eyebrow">Assignments</p>
          <h3 class="panel-title">Active entry allocation</h3>
          <p class="panel-subtitle">Entries are grouped by pickup location so you can dispatch the visible work for a stop in one action instead of one order at a time.</p>
          <p class="stock-results-note">${escapeHtml(getAssignmentFilterSummary(locationGroups.length, filteredOrders.length))}</p>
        </div>
        <label class="assignment-filter">
          Filter by driver
          <select data-assignment-filter>
            ${renderAssignmentFilterOptions(state.assignmentDriverFilter)}
          </select>
        </label>
      </div>
      <div class="global-location-groups">
        ${page.items.length
      ? page.items.map((group) => renderAssignmentLocationGroup(group, viewerRole)).join("")
      : `<div class="empty-state">No active entries match this driver filter.</div>`
    }
      </div>
      ${renderPaginationControls("assignments", page)}
    </section>
  `;
}

function renderGlobalOrdersSection(viewerRole) {
  const sortedOrders = [...state.snapshot.orders].sort(orderDisplaySort);
  const dateOptions = getGlobalListScheduledDateOptions(sortedOrders);
  const dateFilteredOrders = filterOrdersForScheduledDate(sortedOrders, state.globalListScheduledDate);
  const locationGroups = getGlobalLocationGroups(dateFilteredOrders);
  const filteredLocationGroups = getFilteredGlobalLocationGroups(locationGroups);
  const page = getPaginationData(filteredLocationGroups, "globalEntries", PAGE_SIZES.globalEntries);
  const canExport = viewerRole === "admin" || viewerRole === "sales";
  const canDelete = viewerRole === "admin";
  const priorityCount = getActivePriorityOrders().length;
  const rolloverCount = getActiveCarryOverOrders().length;

  return `
    <section class="table-card">
      <div class="table-toolbar">
        <div>
          <p class="eyebrow">Global List</p>
          <h3 class="panel-title">All visible entries</h3>
          <p class="panel-subtitle">
            ${canDelete
      ? "This view groups every visible entry by pickup location. Expand a location to review the orders there, and admins can still remove entries if needed."
      : "This view groups every visible entry by pickup location. Expand a location to review the orders there."
    }
          </p>
          <p class="stock-results-note">${escapeHtml(getGlobalListSearchSummary(filteredLocationGroups, locationGroups.length, dateFilteredOrders.length))}</p>
        </div>
        <label class="assignment-filter global-list-date-filter">
          Scheduled date
          <div class="stock-search-controls global-list-date-filter-controls">
            <select data-global-list-date-filter>
              ${renderGlobalListDateFilterOptions(dateOptions, sortedOrders.length)}
            </select>
            ${state.globalListScheduledDate
      ? `<button type="button" class="button button-ghost" data-action="clear-global-list-date-filter"${state.busy ? " disabled" : ""}>Clear</button>`
      : ""
    }
          </div>
        </label>
        <label class="stock-search">
          Search entries
          <div class="stock-search-controls">
            <input
              type="search"
              value="${escapeHtml(state.globalListSearchQuery)}"
              placeholder="Location, reference, driver, stock description"
              autocomplete="off"
              spellcheck="false"
              data-global-list-search
            >
            ${state.globalListSearchQuery
      ? `<button type="button" class="button button-ghost" data-action="clear-global-list-search"${state.busy ? " disabled" : ""}>Clear</button>`
      : ""
    }
          </div>
        </label>
        ${canExport
      ? `
              <div class="action-row">
                ${viewerRole === "admin"
        ? `
                      <button
                        class="button button-ghost"
                        data-action="clear-all-order-priorities"
                        ${state.busy || !priorityCount ? "disabled" : ""}
                      >
                        Clear all priority${priorityCount ? ` (${priorityCount})` : ""}
                      </button>
                      <button
                        class="button button-ghost"
                        data-action="clear-order-rollovers"
                        ${state.busy || !rolloverCount ? "disabled" : ""}
                      >
                        Clear all rollover${rolloverCount ? ` (${rolloverCount})` : ""}
                      </button>
                    `
        : ""
      }
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
                  class="button button-secondary"
                  data-action="email-rollover-test"
                  ${state.busy || !state.mailConfigured ? "disabled" : ""}
                >
                  Test Rollover Email
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
      ${canExport && !state.mailConfigured
      ? `<p class="field-note">${escapeHtml(state.mailConfigReason || "Email delivery is not configured yet.")}</p>`
      : ""
    }
      ${dateOptions.length ? renderGlobalListDateFilterPanel(dateOptions, sortedOrders.length) : ""}
      <div class="global-location-groups">
        ${page.items.length
      ? page.items.map((group) => renderGlobalLocationGroup(group, viewerRole)).join("")
      : `<div class="empty-state">${escapeHtml(getGlobalListEmptyStateMessage(dateFilteredOrders.length))}</div>`
    }
      </div>
      ${renderPaginationControls("globalEntries", page)}
    </section>
  `;
}

function renderAdminRolloverPanel() {
  const rolloverCount = getActiveCarryOverOrders().length;

  return `
    <section class="panel">
      <div class="table-toolbar">
        <div>
          <p class="eyebrow">Admin tools</p>
          <h3 class="panel-title">Rollover controls</h3>
          <p class="panel-subtitle">Clear the current rollover markers once the day’s carry-over report has been checked.</p>
        </div>
        <div class="chip-row">
          <span class="chip">${rolloverCount} active rollover entr${rolloverCount === 1 ? "y" : "ies"}</span>
        </div>
      </div>
      <div class="action-row">
        <button
          class="button button-secondary"
          data-action="clear-order-rollovers"
          ${state.busy || !rolloverCount ? "disabled" : ""}
        >
          Clear rollover
        </button>
      </div>
    </section>
  `;
}

function getGlobalLocationGroups(sortedOrders = [...state.snapshot.orders].sort(orderDisplaySort)) {
  const grouped = new Map();

  sortedOrders.forEach((order) => {
    const locationId = String(order.locationId || "").trim();
    const fallbackKey = [order.locationName, order.locationAddress]
      .map((value) => String(value || "").trim().toLowerCase())
      .filter(Boolean)
      .join("|");
    const groupKey = locationId || fallbackKey || `order:${order.id}`;

    if (!grouped.has(groupKey)) {
      grouped.set(groupKey, {
        key: groupKey,
        locationId,
        locationName: order.locationName || "Unknown",
        locationAddress: order.locationAddress || "",
        orders: [],
        activeCount: 0,
        completedCount: 0,
        priorityCount: 0,
      });
    }

    const group = grouped.get(groupKey);
    group.orders.push(order);
    if (order.status === "active") {
      group.activeCount += 1;
    }
    if (order.status === "completed") {
      group.completedCount += 1;
    }
    if (isPriorityOrder(order)) {
      group.priorityCount += 1;
    }
  });

  return Array.from(grouped.values())
    .map((group) => ({
      ...group,
      orders: [...group.orders].sort(orderDisplaySort),
    }))
    .sort((left, right) => {
      const leftHasActive = left.activeCount > 0 ? 0 : 1;
      const rightHasActive = right.activeCount > 0 ? 0 : 1;
      if (leftHasActive !== rightHasActive) {
        return leftHasActive - rightHasActive;
      }

      const nameCompare = String(left.locationName || "").localeCompare(String(right.locationName || ""), undefined, {
        sensitivity: "base",
      });
      if (nameCompare) {
        return nameCompare;
      }

      return String(left.locationAddress || "").localeCompare(String(right.locationAddress || ""), undefined, {
        sensitivity: "base",
      });
    });
}

function renderGlobalLocationGroup(group, viewerRole) {
  const locationRecord = getLocation(group.locationId) || {
    name: group.locationName,
    address: group.locationAddress,
  };
  const navigationUrl = getGoogleMapsNavigateUrl(locationRecord);
  const stopCardKey = buildGlobalLocationGroupKey(group.key);
  const isOpen = isStopCardOpen(stopCardKey);

  return `
    <article class="stop-card global-location-group${isOpen ? " is-open" : " is-collapsed"}">
      <div class="stop-header">
        <div>
          <p class="eyebrow">Pickup location</p>
          <h4 class="stop-title">${escapeHtml(group.locationName)}</h4>
          ${isOpen ? `<p class="stop-address">${escapeHtml(group.locationAddress || "Address not set")}</p>` : ""}
        </div>
        <div class="chip-row">
          <span class="chip">${group.orders.length} entr${group.orders.length === 1 ? "y" : "ies"}</span>
          ${group.activeCount ? `<span class="chip chip-success">${group.activeCount} active</span>` : ""}
          ${group.completedCount ? `<span class="chip">${group.completedCount} completed</span>` : ""}
          ${group.priorityCount ? `<span class="chip chip-priority-high">${group.priorityCount} priority</span>` : ""}
        </div>
      </div>
      <div class="action-row stop-actions">
        ${navigationUrl
      ? `
              <a
                class="button button-primary"
                href="${escapeHtml(navigationUrl)}"
                target="_blank"
                rel="noreferrer noopener"
              >
                Navigate
              </a>
            `
      : ""
    }
        <button
          type="button"
          class="button button-ghost"
          data-action="toggle-stop-card"
          data-stop-card-key="${escapeHtml(stopCardKey)}"
          ${state.busy ? " disabled" : ""}
        >
          ${isOpen ? "Hide entries" : "Show entries"}
        </button>
      </div>
      ${isOpen
      ? `
            <div class="location-group-orders">
              ${group.orders.map((order) => renderGlobalOrderCard(order, viewerRole)).join("")}
            </div>
          `
      : ""
    }
    </article>
  `;
}

function renderGlobalOrderCard(order, viewerRole) {
  const canDelete = viewerRole === "admin";
  const canEdit = viewerRole === "admin" || (viewerRole === "sales" && order.status === "active");
  const isPriority = isPriorityOrder(order);
  const referenceLines = getOrderListReferenceLines(order);
  const createdAt = formatDateTime(order.createdAt);

  return `
    <div class="order-card${isPriority ? " order-card-priority" : ""}">
      <div class="stop-header">
        <div>
          <strong>${escapeHtml(getOrderPrimaryDisplay(order))}</strong>
          <div class="order-meta">
            ${referenceLines.length
      ? referenceLines.map((line) => `<span>${escapeHtml(line)}</span>`).join("")
      : '<span>No order references</span>'
    }
          </div>
        </div>
        <div class="chip-row">
          ${renderTypeChip(order.entryType)}
          ${renderOrderScheduledChip(order)}
          ${renderOrderPriorityChip(order)}
          ${renderOrderPickupChip(order)}
          ${order.moveToFactory ? '<span class="chip chip-warning">Factory move</span>' : ""}
          ${renderOrderFlagChip(order)}
          ${order.driverUserId
      ? `<span class="chip">Driver: ${escapeHtml(getDriverDisplayName(order))}</span>`
      : '<span class="chip chip-warning">Unassigned</span>'
    }
          ${renderStatusChip(order.status)}
        </div>
      </div>
      ${renderOrderStockDetails(order)}
      ${renderOrderNotice(order)}
      <p class="order-card-meta-note">
        ${escapeHtml(`Created by ${order.createdByName || "Unknown"}${createdAt ? ` on ${createdAt}` : ""}`)}
      </p>
      ${canEdit || canDelete
      ? `
            <div class="action-row">
              ${canEdit
        ? `
                    <button
                      class="button button-secondary"
                      data-action="edit-order"
                      data-order-id="${order.id}"
                      ${state.busy ? " disabled" : ""}
                    >
                      Edit
                    </button>
                  `
        : ""
      }
              <button
                class="button button-danger"
                data-action="delete-order"
                data-order-id="${order.id}"
                data-order-reference="${escapeHtml(getOrderPrimaryDisplay(order))}"
                ${state.busy ? " disabled" : ""}
              >
                Delete
              </button>
            </div>
          `
      : ""
    }
    </div>
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
              ${plan.priorityStopCount
          ? `<span class="chip chip-priority-high">${plan.priorityStopCount} priority stop${plan.priorityStopCount === 1 ? "" : "s"}</span>`
          : ""
        }
              ${duplicateCount
          ? `<span class="chip chip-warning">${duplicateCount} duplicate order${duplicateCount === 1 ? "" : "s"}</span>`
          : ""
        }
            </div>
          </div>
          <p class="panel-subtitle">Active pickup route sequence.</p>
          ${plan.stops.length
          ? plan.stops.map((stop, index) => renderStopCard(stop, index, viewerRole, driver.id)).join("")
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
      ${viewerRole === "admin" ? renderAdminDriverLocationSection() : ""}
      <div class="panel-grid">
        ${content || '<div class="empty-state">No driver accounts have been created yet.</div>'}
      </div>
      ${renderPaginationControls("driverLists", page)}
    </section>
  `;
}

function renderAdminDriverLocationSection() {
  const totalDrivers = getDriverUsers().length;
  const mappedDrivers = getDriversWithRecordedPositions().length;
  const awaitingDrivers = Math.max(totalDrivers - mappedDrivers, 0);

  return `
    <section class="route-canvas-card admin-driver-map-card">
      <div class="table-toolbar">
        <div>
          <p class="eyebrow">Driver locations</p>
          <h3 class="panel-title">Last recorded driver positions</h3>
          <p class="panel-subtitle">Markers update from the most recent location the driver allowed ${escapeHtml(APP_NAME)} to record.</p>
        </div>
        <div class="chip-row">
          <span class="chip">${mappedDrivers} mapped</span>
          <span class="chip">${awaitingDrivers} awaiting location</span>
        </div>
      </div>
      <div class="route-canvas-wrap">
        <div id="admin-driver-location-map" class="route-map" aria-label="Map of drivers and their last recorded positions"></div>
      </div>
      <p id="admin-driver-location-status" class="route-map-status hidden"></p>
    </section>
  `;
}

function renderCompletedOrders(driverUserId, viewerRole = state.snapshot.user?.role || "") {
  const completed = getCompletedOrders(driverUserId).sort(orderDisplaySort);
  const page = getPaginationData(completed, "completedEntries", PAGE_SIZES.completedEntries);
  const canEdit = viewerRole === "admin";

  return `
    <section class="table-card">
      <p class="eyebrow">Completed</p>
      <h3 class="panel-title">Finished entries</h3>
      <div class="table-scroll">
        <table class="responsive-stack">
          <thead>
            <tr>
              <th>Quote</th>
              <th>Other references</th>
              <th>Type</th>
              <th>Completed</th>
              ${canEdit ? "<th>Action</th>" : ""}
            </tr>
          </thead>
          <tbody>
            ${page.items.length
      ? page.items
        .map(
          (order) => `
                        <tr>
                          <td data-label="Quote">${escapeHtml(getOrderPrimaryDisplay(order))}</td>
                          <td data-label="Other references">${renderOrderListReferenceSummary(order)}</td>
                          <td data-label="Type">${renderTypeChip(order.entryType)}</td>
                          <td data-label="Completed">
                            ${escapeHtml(formatDateTime(order.completedAt) || "Not completed")}<br>
                            <span class="muted">${escapeHtml(getOrderCompletionLabel(order) || "Completed")}</span>
                          </td>
                          ${canEdit
              ? `
                                <td data-label="Action">
                                  <button
                                    class="button button-secondary"
                                    data-action="edit-order"
                                    data-order-id="${order.id}"
                                    ${state.busy ? " disabled" : ""}
                                  >
                                    Edit
                                  </button>
                                </td>
                              `
              : ""
            }
                        </tr>
                      `,
        )
        .join("")
      : `
                  <tr>
                    <td colspan="${canEdit ? "5" : "4"}">No completed entries yet.</td>
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
      <td data-label="Quote">${escapeHtml(getOrderPrimaryDisplay(order))}</td>
      <td data-label="Other references">${renderOrderListReferenceSummary(order)}</td>
      <td data-label="Type">
        <div class="chip-row">
          ${renderTypeChip(order.entryType)}
          ${renderOrderPriorityChip(order)}
          ${renderOrderPickupChip(order)}
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
      ${canDelete
      ? `
            <td data-label="Actions">
              <div class="action-row">
                <button
                  class="button button-danger"
                  data-action="delete-order"
                  data-order-id="${order.id}"
                  data-order-reference="${escapeHtml(getOrderPrimaryDisplay(order))}"
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

function renderAssignmentLocationGroup(group, viewerRole) {
  const locationRecord = getLocation(group.locationId) || {
    name: group.locationName,
    address: group.locationAddress,
  };
  const navigationUrl = getGoogleMapsNavigateUrl(locationRecord);
  const stopCardKey = buildAssignmentLocationGroupKey(group.key);
  const isOpen = isStopCardOpen(stopCardKey);

  return `
    <article class="stop-card${group.priorityCount ? " stop-card-priority" : ""}${isOpen ? " is-open" : " is-collapsed"}">
      <div class="stop-header">
        <div>
          <p class="eyebrow">Pickup location</p>
          <h4 class="stop-title">${escapeHtml(group.locationName || "Unknown")}</h4>
          ${isOpen ? `<p class="stop-address">${escapeHtml(group.locationAddress || "Address not set")}</p>` : ""}
          <p class="assignment-location-current-drivers">${escapeHtml(getAssignmentLocationDriverLine(group))}</p>
        </div>
        <div class="chip-row">
          <span class="chip">${group.orders.length} entr${group.orders.length === 1 ? "y" : "ies"}</span>
          ${group.unassignedCount ? `<span class="chip chip-warning">${group.unassignedCount} unassigned</span>` : ""}
          ${group.assignedCount ? `<span class="chip chip-success">${group.assignedCount} assigned</span>` : ""}
          ${group.priorityCount ? `<span class="chip chip-priority-high">${group.priorityCount} priority</span>` : ""}
        </div>
      </div>
      <div class="assignment-location-toolbar">
        <div class="assignment-control">
          <select data-assignment-location-driver data-group-key="${escapeHtml(group.key)}">
            ${renderAssignmentLocationDriverOptions(group.selectedDriverId, group.hasMixedDriverSelection)}
          </select>
          ${viewerRole === "admin"
      ? `
                <label class="inline-check assignment-inline">
                  <input type="checkbox" data-assignment-location-allow-duplicate data-group-key="${escapeHtml(group.key)}">
                  Admin override
                </label>
              `
      : ""
    }
          <p class="field-note">
            ${escapeHtml(`This applies to ${group.orders.length} visible active entr${group.orders.length === 1 ? "y" : "ies"} at this pickup location.`)}
          </p>
        </div>
        <div class="action-row stop-actions">
          ${navigationUrl
      ? `
                <a
                  class="button button-secondary"
                  href="${escapeHtml(navigationUrl)}"
                  target="_blank"
                  rel="noreferrer noopener"
                >
                  Navigate
                </a>
              `
      : ""
    }
          <button
            class="button button-primary"
            data-action="save-location-assignment"
            data-group-key="${escapeHtml(group.key)}"
            ${state.busy ? " disabled" : ""}
          >
            Save location
          </button>
          <button
            type="button"
            class="button button-ghost"
            data-action="toggle-stop-card"
            data-stop-card-key="${escapeHtml(stopCardKey)}"
            ${state.busy ? " disabled" : ""}
          >
            ${isOpen ? "Hide entries" : "Show entries"}
          </button>
        </div>
      </div>
      ${isOpen
      ? `
            <div class="location-group-orders">
              ${group.orders.map((order) => renderGlobalOrderCard(order, viewerRole)).join("")}
            </div>
          `
      : ""
    }
    </article>
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

function getVisibleOrderReferenceLines(order) {
  const invoiceNumber = String(order?.invoiceNumber || "").trim();
  const poNumber = String(order?.poNumber || "").trim();
  const lines = [];
  const salesOrderNumber = getOrderSalesOrderNumber(order);

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

function getStockItemDisplayLines(value) {
  const text = String(value || "").replace(/\r/g, "\n").trim();
  if (!text) {
    return [];
  }

  return text
    .split(/\n+/)
    .flatMap((line) => splitStockItemDisplayLine(line))
    .map((line) => line.trim())
    .filter(Boolean);
}

function splitStockItemDisplayLine(value) {
  const line = String(value || "").replace(/\s+/g, " ").trim();
  if (!line) {
    return [];
  }

  const parts = line
    .split(/\s+(?=\d+\s*x\b)/i)
    .map((part) => part.trim())
    .filter(Boolean);

  return parts.length > 1 ? parts : [line];
}

function renderStockItemDisplayName(value, options = {}) {
  const {
    containerClass = "",
    lineClass = "",
    lineTag = "strong",
    emptyLabel = "Unnamed stock item",
  } = options;
  const displayLines = getStockItemDisplayLines(value);
  const safeTag = lineTag === "h4" ? "h4" : "strong";
  const combinedLineClass = ["stock-item-name-line", lineClass].filter(Boolean).join(" ");
  const singleLine = displayLines[0] || emptyLabel;

  if (displayLines.length <= 1) {
    return `<${safeTag}${combinedLineClass ? ` class="${combinedLineClass}"` : ""}>${escapeHtml(singleLine)}</${safeTag}>`;
  }

  return `
    <div class="${["stock-item-name-list", containerClass].filter(Boolean).join(" ")}">
      ${displayLines.map((line) => `<${safeTag}${combinedLineClass ? ` class="${combinedLineClass}"` : ""}>${escapeHtml(line)}</${safeTag}>`).join("")}
    </div>
  `;
}

function getOrderPrimaryDisplay(order) {
  const quoteNumber = getOrderQuoteNumber(order);
  if (quoteNumber) {
    return `Quote ${quoteNumber}`;
  }

  const orderNumber = String(order?.orderNumber || "").trim();
  return orderNumber ? `Entry ${orderNumber}` : "Entry";
}

function getOrderListReferenceLines(order) {
  return getVisibleOrderReferenceLines(order);
}

function renderOrderListReferenceSummary(order, emptyLabel = "No order references") {
  const lines = getOrderListReferenceLines(order);
  if (!lines.length) {
    return `<span class="muted">${escapeHtml(emptyLabel)}</span>`;
  }

  return lines.map((line) => `<span class="muted">${escapeHtml(line)}</span>`).join("<br>");
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
  const allowQr = viewerRole === "admin" || viewerRole === "logistics";
  const isEditing = state.editingStockItemId === item.id;
  const activityBadge = renderRecentStockItemBadge(latestMovementByItemId.get(item.id));
  const actions = [];

  if (allowDelete) {
    actions.push(
      `<button class="button button-secondary" data-action="edit-stock-item" data-stock-item-id="${item.id}"${state.busy || isEditing ? " disabled" : ""}>${isEditing ? "Editing" : "Edit"}</button>`,
    );
  }

  if (allowQr) {
    actions.push(
      `<button class="button button-secondary" data-action="open-stock-qr" data-stock-item-id="${item.id}"${state.busy ? " disabled" : ""}>QR</button>`,
    );
  }

  if (allowDelete) {
    actions.push(
      `<button class="button button-danger" data-action="delete-stock-item" data-stock-item-id="${item.id}"${state.busy ? " disabled" : ""}>Delete</button>`,
    );
  }

  return `
    <tr>
      <td data-label="Item">
        ${renderStockItemDisplayName(item.name)}
        ${activityBadge}
      </td>
      <td data-label="References">${renderReferenceSummary(item)}</td>
      <td data-label="Stock code">${escapeHtml(item.sku || "Not set")}</td>
      <td data-label="Unit">${escapeHtml(String(item.onHandQuantity || 0))}</td>
      <td data-label="Notes">${escapeHtml(item.notes || "None")}</td>
      <td data-label="Updated">${escapeHtml(formatDateTime(item.updatedAt || item.createdAt) || "Not updated")}</td>
      <td data-label="Actions">
        ${actions.length
      ? `
              <div class="action-row">
                ${actions.join("")}
              </div>
            `
      : '<span class="muted">View only</span>'
    }
      </td>
    </tr>
  `;
}

function renderStockItemCard(item, viewerRole, latestMovementByItemId = getLatestStockMovementByItemId()) {
  const allowDelete = viewerRole === "admin";
  const allowQr = viewerRole === "admin" || viewerRole === "logistics";
  const isOpen = Boolean(state.stockOpenItemCards[item.id]);
  const isEditing = state.editingStockItemId === item.id;
  const referenceSummary = getReferenceLines(item).join(" | ") || "No order references set.";
  const activityBadge = renderRecentStockItemBadge(latestMovementByItemId.get(item.id));
  const unitLabel = getStockUnitLabel(item);
  const actions = [];

  if (allowDelete) {
    actions.push(
      `<button class="button button-secondary" data-action="edit-stock-item" data-stock-item-id="${item.id}"${state.busy || isEditing ? " disabled" : ""}>${isEditing ? "Editing" : "Edit"}</button>`,
    );
  }

  if (allowQr) {
    actions.push(
      `<button class="button button-secondary" data-action="open-stock-qr" data-stock-item-id="${item.id}"${state.busy ? " disabled" : ""}>QR</button>`,
    );
  }

  if (allowDelete) {
    actions.push(
      `<button class="button button-danger" data-action="delete-stock-item" data-stock-item-id="${item.id}"${state.busy ? " disabled" : ""}>Delete</button>`,
    );
  }

  return `
    <article class="stock-item-mobile-card${isOpen ? " is-open" : ""}">
      <div class="stock-item-mobile-head">
        <div class="stock-item-mobile-copy">
          ${renderStockItemDisplayName(item.name, {
    containerClass: "stock-item-mobile-title-list",
    lineClass: "stock-item-mobile-title",
    lineTag: "h4",
  })}
          <p class="stock-item-mobile-summary">${escapeHtml(referenceSummary)}</p>
          ${activityBadge}
        </div>
        <div class="stock-item-mobile-side">
          <span class="stock-item-mobile-onhand">${escapeHtml(String(item.onHandQuantity || 0))} ${escapeHtml(unitLabel)}</span>
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
      ${isOpen
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
                ${actions.length ? actions.join("") : '<span class="muted">View only</span>'}
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
  const unitLabel = getStockUnitLabel(movement);

  return `
    <tr>
      <td data-label="Logged">${escapeHtml(formatDateTime(movement.createdAt) || "")}</td>
      <td data-label="Item">
        <strong>${escapeHtml(movement.itemName || "Unknown")}</strong><br>
        <span class="muted">${escapeHtml(movement.sku || "No stock code")}</span><br>
        ${renderReferenceSummary(movement)}
      </td>
      <td data-label="Type">${movement.movementType === "in" ? '<span class="chip chip-success">Stock in</span>' : '<span class="chip chip-warning">Stock out</span>'}</td>
      <td data-label="Quantity">${escapeHtml(String(movement.quantity || 0))} ${escapeHtml(unitLabel)}</td>
      <td data-label="Supplier / driver">${escapeHtml(getStockMovementPartyLabel(movement))}</td>
      <td data-label="Notes">${escapeHtml(movement.notes || "None")}</td>
      <td data-label="Recorded by">${escapeHtml(movement.createdByName || "Unknown")}</td>
      <td data-label="Actions">
        ${canEdit
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

function renderStopCard(stop, index, viewerRole, driverUserId = "") {
  const allowComplete = viewerRole === "admin" || viewerRole === "driver";
  const allowDelete = viewerRole === "admin";
  const allowFlag = viewerRole === "admin" || viewerRole === "driver";
  const allowPriority = viewerRole === "admin";
  const allowTransfer = viewerRole === "admin" || viewerRole === "driver";
  const allowNavigate = viewerRole === "admin" || viewerRole === "driver" || viewerRole === "sales";
  const navigationUrl = allowNavigate ? getGoogleMapsNavigateUrl(stop.location) : "";
  const stopCardKey = buildStopCardKey(driverUserId, stop.id);
  const isOpen = isStopCardOpen(stopCardKey);
  const legLabel = stop.hasCoordinates && stop.legKm !== null
    ? `${stop.legKm.toFixed(1)} km leg`
    : "Coordinates pending";

  return `
    <article class="stop-card${stop.isPriority ? " stop-card-priority" : ""}${isOpen ? " is-open" : " is-collapsed"}">
      <div class="stop-header">
        <div>
          <p class="eyebrow">Stop ${index + 1}</p>
          <h4 class="stop-title">${escapeHtml(stop.location.name)}</h4>
          ${isOpen ? `<p class="stop-address">${escapeHtml(stop.location.address)}</p>` : ""}
        </div>
        <div class="chip-row">
          ${stop.isPriority ? '<span class="chip chip-priority-high">Priority stop</span>' : ""}
          <span class="chip">${legLabel}</span>
          <span class="chip">${stop.orders.length} order${stop.orders.length === 1 ? "" : "s"}</span>
        </div>
      </div>
      <div class="action-row stop-actions">
        ${navigationUrl
      ? `
              <a
                class="button button-primary"
                href="${escapeHtml(navigationUrl)}"
                target="_blank"
                rel="noreferrer noopener"
              >
                Navigate
              </a>
            `
      : ""
    }
        <button
          type="button"
          class="button button-ghost"
          data-action="toggle-stop-card"
          data-stop-card-key="${escapeHtml(stopCardKey)}"
          ${state.busy ? " disabled" : ""}
        >
          ${isOpen ? "Hide orders" : "Show orders"}
        </button>
      </div>
      ${isOpen
      ? `
            <div class="stop-orders">
        ${stop.orders
        .map((order) => {
          const isFlagging = state.flaggingOrderId === order.id;
          const isTransferring = state.transferringOrderId === order.id;
          const canFlag = allowFlag && order.status === "active";
          const canTransfer = allowTransfer && order.status === "active" && getTransferDriverChoices(order).length > 0;
          const isPriority = isPriorityOrder(order);
          const pickedUp = isOrderPickedUp(order);
          const canMarkPickedUp = allowComplete && order.status === "active" && !pickedUp;
          const canDropAtOffice = allowComplete && order.status === "active" && pickedUp;
          const canDropAtFactory = allowComplete && order.status === "active" && pickedUp && order.moveToFactory;
          const canEdit = (viewerRole === "admin" || viewerRole === "sales") && order.status === "active";

          return `
              <div class="order-card${isPriority ? " order-card-priority" : ""}">
                <strong>${escapeHtml(getOrderPrimaryDisplay(order))}</strong>
                <div class="order-meta">
                  ${getOrderListReferenceLines(order).map((line) => `<span>${escapeHtml(line)}</span>`).join("")}
                </div>
                ${renderOrderStockDetails(order)}
                <div class="chip-row">
                  ${renderTypeChip(order.entryType)}
                  ${renderOrderScheduledChip(order)}
                  ${renderOrderPriorityChip(order)}
                  ${renderOrderPickupChip(order)}
                  ${order.moveToFactory ? '<span class="chip chip-warning">Factory move</span>' : ""}
                  ${renderOrderFlagChip(order)}
                  <span class="chip">Created by ${escapeHtml(order.createdByName)}</span>
                  ${renderStatusChip(order.status)}
                </div>
                ${renderOrderNotice(order)}
                <p>${escapeHtml(order.locationName || stop.location.name)}<br>${escapeHtml(order.locationAddress || stop.location.address)}</p>
                ${allowComplete || allowDelete || canFlag || allowPriority
              ? `
                      <div class="action-row">
                        ${canEdit
                ? `
                              <button
                                class="button button-secondary"
                                data-action="edit-order"
                                data-order-id="${order.id}"
                                ${state.busy ? " disabled" : ""}
                              >
                                Edit
                              </button>
                            `
                : ""
              }
                        ${allowPriority
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
                        ${canMarkPickedUp
                ? `
                              <button
                                class="button button-primary"
                                data-action="pick-up-order"
                                data-order-id="${order.id}"
                                ${state.busy ? " disabled" : ""}
                              >
                                Picked up
                              </button>
                            `
                : ""
              }
                        ${canTransfer
                ? `
                              <button
                                class="button button-secondary"
                                data-action="toggle-transfer-order"
                                data-order-id="${order.id}"
                                ${state.busy ? " disabled" : ""}
                              >
                                ${isTransferring ? "Cancel transfer" : "Transfer"}
                              </button>
                            `
                : ""
              }
                        ${canDropAtOffice
                ? `
                              <button
                                class="button button-primary"
                                data-action="complete-order"
                                data-order-id="${order.id}"
                                data-completion-type="office"
                                ${state.busy ? " disabled" : ""}
                              >
                                ${escapeHtml(getOrderCompletionTypeLabel("office", order) || "Dropped at office")}
                              </button>
                            `
                : ""
              }
                        ${canDropAtFactory
                ? `
                              <button
                                class="button button-secondary"
                                data-action="complete-order"
                                data-order-id="${order.id}"
                                data-completion-type="factory"
                                ${state.busy ? " disabled" : ""}
                              >
                                Dropped at factory
                              </button>
                            `
                : ""
              }
                        ${canFlag
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
                        ${allowDelete
                ? `
                              <button class="button button-danger" data-action="delete-order" data-order-id="${order.id}" data-order-reference="${escapeHtml(getOrderPrimaryDisplay(order))}"${state.busy ? " disabled" : ""}>
                                Delete
                              </button>
                            `
                : ""
              }
                      </div>
                    `
              : ""
            }
                ${isTransferring ? renderOrderTransferForm(order) : ""}
                ${isFlagging ? renderOrderFlagForm(order) : ""}
              </div>
            `;
        })
        .join("")}
            </div>
          `
      : ""
    }
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

function hasCarryOver(order) {
  return Boolean(order?.driverUserId) && Number(order?.carryOverCount || 0) > 0;
}

function getActivePriorityOrders() {
  return getActiveOrders().filter((order) => isPriorityOrder(order));
}

function getActiveCarryOverOrders() {
  return getActiveOrders().filter((order) => hasCarryOver(order));
}

function renderOrderPriorityChip(order) {
  return isPriorityOrder(order) ? '<span class="chip chip-priority-high">Priority stop</span>' : "";
}

function isOrderPickedUp(order) {
  return Boolean(String(order?.pickedUpAt || "").trim());
}

function renderOrderPickupChip(order) {
  return isOrderPickedUp(order) ? '<span class="chip chip-success">Picked up</span>' : "";
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
        ${hasFlag
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

function renderOrderTransferForm(order) {
  const transferDrivers = getTransferDriverChoices(order);
  if (!transferDrivers.length) {
    return "";
  }

  const currentDriverName = getDriverDisplayName(order);
  const transferNote = state.snapshot.user?.role === "driver"
    ? `Move this active entry from ${currentDriverName} to another active driver. The admin email inbox will be notified when you confirm the transfer.`
    : `Move this active entry from ${currentDriverName} to another active driver.`;

  return `
    <div class="order-flag-form">
      <label>
        Transfer to driver
        <select data-transfer-driver data-order-id="${order.id}">
          ${transferDrivers.map((driver, index) => `
            <option value="${driver.id}"${index === 0 ? " selected" : ""}>${escapeHtml(driver.name)}</option>
          `).join("")}
        </select>
      </label>
      <p class="field-note">
        ${escapeHtml(transferNote)}
      </p>
      <div class="action-row">
        <button
          type="button"
          class="button button-primary"
          data-action="save-transfer-order"
          data-order-id="${order.id}"
          ${state.busy ? " disabled" : ""}
        >
          Transfer item
        </button>
        <button
          type="button"
          class="button button-ghost"
          data-action="cancel-transfer-order"
          ${state.busy ? " disabled" : ""}
        >
          Cancel
        </button>
      </div>
    </div>
  `;
}

function renderOrderFlagTypeOptions(selectedFlagType = "") {
  return Object.entries(ORDER_FLAG_LABELS)
    .map(
      ([value, label]) => `<option value="${value}"${value === selectedFlagType ? " selected" : ""}>${escapeHtml(label)}</option>`,
    )
    .join("");
}

function isDeliveryOrder(order) {
  return String(order?.entryType || "").trim().toLowerCase() === "delivery";
}

function getOrderCompletionTypeLabel(completionType, order = null) {
  const normalized = String(completionType || "").trim();
  if (normalized === "office" && isDeliveryOrder(order)) {
    return "Delivered to client";
  }

  return ORDER_COMPLETION_LABELS[normalized] || "";
}

function getOrderCompletionLabel(order) {
  return getOrderCompletionTypeLabel(order?.completionType, order);
}

function getOrderCompletionNoticeText(order) {
  const label = getOrderCompletionLabel(order);
  if (!label) {
    return "";
  }

  const completedAt = formatDateTime(order?.completedAt);
  const completedBy = String(order?.completedByName || "").trim();
  const parts = [`Completion: ${label}.`];

  if (completedAt || completedBy) {
    parts.push(`Logged ${[completedAt, completedBy ? `by ${completedBy}` : ""].filter(Boolean).join(" ")}.`);
  }

  return parts.join(" ");
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

function renderAssignmentLocationDriverOptions(selectedDriverId = "", hasMixedSelection = false) {
  const drivers = getActiveDriverUsers();
  const selectedDriver = selectedDriverId ? getUser(selectedDriverId) : null;
  const options = [];

  if (hasMixedSelection) {
    options.push('<option value="" selected disabled>Choose a driver</option>');
  }

  options.push(`<option value="__unassigned__"${!hasMixedSelection && !selectedDriverId ? " selected" : ""}>Unassigned</option>`);

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
    return options.join("");
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
      labelParts.push(...getReferenceLines(item));
      if (item.sku) {
        labelParts.push(item.sku);
      }
      const label = labelParts.join(" | ");
      return `<option value="${item.id}"${item.id === selectedStockItemId ? " selected" : ""}>${escapeHtml(label)}</option>`;
    })
    .join("");
}

function isSystemOfficeLocation(location) {
  return String(location?.name || "").trim().toLowerCase() === "office";
}

function renderLocationOptions(selectedLocationId = "") {
  if (!state.snapshot.locations.length) {
    return '<option value="">Create a location first</option>';
  }

  const sortedLocations = [...state.snapshot.locations].sort((left, right) => {
    const leftOfficeRank = isSystemOfficeLocation(left) ? 0 : 1;
    const rightOfficeRank = isSystemOfficeLocation(right) ? 0 : 1;
    if (leftOfficeRank !== rightOfficeRank) {
      return leftOfficeRank - rightOfficeRank;
    }

    return String(left?.name || "").localeCompare(String(right?.name || ""), undefined, {
      sensitivity: "base",
    });
  });

  return [
    `<option value=""${selectedLocationId ? "" : " selected"}>Select a location</option>`,
    ...sortedLocations.map((location) => {
      const typeLabel = !isSystemOfficeLocation(location) && location.locationType
        ? ` - ${escapeHtml(capitalize(location.locationType))}`
        : "";
      return `<option value="${location.id}"${location.id === selectedLocationId ? " selected" : ""}>${escapeHtml(location.name)}${typeLabel}</option>`;
    }),
  ].join("");
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

function getDriversWithRecordedPositions() {
  return getDriverUsers()
    .map((driver) => {
      const rawLat = driver.lastKnownLat;
      const rawLng = driver.lastKnownLng;
      return {
        ...driver,
        lat: rawLat === null || rawLat === "" ? Number.NaN : Number(rawLat),
        lng: rawLng === null || rawLng === "" ? Number.NaN : Number(rawLng),
      };
    })
    .filter((driver) => Number.isFinite(driver.lat) && Number.isFinite(driver.lng));
}

function hasDriverLocationTrackingFields() {
  const drivers = getDriverUsers();
  if (!drivers.length) {
    return true;
  }

  return drivers.some((driver) => (
    Object.prototype.hasOwnProperty.call(driver, "lastKnownLat")
    || Object.prototype.hasOwnProperty.call(driver, "lastKnownLng")
    || Object.prototype.hasOwnProperty.call(driver, "lastKnownRecordedAt")
  ));
}

function getTransferDriverChoices(order) {
  const currentDriverId = String(order?.driverUserId || "").trim();
  return state.snapshot.users.filter((user) => (
    user.role === "driver"
    && user.active
    && user.id !== currentDriverId
  ));
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

function getEditingOrder() {
  return state.editingOrderId ? getOrder(state.editingOrderId) : null;
}

function getOrdersForDriver(driverUserId) {
  return state.snapshot.orders.filter((order) => order.driverUserId === driverUserId);
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

function getLiveScheduleDate() {
  return String(state.publicState.today || "").trim();
}

function getDefaultScheduledDateValue() {
  return getLiveScheduleDate() || getCurrentLocalDateValue();
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

function getFilteredAssignmentLocationGroups(orders = getFilteredAssignmentOrders()) {
  const grouped = new Map();

  orders.forEach((order) => {
    const locationId = String(order.locationId || "").trim();
    const fallbackKey = [order.locationName, order.locationAddress]
      .map((value) => String(value || "").trim().toLowerCase())
      .filter(Boolean)
      .join("|");
    const groupKey = locationId || fallbackKey || `order:${order.id}`;

    if (!grouped.has(groupKey)) {
      grouped.set(groupKey, {
        key: groupKey,
        locationId,
        locationName: order.locationName || "Unknown",
        locationAddress: order.locationAddress || "",
        orders: [],
        unassignedCount: 0,
        assignedCount: 0,
        priorityCount: 0,
        driverCounts: new Map(),
      });
    }

    const group = grouped.get(groupKey);
    group.orders.push(order);

    if (order.driverUserId) {
      group.assignedCount += 1;
      group.driverCounts.set(order.driverUserId, (group.driverCounts.get(order.driverUserId) || 0) + 1);
    } else {
      group.unassignedCount += 1;
    }

    if (isPriorityOrder(order)) {
      group.priorityCount += 1;
    }
  });

  return Array.from(grouped.values())
    .map((group) => {
      const driverSummaries = Array.from(group.driverCounts.entries())
        .map(([driverId, count]) => ({
          id: driverId,
          name: getUser(driverId)?.name || "Unknown driver",
          count,
        }))
        .sort((left, right) => left.name.localeCompare(right.name, "en-ZA"));
      const hasMixedDriverSelection = driverSummaries.length > 1 || (driverSummaries.length > 0 && group.unassignedCount > 0);
      const selectedDriverId = hasMixedDriverSelection
        ? ""
        : (driverSummaries[0]?.id || "");

      return {
        ...group,
        orders: [...group.orders].sort(orderAssignmentSort),
        driverSummaries,
        hasMixedDriverSelection,
        selectedDriverId,
      };
    })
    .sort((left, right) => {
      const leftUnassignedRank = left.unassignedCount > 0 ? 0 : 1;
      const rightUnassignedRank = right.unassignedCount > 0 ? 0 : 1;
      if (leftUnassignedRank !== rightUnassignedRank) {
        return leftUnassignedRank - rightUnassignedRank;
      }

      const nameCompare = String(left.locationName || "").localeCompare(String(right.locationName || ""), "en-ZA", {
        sensitivity: "base",
      });
      if (nameCompare) {
        return nameCompare;
      }

      return String(left.locationAddress || "").localeCompare(String(right.locationAddress || ""), "en-ZA", {
        sensitivity: "base",
      });
    });
}

function getAssignmentLocationDriverLine(group) {
  const parts = group.driverSummaries.map((driver) => `${driver.name} (${driver.count})`);
  if (group.unassignedCount) {
    parts.push(`Unassigned (${group.unassignedCount})`);
  }

  if (!parts.length) {
    return "Current driver: Unassigned.";
  }

  return `Current split: ${parts.join(", ")}.`;
}

function getAssignmentFilterSummary(filteredGroupCount, filteredCount) {
  const filterValue = String(state.assignmentDriverFilter || "").trim();
  const groupLabel = `${filteredGroupCount} pickup location${filteredGroupCount === 1 ? "" : "s"}`;

  if (!filterValue) {
    return `Showing ${filteredCount} active entr${filteredCount === 1 ? "y" : "ies"} across ${groupLabel}. Unassigned work still appears first.`;
  }

  if (filterValue === "unassigned") {
    return filteredCount
      ? `Showing ${filteredCount} unassigned active entr${filteredCount === 1 ? "y" : "ies"} across ${groupLabel}.`
      : "No unassigned active entries right now.";
  }

  const driver = getUser(filterValue);
  const driverName = driver?.name || "that driver";
  return filteredCount
    ? `Showing ${filteredCount} active entr${filteredCount === 1 ? "y" : "ies"} for ${driverName} across ${groupLabel}.`
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

async function recordDriverPosition(lat, lng) {
  if (!sessionToken) {
    return;
  }

  try {
    await callRpc("record_driver_position", {
      p_token: sessionToken,
      p_lat: lat,
      p_lng: lng,
    });
  } catch (error) {
    const message = normalizeError(error);
    state.driverRouteOriginError = message.includes("latest neon.sql")
      ? `${message} Driver tracking will start working after that update is applied.`
      : "Your location was found, but it could not be saved for admin tracking right now.";
    render();
    console.error("Failed to record driver position.", error);
  }
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
      void recordDriverPosition(position.coords.latitude, position.coords.longitude);
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

function buildStopCardKey(driverUserId, stopId) {
  return [String(driverUserId || "").trim(), String(stopId || "").trim()].filter(Boolean).join(":");
}

function buildGlobalLocationGroupKey(groupKey) {
  return `global-location:${String(groupKey || "").trim()}`;
}

function buildAssignmentLocationGroupKey(groupKey) {
  return `assignment-location:${String(groupKey || "").trim()}`;
}

function isStopCardOpen(stopCardKey) {
  if (!stopCardKey) {
    return true;
  }

  if (stopCardKey.startsWith("global-location:")) {
    return state.driverOpenStops[stopCardKey] === true;
  }

  return state.driverOpenStops[stopCardKey] !== false;
}

function getRoutePlan(driverUserId, options = {}) {
  const scheduledFor = String(options?.scheduledFor || "").trim();
  const activeOrders = getOrdersForDriver(driverUserId)
    .filter((order) => order.status === "active")
    .filter((order) => !scheduledFor || getOrderScheduledForValue(order) === scheduledFor)
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

function drawAdminDriverLocationMap() {
  const Leaflet = window.L;
  const mapEl = document.getElementById("admin-driver-location-map");
  if (!(mapEl instanceof HTMLElement)) {
    return;
  }

  if (!Leaflet) {
    setAdminDriverLocationStatus("The live map is unavailable right now, but driver positions will still appear here once mapping loads.");
    return;
  }

  if (adminDriverMap && adminDriverMapContainer !== mapEl) {
    destroyAdminDriverLocationMap();
  }

  if (!adminDriverMap) {
    adminDriverMap = Leaflet.map(mapEl, {
      zoomControl: true,
      scrollWheelZoom: false,
    });
    adminDriverMapContainer = mapEl;
    Leaflet.tileLayer("https://tile.openstreetmap.org/{z}/{x}/{y}.png", {
      maxZoom: 19,
      attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',
    }).addTo(adminDriverMap);
    adminDriverMapLayers = Leaflet.layerGroup().addTo(adminDriverMap);
  } else if (!adminDriverMapLayers) {
    adminDriverMapLayers = Leaflet.layerGroup().addTo(adminDriverMap);
    adminDriverMapContainer = mapEl;
  }

  adminDriverMapLayers.clearLayers();

  const positionedDrivers = getDriversWithRecordedPositions();
  const bounds = [];

  positionedDrivers.forEach((driver) => {
    bounds.push([driver.lat, driver.lng]);
    Leaflet
      .marker([driver.lat, driver.lng], {
        icon: buildAdminDriverLocationMarkerIcon(getDriverInitials(driver.name)),
      })
      .bindPopup(buildAdminDriverLocationPopup(driver))
      .addTo(adminDriverMapLayers);
  });

  if (bounds.length) {
    adminDriverMap.fitBounds(bounds, { padding: [36, 36] });
  } else {
    adminDriverMap.setView([HUB.lat, HUB.lng], 8);
  }

  setAdminDriverLocationStatus(getAdminDriverLocationStatus(positionedDrivers));
  window.requestAnimationFrame(() => {
    adminDriverMap?.invalidateSize();
  });
}

function destroyAdminDriverLocationMap() {
  if (adminDriverMap) {
    adminDriverMap.remove();
  }
  adminDriverMap = null;
  adminDriverMapContainer = null;
  adminDriverMapLayers = null;
}

function setAdminDriverLocationStatus(message) {
  const statusEl = document.getElementById("admin-driver-location-status");
  if (!(statusEl instanceof HTMLElement)) {
    return;
  }

  statusEl.textContent = message || "";
  statusEl.classList.toggle("hidden", !message);
}

function getAdminDriverLocationStatus(positionedDrivers = getDriversWithRecordedPositions()) {
  const totalDrivers = getDriverUsers().length;
  if (!totalDrivers) {
    return "No driver accounts are available yet.";
  }

  if (!hasDriverLocationTrackingFields()) {
    return "Driver location tracking is not available in this database yet. Apply the latest neon.sql update, then have drivers open their route page and allow location access.";
  }

  if (!positionedDrivers.length) {
    return "No driver positions have been recorded yet. Drivers need to open their route page and allow location access first.";
  }

  const awaitingDrivers = Math.max(totalDrivers - positionedDrivers.length, 0);
  if (!awaitingDrivers) {
    return `Showing the last recorded position for all ${positionedDrivers.length} driver${positionedDrivers.length === 1 ? "" : "s"}.`;
  }

  return `Showing ${positionedDrivers.length} recorded driver location${positionedDrivers.length === 1 ? "" : "s"}. ${awaitingDrivers} driver${awaitingDrivers === 1 ? " is" : "s are"} still awaiting a recorded position.`;
}

function getDriverInitials(name) {
  const initials = String(name || "")
    .trim()
    .split(/\s+/)
    .filter(Boolean)
    .slice(0, 2)
    .map((part) => part.charAt(0).toUpperCase())
    .join("");

  return initials || "D";
}

function buildAdminDriverLocationMarkerIcon(label) {
  const Leaflet = window.L;
  return Leaflet.divIcon({
    className: "route-map-marker-shell",
    html: `<span class="route-map-marker route-map-marker-driver">${escapeHtml(label)}</span>`,
    iconSize: [40, 40],
    iconAnchor: [20, 20],
    popupAnchor: [0, -18],
  });
}

function buildAdminDriverLocationPopup(driver) {
  const activeOrders = getOrdersForDriver(driver.id).filter((order) => order.status === "active");
  const activeStops = new Set(activeOrders.map((order) => order.locationId)).size;
  const recordedAt = formatDateTime(driver.lastKnownRecordedAt) || "Recently";

  return `
    <div class="route-map-popup">
      <p class="eyebrow">Driver</p>
      <h4>${escapeHtml(driver.name || "Unknown")}</h4>
      <p>${escapeHtml(driver.phone || "No phone assigned")}</p>
      <p>${escapeHtml(`${activeOrders.length} active entr${activeOrders.length === 1 ? "y" : "ies"} across ${activeStops} stop${activeStops === 1 ? "" : "s"}`)}</p>
      <p class="muted">Recorded ${escapeHtml(recordedAt)}</p>
    </div>
  `;
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
    return ["Route starts from the driver's current location.", state.driverRouteOriginError || ""].filter(Boolean).join(" ");
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
    .map((order) => getOrderPrimaryDisplay(order))
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

function getOrderCsvReferenceSummary(order) {
  return getVisibleOrderReferenceLines(order).join(" | ");
}

function getOrderCsvDriverIssue(order) {
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
  const currentDriverName = getDriverDisplayName(order);

  if (!pickedUpByUserId || !currentDriverUserId || pickedUpByUserId === currentDriverUserId) {
    return "";
  }

  return [pickedUpByName || "Picked-up driver", currentDriverName || "Assigned driver"].join(" -> ");
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

function shouldMarkOrderAsRolledOver(order) {
  return Boolean(order?.driverUserId) && Number(order?.carryOverCount || 0) > 0;
}

function getOrderCsvScheduleSummary(order) {
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
    getDriverDisplayName(order),
    getOrderPickupCsvValue(order),
    getOrderDriverHandoffCsvValue(order),
    order.locationName || "",
    order.locationAddress || "",
    order.deliveryAddress || "",
    capitalize(order.entryType || ""),
    capitalize(getOrderPriority(order)),
    getOrderCsvReferenceSummary(order),
    order.stockDescription || "",
    order.branding || "",
    order.moveToFactory ? "Yes" : "No",
    getCollectionDestinationLabel(order),
    getOrderCsvDriverIssue(order),
    order.notes || "",
    getOrderCsvScheduleSummary(order),
    capitalize(order.status || "active"),
    getOrderCompletionLabel(order),
    order.createdByName || "",
    formatDateTime(order.createdAt),
    formatDateTime(order.completedAt),
  ];
}

function buildOrdersCsvContent(orders) {
  const lineBreak = "\r\n";
  const rows = [
    [
      "Quote",
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

function exportOrdersCsv() {
  const dateStamp = getLiveScheduleDate() || getCurrentLocalDateValue();
  const sortedOrders = [...filterOrdersForScheduledDate(state.snapshot.orders, dateStamp)].sort(orderDisplaySort);
  if (!sortedOrders.length) {
    showFlash(`There are no entries scheduled for ${formatDateOnly(dateStamp) || "the live date"} yet.`, "error");
    return;
  }

  const csv = buildOrdersCsvContent(sortedOrders);
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");

  anchor.href = url;
  anchor.download = `route-ledger-${dateStamp}.csv`;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
  showFlash(`CSV export downloaded for ${formatDateOnly(dateStamp) || dateStamp}.`, "success");
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

async function sendRolloverTestEmail() {
  if (!state.mailConfigured) {
    showFlash(state.mailConfigReason || "Email delivery is not configured yet.", "error");
    return;
  }

  state.busy = true;
  render();

  try {
    const payload = await requestJson(`${API_ROOT}/export/email/rollover-test`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ token: sessionToken }),
    });

    const sentTo = payload?.sentTo || state.artworkTo;
    const count = Number(payload?.count || 0);
    const csvCount = Number(payload?.csvCount || 0);
    const itemLabel = count === 1 ? "item" : "items";
    const entryLabel = csvCount === 1 ? "entry" : "entries";
    showFlash(`Rollover test email sent to ${sentTo} with ${count} rolled ${itemLabel} and ${csvCount} CSV ${entryLabel}.`, "success");
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

function getOrderStockDetailItems(order) {
  const items = [];
  const stockDescription = String(order?.stockDescription || "").trim();
  const branding = String(order?.branding || "").trim();
  const deliveryAddress = String(order?.deliveryAddress || "").trim();
  const collectionDestination = getCollectionDestinationLabel(order);

  if (stockDescription) {
    items.push({ label: "Stock", value: stockDescription });
  }

  if (branding) {
    items.push({ label: "Branding", value: branding });
  }

  if (deliveryAddress) {
    items.push({ label: "Delivery address", value: deliveryAddress });
  }

  if (collectionDestination) {
    items.push({ label: "Destination", value: collectionDestination });
  }

  return items;
}

function renderOrderStockDetails(order) {
  const items = getOrderStockDetailItems(order);
  if (!items.length) {
    return "";
  }

  return `
    <div class="order-stock-details">
      ${items.map((item) => `
        <span class="order-stock-detail">
          <span class="order-stock-label">${escapeHtml(item.label)}:</span>
          ${escapeHtml(item.value)}
        </span>
      `).join("")}
    </div>
  `;
}

function getOrderNoticeLines(order) {
  const lines = [];
  const notice = String(order?.notes || "").trim();
  const driverFlag = getOrderFlagNoticeText(order);
  const pickupNotice = getOrderPickupNoticeText(order);
  const moveToFactory = getMoveToFactoryText(order);
  const completion = getOrderCompletionNoticeText(order);
  const rolloverNotice = getRolloverNoticeText(order);

  if (driverFlag) {
    lines.push(driverFlag);
  }

  if (pickupNotice) {
    lines.push(pickupNotice);
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

function getOrderPickupNoticeText(order) {
  if (!isOrderPickedUp(order)) {
    return "";
  }

  const pickedUpAt = formatDateTime(order?.pickedUpAt);
  const pickedUpBy = String(order?.pickedUpByName || "").trim();
  const parts = ["Picked up."];

  if (pickedUpAt || pickedUpBy) {
    parts.push(`Logged ${[pickedUpAt, pickedUpBy ? `by ${pickedUpBy}` : ""].filter(Boolean).join(" ")}.`);
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

function renderOrderScheduledChip(order) {
  const scheduledFor = formatDateOnly(getOrderScheduledForValue(order));
  if (!scheduledFor) {
    return "";
  }

  return `<span class="chip">Scheduled: ${escapeHtml(scheduledFor)}</span>`;
}

function renderDriverAssignmentValue(order) {
  if (order?.driverUserId) {
    return escapeHtml(getDriverDisplayName(order));
  }

  return '<span class="chip chip-warning">Unassigned</span>';
}

function getRolloverNoticeText(order) {
  if (!shouldMarkOrderAsRolledOver(order)) {
    return "";
  }

  const carryOverCount = Number(order?.carryOverCount || 0);
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
  const deliveryAddressField = form.querySelector("[data-delivery-address-field]");
  const deliveryAddressInput = form.querySelector('[name="deliveryAddress"]');
  const deliveryAddressNote = form.querySelector("[data-delivery-address-note]");
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
  const isDelivery = entryTypeField.value === "delivery";
  const hasFactoryOptions = Array.from(destinationSelect.options).some((option) => option.value);

  if (deliveryAddressField instanceof HTMLElement) {
    deliveryAddressField.classList.toggle("hidden", !isDelivery);
  }

  if (deliveryAddressInput instanceof HTMLTextAreaElement || deliveryAddressInput instanceof HTMLInputElement) {
    deliveryAddressInput.required = isDelivery;
  }

  if (deliveryAddressNote instanceof HTMLElement) {
    deliveryAddressNote.classList.toggle("hidden", !isDelivery);
    deliveryAddressNote.textContent = isDelivery
      ? "Add the address the driver must deliver this order to."
      : "Switch the entry type to Delivery to add a destination address.";
  }

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
      ? "Leave this unchecked when the collected stock should return to the office."
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

function formatPreciseDateTime(value) {
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
    second: "2-digit",
    timeZone: TIME_ZONE,
  }).format(date);
}
