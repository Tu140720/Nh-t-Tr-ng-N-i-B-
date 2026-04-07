const sessionUrl = "/api/session";
const loginUrl = "/api/login";
const loginVerifyUrl = "/api/login/verify";
const logoutUrl = "/api/logout";
const usersUrl = "/api/users";
const ordersUrl = "/api/orders";
const notificationsUrl = "/api/notifications";
const notificationsReadUrl = "/api/notifications/read";
const notificationsReadAllUrl = "/api/notifications/read-all";
const deliveryCompleteUrl = "/api/delivery/complete";
const orderCreateUrl = "/api/orders/create";
const tapeCalculatorConfigUrl = "/api/tape-calculator/config";

const AUTH_STORAGE_KEY = "telegramMiniAuthToken";
const NOTIFICATION_ARCHIVE_STORAGE_PREFIX = "telegramMiniNotificationArchive";
const NOTIFICATION_ARCHIVE_MAX_ITEMS = 50;
const NOTIFICATION_POLL_INTERVAL_MS = 30_000;
const DEFAULT_PRODUCTION_UNITS = ["cay", "cuon", "kg", "tam", "bao"];

const state = {
  authToken: localStorage.getItem(AUTH_STORAGE_KEY) || "",
  currentUser: null,
  users: [],
  orders: [],
  notifications: [],
  notificationArchive: [],
  notificationFilter: "all",
  challengeToken: "",
  orderKind: "transport",
  trackingKind: "transport",
  tapeConfig: null,
  tapeLoaded: false,
  activeSection: "",
  createMenuOpen: false,
  productionCreateOpen: false,
};

let notificationPollTimer = null;

const loginPanel = document.querySelector("#login-panel");
const workspacePanel = document.querySelector("#workspace-panel");
const loginForm = document.querySelector("#login-form");
const loginUsername = document.querySelector("#login-username");
const loginPassword = document.querySelector("#login-password");
const loginOtp = document.querySelector("#login-otp");
const otpBlock = document.querySelector("#otp-block");
const otpNote = document.querySelector("#otp-note");
const loginSubmit = document.querySelector("#login-submit");
const logoutButton = document.querySelector("#logout-button");
const heroUserName = document.querySelector("#hero-user-name");
const heroUserMeta = document.querySelector("#hero-user-meta");
const actionButtons = Array.from(document.querySelectorAll("[data-section]"));
const createActionPanel = document.querySelector("#create-action-panel");
const openProductionCreateButton = document.querySelector("#open-production-create-button");
const openTransportCreateButton = document.querySelector("#open-transport-create-button");
const notificationFilter = document.querySelector("#notification-filter");
const notificationList = document.querySelector("#notification-list");
const notificationCount = document.querySelector("#notification-count");
const notificationHistoryList = document.querySelector("#notification-history-list");
const notificationHistoryCount = document.querySelector("#notification-history-count");
const markAllNotificationsReadButton = document.querySelector("#mark-all-notifications-read");
const clearNotificationHistoryButton = document.querySelector("#clear-notification-history");

const deliveryPanel = document.querySelector("#delivery-panel");
const tapePanel = document.querySelector("#tape-panel");
const createPanel = document.querySelector("#create-panel");
const trackingPanel = document.querySelector("#tracking-panel");

const deliveryForm = document.querySelector("#delivery-form");
const deliveryOrderId = document.querySelector("#delivery-order-id");
const deliveryCustomerName = document.querySelector("#delivery-customer-name");
const deliveryAddress = document.querySelector("#delivery-address");
const deliverySalesUser = document.querySelector("#delivery-sales-user");
const deliveryResultStatus = document.querySelector("#delivery-result-status");
const deliveryCompletedAt = document.querySelector("#delivery-completed-at");
const deliveryPaymentStatus = document.querySelector("#delivery-payment-status");
const deliveryPaymentMethod = document.querySelector("#delivery-payment-method");
const deliveryNote = document.querySelector("#delivery-note");
const deliverySubmit = document.querySelector("#delivery-submit");

const tapeType = document.querySelector("#tape-type");
const tapeOrderQuantity = document.querySelector("#tape-order-quantity");
const tapeCoreType = document.querySelector("#tape-core-type");
const tapePackaging = document.querySelector("#tape-packaging");
const tapeJumboHeight = document.querySelector("#tape-jumbo-height");
const tapeCoreWidth = document.querySelector("#tape-core-width");
const tapeFinishedQuantity = document.querySelector("#tape-finished-quantity");
const tapeRollsPerHand = document.querySelector("#tape-rolls-per-hand");
const tapeRemainingMm = document.querySelector("#tape-remaining-mm");
const tapeHandsNeeded = document.querySelector("#tape-hands-needed");
const tapeTotalProduced = document.querySelector("#tape-total-produced");
const tapeExtraProduced = document.querySelector("#tape-extra-produced");
const tapeSourceBadge = document.querySelector("#tape-source-badge");

const createKindButtons = Array.from(document.querySelectorAll("[data-order-kind]"));
const createForm = document.querySelector("#create-form");
const createOrderId = document.querySelector("#create-order-id");
const createCustomerName = document.querySelector("#create-customer-name");
const createSalesUser = document.querySelector("#create-sales-user");
const createDeliveryUser = document.querySelector("#create-delivery-user");
const createPlannedAt = document.querySelector("#create-planned-at");
const createAddress = document.querySelector("#create-address");
const createOrderItems = document.querySelector("#create-order-items");
const createOrderValue = document.querySelector("#create-order-value");
const createNote = document.querySelector("#create-note");
const createSubmit = document.querySelector("#create-submit");
const transportOnlyFields = Array.from(document.querySelectorAll(".transport-only"));

const trackingKindButtons = Array.from(document.querySelectorAll("[data-track-kind]"));
const trackingSearch = document.querySelector("#tracking-search");
const trackingList = document.querySelector("#tracking-list");
const refreshOrdersButton = document.querySelector("#refresh-orders-button");

const productionCreateScreen = document.querySelector("#production-create-screen");
const productionCreateForm = document.querySelector("#production-create-form");
const productionCreateOrderId = document.querySelector("#production-create-order-id");
const productionCreateCustomerName = document.querySelector("#production-create-customer-name");
const productionCreateRequester = document.querySelector("#production-create-requester");
const productionCreateCreatedAt = document.querySelector("#production-create-created-at");
const productionCreateItemsList = document.querySelector("#production-create-items-list");
const productionCreateItemsEmpty = document.querySelector("#production-create-items-empty");
const productionCreateAddItemButton = document.querySelector("#production-create-add-item-button");
const productionCreateNote = document.querySelector("#production-create-note");
const productionCreateTurnaroundOther = document.querySelector("#production-create-turnaround-other");
const productionCreateRequesterSignature = document.querySelector("#production-create-requester-signature");
const productionCreatePerformerSignature = document.querySelector("#production-create-performer-signature");
const productionCreatePackagingSignature = document.querySelector("#production-create-packaging-signature");
const productionCreateReceiptSignature = document.querySelector("#production-create-receipt-signature");
const closeProductionCreateButton = document.querySelector("#close-production-create-button");
const cancelProductionCreateButton = document.querySelector("#cancel-production-create-button");
const submitProductionCreateButton = document.querySelector("#submit-production-create-button");

const toastStack = document.querySelector("#toast-stack");

const TelegramWebApp = window.Telegram?.WebApp || null;

init().catch((error) => {
  console.error(error);
  showToast(error.message || "Khong the khoi dong mini app.", "error");
});

async function init() {
  initTelegramShell();
  bindEvents();
  renderHeroUser();
  if (state.authToken) {
    const restored = await restoreSession();
    if (restored) {
      return;
    }
  }
  showLoginPanel();
}

function initTelegramShell() {
  updateTelegramViewportMetrics();
  window.addEventListener("resize", updateTelegramViewportMetrics);
  window.visualViewport?.addEventListener("resize", updateTelegramViewportMetrics);

  if (!TelegramWebApp) {
    heroUserMeta.textContent = "Dang test tren browser";
    return;
  }

  TelegramWebApp.ready();
  requestTelegramFullscreen();
  if (TelegramWebApp.onEvent) {
    try {
      TelegramWebApp.onEvent("viewportChanged", updateTelegramViewportMetrics);
    } catch {
      void 0;
    }
  }
  if (TelegramWebApp.setHeaderColor) {
    try {
      TelegramWebApp.setHeaderColor("#101b27");
    } catch {
      void 0;
    }
  }
  heroUserMeta.textContent = "Dang mo trong Telegram";
}

function bindEvents() {
  loginForm?.addEventListener("submit", handleLoginSubmit);
  logoutButton?.addEventListener("click", handleLogout);
  notificationFilter?.addEventListener("change", () => {
    state.notificationFilter = notificationFilter.value || "all";
    renderNotifications();
  });
  markAllNotificationsReadButton?.addEventListener("click", handleMarkAllNotificationsRead);
  clearNotificationHistoryButton?.addEventListener("click", handleClearNotificationHistory);
  openTransportCreateButton?.addEventListener("click", () => {
    closeCreateMenu();
    setOrderKind("transport");
    openSection("create");
  });
  openProductionCreateButton?.addEventListener("click", async () => {
    closeCreateMenu();
    await openProductionCreateScreen();
  });
  closeProductionCreateButton?.addEventListener("click", closeProductionCreateScreen);
  cancelProductionCreateButton?.addEventListener("click", closeProductionCreateScreen);
  productionCreateAddItemButton?.addEventListener("click", () => addProductionCreateItemRow());
  productionCreateRequester?.addEventListener("change", syncProductionCreateRequesterSignature);
  productionCreateForm?.addEventListener("submit", handleProductionCreateSubmit);
  actionButtons.forEach((button) => {
    button.addEventListener("click", () => {
      const section = button.dataset.section || "";
      if (!isSectionAllowed(section)) {
        showToast("Ban khong co quyen dung muc nay.", "error");
        return;
      }
      if (section === "create") {
        toggleCreateMenu();
        return;
      }
      closeCreateMenu();
      if (state.activeSection === section) {
        hideAllFeaturePanels();
        return;
      }
      if (TelegramWebApp?.HapticFeedback?.impactOccurred) {
        TelegramWebApp.HapticFeedback.impactOccurred("light");
      }
      openSection(section);
    });
  });

  deliveryOrderId?.addEventListener("change", autofillDeliveryForm);
  deliveryPaymentStatus?.addEventListener("change", syncDeliveryPaymentMethodState);
  deliveryForm?.addEventListener("submit", handleDeliverySubmit);

  [tapeType, tapeOrderQuantity, tapeCoreType, tapePackaging, tapeJumboHeight, tapeCoreWidth].forEach((field) => {
    field?.addEventListener("input", updateTapeResults);
    field?.addEventListener("change", updateTapeResults);
  });

  createKindButtons.forEach((button) => {
    button.addEventListener("click", () => setOrderKind(button.dataset.orderKind || "transport"));
  });
  createForm?.addEventListener("submit", handleCreateSubmit);

  trackingKindButtons.forEach((button) => {
    button.addEventListener("click", () => setTrackingKind(button.dataset.trackKind || "transport"));
  });
  trackingSearch?.addEventListener("input", renderTrackingOrders);
  refreshOrdersButton?.addEventListener("click", async () => {
    refreshOrdersButton.disabled = true;
    try {
      await loadOrders();
      showToast("Da tai lai danh sach don.", "success");
    } catch (error) {
      showToast(error.message || "Khong the tai lai don hang.", "error");
    } finally {
      refreshOrdersButton.disabled = false;
    }
  });
}

async function restoreSession() {
  try {
    const payload = await fetchJson(sessionUrl);
    if (!payload.authenticated || !payload.currentUser) {
      clearAuth();
      showLoginPanel();
      return false;
    }
    state.currentUser = payload.currentUser;
    await bootstrapWorkspace();
    return true;
  } catch {
    clearAuth();
    showLoginPanel();
    return false;
  }
}

async function bootstrapWorkspace() {
  showWorkspacePanel();
  renderHeroUser();
  await Promise.all([loadUsers(), loadOrders()]);
  await loadNotifications({ silent: true });
  applyPermissionState();
  populateCreateUserOptions();
  populateDeliverySalesOptions();
  populateDeliveryOrders();
  setOrderKind(state.orderKind);
  setTrackingKind(state.trackingKind);
  ensureDefaultSection();
  startNotificationPolling();
}

function showLoginPanel() {
  loginPanel?.classList.remove("hidden");
  workspacePanel?.classList.add("hidden");
  hideAllFeaturePanels();
  closeProductionCreateScreen({ silent: true });
  stopNotificationPolling();
  resetLoginState();
  renderHeroUser();
  renderNotifications();
}

function showWorkspacePanel() {
  loginPanel?.classList.add("hidden");
  workspacePanel?.classList.remove("hidden");
  requestTelegramFullscreen();
}

function hideAllFeaturePanels() {
  [deliveryPanel, tapePanel, createPanel, trackingPanel].forEach((panel) => panel?.classList.add("hidden"));
  actionButtons.forEach((button) => button.classList.remove("is-active"));
  closeCreateMenu();
  state.activeSection = "";
}

function ensureDefaultSection() {
  hideAllFeaturePanels();
}

async function openSection(section) {
  hideAllFeaturePanels();
  state.activeSection = section;
  actionButtons.forEach((button) => {
    button.classList.toggle("is-active", button.dataset.section === section);
  });

  if (section === "delivery") {
    deliveryPanel?.classList.remove("hidden");
    await loadOrders();
    populateDeliveryOrders();
    autofillDeliveryForm();
    deliveryPanel?.scrollIntoView({ behavior: "smooth", block: "start" });
    return;
  }

  if (section === "tape") {
    tapePanel?.classList.remove("hidden");
    await loadTapeConfig();
    tapePanel?.scrollIntoView({ behavior: "smooth", block: "start" });
    return;
  }

  if (section === "create") {
    createPanel?.classList.remove("hidden");
    populateCreateUserOptions();
    setOrderKind(state.orderKind);
    createPanel?.scrollIntoView({ behavior: "smooth", block: "start" });
    return;
  }

  if (section === "tracking") {
    trackingPanel?.classList.remove("hidden");
    await loadOrders();
    renderTrackingOrders();
    trackingPanel?.scrollIntoView({ behavior: "smooth", block: "start" });
  }
}

function toggleCreateMenu() {
  if (state.createMenuOpen) {
    closeCreateMenu();
    return;
  }
  state.activeSection = "";
  hideAllFeaturePanels();
  state.createMenuOpen = true;
  createActionPanel?.classList.remove("hidden");
  actionButtons.forEach((button) => {
    button.classList.toggle("is-active", button.dataset.section === "create");
  });
}

function closeCreateMenu() {
  state.createMenuOpen = false;
  createActionPanel?.classList.add("hidden");
  actionButtons.forEach((button) => {
    if (button.dataset.section === "create" && state.activeSection !== "create") {
      button.classList.remove("is-active");
    }
  });
}

async function openProductionCreateScreen() {
  if (!canCreateOrder()) {
    showToast("Ban khong co quyen tao phieu san xuat.", "error");
    return;
  }
  await Promise.all([loadUsers(), loadOrders()]);
  state.productionCreateOpen = true;
  requestTelegramFullscreen();
  setTelegramVerticalSwipesDisabled(true);
  document.body.classList.add("production-create-open");
  productionCreateScreen?.classList.remove("hidden");
  productionCreateScreen?.setAttribute("aria-hidden", "false");
  resetProductionCreateForm();
  if (productionCreateForm) {
    productionCreateForm.scrollTop = 0;
    productionCreateForm.scrollLeft = 0;
  }
  productionCreateCustomerName?.focus();
}

function closeProductionCreateScreen(options = {}) {
  state.productionCreateOpen = false;
  setTelegramVerticalSwipesDisabled(false);
  document.body.classList.remove("production-create-open");
  productionCreateScreen?.classList.add("hidden");
  productionCreateScreen?.setAttribute("aria-hidden", "true");
  if (!options.silent) {
    closeCreateMenu();
  }
}

async function handleLoginSubmit(event) {
  event.preventDefault();
  loginSubmit.disabled = true;

  try {
    if (state.challengeToken) {
      const otpCode = String(loginOtp?.value || "").trim();
      if (!otpCode) {
        throw new Error("Can nhap ma OTP.");
      }
      const payload = await postJson(
        loginVerifyUrl,
        {
          challengeToken: state.challengeToken,
          code: otpCode,
        },
        { skipAuth: true },
      );
      finishLogin(payload.token, payload.currentUser);
      await bootstrapWorkspace();
      showToast("Dang nhap thanh cong.", "success");
      return;
    }

    const username = String(loginUsername?.value || "").trim();
    const password = String(loginPassword?.value || "");
    if (!username || !password) {
      throw new Error("Can nhap tai khoan va mat khau.");
    }

    const payload = await postJson(
      loginUrl,
      {
        username,
        password,
      },
      { skipAuth: true },
    );

    if (payload.otp_required) {
      state.challengeToken = String(payload.challengeToken || "").trim();
      otpBlock?.classList.remove("hidden");
      otpNote.textContent = payload.delivery || "Ma OTP da duoc gui.";
      loginSubmit.textContent = "Xac minh OTP";
      loginOtp?.focus();
      showToast("Nhap ma OTP de hoan tat dang nhap.", "success");
      return;
    }

    finishLogin(payload.token, payload.currentUser);
    await bootstrapWorkspace();
    showToast("Dang nhap thanh cong.", "success");
  } catch (error) {
    showToast(error.message || "Khong the dang nhap.", "error");
  } finally {
    loginSubmit.disabled = false;
  }
}

function finishLogin(token, currentUser) {
  state.authToken = String(token || "").trim();
  state.currentUser = currentUser || null;
  state.challengeToken = "";
  if (state.authToken) {
    localStorage.setItem(AUTH_STORAGE_KEY, state.authToken);
    localStorage.setItem("authToken", state.authToken);
  }
  showWorkspacePanel();
  resetLoginState();
  renderHeroUser();
}

async function handleLogout() {
  logoutButton.disabled = true;
  try {
    if (state.authToken) {
      await postJson(logoutUrl, {});
    }
  } catch {
    void 0;
  } finally {
    logoutButton.disabled = false;
    clearAuth();
    state.currentUser = null;
    state.users = [];
    state.orders = [];
    state.notifications = [];
    state.notificationArchive = [];
    closeProductionCreateScreen({ silent: true });
    showLoginPanel();
    showToast("Da dang xuat.", "success");
  }
}

function clearAuth() {
  state.authToken = "";
  state.challengeToken = "";
  stopNotificationPolling();
  localStorage.removeItem(AUTH_STORAGE_KEY);
  localStorage.removeItem("authToken");
}

function resetLoginState() {
  loginForm?.reset();
  otpBlock?.classList.add("hidden");
  otpNote.textContent = "";
  loginSubmit.textContent = "Dang nhap";
}

async function loadUsers() {
  const payload = await fetchJson(usersUrl);
  state.users = Array.isArray(payload.users) ? payload.users : [];
  if (payload.currentUser) {
    state.currentUser = payload.currentUser;
    renderHeroUser();
  }
}

async function loadOrders() {
  if (!canViewOrders()) {
    state.orders = [];
    renderTrackingOrders();
    populateDeliveryOrders();
    return;
  }
  const payload = await fetchJson(ordersUrl);
  state.orders = Array.isArray(payload.orders) ? payload.orders : [];
  populateDeliveryOrders();
  renderTrackingOrders();
}

function extractOrderSequenceValue(value) {
  const match = String(value || "")
    .trim()
    .toUpperCase()
    .match(/^DH-(\d+)$/);
  return match ? Number(match[1]) : NaN;
}

function formatOrderSequenceValue(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric) || numeric < 0) {
    return "003000";
  }
  return String(Math.trunc(numeric)).padStart(6, "0");
}

function getNextProductionOrderSequence() {
  const numericValues = state.orders
    .map((item) => extractOrderSequenceValue(item?.order_id || ""))
    .filter((value) => Number.isFinite(value));
  const maxValue = numericValues.length ? Math.max(...numericValues) : 2999;
  return formatOrderSequenceValue(Math.max(3000, maxValue + 1));
}

function updateTelegramViewportMetrics() {
  const viewportHeight = Number(TelegramWebApp?.viewportHeight || window.visualViewport?.height || window.innerHeight || 0);
  const stableHeight = Number(TelegramWebApp?.viewportStableHeight || viewportHeight || window.innerHeight || 0);
  const viewportWidth = Number(window.visualViewport?.width || window.innerWidth || 0);
  if (stableHeight > 0) {
    document.documentElement.style.setProperty("--tg-app-height", `${Math.round(stableHeight)}px`);
  }
  if (viewportHeight > 0) {
    document.documentElement.style.setProperty("--tg-app-live-height", `${Math.round(viewportHeight)}px`);
  }
  if (viewportWidth > 0) {
    document.documentElement.style.setProperty("--tg-app-width", `${Math.round(viewportWidth)}px`);
  }
}

function requestTelegramFullscreen() {
  updateTelegramViewportMetrics();
  if (!TelegramWebApp) {
    return;
  }
  try {
    TelegramWebApp.expand();
  } catch {
    void 0;
  }
  try {
    TelegramWebApp.requestFullscreen?.();
  } catch {
    void 0;
  }
}

function setTelegramVerticalSwipesDisabled(disabled) {
  if (!TelegramWebApp) {
    return;
  }
  try {
    if (disabled) {
      TelegramWebApp.disableVerticalSwipes?.();
    } else {
      TelegramWebApp.enableVerticalSwipes?.();
    }
  } catch {
    void 0;
  }
}

function buildProductionUnitOptions(selectedUnit = "") {
  const normalizedSelected = String(selectedUnit || DEFAULT_PRODUCTION_UNITS[0] || "").trim();
  return DEFAULT_PRODUCTION_UNITS.map(
    (unit) => `<option value="${escapeHtml(unit)}" ${unit === normalizedSelected ? "selected" : ""}>${escapeHtml(unit)}</option>`,
  ).join("");
}

function resetProductionCreateForm() {
  productionCreateForm?.reset();
  populateProductionCreateRequesterOptions();
  if (productionCreateOrderId) {
    productionCreateOrderId.value = getNextProductionOrderSequence();
  }
  if (productionCreateCreatedAt) {
    productionCreateCreatedAt.value = formatDateTime(new Date().toISOString());
  }
  if (productionCreateItemsList) {
    productionCreateItemsList.innerHTML = "";
  }
  if (productionCreatePerformerSignature) {
    productionCreatePerformerSignature.textContent = "-";
  }
  if (productionCreatePackagingSignature) {
    productionCreatePackagingSignature.textContent = "-";
  }
  if (productionCreateReceiptSignature) {
    productionCreateReceiptSignature.textContent = "-";
  }
  addProductionCreateItemRow();
  syncProductionCreateRequesterSignature();
  syncProductionCreateDraft();
}

function populateProductionCreateRequesterOptions() {
  if (!productionCreateRequester) {
    return;
  }
  const salesUsers = state.users
    .filter((user) => String(user?.department || "").trim().toLowerCase() === "sales")
    .sort((left, right) => String(left?.name || "").localeCompare(String(right?.name || ""), "vi"));

  const canChoose = canChooseSalesAssignee();
  const currentDepartment = String(state.currentUser?.department || "").trim().toLowerCase();
  const defaultRequester = canChoose
    ? String(
        (currentDepartment === "sales" ? state.currentUser?.id : salesUsers[0]?.id) || state.currentUser?.id || "",
      ).trim()
    : currentDepartment === "sales"
      ? String(state.currentUser?.id || "").trim()
      : "";
  const availableUsers = canChoose
    ? salesUsers
    : salesUsers.filter((user) => String(user?.id || "").trim() === defaultRequester);

  productionCreateRequester.innerHTML = availableUsers.length
    ? availableUsers
        .map((user) => {
          const code = user?.employee_code || user?.username || user?.id || "-";
          const selected = String(user?.id || "").trim() === defaultRequester ? "selected" : "";
          return `<option value="${escapeHtml(user.id)}" ${selected}>${escapeHtml(`${user.name || user.username || "-"} • ${code}`)}</option>`;
        })
        .join("")
    : '<option value="">Khong co nguoi yeu cau phu hop</option>';

  productionCreateRequester.disabled = !canChoose;
  syncProductionCreateRequesterSignature();
}

function syncProductionCreateRequesterSignature() {
  if (!productionCreateRequesterSignature) {
    return;
  }
  const selectedText = String(productionCreateRequester?.selectedOptions?.[0]?.textContent || "").trim();
  productionCreateRequesterSignature.textContent = selectedText ? selectedText.split("•")[0].trim() : "-";
}

function addProductionCreateItemRow(values = {}) {
  if (!productionCreateItemsList) {
    return;
  }

  const row = document.createElement("div");
  row.className = "production-create-row";
  row.setAttribute("data-production-create-row", "true");
  row.innerHTML = `
    <div class="production-create-index"><span data-field="index-label">1</span></div>
    <input class="production-create-code" data-field="code" type="text" value="${escapeHtml(values.code || "")}" />
    <div class="production-create-name-wrap">
      <textarea class="production-create-name" data-field="name" rows="1" placeholder="Ten hang hoa">${escapeHtml(values.name || "")}</textarea>
    </div>
    <input class="production-create-norm" data-field="norm" type="text" value="${escapeHtml(values.norm || "")}" />
    <select class="production-create-unit" data-field="unit">${buildProductionUnitOptions(values.unit || "")}</select>
    <input class="production-create-quantity" data-field="quantity" type="number" min="0" step="1" value="${escapeHtml(values.quantity || "")}" />
    <input class="production-create-done" data-field="done" type="number" min="0" step="1" value="${escapeHtml(values.done || "0")}" />
    <input class="production-create-missing" data-field="missing" type="number" min="0" step="1" value="${escapeHtml(values.missing || "0")}" />
    <input class="production-create-extra" data-field="extra" type="number" min="0" step="1" value="${escapeHtml(values.extra || "0")}" />
    <input class="production-create-team" data-field="team" type="number" min="0" step="1" value="${escapeHtml(values.team || "")}" />
  `;

  row.querySelectorAll('input[data-field="quantity"], input[data-field="done"]').forEach((input) => {
    input.addEventListener("input", () => updateProductionCreateDerivedFields(row));
    input.addEventListener("change", () => updateProductionCreateDerivedFields(row));
  });
  row.querySelector('[data-field="name"]')?.addEventListener("input", () => {
    autoResizeProductionCreateName(row.querySelector('[data-field="name"]'));
  });

  productionCreateItemsList.append(row);
  autoResizeProductionCreateName(row.querySelector('[data-field="name"]'));
  updateProductionCreateDerivedFields(row);
  syncProductionCreateDraft();
  row.querySelector('[data-field="code"]')?.focus();
}

function autoResizeProductionCreateName(input) {
  if (!input) {
    return;
  }
  input.style.height = "0px";
  input.style.height = `${Math.max(56, input.scrollHeight)}px`;
}

function updateProductionCreateDerivedFields(row) {
  if (!row) {
    return;
  }
  const quantity = Math.max(0, Number(row.querySelector('[data-field="quantity"]')?.value || 0));
  const done = Math.max(0, Number(row.querySelector('[data-field="done"]')?.value || 0));
  const missing = Math.max(0, quantity - done);
  const extra = Math.max(0, done - quantity);
  const missingField = row.querySelector('[data-field="missing"]');
  const extraField = row.querySelector('[data-field="extra"]');
  if (missingField) {
    missingField.value = String(missing);
  }
  if (extraField) {
    extraField.value = String(extra);
  }
}

function syncProductionCreateDraft() {
  const rows = Array.from(productionCreateItemsList?.querySelectorAll("[data-production-create-row]") || []);
  rows.forEach((row, index) => {
    const indexNode = row.querySelector('[data-field="index-label"]');
    if (indexNode) {
      indexNode.textContent = String(index + 1);
    }
    autoResizeProductionCreateName(row.querySelector('[data-field="name"]'));
  });
  if (productionCreateItemsEmpty) {
    productionCreateItemsEmpty.classList.toggle("hidden", rows.length > 0);
  }
}

function buildProductionCreateItemsSummary() {
  const rows = Array.from(productionCreateItemsList?.querySelectorAll("[data-production-create-row]") || []);
  return rows
    .map((row, index) => {
      const code = String(row.querySelector('[data-field="code"]')?.value || "").trim() || "-";
      const name = String(row.querySelector('[data-field="name"]')?.value || "").trim();
      const norm = String(row.querySelector('[data-field="norm"]')?.value || "").trim() || "-";
      const unit = String(row.querySelector('[data-field="unit"]')?.value || "").trim() || "-";
      const quantity = String(row.querySelector('[data-field="quantity"]')?.value || "").trim() || "0";
      const done = String(row.querySelector('[data-field="done"]')?.value || "").trim() || "0";
      const missing = String(row.querySelector('[data-field="missing"]')?.value || "").trim() || "0";
      const extra = String(row.querySelector('[data-field="extra"]')?.value || "").trim() || "0";
      const team = String(row.querySelector('[data-field="team"]')?.value || "").trim() || "-";
      if (!name && !code) {
        return "";
      }
      return `${index + 1}. ${code} | ${name} | DM ${norm} | ${unit} | SL ${quantity} | Da SX ${done} | Thieu ${missing} | Du ${extra} | To ${team}`;
    })
    .filter(Boolean)
    .join("\n");
}

function buildProductionCreateNote() {
  const note = String(productionCreateNote?.value || "").trim();
  const turnaroundValue = String(
    productionCreateForm?.querySelector('input[name="production-create-turnaround"]:checked')?.value || "",
  ).trim();
  const turnaround =
    turnaroundValue === "Khac"
      ? `Khac: ${String(productionCreateTurnaroundOther?.value || "").trim()}`
      : turnaroundValue;
  return [note, turnaround ? `De nghi tra hang trong: ${turnaround}` : ""].filter(Boolean).join("\n");
}

async function handleProductionCreateSubmit(event) {
  event.preventDefault();
  if (!canCreateOrder()) {
    showToast("Ban khong co quyen tao phieu san xuat.", "error");
    return;
  }

  const customerName = String(productionCreateCustomerName?.value || "").trim();
  const requesterId = String(productionCreateRequester?.value || "").trim();
  const orderItems = buildProductionCreateItemsSummary();
  if (!customerName || !requesterId) {
    showToast("Can nhap khach hang va nguoi yeu cau.", "error");
    return;
  }
  if (!orderItems) {
    showToast("Can co it nhat mot dong san xuat.", "error");
    return;
  }

  submitProductionCreateButton && (submitProductionCreateButton.disabled = true);
  try {
    const payload = await postJson(orderCreateUrl, {
      order_id: normalizeOrderIdValue(productionCreateOrderId?.value || ""),
      customer_name: customerName,
      sales_user_id: requesterId,
      delivery_user_id: "",
      planned_delivery_at: "",
      delivery_address: "",
      order_items: orderItems,
      order_value: "",
      note: buildProductionCreateNote(),
      order_kind: "production",
    });
    applyOrderPayload(payload);
    await loadOrders();
    await loadNotifications({ silent: true });
    closeProductionCreateScreen({ silent: true });
    showToast("Da tao phieu san xuat.", "success");
    await openSection("tracking");
    setTrackingKind("production");
  } catch (error) {
    showToast(error.message || "Khong the tao phieu san xuat.", "error");
  } finally {
    submitProductionCreateButton && (submitProductionCreateButton.disabled = false);
  }
}

async function loadNotifications(options = {}) {
  if (!state.authToken) {
    state.notifications = [];
    state.notificationArchive = [];
    renderNotifications();
    return;
  }

  try {
    const payload = await fetchJson(notificationsUrl);
    state.notifications = Array.isArray(payload.notifications) ? payload.notifications : [];
    mergeNotificationArchive(state.notifications);
    renderNotifications();
  } catch (error) {
    if (!options.silent) {
      throw error;
    }
    state.notifications = [];
    state.notificationArchive = loadNotificationArchive();
    renderNotifications();
  }
}

function startNotificationPolling() {
  stopNotificationPolling();
  if (!state.authToken) {
    return;
  }
  notificationPollTimer = window.setInterval(() => {
    loadNotifications({ silent: true }).catch(() => void 0);
  }, NOTIFICATION_POLL_INTERVAL_MS);
}

function stopNotificationPolling() {
  if (notificationPollTimer) {
    window.clearInterval(notificationPollTimer);
    notificationPollTimer = null;
  }
}

async function loadTapeConfig() {
  if (state.tapeLoaded) {
    return;
  }

  try {
    const payload = await fetchJson(tapeCalculatorConfigUrl);
    state.tapeConfig = payload || {};
    state.tapeLoaded = true;
    applyTapeConfig();
  } catch (error) {
    tapeSourceBadge.textContent = "Khong tai duoc config";
    throw error;
  }
}

function applyTapeConfig() {
  const products = Array.isArray(state.tapeConfig?.products) ? state.tapeConfig.products : [];
  const cores = Array.isArray(state.tapeConfig?.cores) ? state.tapeConfig.cores : [];
  const defaults = state.tapeConfig?.defaults || {};

  tapeType.innerHTML = products.length
    ? products
        .map((item) => `<option value="${escapeHtml(item.code)}">${escapeHtml(item.code)}</option>`)
        .join("")
    : '<option value="">Khong co du lieu</option>';

  tapeCoreType.innerHTML = cores.length
    ? cores
        .map((item) => `<option value="${escapeHtml(item.code)}">${escapeHtml(item.code)}</option>`)
        .join("")
    : '<option value="">Khong co du lieu</option>';

  tapeType.value = products.some((item) => item.code === defaults.tape_type) ? defaults.tape_type : products[0]?.code || "";
  tapeCoreType.value = cores.some((item) => item.code === defaults.core_type) ? defaults.core_type : cores[0]?.code || "";
  tapeOrderQuantity.value = String(defaults.order_quantity || 0);
  tapePackaging.value = String(defaults.packaging || 0);
  tapeJumboHeight.value = String(defaults.jumbo_height || products[0]?.jumbo_height || 0);
  tapeCoreWidth.value = String(defaults.core_width || cores[0]?.width_mm || 0);
  tapeSourceBadge.textContent = state.tapeConfig?.fallback ? "Dang dung cache" : "Dang dung sheet live";
  updateTapeResults();
}

function updateTapeResults() {
  const products = Array.isArray(state.tapeConfig?.products) ? state.tapeConfig.products : [];
  const cores = Array.isArray(state.tapeConfig?.cores) ? state.tapeConfig.cores : [];
  const selectedProduct = products.find((item) => item.code === tapeType?.value) || null;
  const selectedCore = cores.find((item) => item.code === tapeCoreType?.value) || null;
  const orderQuantity = Math.max(0, Number(tapeOrderQuantity?.value || 0));
  const packaging = Math.max(0, Number(tapePackaging?.value || 0));
  const jumboHeight = Math.max(0, Number(tapeJumboHeight?.value || selectedProduct?.jumbo_height || 0));
  const coreWidth = Math.max(0, Number(tapeCoreWidth?.value || selectedCore?.width_mm || 0));

  if (selectedProduct && Number(tapeJumboHeight?.value || 0) !== Number(selectedProduct.jumbo_height || 0)) {
    tapeJumboHeight.value = String(selectedProduct.jumbo_height || 0);
  }
  if (selectedCore && Number(tapeCoreWidth?.value || 0) !== Number(selectedCore.width_mm || 0)) {
    tapeCoreWidth.value = String(selectedCore.width_mm || 0);
  }

  const finishedQuantity = orderQuantity * packaging;
  const rollsPerHand = coreWidth > 0 ? Math.floor(jumboHeight / coreWidth) : 0;
  const remainingMm = rollsPerHand > 0 ? jumboHeight - rollsPerHand * coreWidth : 0;
  const handsNeeded = rollsPerHand > 0 ? Math.ceil(finishedQuantity / rollsPerHand) : 0;
  const totalProduced = rollsPerHand * handsNeeded;
  const extraProduced = Math.max(0, totalProduced - finishedQuantity);

  tapeFinishedQuantity.textContent = formatNumber(finishedQuantity);
  tapeRollsPerHand.textContent = formatNumber(rollsPerHand);
  tapeRemainingMm.textContent = formatNumber(remainingMm);
  tapeHandsNeeded.textContent = formatNumber(handsNeeded);
  tapeTotalProduced.textContent = formatNumber(totalProduced);
  tapeExtraProduced.textContent = formatNumber(extraProduced);
}

function populateCreateUserOptions() {
  const salesUsers = state.users
    .filter((user) => String(user?.department || "").trim().toLowerCase() === "sales")
    .sort((left, right) => String(left?.name || "").localeCompare(String(right?.name || ""), "vi"));

  const deliveryUsers = state.users
    .filter((user) => {
      const department = String(user?.department || "").trim().toLowerCase();
      return ["operations", "transport", "logistics", "delivery"].includes(department);
    })
    .sort((left, right) => String(left?.name || "").localeCompare(String(right?.name || ""), "vi"));

  createSalesUser.innerHTML = salesUsers.length
    ? salesUsers
        .map((user) => `<option value="${escapeHtml(user.id)}">${escapeHtml(buildUserLabel(user))}</option>`)
        .join("")
    : '<option value="">Khong co NV kinh doanh</option>';

  createDeliveryUser.innerHTML = deliveryUsers.length
    ? deliveryUsers
        .map((user) => `<option value="${escapeHtml(user.id)}">${escapeHtml(buildUserLabel(user))}</option>`)
        .join("")
    : '<option value="">Khong co NV giao hang</option>';

  if (!canChooseSalesAssignee() && state.currentUser?.id) {
    createSalesUser.value = String(state.currentUser.id || "");
    createSalesUser.disabled = true;
  } else {
    createSalesUser.disabled = false;
  }
}

function populateDeliverySalesOptions() {
  const salesUsers = state.users
    .filter((user) => String(user?.department || "").trim().toLowerCase() === "sales")
    .sort((left, right) => String(left?.name || "").localeCompare(String(right?.name || ""), "vi"));

  deliverySalesUser.innerHTML = salesUsers.length
    ? salesUsers
        .map((user) => `<option value="${escapeHtml(user.id)}">${escapeHtml(buildUserLabel(user))}</option>`)
        .join("")
    : '<option value="">Khong co NV kinh doanh</option>';
}

function populateDeliveryOrders() {
  const options = getPendingTransportOrders();
  deliveryOrderId.innerHTML = options.length
    ? options
        .map((order) => `<option value="${escapeHtml(order.order_id || "")}">${escapeHtml(buildDeliveryOptionLabel(order))}</option>`)
        .join("")
    : '<option value="">Khong co don can giao</option>';
}

function autofillDeliveryForm() {
  const orderId = normalizeOrderIdValue(deliveryOrderId?.value || "");
  const order = state.orders.find((item) => normalizeOrderIdValue(item?.order_id || "") === orderId);
  if (!order) {
    deliveryCustomerName.value = "";
    deliveryAddress.value = "";
    if (deliverySalesUser.options.length) {
      deliverySalesUser.selectedIndex = 0;
    }
    return;
  }

  deliveryCustomerName.value = order.customer_name || "";
  deliveryAddress.value = order.delivery_address || "";
  if (order.sales_user_id) {
    deliverySalesUser.value = order.sales_user_id;
  }
  deliveryCompletedAt.value = toLocalDateTimeInputValue(new Date().toISOString());
}

function syncDeliveryPaymentMethodState() {
  const isPaid = String(deliveryPaymentStatus?.value || "").trim().toLowerCase() === "paid";
  deliveryPaymentMethod.disabled = !isPaid;
  if (!isPaid) {
    deliveryPaymentMethod.value = "";
  }
}

async function handleDeliverySubmit(event) {
  event.preventDefault();
  if (!canCompleteDelivery()) {
    showToast("Ban khong co quyen cap nhat giao hang.", "error");
    return;
  }

  deliverySubmit.disabled = true;
  try {
    const payload = await postJson(deliveryCompleteUrl, {
      order_id: normalizeOrderIdValue(deliveryOrderId?.value || ""),
      customer_name: deliveryCustomerName?.value || "",
      sales_user_id: deliverySalesUser?.value || "",
      result_status: deliveryResultStatus?.value || "delivered",
      completed_at: deliveryCompletedAt?.value || "",
      payment_status: deliveryPaymentStatus?.value || "unpaid",
      payment_method: deliveryPaymentMethod?.value || "",
      note: deliveryNote?.value || "",
    });
    applyOrderPayload(payload);
    deliveryForm.reset();
    syncDeliveryPaymentMethodState();
    populateDeliveryOrders();
    autofillDeliveryForm();
    renderTrackingOrders();
    showToast("Da cap nhat giao hang.", "success");
  } catch (error) {
    showToast(error.message || "Khong the cap nhat giao hang.", "error");
  } finally {
    deliverySubmit.disabled = false;
  }
}

function setOrderKind(kind) {
  state.orderKind = kind === "production" ? "production" : "transport";
  createKindButtons.forEach((button) => {
    button.classList.toggle("is-active", button.dataset.orderKind === state.orderKind);
  });

  const isTransport = state.orderKind === "transport";
  transportOnlyFields.forEach((field) => field.classList.toggle("hidden", !isTransport));
  createOrderId.placeholder = isTransport ? "VD: DH-0002" : "Bo trong de he thong tu cap ma";
}

async function handleCreateSubmit(event) {
  event.preventDefault();
  if (!canCreateOrder()) {
    showToast("Ban khong co quyen tao don.", "error");
    return;
  }

  createSubmit.disabled = true;
  try {
    const isTransport = state.orderKind === "transport";
    const payload = await postJson(orderCreateUrl, {
      order_id: normalizeOrderIdValue(createOrderId?.value || ""),
      customer_name: createCustomerName?.value || "",
      sales_user_id: createSalesUser?.value || "",
      delivery_user_id: isTransport ? createDeliveryUser?.value || "" : "",
      planned_delivery_at: isTransport ? createPlannedAt?.value || "" : "",
      delivery_address: isTransport ? createAddress?.value || "" : "",
      order_items: createOrderItems?.value || "",
      order_value: isTransport ? createOrderValue?.value || "" : "",
      note: createNote?.value || "",
      order_kind: state.orderKind,
    });
    applyOrderPayload(payload);
    createForm.reset();
    setOrderKind(state.orderKind);
    populateDeliveryOrders();
    setTrackingKind(state.orderKind === "production" ? "production" : "transport");
    renderTrackingOrders();
    showToast(state.orderKind === "production" ? "Da tao phieu san xuat." : "Da tao don van chuyen.", "success");
    await openSection("tracking");
  } catch (error) {
    showToast(error.message || "Khong the tao don.", "error");
  } finally {
    createSubmit.disabled = false;
  }
}

function setTrackingKind(kind) {
  state.trackingKind = kind === "production" ? "production" : "transport";
  trackingKindButtons.forEach((button) => {
    button.classList.toggle("is-active", button.dataset.trackKind === state.trackingKind);
  });
  renderTrackingOrders();
}

function renderTrackingOrders() {
  if (!trackingList) {
    return;
  }

  if (!canViewOrders()) {
    trackingList.innerHTML = '<article class="order-card"><p class="empty-state">Ban khong co quyen xem don hang.</p></article>';
    return;
  }

  const keyword = String(trackingSearch?.value || "").trim().toLowerCase();
  const visibleOrders = state.orders.filter((order) => {
    const isProduction = isProductionOrder(order);
    if (state.trackingKind === "production" && !isProduction) {
      return false;
    }
    if (state.trackingKind === "transport" && isProduction) {
      return false;
    }

    if (!keyword) {
      return true;
    }

    const haystack = [
      order.order_id,
      order.customer_name,
      order.sales_user_name,
      order.delivery_user_name,
      order.production_claimed_by_names,
      order.note,
      order.delivery_address,
    ]
      .join(" ")
      .toLowerCase();
    return haystack.includes(keyword);
  });

  if (!visibleOrders.length) {
    trackingList.innerHTML = '<article class="order-card"><p class="empty-state">Khong co don phu hop.</p></article>';
    return;
  }

  trackingList.innerHTML = visibleOrders.map((order) => buildOrderCardMarkup(order)).join("");
  trackingList.querySelectorAll("[data-track-complete]").forEach((button) => {
    button.addEventListener("click", async () => {
      const orderId = button.getAttribute("data-track-complete") || "";
      const order = state.orders.find((item) => normalizeOrderIdValue(item?.order_id || "") === normalizeOrderIdValue(orderId));
      if (!order) {
        return;
      }
      await openSection("delivery");
      deliveryOrderId.value = order.order_id || "";
      autofillDeliveryForm();
      window.scrollTo({ top: 0, behavior: "smooth" });
    });
  });
}

function buildOrderCardMarkup(order) {
  const isProduction = isProductionOrder(order);
  const status = isProduction ? getProductionStatus(order) : getTransportStatus(order);
  const factLines = isProduction
    ? [
        `Nguoi yeu cau: ${order.sales_user_name || "-"}`,
        order.production_claimed_by_names ? `Dang nhan boi: ${order.production_claimed_by_names}` : "Chua co nguoi nhan",
        order.created_at ? `Ngay tao: ${formatDateTime(order.created_at)}` : "",
        order.note ? `Ghi chu: ${order.note}` : "",
      ].filter(Boolean)
    : [
        `NVKD: ${order.sales_user_name || "-"}`,
        `NV giao: ${order.delivery_user_name || "-"}`,
        order.planned_delivery_at ? `Can giao: ${formatDateTime(order.planned_delivery_at)}` : "",
        order.completed_at ? `Hoan thanh: ${formatDateTime(order.completed_at)}` : "Chua giao",
        `Thanh toan: ${labelPaymentStatus(order.payment_status || "unpaid")}`,
        order.delivery_address ? `Dia chi: ${order.delivery_address}` : "",
      ].filter(Boolean);

  return `
    <article class="order-card">
      <div class="order-head">
        <div class="order-title">
          <strong>${escapeHtml(order.order_id || "-")}</strong>
          <span>${escapeHtml(order.customer_name || order.document_title || "-")}</span>
        </div>
        <div class="order-tags">
          <span class="tag ${escapeHtml(status.tone)}">${escapeHtml(status.label)}</span>
          ${isProduction ? '<span class="tag">San xuat</span>' : '<span class="tag">Van chuyen</span>'}
        </div>
      </div>
      <div class="order-meta">${factLines.map((line) => escapeHtml(line)).join("<br />")}</div>
      ${
        !isProduction && canCompleteDelivery() && String(order.status || "").trim().toLowerCase() !== "completed"
          ? `<div class="order-actions"><button class="ghost-button" data-track-complete="${escapeHtml(order.order_id || "")}" type="button">Hoan tat giao</button></div>`
          : ""
      }
    </article>
  `;
}

function renderNotifications() {
  if (!notificationList || !notificationCount || !notificationHistoryList || !notificationHistoryCount) {
    return;
  }

  const visibleNotifications = state.notifications.filter((item) => {
    if (state.notificationFilter === "unread") {
      return !String(item?.read_at || "").trim();
    }
    return true;
  });

  notificationCount.textContent = String(visibleNotifications.length);
  notificationList.innerHTML = visibleNotifications.length
    ? visibleNotifications.map((item) => buildNotificationCardMarkup(item)).join("")
    : '<article class="notification-empty-card">Chua co thong bao nao</article>';

  notificationHistoryCount.textContent = String(state.notificationArchive.length);
  notificationHistoryList.innerHTML = state.notificationArchive.length
    ? state.notificationArchive.map((item) => buildNotificationCardMarkup(item, { history: true })).join("")
    : '<article class="notification-empty-card">Chua co lich su thong bao</article>';

  notificationList.querySelectorAll("[data-notification-read]").forEach((node) => {
    node.addEventListener("click", async () => {
      const notificationId = node.getAttribute("data-notification-read") || "";
      if (!notificationId) {
        return;
      }
      await markNotificationRead(notificationId);
    });
  });
}

function buildNotificationCardMarkup(item, options = {}) {
  const isUnread = !String(item?.read_at || "").trim() && !options.history;
  const title = String(item?.title || item?.type || "Thong bao").trim();
  const message = String(item?.message || "").trim() || "Khong co noi dung.";
  const dateValue = item?.created_at || item?.read_at || "";

  return `
    <article class="notification-item-card${isUnread ? " is-unread" : ""}" ${isUnread ? `data-notification-read="${escapeHtml(item?.id || "")}"` : ""}>
      <div class="notification-item-headline">
        <strong class="notification-item-title">${escapeHtml(title)}</strong>
        <span class="notification-item-time">${escapeHtml(formatCompactDateTime(dateValue))}</span>
      </div>
      <div class="notification-item-copy">${escapeHtml(message)}</div>
    </article>
  `;
}

async function markNotificationRead(notificationId) {
  const target = state.notifications.find((item) => String(item?.id || "").trim() === String(notificationId || "").trim());
  if (!target || String(target?.read_at || "").trim()) {
    return;
  }

  try {
    const payload = await postJson(notificationsReadUrl, {
      notification_id: notificationId,
    });
    state.notifications = Array.isArray(payload.notifications) ? payload.notifications : state.notifications;
    mergeNotificationArchive(state.notifications);
    renderNotifications();
  } catch (error) {
    showToast(error.message || "Khong the danh dau da doc.", "error");
  }
}

async function handleMarkAllNotificationsRead() {
  if (!state.notifications.some((item) => !String(item?.read_at || "").trim())) {
    return;
  }

  markAllNotificationsReadButton.disabled = true;
  try {
    const payload = await postJson(notificationsReadAllUrl, {});
    state.notifications = Array.isArray(payload.notifications) ? payload.notifications : [];
    mergeNotificationArchive(state.notifications);
    renderNotifications();
    showToast("Da danh dau tat ca thong bao la da doc.", "success");
  } catch (error) {
    showToast(error.message || "Khong the cap nhat thong bao.", "error");
  } finally {
    markAllNotificationsReadButton.disabled = false;
  }
}

function handleClearNotificationHistory() {
  state.notificationArchive = [];
  saveNotificationArchive();
  renderNotifications();
  showToast("Da xoa lich su thong bao.", "success");
}

function mergeNotificationArchive(items) {
  const archive = loadNotificationArchive();
  const map = new Map(archive.map((item) => [String(item?.id || "").trim(), item]));
  items.forEach((item) => {
    const key = String(item?.id || "").trim();
    if (!key) {
      return;
    }
    map.set(key, item);
  });
  state.notificationArchive = Array.from(map.values())
    .sort((left, right) => String(right?.created_at || "").localeCompare(String(left?.created_at || "")))
    .slice(0, NOTIFICATION_ARCHIVE_MAX_ITEMS);
  saveNotificationArchive();
}

function loadNotificationArchive() {
  const storageKey = getNotificationArchiveStorageKey();
  if (!storageKey) {
    return [];
  }
  try {
    const raw = localStorage.getItem(storageKey);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function saveNotificationArchive() {
  const storageKey = getNotificationArchiveStorageKey();
  if (!storageKey) {
    return;
  }
  localStorage.setItem(storageKey, JSON.stringify(state.notificationArchive));
}

function getNotificationArchiveStorageKey() {
  const ownerKey = String(state.currentUser?.id || state.currentUser?.username || "").trim();
  return ownerKey ? `${NOTIFICATION_ARCHIVE_STORAGE_PREFIX}:${ownerKey}` : "";
}

function getTransportStatus(order) {
  const normalized = String(order?.status || "").trim().toLowerCase();
  if (normalized === "completed" || order?.completed_at) {
    return { label: "Da giao", tone: "success" };
  }
  if (normalized === "assigned") {
    return { label: "Dang giao", tone: "warn" };
  }
  return { label: "Dang xu ly", tone: "" };
}

function getProductionStatus(order) {
  if (order?.production_receipt_completed_at) {
    return { label: "Da nhan du hang", tone: "success" };
  }
  if (order?.production_packaging_completed_at) {
    return { label: "Da dong goi", tone: "warn" };
  }
  if (String(order?.production_claimed_by_names || "").trim()) {
    return { label: "Dang san xuat", tone: "" };
  }
  return { label: "Chua nhan", tone: "danger" };
}

function applyPermissionState() {
  actionButtons.forEach((button) => {
    const allowed = isSectionAllowed(button.dataset.section || "");
    button.classList.toggle("is-disabled", !allowed);
    button.setAttribute("aria-disabled", allowed ? "false" : "true");
  });
}

function isSectionAllowed(section) {
  if (!state.currentUser) {
    return false;
  }
  if (section === "delivery") {
    return canCompleteDelivery();
  }
  if (section === "create") {
    return canCreateOrder();
  }
  if (section === "tracking") {
    return canViewOrders();
  }
  if (section === "tape") {
    return true;
  }
  return false;
}

function canCreateOrder() {
  const role = String(state.currentUser?.role || "").trim().toLowerCase();
  const department = String(state.currentUser?.department || "").trim().toLowerCase();
  return role === "admin" || role === "director" || role === "manager" || department === "sales";
}

function canChooseSalesAssignee() {
  const role = String(state.currentUser?.role || "").trim().toLowerCase();
  return role === "admin" || role === "director" || role === "manager";
}

function canViewOrders() {
  if (!state.currentUser) {
    return false;
  }
  const role = String(state.currentUser?.role || "").trim().toLowerCase();
  const department = String(state.currentUser?.department || "").trim().toLowerCase();
  return (
    role === "admin" ||
    role === "director" ||
    role === "manager" ||
    department === "sales" ||
    department === "production" ||
    ["operations", "transport", "logistics", "delivery"].includes(department)
  );
}

function canCompleteDelivery() {
  const role = String(state.currentUser?.role || "").trim().toLowerCase();
  const department = String(state.currentUser?.department || "").trim().toLowerCase();
  return role === "admin" || role === "director" || ["operations", "transport", "logistics", "delivery"].includes(department);
}

function isProductionOrder(order) {
  const explicitKind = String(order?.order_kind || "").trim().toLowerCase();
  if (explicitKind === "production") {
    return true;
  }
  if (explicitKind === "transport") {
    return false;
  }
  return !String(order?.delivery_user_id || "").trim();
}

function getPendingTransportOrders() {
  return state.orders
    .filter((order) => !isProductionOrder(order))
    .filter((order) => String(order?.status || "").trim().toLowerCase() !== "completed")
    .sort((left, right) => String(right?.created_at || "").localeCompare(String(left?.created_at || "")));
}

function applyOrderPayload(payload) {
  if (Array.isArray(payload?.orders)) {
    state.orders = payload.orders;
  }
}

function renderHeroUser() {
  document.body.classList.toggle("telegram-authenticated", Boolean(state.currentUser));
  if (!state.currentUser) {
    heroUserName.textContent = "Chua dang nhap";
    heroUserMeta.textContent = TelegramWebApp ? "Dang mo trong Telegram" : "Dang test tren browser";
    return;
  }

  heroUserName.textContent = state.currentUser.name || state.currentUser.username || "Da dang nhap";
  const parts = [
    labelRole(state.currentUser.role),
    labelDepartment(state.currentUser.department),
    labelAccessLevel(state.currentUser.policy?.max_access_level || "basic"),
  ].filter(Boolean);
  heroUserMeta.textContent = parts.join(" | ");
}

function labelRole(role) {
  const normalized = String(role || "").trim().toLowerCase();
  if (normalized === "admin") return "Admin";
  if (normalized === "director") return "Giam doc";
  if (normalized === "manager") return "Quan ly";
  return "Nhan vien";
}

function labelDepartment(department) {
  const normalized = String(department || "").trim().toLowerCase();
  if (normalized === "sales") return "Kinh doanh";
  if (normalized === "production") return "San xuat";
  if (normalized === "operations" || normalized === "transport" || normalized === "logistics" || normalized === "delivery") {
    return "Van chuyen";
  }
  if (normalized === "finance") return "Tai chinh";
  if (normalized === "executive") return "Ban giam doc";
  return department || "-";
}

function labelAccessLevel(level) {
  const normalized = String(level || "").trim().toLowerCase();
  if (normalized === "advanced") return "Nang cao";
  if (normalized === "sensitive") return "Ca nhan";
  return "Co ban";
}

function labelPaymentStatus(value) {
  return String(value || "").trim().toLowerCase() === "paid" ? "Da thanh toan" : "Chua thanh toan";
}

function normalizeOrderIdValue(value) {
  const core = String(value || "")
    .trim()
    .replace(/^dh[-\s]*/i, "")
    .replace(/[^0-9a-z-]/gi, "")
    .trim()
    .toUpperCase();
  return core ? `DH-${core}` : "";
}

function buildUserLabel(user) {
  const code = user?.employee_code || user?.username || user?.id || "-";
  return `${user?.name || user?.username || "-"} • ${code}`;
}

function buildDeliveryOptionLabel(order) {
  return `${order.order_id || "-"} • ${order.customer_name || "-"} • ${order.delivery_user_name || "Chua giao NV"}`;
}

function formatCompactDateTime(value) {
  const date = new Date(value);
  if (!Number.isFinite(date.getTime())) {
    return "-";
  }
  return new Intl.DateTimeFormat("vi-VN", {
    day: "2-digit",
    month: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  }).format(date);
}

function formatDateTime(value) {
  const date = new Date(value);
  if (!Number.isFinite(date.getTime())) {
    return value || "-";
  }
  return new Intl.DateTimeFormat("vi-VN", {
    dateStyle: "medium",
    timeStyle: "short",
  }).format(date);
}

function toLocalDateTimeInputValue(value) {
  const date = value instanceof Date ? value : new Date(value);
  if (!Number.isFinite(date.getTime())) {
    return "";
  }
  const offset = date.getTimezoneOffset();
  const local = new Date(date.getTime() - offset * 60_000);
  return local.toISOString().slice(0, 16);
}

function formatNumber(value) {
  return new Intl.NumberFormat("vi-VN").format(Number(value || 0));
}

function showToast(message, tone = "success") {
  if (!toastStack) {
    return;
  }
  const element = document.createElement("div");
  element.className = `toast ${tone === "error" ? "error" : "success"}`;
  element.textContent = String(message || "");
  toastStack.append(element);
  setTimeout(() => {
    element.remove();
  }, 3200);
}

async function fetchJson(url, options = {}) {
  const response = await fetch(url, {
    cache: "no-store",
    headers: {
      Accept: "application/json",
      ...(state.authToken && !options.skipAuth ? { Authorization: `Bearer ${state.authToken}` } : {}),
    },
  });
  const payload = await response.json();
  if (!response.ok) {
    if (response.status === 401 && !options.skipAuth) {
      clearAuth();
      state.currentUser = null;
      showLoginPanel();
    }
    throw new Error(payload.error || "Yeu cau that bai.");
  }
  return payload;
}

async function postJson(url, data, options = {}) {
  const response = await fetch(url, {
    method: "POST",
    cache: "no-store",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json",
      ...(state.authToken && !options.skipAuth ? { Authorization: `Bearer ${state.authToken}` } : {}),
    },
    body: JSON.stringify(data || {}),
  });
  const payload = await response.json();
  if (!response.ok) {
    if (response.status === 401 && !options.skipAuth) {
      clearAuth();
      state.currentUser = null;
      showLoginPanel();
    }
    throw new Error(payload.error || "Yeu cau that bai.");
  }
  return payload;
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}
