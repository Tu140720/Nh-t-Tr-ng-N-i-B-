const healthUrl = "/api/health";
const askUrl = "/api/ask";
const sessionUrl = "/api/session";
const loginUrl = "/api/login";
const loginVerifyUrl = "/api/login/verify";
const logoutUrl = "/api/logout";
const docsUrl = "/api/internal-docs";
const sourcesUrl = "/api/sources";
const usersUrl = "/api/users";
const userAccessUrl = "/api/users/access";
const userProfileUrl = "/api/users/profile";
const userPasswordUrl = "/api/users/password";
const userCreateUrl = "/api/users/create";
const userDeleteUrl = "/api/users/delete";
const checkDocsUrl = "/api/internal-docs/check";
const syncDocsUrl = "/api/internal-docs/sync";
const createFolderUrl = "/api/internal-folders";
const renameFolderUrl = "/api/internal-folders/rename";
const deleteFolderUrl = "/api/internal-folders/delete";
const docContentUrl = "/api/internal-docs/content";
const docDownloadUrl = "/api/internal-docs/download";
const docDeleteUrl = "/api/internal-docs/delete";
const docAccessUrl = "/api/internal-docs/access";
const docReviewUrl = "/api/internal-docs/review";
const auditLogsUrl = "/api/audit-logs";
const dashboardUrl = "/api/dashboard";
const syncOneSourceUrl = "/api/sources/sync-one";
const sourceSyncableUrl = "/api/sources/syncable";
const deleteSourceUrl = "/api/sources/delete";
const reviewSourceUrl = "/api/sources/review";
const orderProductsUrl = "/api/order-products";
const orderProductsSaveUrl = "/api/order-products";
const tapeCalculatorConfigUrl = "/api/tape-calculator/config";
const orderCreateUrl = "/api/orders/create";
const orderUpdateUrl = "/api/orders/update";
const orderBackfillUrl = "/api/orders/backfill-details";
const orderDeleteUrl = "/api/orders/delete";
const orderProductionClaimUrl = "/api/orders/production-claim";
const orderProductionProgressUrl = "/api/orders/production-progress";
const orderProductionPackagingCompleteUrl = "/api/orders/production-packaging-complete";
const orderProductionReceiptCompleteUrl = "/api/orders/production-receipt-complete";
const ordersUrl = "/api/orders";
const salesInventoryUrl = "/api/sales-inventory";
const salesInventoryReceiveUrl = "/api/sales-inventory/receive";
const salesInventoryTransferUrl = "/api/sales-inventory/transfer-from-production";
const deliveryCompleteUrl = "/api/delivery/complete";
const notificationsUrl = "/api/notifications";
const notificationsStreamUrl = "/api/notifications/stream";
const notificationsReadUrl = "/api/notifications/read";
const notificationsReadAllUrl = "/api/notifications/read-all";
const notificationsDeleteUrl = "/api/notifications/delete";

const DEFAULT_TAPE_CALCULATOR_PRODUCTS = [
  { code: "BOPP45", jumbo_height: 1260 },
  { code: "BOPP48", jumbo_height: 1260 },
  { code: "BOPP50", jumbo_height: 1260 },
  { code: "BOPP57", jumbo_height: 1260 },
  { code: "BOPP60", jumbo_height: 1260 },
  { code: "HDV51", jumbo_height: 1260 },
  { code: "IN1T50", jumbo_height: 980 },
  { code: "IN2T50", jumbo_height: 980 },
  { code: "IN1TS50", jumbo_height: 980 },
  { code: "IN2TS50", jumbo_height: 980 },
];

const DEFAULT_TAPE_CALCULATOR_CORES = [
  { code: "G10", width_mm: 10 },
  { code: "G12", width_mm: 12 },
  { code: "G15", width_mm: 15 },
  { code: "G18", width_mm: 18 },
  { code: "G20", width_mm: 20 },
  { code: "G25", width_mm: 25 },
  { code: "G46", width_mm: 46 },
  { code: "G70", width_mm: 70 },
  { code: "N36", width_mm: 36 },
];

const DEFAULT_TAPE_CALCULATOR_FORM = {
  tape_type: "BOPP60",
  order_quantity: 212,
  core_type: "G15",
  packaging: 20,
  finished_quantity: 4240,
  jumbo_height: 1260,
  core_width: 15,
};

let tapeCalculatorProducts = [...DEFAULT_TAPE_CALCULATOR_PRODUCTS];
let tapeCalculatorCores = [...DEFAULT_TAPE_CALCULATOR_CORES];
let tapeCalculatorDefaults = { ...DEFAULT_TAPE_CALCULATOR_FORM };
let tapeCalculatorConfigLoaded = false;
let tapeCalculatorConfigLoadingPromise = null;

const askForm = document.querySelector("#ask-form");
const questionInput = document.querySelector("#question-input");
const submitButton = document.querySelector("#submit-button");
const chatThread = document.querySelector("#chat-thread");
const emptyState = document.querySelector("#empty-state");
const historyList = document.querySelector("#history-list");
const historyCount = document.querySelector("#history-count");
const systemStatus = document.querySelector("#system-status");
const notificationList = document.querySelector("#notification-list");
const notificationCount = document.querySelector("#notification-count");
const notificationFilter = document.querySelector("#notification-filter");
const notificationHistoryList = document.querySelector("#notification-history-list");
const notificationHistoryCount = document.querySelector("#notification-history-count");
const clearNotificationHistoryButton = document.querySelector("#clear-notification-history");
const markAllNotificationsReadButton = document.querySelector("#mark-all-notifications-read");
const currentUserName = document.querySelector("#current-user-name");
const currentUserRole = document.querySelector("#current-user-role");
const currentUserMeta = document.querySelector("#current-user-meta");
const currentUserAvatar = document.querySelector("#current-user-avatar");
const openPermissionsModalButton = document.querySelector("#open-permissions-modal");
const openDeliveryModalButton = document.querySelector("#open-delivery-modal");
const openTapeCalculatorModalButton = document.querySelector("#open-tape-calculator-modal");
const orderCtaMenu = document.querySelector("#order-cta-menu");
const openOrderMenuButton = document.querySelector("#open-order-menu");
const orderCtaDropdown = document.querySelector("#order-cta-dropdown");
const openProductionOrderModalButton = document.querySelector("#open-production-order-modal");
const openTransportOrderModalButton = document.querySelector("#open-transport-order-modal");
const openOrderProductsModalButton = document.querySelector("#open-order-products-modal");
const ordersBoardMenu = document.querySelector("#orders-board-menu");
const openOrdersBoardMenuButton = document.querySelector("#open-orders-board-menu");
const ordersBoardDropdown = document.querySelector("#orders-board-dropdown");
const openProductionOrdersBoardModalButton = document.querySelector("#open-production-orders-board-modal");
const openTransportOrdersBoardModalButton = document.querySelector("#open-transport-orders-board-modal");
const openPasswordModalButton = document.querySelector("#open-password-modal");
const openLoginModalButton = document.querySelector("#open-login-modal");
const logoutButton = document.querySelector("#logout-button");
const sourceCountPill = document.querySelector("#source-count-pill");
const sourceList = document.querySelector("#source-list");
const docCountPill = document.querySelector("#doc-count-pill");
const docList = document.querySelector("#doc-list");
const newChatButton = document.querySelector("#new-chat-button");
const openImportModalButton = document.querySelector("#open-import-modal");
const openDashboardModalButton = document.querySelector("#open-dashboard-modal");
const syncSheetButton = document.querySelector("#sync-sheet-button");
const openLibraryModalButton = document.querySelector("#open-library-modal");
const openProductionLibraryModalButton = document.querySelector("#open-production-library-modal");
const openSalesLibraryModalButton = document.querySelector("#open-sales-library-modal");
const openSalesProductsModalButton = document.querySelector("#open-sales-products-modal");
const openUsersModalButton = document.querySelector("#open-users-modal");
const openAuditModalButton = document.querySelector("#open-audit-modal");
const closeImportModalButton = document.querySelector("#close-import-modal");
const cancelImportButton = document.querySelector("#cancel-import-button");
const importModal = document.querySelector("#import-modal");
const importBackdrop = document.querySelector("#import-modal-backdrop");
const libraryModal = document.querySelector("#library-modal");
const libraryBackdrop = document.querySelector("#library-modal-backdrop");
const closeLibraryModalButton = document.querySelector("#close-library-modal");
const libraryModalKicker = document.querySelector("#library-modal-kicker");
const libraryModalTitle = document.querySelector("#library-modal-title");
const libraryToolbar = document.querySelector("#library-toolbar");
const librarySearch = document.querySelector("#library-search");
const libraryAccessFilter = document.querySelector("#library-access-filter");
const libraryFolderFilter = document.querySelector("#library-folder-filter");
const libraryStatusFilter = document.querySelector("#library-status-filter");
const libraryDocumentsWrap = document.querySelector("#library-documents-wrap");
const productionStockPanel = document.querySelector("#production-stock-panel");
const productionStockBody = document.querySelector("#production-stock-body");
const productionStockHead = document.querySelector("#production-stock-head");
const productionStockCount = document.querySelector("#production-stock-count");
const productionStockHint = document.querySelector("#production-stock-hint");
const productionStockViewMode = document.querySelector("#production-stock-view-mode");
const salesStockPanel = document.querySelector("#sales-stock-panel");
const salesStockBody = document.querySelector("#sales-stock-body");
const salesStockHead = document.querySelector("#sales-stock-head");
const salesStockCount = document.querySelector("#sales-stock-count");
const salesStockHint = document.querySelector("#sales-stock-hint");
const salesStockViewMode = document.querySelector("#sales-stock-view-mode");
const salesStockReceiveForm = document.querySelector("#sales-stock-receive-form");
const salesStockCodeInput = document.querySelector("#sales-stock-code");
const salesStockNameInput = document.querySelector("#sales-stock-name");
const salesStockUnitInput = document.querySelector("#sales-stock-unit");
const salesStockQuantityInput = document.querySelector("#sales-stock-quantity");
const salesStockSupplierInput = document.querySelector("#sales-stock-supplier");
const salesStockReceivedAtInput = document.querySelector("#sales-stock-received-at");
const salesStockNoteInput = document.querySelector("#sales-stock-note");
const salesStockReceiveButton = document.querySelector("#sales-stock-receive-button");
const salesStockTransferToggleButton = document.querySelector("#sales-stock-transfer-toggle");
const salesStockTransferPanel = document.querySelector("#sales-stock-transfer-panel");
const salesStockTransferAtInput = document.querySelector("#sales-stock-transfer-at");
const salesStockTransferNoteInput = document.querySelector("#sales-stock-transfer-note");
const salesStockTransferList = document.querySelector("#sales-stock-transfer-list");
const salesStockTransferConfirmButton = document.querySelector("#sales-stock-transfer-confirm");
const libraryTableBody = document.querySelector("#library-table-body");
const documentViewerModal = document.querySelector("#document-viewer-modal");
const documentViewerBackdrop = document.querySelector("#document-viewer-backdrop");
const closeDocumentViewerButton = document.querySelector("#close-document-viewer");
const closeDocumentViewerFooterButton = document.querySelector("#close-document-viewer-button");
const documentViewerTitle = document.querySelector("#document-viewer-title");
const documentViewerMeta = document.querySelector("#document-viewer-meta");
const documentViewerBody = document.querySelector("#document-viewer-body");
const auditModal = document.querySelector("#audit-modal");
const auditBackdrop = document.querySelector("#audit-modal-backdrop");
const closeAuditModalButton = document.querySelector("#close-audit-modal");
const auditLogList = document.querySelector("#audit-log-list");
const dashboardModal = document.querySelector("#dashboard-modal");
const dashboardBackdrop = document.querySelector("#dashboard-modal-backdrop");
const closeDashboardModalButton = document.querySelector("#close-dashboard-modal");
const dashboardStatsGrid = document.querySelector("#dashboard-stats-grid");
const dashboardFolders = document.querySelector("#dashboard-folders");
const dashboardUploaders = document.querySelector("#dashboard-uploaders");
const dashboardActions = document.querySelector("#dashboard-actions");
const shareModal = document.querySelector("#share-modal");
const shareBackdrop = document.querySelector("#share-modal-backdrop");
const closeShareModalButton = document.querySelector("#close-share-modal");
const cancelShareModalButton = document.querySelector("#cancel-share-modal");
const confirmShareModalButton = document.querySelector("#confirm-share-modal");
const shareModalTitle = document.querySelector("#share-modal-title");
const shareOwnerSelect = document.querySelector("#share-owner-select");
const shareUserSearch = document.querySelector("#share-user-search");
const shareModalUserList = document.querySelector("#share-modal-user-list");
const deliveryModal = document.querySelector("#delivery-modal");
const deliveryBackdrop = document.querySelector("#delivery-modal-backdrop");
const deliveryForm = document.querySelector("#delivery-form");
const closeDeliveryModalButton = document.querySelector("#close-delivery-modal");
const cancelDeliveryModalButton = document.querySelector("#cancel-delivery-modal");
const confirmDeliveryButton = document.querySelector("#confirm-delivery-button");
const deliveryOrderId = document.querySelector("#delivery-order-id");
const deliveryCustomerName = document.querySelector("#delivery-customer-name");
const deliveryAddress = document.querySelector("#delivery-address");
const deliverySalesUser = document.querySelector("#delivery-sales-user");
const deliveryResultStatus = document.querySelector("#delivery-result-status");
const deliveryCompletedAt = document.querySelector("#delivery-completed-at");
const deliveryPaymentStatus = document.querySelector("#delivery-payment-status");
const deliveryPaymentMethod = document.querySelector("#delivery-payment-method");
const deliveryNote = document.querySelector("#delivery-note");
const ordersBoardModal = document.querySelector("#orders-board-modal");
const ordersBoardBackdrop = document.querySelector("#orders-board-modal-backdrop");
const closeOrdersBoardModalButton = document.querySelector("#close-orders-board-modal");
const ordersBoardTitle = document.querySelector("#orders-board-title");
const ordersSearch = document.querySelector("#orders-search");
const ordersDateFromFilter = document.querySelector("#orders-date-from-filter");
const ordersDateToFilter = document.querySelector("#orders-date-to-filter");
const ordersStatusFilterLabel = document.querySelector("#orders-status-filter-label");
const ordersStatusFilter = document.querySelector("#orders-status-filter");
const ordersPaymentFilterField = document.querySelector("#orders-payment-filter-field");
const ordersPaymentFilter = document.querySelector("#orders-payment-filter");
const ordersToolbar = document.querySelector("#orders-board-modal .orders-toolbar");
const ordersTable = document.querySelector("#orders-board-modal .orders-table");
const ordersTableBody = document.querySelector("#orders-table-body");
const orderDetailsModal = document.querySelector("#order-details-modal");
const orderDetailsBackdrop = document.querySelector("#order-details-modal-backdrop");
const closeOrderDetailsModalButton = document.querySelector("#close-order-details-modal");
const closeOrderDetailsButton = document.querySelector("#close-order-details-button");
const printOrderDetailsButton = document.querySelector("#print-order-details-button");
const exportOrderDetailsButton = document.querySelector("#export-order-details-button");
const orderDetailsId = document.querySelector("#order-details-id");
const orderDetailsCustomerName = document.querySelector("#order-details-customer-name");
const orderDetailsSalesUser = document.querySelector("#order-details-sales-user");
const orderDetailsDeliveryUser = document.querySelector("#order-details-delivery-user");
const orderDetailsCreatedBy = document.querySelector("#order-details-created-by");
const orderDetailsCreatedAt = document.querySelector("#order-details-created-at");
const orderDetailsPlannedAt = document.querySelector("#order-details-planned-at");
const orderDetailsStatus = document.querySelector("#order-details-status");
const orderDetailsCompletedAt = document.querySelector("#order-details-completed-at");
const orderDetailsPaymentStatus = document.querySelector("#order-details-payment-status");
const orderDetailsPaymentMethod = document.querySelector("#order-details-payment-method");
const orderDetailsAddress = document.querySelector("#order-details-address");
const orderDetailsItemsList = document.querySelector("#order-details-items-list");
const orderDetailsItemsEmpty = document.querySelector("#order-details-items-empty");
const orderDetailsValue = document.querySelector("#order-details-value");
const orderDetailsTotalDisplay = document.querySelector("#order-details-total-display");
const orderDetailsNote = document.querySelector("#order-details-note");
const orderEditModal = document.querySelector("#order-edit-modal");
const orderEditBackdrop = document.querySelector("#order-edit-modal-backdrop");
const orderEditForm = document.querySelector("#order-edit-form");
const closeOrderEditModalButton = document.querySelector("#close-order-edit-modal");
const cancelOrderEditModalButton = document.querySelector("#cancel-order-edit-modal");
const confirmOrderEditButton = document.querySelector("#confirm-order-edit-button");
const orderEditId = document.querySelector("#order-edit-id");
const orderEditCustomerName = document.querySelector("#order-edit-customer-name");
const orderEditAddress = document.querySelector("#order-edit-address");
const orderEditItems = document.querySelector("#order-edit-items");
const orderEditValue = document.querySelector("#order-edit-value");
const orderEditPaymentStatus = document.querySelector("#order-edit-payment-status");
const orderEditPaymentMethod = document.querySelector("#order-edit-payment-method");
const orderEditNote = document.querySelector("#order-edit-note");
const orderProductsModal = document.querySelector("#order-products-modal");
const orderProductsBackdrop = document.querySelector("#order-products-modal-backdrop");
const closeOrderProductsModalButton = document.querySelector("#close-order-products-modal");
const cancelOrderProductsModalButton = document.querySelector("#cancel-order-products-modal");
const addOrderProductButton = document.querySelector("#add-order-product-button");
const saveOrderProductsButton = document.querySelector("#save-order-products-button");
const orderProductsList = document.querySelector("#order-products-list");
const orderModal = document.querySelector("#order-modal");
const orderBackdrop = document.querySelector("#order-modal-backdrop");
const productionOrderSheet = document.querySelector("#production-order-sheet");
const productionCompletionStamp = document.querySelector("#production-completion-stamp");
const tapeCalculatorModal = document.querySelector("#tape-calculator-modal");
const tapeCalculatorBackdrop = document.querySelector("#tape-calculator-modal-backdrop");
const closeTapeCalculatorModalButton = document.querySelector("#close-tape-calculator-modal");
const cancelTapeCalculatorModalButton = document.querySelector("#cancel-tape-calculator-modal");
const orderForm = document.querySelector("#order-form");
const orderModalTitle = orderModal?.querySelector(".modal-head h2");
const orderStepButtons = Array.from(document.querySelectorAll("[data-order-step]"));
const orderStepPanels = Array.from(document.querySelectorAll("[data-order-step-panel]"));
const productionOrderIdInput = document.querySelector("#production-order-id");
const productionCustomerNameInput = document.querySelector("#production-customer-name");
const productionRequesterField = document.querySelector("#production-requester-field");
const productionRequesterSelect = document.querySelector("#production-requester");
const productionRequesterSignature = document.querySelector("#production-requester-signature");
const productionPerformerSignature = document.querySelector("#production-performer-signature");
const productionPackagingSignature = document.querySelector("#production-packaging-signature");
const productionReceiptSignature = document.querySelector("#production-receipt-signature");
const productionCreatedAtInput = document.querySelector("#production-created-at");
const productionOrderItemsList = document.querySelector("#production-order-items-list");
const productionOrderItemsEmpty = document.querySelector("#production-order-items-empty");
const productionAddItemButton = document.querySelector("#production-add-item-button");
const productionNoteInput = document.querySelector("#production-note");
const productionTurnaroundOtherInput = document.querySelector("#production-turnaround-other");
const closeOrderModalButton = document.querySelector("#close-order-modal");
const cancelOrderModalButton = document.querySelector("#cancel-order-modal");
const printProductionOrderButton = document.querySelector("#print-production-order-button");
const receiveAllProductionOrderButton = document.querySelector("#receive-all-production-order-button");
const receivePartialProductionOrderButton = document.querySelector("#receive-partial-production-order-button");
const confirmPartialProductionOrderButton = document.querySelector("#confirm-partial-production-order-button");
const confirmProductionProgressButton = document.querySelector("#confirm-production-progress-button");
const completeProductionProgressButton = document.querySelector("#complete-production-progress-button");
const completeProductionPackagingButton = document.querySelector("#complete-production-packaging-button");
const completeProductionReceiptButton = document.querySelector("#complete-production-receipt-button");
const editCompletedProductionOrderButton = document.querySelector("#edit-completed-production-order-button");
const productionCompleteChip = document.querySelector("#production-complete-chip");
const productionPackagingConfirmModal = document.querySelector("#production-packaging-confirm-modal");
const productionPackagingConfirmBackdrop = document.querySelector("#production-packaging-confirm-backdrop");
const closeProductionPackagingConfirmModalButton = document.querySelector("#close-production-packaging-confirm-modal");
const cancelProductionPackagingConfirmButton = document.querySelector("#cancel-production-packaging-confirm-button");
const confirmProductionPackagingCompleteButton = document.querySelector("#confirm-production-packaging-complete-button");
const productionReceiptConfirmModal = document.querySelector("#production-receipt-confirm-modal");
const productionReceiptConfirmBackdrop = document.querySelector("#production-receipt-confirm-backdrop");
const closeProductionReceiptConfirmModalButton = document.querySelector("#close-production-receipt-confirm-modal");
const cancelProductionReceiptConfirmButton = document.querySelector("#cancel-production-receipt-confirm-button");
const confirmProductionReceiptCompleteButton = document.querySelector("#confirm-production-receipt-complete-button");
const confirmOrderButton = document.querySelector("#confirm-order-button");
const ordersAssigneeHeading = document.querySelector("#orders-assignee-heading");
const ordersProductionStatusHeading = document.querySelector("#orders-production-status-heading");
const ordersDeliveryHeading = document.querySelector("#orders-delivery-heading");
const ordersPaymentHeading = document.querySelector("#orders-payment-heading");
const orderIdInput = document.querySelector("#order-id");
const orderCustomerName = document.querySelector("#order-customer-name");
const orderSalesUser = document.querySelector("#order-sales-user");
const orderDeliveryUser = document.querySelector("#order-delivery-user");
const orderPlannedAt = document.querySelector("#order-planned-at");
const orderCreatedAt = document.querySelector("#order-created-at");
const orderAddress = document.querySelector("#order-address");
const orderAddItemButton = document.querySelector("#order-add-item-button");
const checkOrderInventoryButton = document.querySelector("#check-order-inventory-button");
const orderItemsEmpty = document.querySelector("#order-items-empty");
const orderItemsList = document.querySelector("#order-items-list");
const orderItems = document.querySelector("#order-items");
const orderValue = document.querySelector("#order-value");
const orderTotalDisplay = document.querySelector("#order-total-display");
const orderNote = document.querySelector("#order-note");
const tapeType = document.querySelector("#tape-type");
const tapeOrderQuantity = document.querySelector("#tape-order-quantity");
const tapeCoreType = document.querySelector("#tape-core-type");
const tapePackaging = document.querySelector("#tape-packaging");
const tapeFinishedQuantity = document.querySelector("#tape-finished-quantity");
const tapeJumboHeight = document.querySelector("#tape-jumbo-height");
const tapeCoreWidth = document.querySelector("#tape-core-width");
const tapeRollsPerHand = document.querySelector("#tape-rolls-per-hand");
const tapeRemainingMm = document.querySelector("#tape-remaining-mm");
const tapeHandsNeeded = document.querySelector("#tape-hands-needed");
const tapeTotalProduced = document.querySelector("#tape-total-produced");
const tapeExtraProduced = document.querySelector("#tape-extra-produced");
const loginModal = document.querySelector("#login-modal");
const loginBackdrop = document.querySelector("#login-modal-backdrop");
const loginForm = document.querySelector("#login-form");
const loginUsername = document.querySelector("#login-username");
const loginPassword = document.querySelector("#login-password");
const loginOtpField = document.querySelector("#login-otp-field");
const loginOtp = document.querySelector("#login-otp");
const loginOtpHelp = document.querySelector("#login-otp-help");
const loginSubmitButton = document.querySelector("#login-submit-button");
const importForm = document.querySelector("#import-form");
const importTitle = document.querySelector("#import-title");
const importStorageLevel = document.querySelector("#import-storage-level");
const importFolder = document.querySelector("#import-folder");
const importCustomFolderField = document.querySelector("#import-custom-folder-field");
const importCustomFolder = document.querySelector("#import-custom-folder");
const privateSharePanel = document.querySelector("#private-share-panel");
const privateShareList = document.querySelector("#private-share-list");
const importOwner = document.querySelector("#import-owner");
const importOwnerField = importOwner?.closest(".field");
const importStatus = document.querySelector("#import-status");
const importStatusField = importStatus?.closest(".field");
const importTopicKey = document.querySelector("#import-topic-key");
const importTopicKeyField = importTopicKey?.closest(".field");
const importFile = document.querySelector("#import-file");
const importSheetUrl = document.querySelector("#import-sheet-url");
const confirmImportButton = document.querySelector("#confirm-import-button");
const permissionsModal = document.querySelector("#permissions-modal");
const permissionsBackdrop = document.querySelector("#permissions-modal-backdrop");
const permissionsList = document.querySelector("#permissions-list");
const permissionsUserSearch = document.querySelector("#permissions-user-search");
const createUserButton = document.querySelector("#create-user-button");
const closePermissionsModalButton = document.querySelector("#close-permissions-modal");
const closePermissionsButton = document.querySelector("#close-permissions-button");
const passwordModal = document.querySelector("#password-modal");
const passwordBackdrop = document.querySelector("#password-modal-backdrop");
const passwordForm = document.querySelector("#password-form");
const closePasswordModalButton = document.querySelector("#close-password-modal");
const cancelPasswordButton = document.querySelector("#cancel-password-button");
const currentPasswordInput = document.querySelector("#current-password");
const newPasswordInput = document.querySelector("#new-password");
const confirmPasswordInput = document.querySelector("#confirm-password");
const confirmPasswordButton = document.querySelector("#confirm-password-button");
const fileImportPanel = document.querySelector("#file-import-panel");
const sheetImportPanel = document.querySelector("#sheet-import-panel");
const checkSheetButton = document.querySelector("#check-sheet-button");
const sheetCheckResult = document.querySelector("#sheet-check-result");
const toast = document.querySelector("#toast");
const askAccessFilter = document.querySelector("#ask-access-filter");
const askFolderFilter = document.querySelector("#ask-folder-filter");
const askFilterHint = document.querySelector("#ask-filter-hint");
const libraryFilterHint = document.querySelector("#library-filter-hint");
const composerAttachments = document.querySelector("#composer-attachments");
const pastedImagePreview = document.querySelector("#pasted-image-preview");
const pastedImageName = document.querySelector("#pasted-image-name");
const pastedImageSize = document.querySelector("#pasted-image-size");
const removePastedImageButton = document.querySelector("#remove-pasted-image");
const ocrPreview = document.querySelector("#ocr-preview");
const ocrPreviewBody = document.querySelector("#ocr-preview-body");
const refreshOcrButton = document.querySelector("#refresh-ocr-button");
const dropHint = document.querySelector("#drop-hint");
const ocrDecision = document.querySelector("#ocr-decision");
const confirmOcrButton = document.querySelector("#confirm-ocr-button");
const skipOcrButton = document.querySelector("#skip-ocr-button");

const conversationHistory = [];
const NOTIFICATION_ARCHIVE_RETENTION_MS = 7 * 24 * 60 * 60 * 1000;
let pendingAssistantMessage = null;
let importMode = "file";
let toastTimer = null;
let pastedImage = null;
let dragDepth = 0;
let isSheetValidated = false;
let authToken = localStorage.getItem("authToken") || "";
let currentUser = null;
let allUsers = [];
let pendingLoginChallenge = "";
let selectedPermissionUserId = "";
let permissionsUserSearchValue = "";
let isCreatingManagedUser = false;
let documentLibrary = [];
let sourceLibrary = [];
let auditEntries = [];
let notifications = [];
let notificationArchive = [];
let notificationArchiveClearedAt = 0;
let notificationFilterValue = "all";
let orders = [];
let productionStockSummary = [];
let productionStockView = "summary";
let salesInventoryReceipts = [];
let salesInventoryTransfers = [];
let salesStockSummary = [];
let salesStockView = "summary";
let isSalesStockTransferPanelOpen = false;
let activeEditableOrderId = "";
let activeEditableOrderRecord = null;
let activeOrderDraftEditId = "";
let activeOrderDraftEditRecord = null;
let activeOrderCreateKind = "production";
let activeOrdersBoardKind = "production";
let isOrderModalReadOnly = false;
let activeProductionClaimMode = "";
let hasShownSessionExpiredToast = false;
let selectedHistoryQuestion = "";
let dashboardStats = null;
let activePrivateDocument = null;
let notificationPollTimer = null;
let notificationStream = null;
let hasBootstrappedNotifications = false;
let activeNotificationOnViewerClose = null;
let activeOrderDetails = null;
let notificationUiStateOwnerKey = "";
const hiddenNotificationIds = new Set();
const notificationAutoHideTimers = new Map();
let activeLibraryPreset = "general";
const seenNotificationIds = new Set();
let orderProducts = [];
const OPEN_ORDER_PREVIEW_STORAGE_KEY = "open-order-preview";

const STORAGE_FOLDER_OPTIONS = {
  basic: [
    { value: "internal", label: "internal" },
    { value: "van-chuyen", label: "Vận Chuyển" },
  ],
  advanced: [
    { value: "ke-toan", label: "Kế Toán" },
    { value: "phong-kinh-doanh", label: "phòng kinh doanh" },
    { value: "phong-san-xuat", label: "phòng Sản Xuất" },
  ],
};

function getDefaultOrderProducts() {
  return [
    { id: "foam-10", name: "Xốp nổ 10cm", units: ["cuộn"], default_unit: "cuộn" },
    { id: "foam-20", name: "Xốp nổ 20cm", units: ["cuộn"], default_unit: "cuộn" },
    { id: "foam-30", name: "Xốp nổ 30cm", units: ["cuộn"], default_unit: "cuộn" },
    { id: "foam-40", name: "Xốp nổ 40cm", units: ["cuộn"], default_unit: "cuộn" },
    { id: "foam-50", name: "Xốp nổ 50cm", units: ["cuộn"], default_unit: "cuộn" },
    { id: "foam-60", name: "Xốp nổ 60cm", units: ["cuộn"], default_unit: "cuộn" },
    { id: "tape-1kg", name: "Băng dính vàng 1kg", units: ["cuộn"], default_unit: "cuộn" },
    { id: "tape-500g", name: "Băng dính vàng 500g", units: ["cuộn"], default_unit: "cuộn" },
    { id: "core-n3", name: "Lõi nhựa N3", units: ["cây"], default_unit: "cây" },
    { id: "core-n5", name: "Lõi nhựa N5", units: ["cây"], default_unit: "cây" },
    { id: "core-n6", name: "Lõi nhựa N6", units: ["cây"], default_unit: "cây" },
  ];
}

function getSortedOrderProducts(products = null) {
  const source = Array.isArray(products) ? products : orderProducts.length ? orderProducts : getDefaultOrderProducts();
  return [...source].sort((left, right) =>
    String(left?.name || "").localeCompare(String(right?.name || ""), "vi", { numeric: true, sensitivity: "base" }),
  );
}

const DOCUMENT_SECTION_TEMPLATE = [
  {
    level: "basic",
    folders: ["internal", "Vận Chuyển"],
  },
  {
    level: "advanced",
    folders: ["Kế Toán", "phòng kinh doanh", "phòng Sản Xuất"],
  },
  {
    level: "sensitive",
    folders: ["dữ liệu cá nhân"],
  },
];

const SOURCE_SECTION_TEMPLATE = [
  { level: "basic" },
  { level: "advanced" },
];
const dynamicFolderOptions = {
  basic: [],
  advanced: [],
};
const customDocumentFolders = {
  basic: [],
  advanced: [],
};
const ACCESS_LEVELS = ["basic", "advanced", "sensitive"];

newChatButton.addEventListener("click", () => {
  conversationHistory.length = 0;
  selectedHistoryQuestion = "";
  pendingAssistantMessage = null;
  chatThread.innerHTML = "";
  chatThread.append(emptyState);
  emptyState.classList.remove("hidden");
  renderHistory();
  clearPastedImage();
  questionInput.value = "";
  questionInput.focus();
});

for (const button of document.querySelectorAll(".tab-button")) {
  button.addEventListener("click", () => setImportMode(button.dataset.mode || "file"));
}

openImportModalButton.addEventListener("click", openImportModal);
openDashboardModalButton?.addEventListener("click", openDashboardModal);
syncSheetButton.addEventListener("click", syncSheetDocumentsNow);
openLibraryModalButton?.addEventListener("click", () => {
  openLibraryModal("general");
});
openProductionLibraryModalButton?.addEventListener("click", () => {
  openLibraryModal("production");
});
openSalesLibraryModalButton?.addEventListener("click", () => {
  openLibraryModal("sales");
});
openAuditModalButton?.addEventListener("click", openAuditModal);
openSalesProductsModalButton?.addEventListener("click", openOrderProductsModal);
openUsersModalButton?.addEventListener("click", openPermissionsModal);
openPermissionsModalButton.addEventListener("click", openPermissionsModal);
openDeliveryModalButton?.addEventListener("click", openDeliveryModal);
openOrderMenuButton?.addEventListener("click", (event) => {
  event.preventDefault();
  event.stopPropagation();
  setOrderCtaMenuOpen(!orderCtaMenu?.classList.contains("is-open"));
});
openProductionOrderModalButton?.addEventListener("click", async () => {
  setOrderCtaMenuOpen(false);
  await openOrderModal("production");
});
openTransportOrderModalButton?.addEventListener("click", async () => {
  setOrderCtaMenuOpen(false);
  await openOrderModal("transport");
});
orderCtaMenu?.addEventListener("focusout", (event) => {
  if (!orderCtaMenu.contains(event.relatedTarget)) {
    setOrderCtaMenuOpen(false);
  }
});
document.addEventListener("pointerdown", (event) => {
  if (!orderCtaMenu?.contains(event.target)) {
    setOrderCtaMenuOpen(false);
  }
  if (!ordersBoardMenu?.contains(event.target)) {
    setOrdersBoardMenuOpen(false);
  }
});
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") {
    setOrderCtaMenuOpen(false);
    setOrdersBoardMenuOpen(false);
  }
});
window.addEventListener("resize", () => {
  if (orderCtaMenu?.classList.contains("is-open")) {
    positionOrderCtaDropdown();
  }
  if (ordersBoardMenu?.classList.contains("is-open")) {
    positionOrdersBoardDropdown();
  }
});
window.addEventListener("scroll", () => {
  if (orderCtaMenu?.classList.contains("is-open")) {
    positionOrderCtaDropdown();
  }
  if (ordersBoardMenu?.classList.contains("is-open")) {
    positionOrdersBoardDropdown();
  }
}, { passive: true });
openOrderProductsModalButton?.addEventListener("click", openOrderProductsModal);
openTapeCalculatorModalButton?.addEventListener("click", openTapeCalculatorModal);
openOrdersBoardMenuButton?.addEventListener("click", (event) => {
  event.preventDefault();
  event.stopPropagation();
  setOrdersBoardMenuOpen(!ordersBoardMenu?.classList.contains("is-open"));
});
openProductionOrdersBoardModalButton?.addEventListener("click", () => {
  setOrdersBoardMenuOpen(false);
  openOrdersBoardModal("production");
});
openTransportOrdersBoardModalButton?.addEventListener("click", () => {
  setOrdersBoardMenuOpen(false);
  openOrdersBoardModal("transport");
});
openLoginModalButton?.addEventListener("click", showLoginModal);
openPasswordModalButton.addEventListener("click", openPasswordModal);
logoutButton.addEventListener("click", logoutCurrentUser);
closeImportModalButton.addEventListener("click", closeImportModal);
cancelImportButton.addEventListener("click", closeImportModal);
closeDashboardModalButton?.addEventListener("click", closeDashboardModal);
closeLibraryModalButton?.addEventListener("click", closeLibraryModal);
closeDocumentViewerButton?.addEventListener("click", closeDocumentViewerModal);
closeDocumentViewerFooterButton?.addEventListener("click", closeDocumentViewerModal);
closeAuditModalButton?.addEventListener("click", closeAuditModal);
closeShareModalButton?.addEventListener("click", closeShareModal);
cancelShareModalButton?.addEventListener("click", closeShareModal);
closeDeliveryModalButton?.addEventListener("click", closeDeliveryModal);
cancelDeliveryModalButton?.addEventListener("click", closeDeliveryModal);
closeTapeCalculatorModalButton?.addEventListener("click", closeTapeCalculatorModal);
cancelTapeCalculatorModalButton?.addEventListener("click", closeTapeCalculatorModal);
closeOrdersBoardModalButton?.addEventListener("click", closeOrdersBoardModal);
closeOrderProductsModalButton?.addEventListener("click", closeOrderProductsModal);
cancelOrderProductsModalButton?.addEventListener("click", closeOrderProductsModal);
addOrderProductButton?.addEventListener("click", () => appendOrderProductEditorRow());
saveOrderProductsButton?.addEventListener("click", saveOrderProductsCatalog);
closeOrderDetailsModalButton?.addEventListener("click", closeOrderDetailsModal);
closeOrderDetailsButton?.addEventListener("click", closeOrderDetailsModal);
printOrderDetailsButton?.addEventListener("click", printActiveOrderDetails);
exportOrderDetailsButton?.addEventListener("click", exportActiveOrderDetails);
closeOrderEditModalButton?.addEventListener("click", closeOrderEditModal);
cancelOrderEditModalButton?.addEventListener("click", closeOrderEditModal);
closeOrderModalButton?.addEventListener("click", closeOrderModal);
cancelOrderModalButton?.addEventListener("click", closeOrderModal);
printProductionOrderButton?.addEventListener("click", printProductionOrderSheet);
receiveAllProductionOrderButton?.addEventListener("click", () => {
  submitProductionClaim("all").catch((error) => {
    showToast(error.message || "Không thể nhận toàn bộ phiếu sản xuất.", "error");
  });
});
receivePartialProductionOrderButton?.addEventListener("click", () => {
  setProductionClaimMode(activeProductionClaimMode === "partial" ? "" : "partial");
});
confirmPartialProductionOrderButton?.addEventListener("click", () => {
  submitProductionClaim("partial").catch((error) => {
    showToast(error.message || "Không thể nhận một phần phiếu sản xuất.", "error");
  });
});
confirmProductionProgressButton?.addEventListener("click", () => {
  submitProductionProgress("confirm").catch((error) => {
    showToast(error.message || "Không thể xác nhận tiến độ sản xuất.", "error");
  });
});
completeProductionProgressButton?.addEventListener("click", () => {
  submitProductionProgress("complete").catch((error) => {
    showToast(error.message || "Không thể xác nhận hoàn thành sản xuất.", "error");
  });
});
completeProductionPackagingButton?.addEventListener("click", openProductionPackagingConfirmModal);
closeProductionPackagingConfirmModalButton?.addEventListener("click", closeProductionPackagingConfirmModal);
cancelProductionPackagingConfirmButton?.addEventListener("click", closeProductionPackagingConfirmModal);
productionPackagingConfirmBackdrop?.addEventListener("click", closeProductionPackagingConfirmModal);
confirmProductionPackagingCompleteButton?.addEventListener("click", () => {
  submitProductionPackagingComplete().catch((error) => {
    showToast(error.message || "Không thể hoàn tất đóng gói.", "error");
  });
});
completeProductionReceiptButton?.addEventListener("click", openProductionReceiptConfirmModal);
closeProductionReceiptConfirmModalButton?.addEventListener("click", closeProductionReceiptConfirmModal);
cancelProductionReceiptConfirmButton?.addEventListener("click", closeProductionReceiptConfirmModal);
productionReceiptConfirmBackdrop?.addEventListener("click", closeProductionReceiptConfirmModal);
editCompletedProductionOrderButton?.addEventListener("click", () => {
  if (!activeOrderDetails) {
    return;
  }
  openProductionOrderEditModal(activeOrderDetails).catch((error) => {
    showToast(error.message || "Không thể mở phiếu sản xuất để sửa.", "error");
  });
});
confirmProductionReceiptCompleteButton?.addEventListener("click", () => {
  submitProductionReceiptComplete().catch((error) => {
    showToast(error.message || "Không thể xác nhận nhận đủ hàng.", "error");
  });
});
orderStepButtons.forEach((button) => {
  button.addEventListener("click", () => setOrderFormStep(button.dataset.orderStep || "1"));
});
deliveryOrderId?.addEventListener("change", autofillDeliveryOrderDetails);
closePermissionsModalButton.addEventListener("click", closePermissionsModal);
closePermissionsButton.addEventListener("click", closePermissionsModal);
permissionsUserSearch?.addEventListener("input", () => {
  permissionsUserSearchValue = String(permissionsUserSearch.value || "").trim();
  renderPermissionsPanel();
});
createUserButton?.addEventListener("click", () => {
  isCreatingManagedUser = true;
  selectedPermissionUserId = "";
  renderPermissionsPanel();
});
permissionsList?.addEventListener("change", handleCreateUserPresetChange);
permissionsList?.addEventListener("change", handleEditableUserProfilePresetChange);
permissionsList?.addEventListener("click", handlePermissionDirectoryClick);
permissionsList?.addEventListener("keydown", handlePermissionDirectoryKeydown);
closePasswordModalButton.addEventListener("click", closePasswordModal);
cancelPasswordButton.addEventListener("click", closePasswordModal);
removePastedImageButton.addEventListener("click", clearPastedImage);
checkSheetButton.addEventListener("click", validateSheetLink);
loginForm.addEventListener("submit", handleLoginSubmit);
passwordForm.addEventListener("submit", handlePasswordSubmit);
deliveryForm?.addEventListener("submit", submitDeliveryCompletion);
orderForm?.addEventListener("submit", submitOrderCreate);
orderItemsList?.addEventListener("input", syncOrderDraftTotals);
orderItemsList?.addEventListener("change", syncOrderDraftTotals);
orderItemsList?.addEventListener("input", handleOrderItemNameInput);
orderItemsList?.addEventListener("focusin", handleOrderItemNameFocus);
orderItemsList?.addEventListener("keydown", handleOrderItemNameKeydown);
orderItemsList?.addEventListener("click", handleOrderItemSuggestionClick);
orderItemsList?.addEventListener("click", handleOrderItemToggleClick);
orderItemsList?.addEventListener("input", handleOrderItemSearchInput);
orderItemsList?.addEventListener("keydown", handleOrderItemSearchKeydown);
document.addEventListener("pointerdown", handleOrderItemSuggestionOutsideClick);
orderAddItemButton?.addEventListener("click", () => addOrderItemRow());
productionOrderItemsList?.addEventListener("input", syncProductionOrderDraft);
productionOrderItemsList?.addEventListener("change", syncProductionOrderDraft);
productionOrderItemsList?.addEventListener("input", handleOrderItemNameInput);
productionOrderItemsList?.addEventListener("focusin", handleOrderItemNameFocus);
productionOrderItemsList?.addEventListener("keydown", handleOrderItemNameKeydown);
productionOrderItemsList?.addEventListener("click", handleOrderItemSuggestionClick);
productionOrderItemsList?.addEventListener("click", handleOrderItemToggleClick);
productionOrderItemsList?.addEventListener("click", handleProductionOrderNameClick);
productionOrderItemsList?.addEventListener("click", handleProductionClaimRowClick);
productionOrderItemsList?.addEventListener("change", handleProductionClaimCheckboxChange);
productionOrderItemsList?.addEventListener("input", handleOrderItemSearchInput);
productionOrderItemsList?.addEventListener("keydown", handleOrderItemSearchKeydown);
productionAddItemButton?.addEventListener("click", () => addProductionOrderItemRow());
productionRequesterSelect?.addEventListener("change", syncProductionRequesterSignature);
orderEditForm?.addEventListener("submit", submitOrderUpdate);
populateTapeCalculatorOptions();
syncOrderDraftTotals();
[tapeOrderQuantity, tapePackaging, tapeJumboHeight, tapeCoreWidth, tapeType, tapeCoreType].forEach((field) => {
  field?.addEventListener("input", handleTapeCalculatorInput);
  field?.addEventListener("focus", handleTapeCalculatorFocus);
  field?.addEventListener("keydown", handleTapeCalculatorKeydown);
  field?.addEventListener("change", updateTapeCalculatorResults);
});
document.addEventListener("click", handleTapeCalculatorSuggestionClick);
document.addEventListener("click", handleTapeCalculatorToggleClick);
document.addEventListener("pointerdown", handleTapeCalculatorOutsideClick);
updateTapeCalculatorResults();
notificationFilter?.addEventListener("change", () => {
  notificationFilterValue = notificationFilter.value || "all";
  renderNotifications();
});
markAllNotificationsReadButton?.addEventListener("click", markAllNotificationsRead);
clearNotificationHistoryButton?.addEventListener("click", clearNotificationHistory);
ordersSearch?.addEventListener("input", renderOrdersBoard);
ordersDateFromFilter?.addEventListener("change", renderOrdersBoard);
ordersDateToFilter?.addEventListener("change", renderOrdersBoard);
ordersStatusFilter?.addEventListener("change", renderOrdersBoard);
ordersPaymentFilter?.addEventListener("change", renderOrdersBoard);
librarySearch?.addEventListener("input", renderDocumentLibrary);
libraryAccessFilter?.addEventListener("change", () => {
  applyAccessFilterPolicy();
  syncFolderFilterOptions();
  renderDocumentLibrary();
});
libraryFolderFilter?.addEventListener("change", renderDocumentLibrary);
libraryStatusFilter?.addEventListener("change", renderDocumentLibrary);
productionStockViewMode?.addEventListener("change", () => {
  productionStockView = String(productionStockViewMode.value || "summary").trim().toLowerCase();
  renderProductionStockPanel();
});
salesStockViewMode?.addEventListener("change", () => {
  salesStockView = String(salesStockViewMode.value || "summary").trim().toLowerCase();
  renderSalesStockPanel();
});
salesStockReceiveForm?.addEventListener("submit", submitSalesStockReceive);
salesStockTransferToggleButton?.addEventListener("click", toggleSalesStockTransferPanel);
salesStockTransferConfirmButton?.addEventListener("click", submitSalesStockTransferFromProduction);
checkOrderInventoryButton?.addEventListener("click", () => {
  checkTransportOrderInventory({ showSuccessToast: true });
});
askAccessFilter?.addEventListener("change", () => {
  applyAccessFilterPolicy();
  syncFolderFilterOptions();
});
shareUserSearch?.addEventListener("input", renderShareModalUsers);
shareOwnerSelect?.addEventListener("change", renderShareModalUsers);
confirmShareModalButton?.addEventListener("click", submitShareModal);
refreshOcrButton.addEventListener("click", () => {
  if (pastedImage) {
    populateOcrPreview(true).catch((error) => {
      showToast(error.message || "Không thể đọc lại nội dung ảnh.", "error");
    });
  }
});
confirmOcrButton.addEventListener("click", () => {
  if (pastedImage) {
    populateOcrPreview(true).catch((error) => {
      showToast(error.message || "Không thể đọc nội dung ảnh.", "error");
    });
  }
});
skipOcrButton.addEventListener("click", () => {
  if (!pastedImage) {
    return;
  }

  pastedImage.extractedText = "";
  pastedImage.ocrSkipped = true;
  ocrDecision.classList.add("hidden");
  ocrPreview.classList.add("hidden");
  showToast("Ảnh sẽ được gửi mà không đọc nội dung.", "success");
});
importTitle.addEventListener("input", invalidateSheetValidation);
importSheetUrl.addEventListener("input", invalidateSheetValidation);
importOwner.addEventListener("input", invalidateSheetValidation);
importStorageLevel.addEventListener("change", () => {
  updateFolderOptions();
  renderPrivateShareOptions();
  invalidateSheetValidation();
});
importFolder.addEventListener("change", () => {
  importOwner.value = resolveFolderOwnerLabel(importFolder.value);
  toggleCustomFolderInput();
  invalidateSheetValidation();
});
importCustomFolder?.addEventListener("input", invalidateSheetValidation);

questionInput.addEventListener("input", autoResizeTextarea);
questionInput.addEventListener("paste", handleComposerPaste);
questionInput.addEventListener("dragenter", handleComposerDragEnter);
questionInput.addEventListener("dragover", handleComposerDragOver);
questionInput.addEventListener("dragleave", handleComposerDragLeave);
questionInput.addEventListener("drop", handleComposerDrop);
questionInput.addEventListener("keydown", (event) => {
  if (event.key === "Enter" && !event.shiftKey) {
    event.preventDefault();
    askForm.requestSubmit();
  }
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && !importModal.classList.contains("hidden")) {
    closeImportModal();
  }
});

askForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  const question = questionInput.value.trim();

  if (!question && !pastedImage) {
    questionInput.focus();
    return;
  }

  appendUserMessage(question, pastedImage);
  selectedHistoryQuestion = question || "Ảnh đã dán";
  setLoadingState(true);
  questionInput.value = "";
  autoResizeTextarea();
  pendingAssistantMessage = appendAssistantLoadingMessage();

  try {
    const result = await postJson(askUrl, {
      question,
      filters: {
        access_level: askAccessFilter.value.trim(),
        folder: askFolderFilter.value.trim(),
      },
      image: pastedImage
        ? {
            mimeType: pastedImage.mimeType,
            dataBase64: pastedImage.dataBase64,
            extractedText: pastedImage.extractedText || "",
          }
        : null,
    });
    hydrateAssistantMessage(pendingAssistantMessage, result);
    conversationHistory.unshift({
      question: question || "Ảnh đã dán",
      route: result.route || "internal",
      answer: result.answer || "",
    });
    conversationHistory.splice(8);
    renderHistory();
    clearPastedImage();
  } catch (error) {
    hydrateAssistantMessage(pendingAssistantMessage, {
      route: "error",
      answer: error.message || "Không thể xử lý câu hỏi.",
      sources: [],
    });
  } finally {
    pendingAssistantMessage = null;
    setLoadingState(false);
  }
});

importForm.addEventListener("submit", async (event) => {
  event.preventDefault();

  const title = importTitle.value.trim();
  const effective_date = getTodayDateValue();
  const status = "active";
  const topic_key = buildTopicKey(title);
  const storage_level = String(importStorageLevel.value || "basic").trim().toLowerCase();
  const storage_folder_key = importFolder.value.trim();
  const custom_folder_name = importCustomFolder.value.trim();
  const shared_with_users = getSelectedSharedUsers().join(",");
  if (!title) {
    importTitle.focus();
    return;
  }
  if (!storage_folder_key || (isCustomFolderOption(storage_folder_key) && !custom_folder_name)) {
    showToast("Hãy nhập đủ metadata bắt buộc cho tài liệu.", "error");
    return;
  }

  confirmImportButton.disabled = true;

  try {
    const payload = {
      title,
      effective_date,
      status,
      topic_key,
      storage_level,
      storage_folder_key,
      custom_folder_name,
      shared_with_users,
    };

    if (importMode === "file") {
      const file = importFile.files?.[0];
      if (!file) {
        importFile.focus();
        throw new Error("Hãy chọn một file để tải lên.");
      }

      const result = await postJson(docsUrl, {
        ...payload,
        sourceType: "file",
        fileName: file.name,
        contentBase64: await fileToBase64(file),
      });
      applyWorkspacePayload(result);
      closeImportModal();
      showToast(
        result.saved?.metadata?.status === "draft"
          ? `Đã gửi chờ duyệt: ${result.saved?.title || title}`
          : `Đã thêm tài liệu: ${result.saved?.title || title}`,
        "success",
      );
    } else {
      const sheetUrl = importSheetUrl.value.trim();
      if (!sheetUrl) {
        importSheetUrl.focus();
        throw new Error("Hãy nhập link Google Sheet.");
      }
      if (!isSheetValidated) {
        throw new Error("Hãy kiểm tra link Google Sheet trước khi lưu.");
      }

      const result = await postJson(sourcesUrl, {
        ...payload,
        sheetUrl,
      });
      applyWorkspacePayload(result);
      closeImportModal();
      showToast(
        result.saved?.source?.status === "draft"
          ? `Đã gửi nguồn chờ duyệt: ${result.saved?.source?.title || title}`
          : `Đã thêm nguồn đồng bộ: ${result.saved?.source?.title || title}`,
        "success",
      );
    }
  } catch (error) {
    showToast(error.message || "Không thể thêm tài liệu.", "error");
  } finally {
    confirmImportButton.disabled = false;
  }
});

void bootApp();

async function bootApp() {
  await initializeAuth();
  autoResizeTextarea();
  setImportMode("file");
  updateFolderOptions();
  renderPrivateShareOptions();
  importOwnerField?.classList.add("hidden");
  importStatusField?.classList.add("hidden");
  importTopicKeyField?.classList.add("hidden");
}

async function loadHealth() {
  try {
    const health = await fetchJson(healthUrl);
    systemStatus.innerHTML = "";
    const pills = [
      createMiniPill(`${health.sourceCount || 0} nguồn`, "neutral"),
      createMiniPill(`${health.documentCount} tài liệu`, "neutral"),
    ];

    if (health.sheetSync?.lastFinishedAt) {
      pills.push(createMiniPill(`Sheet ${health.sheetSync.lastFinishedAt}`, "neutral"));
    }

    systemStatus.append(...pills);
  } catch {
    systemStatus.innerHTML = "";
    systemStatus.append(createMiniPill("Không lấy được trạng thái", "neutral"));
  }
}

async function initializeAuth() {
  const session = await fetchSession();
  if (!session.authenticated) {
    stopNotificationStream();
    stopNotificationPolling();
    showLoginModal();
    return;
  }

  currentUser = session.currentUser || null;
  await loadUsers();
  await loadOrderProducts();
  hideLoginModal();
  await Promise.all([loadHealth(), loadSources(), loadDocuments(), loadNotifications()]);
  startNotificationStream();
  startNotificationPolling();
  await restoreOpenOrderPreviewIfNeeded();
}

async function fetchSession() {
  try {
    const session = await fetchJson(sessionUrl);
    if (!session?.authenticated) {
      authToken = "";
      localStorage.removeItem("authToken");
    }
    return session;
  } catch {
    authToken = "";
    localStorage.removeItem("authToken");
    return { authenticated: false, currentUser: null };
  }
}

async function loadUsers() {
  const payload = await fetchJson(usersUrl);
  currentUser = payload.currentUser || null;
  allUsers = payload.users || [];
  renderCurrentUser();
  renderPermissionsPanel();
  renderPrivateShareOptions();
}

async function loadNotifications() {
  ensureNotificationUiStateLoaded();
  try {
    const payload = await fetchJson(notificationsUrl);
    const nextNotifications = payload.notifications || [];
    announceIncomingNotifications(nextNotifications);
    notifications = nextNotifications;
    mergeNotificationArchiveItems(nextNotifications);
    renderNotifications();
  } catch {
    notifications = [];
    renderNotifications();
  }
}

function persistOpenOrderPreview(order) {
  try {
    const orderId = String(order?.order_id || "").trim();
    if (!orderId) {
      return;
    }
    sessionStorage.setItem(
      OPEN_ORDER_PREVIEW_STORAGE_KEY,
      JSON.stringify({
        order_id: orderId,
      }),
    );
  } catch {}
}

function clearPersistedOpenOrderPreview() {
  try {
    sessionStorage.removeItem(OPEN_ORDER_PREVIEW_STORAGE_KEY);
  } catch {}
}

function getPersistedOpenOrderPreviewId() {
  try {
    const raw = sessionStorage.getItem(OPEN_ORDER_PREVIEW_STORAGE_KEY);
    if (!raw) {
      return "";
    }
    const parsed = JSON.parse(raw);
    return String(parsed?.order_id || "").trim();
  } catch {
    return "";
  }
}

async function restoreOpenOrderPreviewIfNeeded() {
  const orderId = getPersistedOpenOrderPreviewId();
  if (!orderId || !canViewOrders()) {
    return;
  }
  await loadOrders();
  const matchedOrder = orders.find((order) => String(order?.order_id || "").trim() === orderId);
  if (!matchedOrder) {
    clearPersistedOpenOrderPreview();
    return;
  }
  await openOrderPreviewModal(matchedOrder);
}

function renderCurrentUser() {
  currentUserName.textContent = currentUser?.name || "Chưa đăng nhập";
  currentUserRole.textContent = currentUser ? labelRole(currentUser.role) : "-";
  currentUserAvatar.className = `user-card-avatar ${currentUser ? `role-${String(currentUser.role || "").toLowerCase()}` : ""}`.trim();
  currentUserAvatar.innerHTML = avatarIconForRole(currentUser?.role);
  currentUserMeta.innerHTML = currentUser
    ? `
      <div class="user-fact">
        <span class="user-fact-label">Mã nhân viên</span>
        <strong class="user-fact-value">${escapeHtml(String(currentUser.employee_code || currentUser.username || "-").toUpperCase())}</strong>
      </div>
      <div class="user-fact">
        <span class="user-fact-label">Bộ phận</span>
        <strong class="user-fact-value">${escapeHtml(labelDepartment(currentUser.department))}</strong>
      </div>
      <div class="user-fact user-fact-wide">
        <span class="user-fact-label">Chức vụ</span>
        <strong class="user-fact-value">${escapeHtml(currentUser.title || labelRole(currentUser.role))}</strong>
      </div>
      <div class="user-fact user-fact-wide">
        <span class="user-fact-label">Mức truy cập</span>
        <strong class="user-fact-value">${escapeHtml(labelAccessLevel(currentUser.policy?.max_access_level || "basic"))}</strong>
      </div>
    `
    : `
      <div class="user-fact user-fact-wide">
        <span class="user-fact-label">Trạng thái</span>
        <strong class="user-fact-value">Bạn cần đăng nhập để sử dụng hệ thống.</strong>
      </div>
    `;
  openPermissionsModalButton?.classList.add("hidden");
  openUsersModalButton?.classList.toggle("hidden", String(currentUser?.role || "").trim().toLowerCase() !== "admin");
  openAuditModalButton.classList.toggle("hidden", !currentUser?.policy?.can_manage_permissions);
  openDashboardModalButton.classList.toggle("hidden", !currentUser?.policy?.can_manage_permissions);
  openImportModalButton.classList.toggle("hidden", !currentUser);
  openLoginModalButton?.classList.toggle("hidden", Boolean(currentUser));
  openImportModalButton.textContent = canCreateManagedContent() ? "Thêm nguồn / file" : "Tải file cá nhân";
  openDeliveryModalButton?.classList.toggle("hidden", !canCompleteDelivery());
  openTapeCalculatorModalButton?.classList.toggle("hidden", !currentUser);
  orderCtaMenu?.classList.toggle("hidden", !canCreateOrder());
  openOrderProductsModalButton?.classList.add("hidden");
  openSalesProductsModalButton?.classList.toggle("hidden", !canManageOrders());
  ordersBoardMenu?.classList.toggle("hidden", !canViewOrders());
  logoutButton.classList.toggle("hidden", !currentUser);
  openPasswordModalButton.classList.toggle("hidden", !currentUser);
  syncSheetButton.classList.toggle(
    "hidden",
    !(currentUser && ["director", "admin"].includes(String(currentUser.role || "").toLowerCase())),
  );
  syncImportPolicy();
  updateFolderOptions();
  renderPrivateShareOptions();
  applyAccessFilterPolicy();
  syncFolderFilterOptions();
  renderDocumentLibrary();
}

function renderNotifications() {
  if (!notificationList || !notificationCount) {
    renderNotificationHistory();
    return;
  }

  const visibleNotifications = notifications.filter((item) => {
    const id = String(item?.id || "").trim();
    if (id && hiddenNotificationIds.has(id)) {
      return false;
    }
    if (notificationFilterValue === "unread") {
      return !String(item?.read_at || "").trim();
    }
    if (notificationFilterValue !== "all") {
      return String(item?.type || "").trim() === notificationFilterValue;
    }
    return true;
  });

  notificationCount.textContent = String(visibleNotifications.length);
  notificationList.innerHTML = "";

  if (!visibleNotifications.length) {
    const empty = document.createElement("div");
    empty.className = "history-item muted";
    empty.textContent = notificationFilterValue === "all" ? "Chưa có thông báo nào" : "Không có thông báo phù hợp";
    notificationList.append(empty);
    renderNotificationHistory();
    return;
  }

  for (const item of visibleNotifications.slice(0, 8)) {
    const node = document.createElement("article");
    node.className = `notification-item${item.read_at ? "" : " is-unread"}`;
    node.innerHTML = `
      <div class="notification-item-head">
        <strong>${escapeHtml(item.title || "Thông báo nội bộ")}</strong>
        <div class="notification-item-head-actions">
          ${item.read_at ? "" : '<span class="mini-pill notification-pill">Mới</span>'}
          <button type="button" class="notification-delete-button" aria-label="Xóa thông báo">
            <span aria-hidden="true">&times;</span>
          </button>
        </div>
      </div>
      <p>${escapeHtml(item.message || "")}</p>
      <small>${escapeHtml(formatNotificationTime(item.created_at || ""))}</small>
    `;
    node.tabIndex = 0;
    node.querySelector(".notification-delete-button")?.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      deleteNotification(item);
    });
    node.addEventListener("click", () => openNotification(item));
    node.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        openNotification(item);
      }
    });
    notificationList.append(node);
  }

  renderNotificationHistory();
}

function getNotificationUiStateStorageKey(suffix) {
  const ownerKey = String(currentUser?.id || "").trim();
  return ownerKey ? `notification:${suffix}:${ownerKey}` : "";
}

function clearNotificationAutoHideTimers() {
  notificationAutoHideTimers.forEach((timerId) => window.clearTimeout(timerId));
  notificationAutoHideTimers.clear();
}

function ensureNotificationUiStateLoaded() {
  const ownerKey = String(currentUser?.id || "").trim();
  if (notificationUiStateOwnerKey === ownerKey) {
    return;
  }

  clearNotificationAutoHideTimers();
  hiddenNotificationIds.clear();
  notificationArchive = [];
  notificationArchiveClearedAt = 0;
  notificationUiStateOwnerKey = ownerKey;

  if (!ownerKey) {
    return;
  }

  const hiddenKey = getNotificationUiStateStorageKey("hidden");
  const archiveKey = getNotificationUiStateStorageKey("archive");
  const archiveClearedAtKey = getNotificationUiStateStorageKey("archive-cleared-at");

  try {
    const hiddenItems = JSON.parse(localStorage.getItem(hiddenKey) || "[]");
    if (Array.isArray(hiddenItems)) {
      hiddenItems.forEach((id) => {
        const normalizedId = String(id || "").trim();
        if (normalizedId) {
          hiddenNotificationIds.add(normalizedId);
        }
      });
    }
  } catch {}

  try {
    const archivedItems = JSON.parse(localStorage.getItem(archiveKey) || "[]");
    if (Array.isArray(archivedItems)) {
      notificationArchive = archivedItems
        .map((item) => normalizeNotificationArchiveItem(item))
        .filter(Boolean);
    }
  } catch {}

  const clearedAtValue = Number(localStorage.getItem(archiveClearedAtKey) || 0);
  notificationArchiveClearedAt = Number.isFinite(clearedAtValue) ? Math.max(0, clearedAtValue) : 0;
}

function persistHiddenNotificationIds() {
  const storageKey = getNotificationUiStateStorageKey("hidden");
  if (!storageKey) {
    return;
  }
  localStorage.setItem(storageKey, JSON.stringify(Array.from(hiddenNotificationIds)));
}

function persistNotificationArchive() {
  const storageKey = getNotificationUiStateStorageKey("archive");
  if (!storageKey) {
    return;
  }
  notificationArchive = pruneNotificationArchive(notificationArchive).slice(0, 120);
  localStorage.setItem(storageKey, JSON.stringify(notificationArchive));
}

function persistNotificationArchiveClearedAt() {
  const storageKey = getNotificationUiStateStorageKey("archive-cleared-at");
  if (!storageKey) {
    return;
  }
  localStorage.setItem(storageKey, String(notificationArchiveClearedAt || 0));
}

function normalizeNotificationArchiveItem(item) {
  const id = String(item?.id || "").trim();
  if (!id) {
    return null;
  }
  return {
    id,
    title: String(item?.title || "").trim(),
    message: String(item?.message || "").trim(),
    type: String(item?.type || "").trim(),
    created_at: String(item?.created_at || "").trim(),
    read_at: String(item?.read_at || "").trim(),
    meta: item?.meta && typeof item.meta === "object" ? item.meta : {},
  };
}

function mergeNotificationArchiveItems(items = []) {
  ensureNotificationUiStateLoaded();
  const archiveMap = new Map(notificationArchive.map((item) => [String(item.id || "").trim(), item]));
  items.forEach((item) => {
    const normalized = normalizeNotificationArchiveItem(item);
    if (!normalized) {
      return;
    }
    const createdAt = new Date(normalized.created_at || 0).getTime();
    if (notificationArchiveClearedAt && Number.isFinite(createdAt) && createdAt <= notificationArchiveClearedAt) {
      return;
    }
    archiveMap.set(normalized.id, {
      ...(archiveMap.get(normalized.id) || {}),
      ...normalized,
    });
  });
  notificationArchive = Array.from(archiveMap.values())
    .sort((left, right) => {
      const rightTime = new Date(right.created_at || 0).getTime();
      const leftTime = new Date(left.created_at || 0).getTime();
      return rightTime - leftTime;
    });
  persistNotificationArchive();
}

function pruneNotificationArchive(items = []) {
  const cutoff = Math.max(Date.now() - NOTIFICATION_ARCHIVE_RETENTION_MS, notificationArchiveClearedAt || 0);
  return items.filter((item) => {
    const createdAt = new Date(item?.created_at || 0).getTime();
    return Number.isFinite(createdAt) && createdAt > cutoff;
  });
}

function removeNotificationFromArchive(notificationId) {
  const normalizedId = String(notificationId || "").trim();
  if (!normalizedId) {
    return;
  }
  notificationArchive = notificationArchive.filter((item) => String(item?.id || "").trim() !== normalizedId);
  persistNotificationArchive();
}

function hideNotificationFromActive(notificationId) {
  const normalizedId = String(notificationId || "").trim();
  if (!normalizedId) {
    return;
  }
  hiddenNotificationIds.add(normalizedId);
  persistHiddenNotificationIds();
  const existingTimer = notificationAutoHideTimers.get(normalizedId);
  if (existingTimer) {
    window.clearTimeout(existingTimer);
    notificationAutoHideTimers.delete(normalizedId);
  }
  renderNotifications();
}

function scheduleNotificationAutoHide(item, delayMs = 30000) {
  const notificationId = String(item?.id || "").trim();
  if (!notificationId) {
    return;
  }
  const existingTimer = notificationAutoHideTimers.get(notificationId);
  if (existingTimer) {
    window.clearTimeout(existingTimer);
  }
  const timerId = window.setTimeout(() => {
    notificationAutoHideTimers.delete(notificationId);
    hideNotificationFromActive(notificationId);
  }, delayMs);
  notificationAutoHideTimers.set(notificationId, timerId);
}

function renderNotificationHistory() {
  if (!notificationHistoryList || !notificationHistoryCount) {
    return;
  }

  ensureNotificationUiStateLoaded();
  notificationArchive = pruneNotificationArchive(notificationArchive);
  persistNotificationArchive();
  notificationHistoryCount.textContent = String(notificationArchive.length);
  notificationHistoryList.innerHTML = "";

  if (!notificationArchive.length) {
    const empty = document.createElement("div");
    empty.className = "history-item muted";
    empty.textContent = "Chưa có lịch sử thông báo";
    notificationHistoryList.append(empty);
    return;
  }

  notificationArchive.slice(0, 20).forEach((item) => {
    const node = document.createElement("article");
    node.className = `notification-item notification-history-item${item.read_at ? "" : " is-unread"}`;
    node.innerHTML = `
      <div class="notification-item-head">
        <strong>${escapeHtml(item.title || "Thông báo nội bộ")}</strong>
      </div>
      <p>${escapeHtml(item.message || "")}</p>
      <small>${escapeHtml(formatNotificationTime(item.created_at || ""))}</small>
    `;
    node.tabIndex = 0;
    node.addEventListener("click", () => openNotification(item, { skipAutoHide: true }));
    node.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        openNotification(item, { skipAutoHide: true });
      }
    });
    notificationHistoryList.append(node);
  });
}

function clearNotificationHistory() {
  ensureNotificationUiStateLoaded();
  notificationArchiveClearedAt = Date.now();
  notificationArchive = [];
  persistNotificationArchive();
  persistNotificationArchiveClearedAt();
  renderNotificationHistory();
  showToast("Đã xóa lịch sử thông báo.", "success");
}

async function reloadWorkspace() {
  await loadUsers();
  await Promise.all([loadHealth(), loadSources(), loadDocuments(), loadNotifications()]);
}

function showLoginModal() {
  loginModal.classList.remove("hidden");
  loginBackdrop.classList.remove("hidden");
  loginModal.setAttribute("aria-hidden", "false");
  resetLoginChallenge();
  loginUsername.focus();
}

function hideLoginModal() {
  loginModal.classList.add("hidden");
  loginBackdrop.classList.add("hidden");
  loginModal.setAttribute("aria-hidden", "true");
  resetLoginChallenge();
}

async function handleLoginSubmit(event) {
  event.preventDefault();
  loginSubmitButton.disabled = true;
  try {
    const payload = pendingLoginChallenge
      ? await postJson(loginVerifyUrl, {
          challengeToken: pendingLoginChallenge,
          code: loginOtp.value.trim(),
        }, { skipAuth: true })
      : await postJson(loginUrl, {
          username: loginUsername.value.trim(),
          password: loginPassword.value,
        }, { skipAuth: true });

    if (payload.otp_required) {
      pendingLoginChallenge = payload.challengeToken || "";
      loginOtpField.classList.remove("hidden");
      loginOtpHelp.classList.remove("hidden");
      loginOtpHelp.textContent = payload.delivery || "Ma xac thuc da duoc gui qua email.";
      loginSubmitButton.textContent = "Xác minh mã";
      loginOtp.value = "";
      loginOtp.focus();
      showToast(payload.delivery || "Ma xac thuc da duoc gui qua email.", "success");
      return;
    }

    authToken = payload.token || "";
    if (authToken) {
      localStorage.setItem("authToken", authToken);
    }
    loginPassword.value = "";
    loginOtp.value = "";
    currentUser = payload.currentUser || null;
    hideLoginModal();
    await reloadWorkspace();
    startNotificationStream();
    startNotificationPolling();
    showToast(`Đã đăng nhập với ${currentUser?.name || "tài khoản"}.`, "success");
  } catch (error) {
    showToast(error.message || "Không thể đăng nhập.", "error");
  } finally {
    loginSubmitButton.disabled = false;
  }
}

function resetLoginChallenge() {
  pendingLoginChallenge = "";
  loginOtpField.classList.add("hidden");
  loginOtpHelp.classList.add("hidden");
  loginOtpHelp.textContent = "";
  loginOtp.value = "";
  loginSubmitButton.textContent = "Đăng nhập";
}

async function logoutCurrentUser() {
  try {
    await postJson(logoutUrl, {}, { skipAuth: false });
  } catch {}
  stopNotificationStream();
  stopNotificationPolling();
  clearNotificationAutoHideTimers();
  authToken = "";
  currentUser = null;
  allUsers = [];
  notifications = [];
  notificationArchive = [];
  notificationArchiveClearedAt = 0;
  hiddenNotificationIds.clear();
  notificationUiStateOwnerKey = "";
  hasBootstrappedNotifications = false;
  seenNotificationIds.clear();
  localStorage.removeItem("authToken");
  renderCurrentUser();
  renderNotifications();
  renderPermissionsPanel();
  sourceList.innerHTML = "";
  docList.innerHTML = "";
  systemStatus.innerHTML = "";
  showLoginModal();
}

function openPermissionsModal() {
  if (String(currentUser?.role || "").trim().toLowerCase() !== "admin") {
    return;
  }

  permissionsModal.classList.remove("hidden");
  permissionsBackdrop.classList.remove("hidden");
  permissionsModal.setAttribute("aria-hidden", "false");
}

function closePermissionsModal() {
  permissionsModal.classList.add("hidden");
  permissionsBackdrop.classList.add("hidden");
  permissionsModal.setAttribute("aria-hidden", "true");
}

function openPasswordModal() {
  if (!currentUser) {
    return;
  }

  passwordModal.classList.remove("hidden");
  passwordBackdrop.classList.remove("hidden");
  passwordModal.setAttribute("aria-hidden", "false");
  currentPasswordInput.focus();
}

async function openDeliveryModal() {
  if (!canCompleteDelivery()) {
    showToast("Chỉ bộ phận vận chuyển hoặc quản trị mới được xác nhận giao hàng.", "error");
    return;
  }

  populateDeliverySalesOptions();
  await loadOrders();
  if (!deliverySalesUser?.value) {
    showToast("Chưa có nhân viên kinh doanh nào trong danh sách để nhận thông báo.", "error");
    return;
  }
  deliveryForm?.reset();
  populateDeliveryOrderOptions();
  if (!orders.some(isPendingDeliveryOrder)) {
    showToast("Hiện không có đơn nào chưa giao hoặc chờ giao lại.", "success");
    return;
  }
  if (deliveryResultStatus) {
    deliveryResultStatus.value = "delivered";
  }
  if (deliveryCompletedAt) {
    deliveryCompletedAt.value = toLocalDateTimeInputValue(new Date());
  }
  if (deliveryPaymentStatus) {
    deliveryPaymentStatus.value = "unpaid";
  }
  if (deliveryPaymentMethod) {
    deliveryPaymentMethod.value = "";
  }
  deliveryModal?.classList.remove("hidden");
  deliveryBackdrop?.classList.remove("hidden");
  deliveryModal?.setAttribute("aria-hidden", "false");
  autofillDeliveryOrderDetails();
  deliveryOrderId?.focus();
}

function closeDeliveryModal() {
  deliveryModal?.classList.add("hidden");
  deliveryBackdrop?.classList.add("hidden");
  deliveryModal?.setAttribute("aria-hidden", "true");
  deliveryForm?.reset();
}

function isPendingDeliveryOrder(item) {
  if (isProductionOrder(item)) {
    return false;
  }
  const status = String(item?.status || "").trim().toLowerCase();
  const resultStatus = String(item?.completion_result_status || "").trim().toLowerCase();
  if (status === "assigned") {
    return true;
  }
  return ["partial", "rescheduled", "failed"].includes(resultStatus);
}

function labelPendingDeliveryOrder(item) {
  const resultStatus = String(item?.completion_result_status || "").trim().toLowerCase();
  if (resultStatus === "rescheduled") {
    return "Chờ giao lại";
  }
  if (resultStatus === "failed") {
    return "Không giao được";
  }
  if (resultStatus === "partial") {
    return "Giao một phần";
  }
  return "Chưa giao";
}

function populateDeliveryOrderOptions() {
  if (!deliveryOrderId) {
    return;
  }

  const previous = String(deliveryOrderId.value || "").trim();
  const pendingOrders = orders
    .filter(isPendingDeliveryOrder)
    .sort((left, right) => String(right.updated_at || right.created_at || "").localeCompare(String(left.updated_at || left.created_at || "")));

  const selectedOrderId = previous || String(pendingOrders[0]?.order_id || "").trim();
  const options = pendingOrders.map((item) => {
      const label = `${item.order_id || "-"} • ${item.customer_name || "Khách hàng"} • ${labelPendingDeliveryOrder(item)}`;
      const selected = String(item.order_id || "").trim() === selectedOrderId ? "selected" : "";
      return `<option value="${escapeHtml(item.order_id || "")}" ${selected}>${escapeHtml(label)}</option>`;
    });

  deliveryOrderId.innerHTML = options.length
    ? options.join("")
    : `<option value="">Không có đơn nào đang chờ giao</option>`;
}

function autofillDeliveryOrderDetails() {
  const normalizedOrderId = normalizeOrderIdValue(deliveryOrderId?.value || "");
  if (!normalizedOrderId) {
    if (deliveryCustomerName) {
      deliveryCustomerName.value = "";
    }
    if (deliveryAddress) {
      deliveryAddress.value = "";
    }
    return;
  }

  const matchedOrder = orders.find(
    (item) => String(item?.order_id || "").trim().toUpperCase() === normalizedOrderId,
  );

  if (!matchedOrder) {
    return;
  }

  if (deliveryCustomerName) {
    deliveryCustomerName.value = matchedOrder.customer_name || "";
  }
  if (deliveryAddress) {
    deliveryAddress.value = matchedOrder.delivery_address || "";
  }
  if (deliverySalesUser && matchedOrder.sales_user_id) {
    deliverySalesUser.value = matchedOrder.sales_user_id;
  }
}

async function openOrdersBoardModal(kind = "production") {
  if (!canViewOrders()) {
    showToast("Bạn không có quyền xem danh sách đơn hàng.", "error");
    return;
  }

  activeOrdersBoardKind = kind === "transport" ? "transport" : "production";
  syncOrdersBoardUI();
  await loadOrders();
  ordersBoardModal?.classList.remove("hidden");
  ordersBoardBackdrop?.classList.remove("hidden");
  ordersBoardModal?.setAttribute("aria-hidden", "false");
  renderOrdersBoard();
  ordersSearch?.focus();
}

function closeOrdersBoardModal() {
  ordersBoardModal?.classList.add("hidden");
  ordersBoardBackdrop?.classList.add("hidden");
  ordersBoardModal?.setAttribute("aria-hidden", "true");
}

function closeOrderDetailsModal() {
  orderDetailsModal?.classList.add("hidden");
  orderDetailsBackdrop?.classList.add("hidden");
  orderDetailsModal?.setAttribute("aria-hidden", "true");
  activeOrderDetails = null;
}

function closeOrderProductsModal() {
  orderProductsModal?.classList.add("hidden");
  orderProductsBackdrop?.classList.add("hidden");
  orderProductsModal?.setAttribute("aria-hidden", "true");
  setRepoTabActive(null);
}

function populateTapeCalculatorOptions() {
  const selectedTapeType = tapeType?.value || tapeCalculatorDefaults.tape_type || DEFAULT_TAPE_CALCULATOR_FORM.tape_type;
  const selectedCoreType = tapeCoreType?.value || tapeCalculatorDefaults.core_type || DEFAULT_TAPE_CALCULATOR_FORM.core_type;

  if (tapeType) {
    tapeType.value = tapeCalculatorProducts.some((item) => item.code === selectedTapeType) ? selectedTapeType : tapeCalculatorProducts[0]?.code || "";
  }

  if (tapeCoreType) {
    tapeCoreType.value = tapeCalculatorCores.some((item) => item.code === selectedCoreType) ? selectedCoreType : tapeCalculatorCores[0]?.code || "";
  }
}

function updateTapeCalculatorResults() {
  const orderQuantity = Math.max(0, Number(tapeOrderQuantity?.value || 0));
  const packaging = Math.max(0, Number(tapePackaging?.value || 0));
  const selectedProduct = tapeCalculatorProducts.find((item) => item.code === tapeType?.value) || null;
  const selectedCore = tapeCalculatorCores.find((item) => item.code === tapeCoreType?.value) || null;
  const jumboHeight = Math.max(0, Number(tapeJumboHeight?.value || selectedProduct?.jumbo_height || 0));
  const coreWidth = Math.max(0, Number(tapeCoreWidth?.value || selectedCore?.width_mm || 0));

  if (selectedProduct && tapeJumboHeight && Number(tapeJumboHeight.value || 0) !== selectedProduct.jumbo_height) {
    tapeJumboHeight.value = String(selectedProduct.jumbo_height);
  }
  if (selectedCore && tapeCoreWidth && Number(tapeCoreWidth.value || 0) !== selectedCore.width_mm) {
    tapeCoreWidth.value = String(selectedCore.width_mm);
  }

  const finishedQuantity = orderQuantity * packaging;
  const rollsPerHand = coreWidth > 0 ? Math.floor(jumboHeight / coreWidth) : 0;
  const remainingMm = rollsPerHand > 0 ? jumboHeight - rollsPerHand * coreWidth : 0;
  const handsNeeded = rollsPerHand > 0 ? Math.ceil(finishedQuantity / rollsPerHand) : 0;
  const totalProduced = rollsPerHand * handsNeeded;
  const extraProduced = Math.max(0, totalProduced - finishedQuantity);

  if (tapeFinishedQuantity) tapeFinishedQuantity.value = finishedQuantity > 0 ? String(finishedQuantity) : "";
  if (tapeRollsPerHand) tapeRollsPerHand.value = rollsPerHand > 0 ? String(rollsPerHand) : "";
  if (tapeRemainingMm) tapeRemainingMm.value = String(remainingMm);
  if (tapeHandsNeeded) tapeHandsNeeded.value = handsNeeded > 0 ? String(handsNeeded) : "";
  if (tapeTotalProduced) tapeTotalProduced.value = totalProduced > 0 ? String(totalProduced) : "";
  if (tapeExtraProduced) tapeExtraProduced.value = String(extraProduced);
}

function applyTapeCalculatorConfig(payload = {}) {
  const productMap = new Map();
  for (const item of payload.products || []) {
    const code = String(item?.code || "").trim().toUpperCase();
    const jumboHeight = Number(item?.jumbo_height || 0);
    if (!code) {
      continue;
    }

    productMap.set(code, {
      code,
      jumbo_height: Number.isFinite(jumboHeight) && jumboHeight > 0 ? jumboHeight : 1260,
    });
  }

  const coreMap = new Map();
  for (const item of payload.cores || []) {
    const code = String(item?.code || "").trim().toUpperCase();
    const widthMm = Number(item?.width_mm || 0);
    if (!code) {
      continue;
    }

    coreMap.set(code, {
      code,
      width_mm: Number.isFinite(widthMm) && widthMm > 0 ? widthMm : Number(code.replace(/^[A-Z]/, "")) || 0,
    });
  }

  tapeCalculatorProducts = [...(productMap.size ? productMap.values() : DEFAULT_TAPE_CALCULATOR_PRODUCTS)];
  tapeCalculatorCores = [...(coreMap.size ? coreMap.values() : DEFAULT_TAPE_CALCULATOR_CORES)];

  const selectedDefaultProduct =
    tapeCalculatorProducts.find((item) => item.code === String(payload.defaults?.tape_type || "").trim().toUpperCase()) ||
    tapeCalculatorProducts.find((item) => item.code === DEFAULT_TAPE_CALCULATOR_FORM.tape_type) ||
    tapeCalculatorProducts[0] ||
    DEFAULT_TAPE_CALCULATOR_PRODUCTS[0];
  const selectedDefaultCore =
    tapeCalculatorCores.find((item) => item.code === String(payload.defaults?.core_type || "").trim().toUpperCase()) ||
    tapeCalculatorCores.find((item) => item.code === DEFAULT_TAPE_CALCULATOR_FORM.core_type) ||
    tapeCalculatorCores[0] ||
    DEFAULT_TAPE_CALCULATOR_CORES[0];

  tapeCalculatorDefaults = {
    tape_type: selectedDefaultProduct?.code || DEFAULT_TAPE_CALCULATOR_FORM.tape_type,
    order_quantity: Math.max(0, Number(payload.defaults?.order_quantity || DEFAULT_TAPE_CALCULATOR_FORM.order_quantity)),
    core_type: selectedDefaultCore?.code || DEFAULT_TAPE_CALCULATOR_FORM.core_type,
    packaging: Math.max(0, Number(payload.defaults?.packaging || DEFAULT_TAPE_CALCULATOR_FORM.packaging)),
    finished_quantity: Math.max(0, Number(payload.defaults?.finished_quantity || DEFAULT_TAPE_CALCULATOR_FORM.finished_quantity)),
    jumbo_height: Math.max(
      0,
      Number(payload.defaults?.jumbo_height || selectedDefaultProduct?.jumbo_height || DEFAULT_TAPE_CALCULATOR_FORM.jumbo_height),
    ),
    core_width: Math.max(
      0,
      Number(payload.defaults?.core_width || selectedDefaultCore?.width_mm || DEFAULT_TAPE_CALCULATOR_FORM.core_width),
    ),
  };

  populateTapeCalculatorOptions();
}

function getTapeCalculatorSource(kind = "product") {
  return kind === "core" ? tapeCalculatorCores : tapeCalculatorProducts;
}

function getTapeCalculatorSuggestions(kind = "product", query = "") {
  const normalizedQuery = String(query || "").trim().toLowerCase();
  const source = getTapeCalculatorSource(kind);
  if (!normalizedQuery) {
    return source.map((item) => item.code);
  }

  return source
    .filter((item) => item.code.toLowerCase().includes(normalizedQuery))
    .map((item) => item.code);
}

function closeTapeCalculatorSuggestions(exceptInput = null) {
  document.querySelectorAll(".tape-calculator-picker").forEach((wrap) => {
    const input = wrap.querySelector("input");
    if (exceptInput && input === exceptInput) {
      return;
    }

    wrap.dataset.expanded = "false";
    wrap.dataset.activeIndex = "-1";
    const panel = wrap.querySelector(".tape-calculator-suggestions");
    if (panel) {
      panel.innerHTML = "";
      panel.classList.add("hidden");
    }
  });
}

function setActiveTapeCalculatorSuggestion(wrap, nextIndex) {
  if (!wrap) {
    return;
  }

  const buttons = Array.from(wrap.querySelectorAll("[data-tape-calculator-option]"));
  if (!buttons.length) {
    wrap.dataset.activeIndex = "-1";
    return;
  }

  const boundedIndex = Math.max(0, Math.min(nextIndex, buttons.length - 1));
  wrap.dataset.activeIndex = String(boundedIndex);
  buttons.forEach((button, index) => {
    button.classList.toggle("active", index === boundedIndex);
  });
  buttons[boundedIndex]?.scrollIntoView({ block: "nearest" });
}

function renderTapeCalculatorSuggestions(input, queryOverride = null) {
  const wrap = input?.closest(".tape-calculator-picker");
  const panel = wrap?.querySelector(".tape-calculator-suggestions");
  const kind = wrap?.dataset.tapeCalculatorKind || "product";
  if (!wrap || !panel || !input) {
    return;
  }

  const query = queryOverride === null ? input.value : queryOverride;
  const suggestions = getTapeCalculatorSuggestions(kind, query);
  if (!suggestions.length) {
    wrap.dataset.expanded = "true";
    wrap.dataset.activeIndex = "-1";
    panel.innerHTML = `<div class="tape-calculator-suggestion-empty">Không tìm thấy dữ liệu phù hợp</div>`;
    panel.classList.remove("hidden");
    return;
  }

  panel.innerHTML = suggestions
    .map(
      (name) => `
        <button type="button" class="tape-calculator-suggestion" data-tape-calculator-option="${escapeHtml(name)}">
          ${buildHighlightedOrderItemLabel(name, query)}
        </button>
      `,
    )
    .join("");
  panel.scrollTop = 0;
  wrap.dataset.expanded = "true";
  panel.classList.remove("hidden");
  setActiveTapeCalculatorSuggestion(wrap, 0);
}

function selectTapeCalculatorSuggestion(input, optionValue) {
  if (!input) {
    return;
  }

  input.value = optionValue;
  closeTapeCalculatorSuggestions();
  updateTapeCalculatorResults();
  input.focus();
  input.setSelectionRange(input.value.length, input.value.length);
}

function handleTapeCalculatorInput(event) {
  const input = event.target?.closest?.("#tape-type, #tape-core-type");
  if (!input) {
    return;
  }

  renderTapeCalculatorSuggestions(input);
  updateTapeCalculatorResults();
}

function handleTapeCalculatorFocus(event) {
  const input = event.target?.closest?.("#tape-type, #tape-core-type");
  if (!input) {
    return;
  }

  closeTapeCalculatorSuggestions(input);
}

function handleTapeCalculatorKeydown(event) {
  const input = event.target?.closest?.("#tape-type, #tape-core-type");
  if (!input) {
    return;
  }

  const wrap = input.closest(".tape-calculator-picker");
  const buttons = Array.from(wrap?.querySelectorAll("[data-tape-calculator-option]") || []);
  const activeIndex = Number(wrap?.dataset.activeIndex || -1);

  if (event.key === "ArrowDown") {
    event.preventDefault();
    if (!buttons.length) {
      renderTapeCalculatorSuggestions(input, "");
      return;
    }
    setActiveTapeCalculatorSuggestion(wrap, activeIndex + 1);
    return;
  }

  if (event.key === "ArrowUp") {
    event.preventDefault();
    if (!buttons.length) {
      renderTapeCalculatorSuggestions(input, "");
      return;
    }
    setActiveTapeCalculatorSuggestion(wrap, activeIndex <= 0 ? 0 : activeIndex - 1);
    return;
  }

  if (event.key === "Enter" && buttons.length && activeIndex >= 0) {
    event.preventDefault();
    selectTapeCalculatorSuggestion(input, buttons[activeIndex].dataset.tapeCalculatorOption || "");
    return;
  }

  if (event.key === "Escape") {
    closeTapeCalculatorSuggestions();
  }
}

function handleTapeCalculatorSuggestionClick(event) {
  const button = event.target?.closest?.("[data-tape-calculator-option]");
  if (!button) {
    return;
  }

  const input = button.closest(".tape-calculator-picker")?.querySelector("input");
  selectTapeCalculatorSuggestion(input, button.dataset.tapeCalculatorOption || "");
}

function handleTapeCalculatorToggleClick(event) {
  const button = event.target?.closest?.("[data-tape-calculator-toggle]");
  if (!button) {
    return;
  }

  const wrap = button.closest(".tape-calculator-picker");
  const input = wrap?.querySelector("input");
  const isExpanded = wrap?.dataset.expanded === "true";
  if (!wrap || !input) {
    return;
  }

  if (isExpanded) {
    closeTapeCalculatorSuggestions();
    return;
  }

  closeTapeCalculatorSuggestions(input);
  renderTapeCalculatorSuggestions(input, "");
  input.focus();
}

function handleTapeCalculatorOutsideClick(event) {
  if (event.target?.closest?.(".tape-calculator-picker")) {
    return;
  }

  closeTapeCalculatorSuggestions();
}

async function loadTapeCalculatorConfig(force = false) {
  if (!force && tapeCalculatorConfigLoaded) {
    return;
  }

  if (!force && tapeCalculatorConfigLoadingPromise) {
    await tapeCalculatorConfigLoadingPromise;
    return;
  }

  tapeCalculatorConfigLoadingPromise = (async () => {
    try {
      const payload = await fetchJson(tapeCalculatorConfigUrl);
      applyTapeCalculatorConfig(payload);
      tapeCalculatorConfigLoaded = true;
    } catch (error) {
      applyTapeCalculatorConfig({});
      tapeCalculatorConfigLoaded = true;
      console.error("Khong the tai config bang tinh bang dinh:", error);
    } finally {
      tapeCalculatorConfigLoadingPromise = null;
    }
  })();

  await tapeCalculatorConfigLoadingPromise;
}

async function openTapeCalculatorModal() {
  if (!currentUser) {
    showToast("Cần đăng nhập để dùng bảng tính.", "error");
    return;
  }

  await loadTapeCalculatorConfig();
  if (tapeOrderQuantity && !Number(tapeOrderQuantity.value || 0)) {
    tapeOrderQuantity.value = String(tapeCalculatorDefaults.order_quantity || DEFAULT_TAPE_CALCULATOR_FORM.order_quantity);
  }
  if (tapePackaging && !Number(tapePackaging.value || 0)) {
    tapePackaging.value = String(tapeCalculatorDefaults.packaging || DEFAULT_TAPE_CALCULATOR_FORM.packaging);
  }
  if (tapeJumboHeight && !Number(tapeJumboHeight.value || 0)) {
    tapeJumboHeight.value = String(tapeCalculatorDefaults.jumbo_height || DEFAULT_TAPE_CALCULATOR_FORM.jumbo_height);
  }
  if (tapeCoreWidth && !Number(tapeCoreWidth.value || 0)) {
    tapeCoreWidth.value = String(tapeCalculatorDefaults.core_width || DEFAULT_TAPE_CALCULATOR_FORM.core_width);
  }
  updateTapeCalculatorResults();
  tapeCalculatorModal?.classList.remove("hidden");
  tapeCalculatorBackdrop?.classList.remove("hidden");
  tapeCalculatorModal?.setAttribute("aria-hidden", "false");
  tapeOrderQuantity?.focus();
}

function closeTapeCalculatorModal() {
  tapeCalculatorModal?.classList.add("hidden");
  tapeCalculatorBackdrop?.classList.add("hidden");
  tapeCalculatorModal?.setAttribute("aria-hidden", "true");
}

function appendOrderProductEditorRow(product = null) {
  if (!orderProductsList) {
    return;
  }

  const nextIndex = orderProductsList.querySelectorAll(".order-product-editor-row").length;
  const row = document.createElement("div");
  row.className = "order-product-editor-row";
  row.innerHTML = `
    <div class="order-product-index">${nextIndex}</div>
    <input class="order-product-id-input" type="hidden" value="${escapeHtml(product?.id || "")}" />
    <input class="order-product-code-input" type="text" placeholder="Mã hàng" value="${escapeHtml(product?.code || "")}" />
    <input class="order-product-name-input" type="text" placeholder="Tên hàng hóa" value="${escapeHtml(product?.name || "")}" />
    <input class="order-product-units-input" type="text" placeholder="Đơn vị, cách nhau bởi dấu phẩy" value="${escapeHtml(Array.isArray(product?.units) ? product.units.join(", ") : "")}" />
    <input class="order-product-default-unit-input" type="text" placeholder="Đơn vị mặc định" value="${escapeHtml(product?.default_unit || "")}" />
    <button type="button" class="ghost-button order-product-delete-button">Xóa</button>
  `;
  row.querySelector(".order-product-delete-button")?.addEventListener("click", () => {
    row.remove();
    syncOrderProductEditorIndexes();
  });
  orderProductsList.appendChild(row);
  syncOrderProductEditorIndexes();
}

function syncOrderProductEditorIndexes() {
  if (!orderProductsList) {
    return;
  }

  Array.from(orderProductsList.querySelectorAll(".order-product-editor-row")).forEach((row, index) => {
    const indexNode = row.querySelector(".order-product-index");
    if (indexNode) {
      indexNode.textContent = String(index + 1);
    }
  });
}

async function openOrderProductsModal() {
  if (!canManageOrders()) {
    return;
  }

  await loadOrderProducts();
  orderProductsList.innerHTML = "";
  const products = getSortedOrderProducts();
  products.forEach((product) => appendOrderProductEditorRow(product));
  orderProductsModal?.classList.remove("hidden");
  orderProductsBackdrop?.classList.remove("hidden");
  orderProductsModal?.setAttribute("aria-hidden", "false");
  setRepoTabActive("sales-products");
}

function collectOrderProductsFromEditor() {
  if (!orderProductsList) {
    return [];
  }

  return getSortedOrderProducts(
    Array.from(orderProductsList.querySelectorAll(".order-product-editor-row"))
    .map((row, index) => {
      const storedId = row.querySelector(".order-product-id-input")?.value?.trim() || "";
      const name = row.querySelector(".order-product-name-input")?.value?.trim() || "";
      const code = row.querySelector(".order-product-code-input")?.value?.trim() || "";
      const unitsRaw = row.querySelector(".order-product-units-input")?.value || "";
      const units = unitsRaw
        .split(",")
        .map((value) => value.trim())
        .filter(Boolean);
      const defaultUnit = row.querySelector(".order-product-default-unit-input")?.value?.trim() || units[0] || "";
      if (!name) {
        return null;
      }
      return {
        id: storedId || `product-${index + 1}-${Date.now()}-${Math.random().toString(16).slice(2, 6)}`,
        code,
        name,
        units: units.length ? units : [defaultUnit || "cây"],
        default_unit: defaultUnit || units[0] || "cây",
      };
    })
    .filter(Boolean),
  );
}

async function saveOrderProductsCatalog() {
  if (!canManageOrders()) {
    showToast("Bạn không có quyền quản lý hàng hóa.", "error");
    return;
  }

  const products = collectOrderProductsFromEditor();
  if (!products.length) {
    showToast("Hãy nhập ít nhất một hàng hóa.", "error");
    return;
  }

  const seenNames = new Set();
  for (const product of products) {
    const normalizedName = String(product?.name || "").trim().toLowerCase();
    if (!normalizedName) {
      continue;
    }
    if (seenNames.has(normalizedName)) {
      showToast("Tên hàng hóa không được trùng nhau. Mã hàng có thể trùng nếu tên khác.", "error");
      return;
    }
    seenNames.add(normalizedName);
  }

  saveOrderProductsButton.disabled = true;
  try {
    const payload = await postJson(orderProductsSaveUrl, { products });
    orderProducts = getSortedOrderProducts(payload.products || products);
    await loadOrderProducts();
    closeOrderProductsModal();
    showToast("Đã cập nhật danh mục hàng hóa.", "success");
  } catch (error) {
    showToast(error.message || "Không thể lưu danh mục hàng hóa.", "error");
  } finally {
    saveOrderProductsButton.disabled = false;
  }
}

function extractOrderCreatedParts(order) {
  const createdAt = String(order?.created_at || "").trim();
  const date = new Date(createdAt);
  if (!Number.isFinite(date.getTime())) {
    return { date: "", day: "", month: "", year: "" };
  }
  return {
    date: date.toISOString().slice(0, 10),
    day: String(date.getDate()).padStart(2, "0"),
    month: String(date.getMonth() + 1).padStart(2, "0"),
    year: String(date.getFullYear()),
  };
}

function populateOrdersDateFilters() {
  return;
}

function normalizeComparisonText(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase();
}

function parseDocumentFallbackDetails(content) {
  const result = {
    order_items: "",
    order_value: "",
    note: "",
    delivery_address: "",
  };

  String(content || "")
    .split(/\r?\n/)
    .forEach((line) => {
      const trimmed = String(line || "").trim();
      if (!trimmed.startsWith("-")) {
        return;
      }

      const bullet = trimmed.replace(/^-+\s*/, "");
      const separatorIndex = bullet.indexOf(":");
      if (separatorIndex < 0) {
        return;
      }

      const label = normalizeComparisonText(bullet.slice(0, separatorIndex));
      const value = bullet.slice(separatorIndex + 1).trim();
      if (!value) {
        return;
      }

      if (!result.order_items && label.includes("hang giao")) {
        result.order_items = value;
      }
      if (!result.order_value && label.includes("gia tri don hang")) {
        result.order_value = value;
      }
      if (!result.delivery_address && label.includes("dia chi giao hang")) {
        result.delivery_address = value;
      }
      if (!result.note && label === "ghi chu") {
        result.note = value;
      }
    });

  return result;
}

function parseOrderItemsSummary(value) {
  return String(value || "")
    .split(/\r?\n/)
    .map((line) => String(line || "").trim())
    .filter(Boolean)
    .map((line, index) => {
      const normalized = line.replace(/^\d+\.\s*/, "").trim();
      const match = normalized.match(/^(.*?)\s+x\s+(\d+)\s+(\S+)\s+•\s+([\d.,\s]+)\s*[đdĐ]?\s*=\s*([\d.,\s]+)\s*[đdĐ]?$/i);
      if (!match) {
        return {
          index: index + 1,
          name: normalized,
          quantity: "",
          unit: "",
          price: "",
          total: "",
        };
      }

      return {
        index: index + 1,
        name: match[1].trim(),
        quantity: match[2].trim(),
        unit: match[3].trim(),
        price: formatOrderDraftMoneyInput(match[4]),
        total: formatOrderDraftMoney(parseOrderDraftMoneyInput(match[5])),
      };
    });
}

function parseProductionOrderItemsSummary(value) {
  return String(value || "")
    .split(/\r?\n/)
    .map((line) => String(line || "").trim())
    .filter(Boolean)
    .map((line, index) => {
      const normalized = line.replace(/^\d+\.\s*/, "").trim();
      const parts = normalized.split("|").map((part) => String(part || "").trim());
      const readValue = (part, patterns = []) => {
        let nextValue = String(part || "").trim();
        patterns.forEach((pattern) => {
          nextValue = nextValue.replace(pattern, "").trim();
        });
        return nextValue;
      };

      if (parts.length >= 17 && parts[0] === "-") {
        return {
          index: index + 1,
          code: parts[1] || "",
          name: parts[2] || "",
          norm: readValue(parts[3], [/^(?:ĐM|DM|\?M)\s*/i]),
          unit: parts[4] || "",
          quantity: readValue(parts[5], [/^SL\s*/i]),
          done: readValue(parts[13], [/^(?:Đã SX|Da SX|\?+ SX(?: SX)?)\s*/i]),
          missing: readValue(parts[14], [/^Thiếu\s*/i, /^Thieu\s*/i, /^Thi\?u\s*/i]),
          extra: readValue(parts[15], [/^Dư\s*/i, /^Du\s*/i, /^D\?\s*/i]),
          team: readValue(parts[16], [/^Tổ\s*/i, /^To\s*/i, /^T\?\s*/i]),
        };
      }

      if (parts.length < 9) {
        return {
          index: index + 1,
          code: "",
          name: normalized,
          norm: "",
          unit: "",
          quantity: "",
          done: "",
          missing: "",
          extra: "",
          team: "",
        };
      }

      return {
        index: index + 1,
        code: parts[0] || "",
        name: parts[1] || "",
        norm: readValue(parts[2], [/^(?:ĐM|DM|\?M)\s*/i]),
        unit: parts[3] || "",
        quantity: readValue(parts[4], [/^SL\s*/i]),
        done: readValue(parts[5], [/^(?:Đã SX|Da SX|\?+ SX(?: SX)?)\s*/i]),
        missing: readValue(parts[6], [/^Thiếu\s*/i, /^Thieu\s*/i, /^Thi\?u\s*/i]),
        extra: readValue(parts[7], [/^Dư\s*/i, /^Du\s*/i, /^D\?\s*/i]),
        team: readValue(parts[8], [/^Tổ\s*/i, /^To\s*/i, /^T\?\s*/i]),
      };
    });
}

function populateOrderDraftItems(orderItemsValue = "") {
  if (!orderItemsList) {
    return;
  }

  orderItemsList.innerHTML = "";
  const parsedItems = parseOrderItemsSummary(orderItemsValue);

  if (!parsedItems.length) {
    syncOrderDraftTotals();
    return;
  }

  parsedItems.forEach((item) => {
    addOrderItemRow(item.name || "", item.unit || "");
    const row = orderItemsList.lastElementChild;
    if (!row) {
      return;
    }

    const quantityInput = row.querySelector('[data-field="quantity"]');
    const priceInput = row.querySelector('[data-field="price"]');
    if (quantityInput) {
      quantityInput.value = String(item.quantity || "1");
    }
    if (priceInput) {
      priceInput.value = String(item.price || "");
    }
  });

  syncOrderDraftTotals();
}

function populateProductionOrderDraftItems(orderItemsValue = "") {
  if (!productionOrderItemsList) {
    return;
  }

  productionOrderItemsList.innerHTML = "";
  const parsedItems = parseProductionOrderItemsSummary(orderItemsValue);

  if (!parsedItems.length) {
    syncProductionOrderDraft();
    return;
  }

  parsedItems.forEach((item) => addProductionOrderItemRow(item));
  syncProductionOrderDraft();
}

function applyProductionTurnaroundFromNote(noteValue = "") {
  const noteLines = String(noteValue || "")
    .split(/\r?\n/)
    .map((line) => String(line || "").trim());
  const turnaroundPrefix = "đề nghị trả hàng trong:";
  let turnaroundValue = "";
  const contentLines = [];

  noteLines.forEach((line) => {
    if (!line) {
      if (contentLines.length && contentLines[contentLines.length - 1]) {
        contentLines.push("");
      }
      return;
    }

    if (!turnaroundValue && line.toLowerCase().startsWith(turnaroundPrefix)) {
      turnaroundValue = line.slice(turnaroundPrefix.length).trim();
      return;
    }

    contentLines.push(line);
  });

  if (productionNoteInput) {
    productionNoteInput.value = contentLines.join("\n").trim();
  }
  if (productionTurnaroundOtherInput) {
    productionTurnaroundOtherInput.value = "";
  }

  const turnaroundOptions = Array.from(orderForm?.querySelectorAll('input[name="production-turnaround"]') || []);
  turnaroundOptions.forEach((input) => {
    input.checked = false;
  });

  const normalizedTurnaround = String(turnaroundValue || "").trim().toUpperCase();
  const matchedDefault =
    normalizedTurnaround && turnaroundOptions.find((input) => String(input.value || "").trim().toUpperCase() === normalizedTurnaround);
  if (matchedDefault) {
    matchedDefault.checked = true;
    return;
  }

  const otherOption = turnaroundOptions.find((input) => String(input.value || "").trim().toLowerCase() === "khác");
  if (turnaroundValue && otherOption) {
    otherOption.checked = true;
    if (productionTurnaroundOtherInput) {
      productionTurnaroundOtherInput.value = turnaroundValue.replace(/^khác\s*:\s*/i, "").trim();
    }
    return;
  }

  const defaultTurnaroundRadio = orderForm?.querySelector('input[name="production-turnaround"][value="72H"]');
  if (defaultTurnaroundRadio) {
    defaultTurnaroundRadio.checked = true;
  }
}

function renderOrderDetailsItems(orderItemsValue, orderValueRaw) {
  if (!orderDetailsItemsList || !orderDetailsItemsEmpty || !orderDetailsValue || !orderDetailsTotalDisplay) {
    return;
  }

  const parsedItems = parseOrderItemsSummary(orderItemsValue);
  orderDetailsItemsList.innerHTML = "";

  if (!parsedItems.length) {
    orderDetailsItemsEmpty.classList.remove("hidden");
    orderDetailsValue.value = "Chưa cập nhật";
    orderDetailsTotalDisplay.textContent = "0 ₫";
    return;
  }

  orderDetailsItemsEmpty.classList.add("hidden");
  orderDetailsItemsList.innerHTML = parsedItems
    .map((item) => {
      return `
        <div class="order-item-row order-details-item-row">
          <div class="order-item-index">${escapeHtml(String(item.index || ""))}</div>
          <div class="order-item-static">${escapeHtml(item.name || "-")}</div>
          <div class="order-item-static">${escapeHtml(String(item.quantity || "-"))}</div>
          <div class="order-item-static">${escapeHtml(item.unit || "-")}</div>
          <div class="order-item-static">${escapeHtml(item.price || "-")}</div>
          <div class="order-item-line-summary">
            <span class="order-item-line-label">Thành tiền</span>
            <div class="order-item-total">${escapeHtml(item.total || "-")}</div>
          </div>
        </div>
      `;
    })
    .join("");

  orderDetailsValue.value = orderValueRaw ? formatOrderValue(orderValueRaw) : "Chưa cập nhật";
  orderDetailsTotalDisplay.textContent = orderValueRaw ? formatOrderValue(orderValueRaw) : "0 ₫";
}

async function hydrateOrderDetailsFallback(order) {
  const documentPaths = [order?.document_path, order?.completion_document_path]
    .map((value) => String(value || "").trim())
    .filter(Boolean);

  const merged = {
    order_items: "",
    order_value: "",
    note: "",
    delivery_address: "",
  };

  for (const documentPath of documentPaths) {
    try {
      const payload = await fetchJson(`${docContentUrl}?path=${encodeURIComponent(documentPath)}`);
      const fallback = parseDocumentFallbackDetails(payload?.document?.content || "");
      if (!merged.order_items && fallback.order_items) merged.order_items = fallback.order_items;
      if (!merged.order_value && fallback.order_value) merged.order_value = fallback.order_value;
      if (!merged.note && fallback.note) merged.note = fallback.note;
      if (!merged.delivery_address && fallback.delivery_address) merged.delivery_address = fallback.delivery_address;
    } catch {}
  }

  return merged;
}

function applyOrderDetailsToModal(order, fallback = {}) {
  if (!order) {
    return;
  }

  const orderItemsValue = String(order.order_items || fallback.order_items || "").trim();
  const orderValueRaw = String(order.order_value || fallback.order_value || "").trim();
  const noteValue = String(order.completion_note || order.note || fallback.note || "").trim();
  const addressValue = String(order.delivery_address || fallback.delivery_address || "").trim();

  if (orderDetailsId) orderDetailsId.value = order.order_id || "";
  if (orderDetailsCustomerName) orderDetailsCustomerName.value = order.customer_name || "";
  if (orderDetailsSalesUser) orderDetailsSalesUser.value = order.sales_user_name || "";
  if (orderDetailsDeliveryUser) orderDetailsDeliveryUser.value = order.delivery_user_name || "";
  if (orderDetailsCreatedBy) orderDetailsCreatedBy.value = order.created_by_user_name || "-";
  if (orderDetailsCreatedAt) orderDetailsCreatedAt.value = order.created_at ? formatNotificationTime(order.created_at) : "-";
  if (orderDetailsPlannedAt) orderDetailsPlannedAt.value = order.planned_delivery_at ? formatNotificationTime(order.planned_delivery_at) : "-";
  if (orderDetailsStatus) orderDetailsStatus.value = labelOrderStatus(order.status || "");
  if (orderDetailsCompletedAt) orderDetailsCompletedAt.value = order.completed_at ? formatNotificationTime(order.completed_at) : "Chưa giao";
  if (orderDetailsPaymentStatus) orderDetailsPaymentStatus.value = labelPaymentStatus(order.payment_status || "unpaid");
  if (orderDetailsPaymentMethod) orderDetailsPaymentMethod.value = labelPaymentMethod(order.payment_method || "");
  if (orderDetailsAddress) orderDetailsAddress.value = addressValue || "Chưa cập nhật địa chỉ giao hàng.";
  renderOrderDetailsItems(orderItemsValue, orderValueRaw);
  if (orderDetailsNote) orderDetailsNote.value = noteValue || "Không có ghi chú.";
}

async function openOrderDetailsModal(order) {
  if (!order) {
    return;
  }

  activeOrderDetails = order;
  applyOrderDetailsToModal(order);

  orderDetailsModal?.classList.remove("hidden");
  orderDetailsBackdrop?.classList.remove("hidden");
  orderDetailsModal?.setAttribute("aria-hidden", "false");

  const needsFallback =
    !String(order.order_items || "").trim() ||
    !String(order.order_value || "").trim() ||
    !String(order.completion_note || order.note || "").trim();

  if (!needsFallback) {
    return;
  }

  const fallback = await hydrateOrderDetailsFallback(order);
  if (activeOrderDetails?.order_id !== order.order_id) {
    return;
  }
  if (fallback.order_items || fallback.order_value || fallback.note || fallback.delivery_address) {
    try {
      const payload = await postJson(orderBackfillUrl, {
        order_id: order.order_id,
        delivery_address: fallback.delivery_address,
        order_items: fallback.order_items,
        order_value: fallback.order_value,
        note: fallback.note,
      }, { showErrorToast: false });
      orders = payload.orders || orders;
      renderOrdersBoard();
      activeOrderDetails =
        orders.find((item) => String(item?.order_id || "").trim() === String(order.order_id || "").trim()) ||
        activeOrderDetails;
    } catch {}
  }
  applyOrderDetailsToModal(activeOrderDetails || order, fallback);
}

function buildOrderDetailsLines(order) {
  if (!order) {
    return [];
  }

  const orderItemsValue = String(order.order_items || "").trim() || "Chưa cập nhật thông tin hàng hóa.";
  const orderValueRaw = String(order.order_value || "").trim();
  const noteValue = String(order.completion_note || order.note || "").trim() || "Không có ghi chú.";

  return [
    `Mã đơn hàng: ${order.order_id || "-"}`,
    `Khách hàng: ${order.customer_name || "-"}`,
    `NVKD phụ trách: ${order.sales_user_name || "-"}`,
    `NV giao hàng: ${order.delivery_user_name || "-"}`,
    `Người tạo đơn: ${order.created_by_user_name || "-"}`,
    `Ngày tạo đơn: ${order.created_at ? formatNotificationTime(order.created_at) : "-"}`,
    `Ngày cần giao: ${order.planned_delivery_at ? formatNotificationTime(order.planned_delivery_at) : "-"}`,
    `Trạng thái giao: ${labelOrderStatus(order.status || "")}`,
    `Hoàn thành giao: ${order.completed_at ? formatNotificationTime(order.completed_at) : "Chưa giao"}`,
    `Thanh toán: ${labelPaymentStatus(order.payment_status || "unpaid")}`,
    `Hình thức thanh toán: ${labelPaymentMethod(order.payment_method || "")}`,
    `Giá trị đơn hàng: ${orderValueRaw ? formatOrderValue(orderValueRaw) : "Chưa cập nhật"}`,
    `Địa chỉ giao hàng: ${order.delivery_address || "Chưa cập nhật địa chỉ giao hàng."}`,
    `Thông tin hàng hóa:\n${orderItemsValue}`,
    `Ghi chú: ${noteValue}`,
  ];
}

function printActiveOrderDetails() {
  if (!activeOrderDetails) {
    showToast("Chưa có đơn hàng nào đang mở.", "error");
    return;
  }

  const printWindow = window.open("", "_blank", "noopener,noreferrer,width=900,height=700");
  if (!printWindow) {
    showToast("Không thể mở cửa sổ in.", "error");
    return;
  }

  const lines = buildOrderDetailsLines(activeOrderDetails)
    .map((line) => `<p>${escapeHtml(line).replace(/\n/g, "<br />")}</p>`)
    .join("");

  printWindow.document.write(`
    <!doctype html>
    <html lang="vi">
      <head>
        <meta charset="utf-8" />
        <title>${escapeHtml(activeOrderDetails.order_id || "Đơn hàng")}</title>
        <style>
          body { font-family: Arial, sans-serif; padding: 24px; color: #111; }
          h1 { margin: 0 0 16px; font-size: 24px; }
          p { margin: 0 0 10px; line-height: 1.55; white-space: pre-wrap; }
        </style>
      </head>
      <body>
        <h1>${escapeHtml(activeOrderDetails.order_id || "Chi tiết đơn hàng")}</h1>
        ${lines}
      </body>
    </html>
  `);
  printWindow.document.close();
  printWindow.focus();
  printWindow.print();
}

async function exportActiveOrderDetails() {
  if (!activeOrderDetails) {
    showToast("Chưa có đơn hàng nào đang mở.", "error");
    return;
  }

  const documentPath = String(activeOrderDetails.document_path || "").trim();
  if (documentPath) {
    try {
      await downloadDocument(documentPath);
      return;
    } catch (error) {
      showToast(error.message || "Không thể tải file đơn hàng.", "error");
      return;
    }
  }

  const content = buildOrderDetailsLines(activeOrderDetails).join("\n\n");
  const blob = new Blob([content], { type: "text/plain;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `${String(activeOrderDetails.order_id || "don-hang").toLowerCase()}.txt`;
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

function formatOrderDraftMoney(value) {
  const amount = Number(value);
  if (!Number.isFinite(amount) || amount <= 0) {
    return "0 ₫";
  }
  return `${Math.round(amount).toLocaleString("vi-VN")} ₫`;
}

function parseOrderDraftMoneyInput(value) {
  const digits = String(value || "").replace(/[^\d]/g, "");
  return digits ? Number(digits) : 0;
}

function formatOrderDraftMoneyInput(value) {
  const amount = parseOrderDraftMoneyInput(value);
  return amount > 0 ? Math.round(amount).toLocaleString("vi-VN") : "";
}

function getOrderItemSuggestions(query = "") {
  const normalizedQuery = String(query || "").trim().toLowerCase();
  const products = getSortedOrderProducts()
    .map((item) => ({
      name: String(item?.name || "").trim(),
      code: String(item?.code || "").trim(),
    }))
    .filter((item) => item.name);

  if (!normalizedQuery) {
    return products;
  }

  return products
    .filter((item) => item.name.toLowerCase().includes(normalizedQuery) || item.code.toLowerCase().includes(normalizedQuery))
    .sort((left, right) => {
      const leftStarts = left.name.toLowerCase().startsWith(normalizedQuery) || left.code.toLowerCase().startsWith(normalizedQuery) ? 0 : 1;
      const rightStarts = right.name.toLowerCase().startsWith(normalizedQuery) || right.code.toLowerCase().startsWith(normalizedQuery) ? 0 : 1;
      if (leftStarts !== rightStarts) {
        return leftStarts - rightStarts;
      }

      return left.name.localeCompare(right.name, "vi", { numeric: true, sensitivity: "base" });
    });
}

function buildHighlightedOrderItemLabel(item, query = "") {
  const code = String(item?.code || "").trim();
  const name = String(item?.name || "");
  const source = code ? `${code} • ${name}` : name;
  const normalizedQuery = String(query || "").trim().toLowerCase();
  if (!normalizedQuery) {
    return escapeHtml(source);
  }

  const matchIndex = source.toLowerCase().indexOf(normalizedQuery);
  if (matchIndex < 0) {
    return escapeHtml(source);
  }

  const before = source.slice(0, matchIndex);
  const match = source.slice(matchIndex, matchIndex + normalizedQuery.length);
  const after = source.slice(matchIndex + normalizedQuery.length);
  return `${escapeHtml(before)}<span class="order-item-suggestion-match">${escapeHtml(match)}</span>${escapeHtml(after)}`;
}

function resolveOrderProductConfig(productName) {
  const normalizedName = String(productName || "").trim().toLowerCase();
  return (orderProducts.length ? orderProducts : getDefaultOrderProducts()).find(
    (item) => String(item?.name || "").trim().toLowerCase() === normalizedName,
  );
}

function buildOrderItemUnitOptions(productName = "", selectedUnit = "") {
  const config = resolveOrderProductConfig(productName);
  const units = Array.isArray(config?.units) && config.units.length ? config.units : ["cây", "cuộn"];
  const normalizedSelected = String(selectedUnit || config?.default_unit || units[0] || "").trim();
  return units
    .map((unit) => `<option value="${escapeHtml(unit)}" ${unit === normalizedSelected ? "selected" : ""}>${escapeHtml(unit)}</option>`)
    .join("");
}

function renderOrderItemEmptyState() {
  if (!orderItemsList || !orderItemsEmpty) {
    return;
  }

  orderItemsEmpty.classList.toggle("hidden", orderItemsList.children.length > 0);
}

function closeOrderItemSuggestions(exceptInput = null) {
  document.querySelectorAll(".order-item-name-wrap").forEach((wrap) => {
    const input = wrap.querySelector('[data-field="name"]');
    if (exceptInput && input === exceptInput) {
      return;
    }

    wrap.dataset.expanded = "false";
    wrap.dataset.activeIndex = "-1";
    const panel = wrap.querySelector(".order-item-suggestions");
    if (panel) {
      panel.innerHTML = "";
      panel.classList.add("hidden");
    }
  });
}

function setActiveOrderItemSuggestion(wrap, nextIndex) {
  if (!wrap) {
    return;
  }

  const buttons = Array.from(wrap.querySelectorAll("[data-order-item-option]"));
  if (!buttons.length) {
    wrap.dataset.activeIndex = "-1";
    return;
  }

  const boundedIndex = Math.max(0, Math.min(nextIndex, buttons.length - 1));
  wrap.dataset.activeIndex = String(boundedIndex);

  buttons.forEach((button, index) => {
    button.classList.toggle("active", index === boundedIndex);
  });

  buttons[boundedIndex]?.scrollIntoView({ block: "nearest" });
}

function renderOrderItemSuggestions(input, queryOverride = null) {
  const wrap = input?.closest(".order-item-name-wrap");
  const panel = wrap?.querySelector(".order-item-suggestions");
  if (!wrap || !panel || !input) {
    return;
  }

  const query = queryOverride === null ? input.value : queryOverride;
  const suggestions = getOrderItemSuggestions(query);
  const searchMarkup = `
    <div class="order-item-suggestions-search">
      <input
        type="text"
        class="order-item-suggestion-search-input"
        data-order-item-search
        placeholder="Tìm sản phẩm..."
        value="${escapeHtml(query)}"
        autocomplete="off"
        spellcheck="false"
      />
    </div>
  `;
  if (!suggestions.length) {
    wrap.dataset.expanded = "true";
    wrap.dataset.activeIndex = "-1";
    panel.innerHTML = `${searchMarkup}<div class="order-item-suggestion-empty">Không tìm thấy hàng hóa phù hợp</div>`;
    panel.classList.remove("hidden");
    return;
  }

  panel.innerHTML =
    searchMarkup +
    suggestions
      .slice(0, 10)
      .map(
        (item) => `
          <button type="button" class="order-item-suggestion" data-order-item-option="${escapeHtml(item.name)}">
            ${buildHighlightedOrderItemLabel(item, query)}
          </button>
        `,
      )
      .join("");
  wrap.dataset.expanded = "true";
  panel.classList.remove("hidden");
  setActiveOrderItemSuggestion(wrap, 0);
}

function selectOrderItemSuggestion(input, productName) {
  if (!input) {
    return;
  }

  const wrap = input.closest(".order-item-name-wrap");
  input.value = productName;
  if (wrap) {
    wrap.dataset.locked = "true";
  }
  syncProductionOrderRowProductDetails(input.closest(".production-order-row"));
  closeOrderItemSuggestions();
  syncOrderDraftTotals();
  syncProductionOrderDraft();
  input.focus();
  input.setSelectionRange(input.value.length, input.value.length);
}

function findOrderProductByName(productName = "") {
  const normalizedName = String(productName || "").trim().toLowerCase();
  if (!normalizedName) {
    return null;
  }

  return (
    getSortedOrderProducts().find((item) => String(item?.name || "").trim().toLowerCase() === normalizedName) || null
  );
}

function findOrderProductByCode(productCode = "") {
  const normalizedCode = String(productCode || "").trim().toLowerCase();
  if (!normalizedCode) {
    return null;
  }

  return (
    getSortedOrderProducts().find((item) => String(item?.code || "").trim().toLowerCase() === normalizedCode) || null
  );
}

function syncProductionOrderRowProductDetails(row) {
  if (!row) {
    return;
  }

  const nameInput = row.querySelector('[data-field="name"]');
  const codeInput = row.querySelector('[data-field="code"]');
  const unitSelect = row.querySelector('[data-field="unit"]');
  let matchedProduct = findOrderProductByName(nameInput?.value || "");

  if (!matchedProduct) {
    matchedProduct = findOrderProductByCode(codeInput?.value || "");
    if (matchedProduct && nameInput && !String(nameInput.value || "").trim()) {
      nameInput.value = String(matchedProduct.name || "").trim();
    }
  }

  if (!matchedProduct) {
    return;
  }

  if (codeInput) {
    codeInput.value = String(matchedProduct.code || "").trim();
  }

  if (unitSelect) {
    unitSelect.innerHTML = buildOrderItemUnitOptions(matchedProduct.name || "", unitSelect.value || matchedProduct.default_unit || "");
    if (!unitSelect.value && matchedProduct.default_unit) {
      unitSelect.value = matchedProduct.default_unit;
    }
  }
}

function handleOrderItemNameInput(event) {
  const input = event.target?.closest?.('[data-field="name"]');
  if (!input) {
    return;
  }

  if (isOrderModalReadOnly) {
    return;
  }

  const wrap = input.closest(".order-item-name-wrap");
  if (wrap?.dataset.locked === "true") {
    return;
  }

  renderOrderItemSuggestions(input);
}

function handleOrderItemNameFocus(event) {
  const input = event.target?.closest?.('[data-field="name"]');
  if (!input) {
    return;
  }

  if (isOrderModalReadOnly) {
    closeOrderItemSuggestions();
    return;
  }

  closeOrderItemSuggestions(input);
}

function handleOrderItemNameKeydown(event) {
  const input = event.target?.closest?.('[data-field="name"]');
  if (!input) {
    return;
  }

  if (isOrderModalReadOnly) {
    return;
  }

  const wrap = input.closest(".order-item-name-wrap");
  const buttons = Array.from(wrap?.querySelectorAll("[data-order-item-option]") || []);
  const activeIndex = Number(wrap?.dataset.activeIndex || -1);

  if (event.key === "ArrowDown") {
    event.preventDefault();
    if (!buttons.length) {
      renderOrderItemSuggestions(input);
      return;
    }
    setActiveOrderItemSuggestion(wrap, activeIndex + 1);
    return;
  }

  if (event.key === "ArrowUp") {
    event.preventDefault();
    if (!buttons.length) {
      renderOrderItemSuggestions(input);
      return;
    }
    setActiveOrderItemSuggestion(wrap, activeIndex <= 0 ? 0 : activeIndex - 1);
    return;
  }

  if (event.key === "Enter" && buttons.length && activeIndex >= 0) {
    event.preventDefault();
    selectOrderItemSuggestion(input, buttons[activeIndex].dataset.orderItemOption || "");
    return;
  }

  if (event.key === "Escape") {
    closeOrderItemSuggestions();
  }
}

function handleOrderItemSuggestionClick(event) {
  if (isOrderModalReadOnly) {
    return;
  }

  const button = event.target?.closest?.("[data-order-item-option]");
  if (!button) {
    return;
  }

  const input = button.closest(".order-item-name-wrap")?.querySelector('[data-field="name"]');
  selectOrderItemSuggestion(input, button.dataset.orderItemOption || "");
}

function handleOrderItemToggleClick(event) {
  if (isOrderModalReadOnly) {
    return;
  }

  const button = event.target?.closest?.("[data-order-item-toggle]");
  if (!button) {
    return;
  }

  const wrap = button.closest(".order-item-name-wrap");
  const input = wrap?.querySelector('[data-field="name"]');
  const isExpanded = wrap?.dataset.expanded === "true";
  if (!input || !wrap) {
    return;
  }

  if (isExpanded) {
    closeOrderItemSuggestions();
    wrap.dataset.locked = "true";
    return;
  }

  wrap.dataset.locked = "false";
  closeOrderItemSuggestions(input);
  renderOrderItemSuggestions(input, "");
  input.focus();
}

function handleProductionOrderNameClick(event) {
  if (isOrderModalReadOnly) {
    return;
  }

  const input = event.target?.closest?.(".production-order-name-wrap [data-field='name']");
  if (!input) {
    return;
  }

  const wrap = input.closest(".production-order-name-wrap");
  if (!wrap) {
    return;
  }

  if (wrap.dataset.expanded === "true") {
    closeOrderItemSuggestions();
    wrap.dataset.locked = "true";
    input.blur();
    return;
  }

  wrap.dataset.locked = "false";
  closeOrderItemSuggestions(input);
  renderOrderItemSuggestions(input, "");
  input.focus();
}

function handleProductionClaimRowClick(event) {
  if (!orderModal?.classList.contains("production-claim-partial-mode")) {
    return;
  }

  const target = event.target;
  if (
    target?.closest?.(".order-item-name-wrap") ||
    target?.closest?.("input:not([data-production-claim-checkbox])") ||
    target?.closest?.("select") ||
    target?.closest?.("textarea") ||
    target?.closest?.("button")
  ) {
    return;
  }

  const row = target?.closest?.(".production-order-row");
  const checkbox = row?.querySelector?.("[data-production-claim-checkbox]");
  if (!row || !checkbox) {
    return;
  }
  if (row.dataset.claimed === "true") {
    return;
  }

  if (target === checkbox) {
    return;
  }

  toggleProductionClaimSelected(row);
}

function handleProductionClaimCheckboxChange(event) {
  const checkbox = event.target?.closest?.("[data-production-claim-checkbox]");
  const row = checkbox?.closest?.(".production-order-row");
  if (!checkbox || !row) {
    return;
  }
  if (row.dataset.claimed === "true") {
    checkbox.checked = false;
    return;
  }

  setProductionClaimSelected(row, checkbox.checked);
}

function setProductionClaimSelected(row, selected) {
  if (!row) {
    return;
  }
  if (row.dataset.claimed === "true") {
    row.dataset.claimSelected = "false";
    const claimedCheckbox = row.querySelector("[data-production-claim-checkbox]");
    if (claimedCheckbox) {
      claimedCheckbox.checked = false;
    }
    return;
  }
  const nextSelected = Boolean(selected);
  row.dataset.claimSelected = nextSelected ? "true" : "false";
  const checkbox = row.querySelector("[data-production-claim-checkbox]");
  if (checkbox) {
    checkbox.checked = nextSelected;
  }
}

function toggleProductionClaimSelected(row) {
  if (!row || !orderModal?.classList.contains("production-claim-partial-mode")) {
    return;
  }
  if (row.dataset.claimed === "true") {
    return;
  }
  setProductionClaimSelected(row, row.dataset.claimSelected !== "true");
}

function handleOrderItemSuggestionOutsideClick(event) {
  if (event.target?.closest?.(".order-item-name-wrap")) {
    return;
  }

  closeOrderItemSuggestions();
}

function handleOrderItemSearchInput(event) {
  if (isOrderModalReadOnly) {
    return;
  }

  const searchInput = event.target?.closest?.("[data-order-item-search]");
  if (!searchInput) {
    return;
  }

  const wrap = searchInput.closest(".order-item-name-wrap");
  const input = wrap?.querySelector('[data-field="name"]');
  if (!wrap || !input) {
    return;
  }

  wrap.dataset.locked = "false";
  renderOrderItemSuggestions(input, searchInput.value || "");
  const nextSearchInput = wrap.querySelector("[data-order-item-search]");
  nextSearchInput?.focus();
  nextSearchInput?.setSelectionRange(nextSearchInput.value.length, nextSearchInput.value.length);
}

function handleOrderItemSearchKeydown(event) {
  if (isOrderModalReadOnly) {
    return;
  }

  const searchInput = event.target?.closest?.("[data-order-item-search]");
  if (!searchInput) {
    return;
  }

  const wrap = searchInput.closest(".order-item-name-wrap");
  const input = wrap?.querySelector('[data-field="name"]');
  const buttons = Array.from(wrap?.querySelectorAll("[data-order-item-option]") || []);
  const activeIndex = Number(wrap?.dataset.activeIndex || -1);
  if (!wrap || !input) {
    return;
  }

  if (event.key === "ArrowDown") {
    event.preventDefault();
    setActiveOrderItemSuggestion(wrap, buttons.length ? activeIndex + 1 : 0);
    return;
  }

  if (event.key === "ArrowUp") {
    event.preventDefault();
    setActiveOrderItemSuggestion(wrap, activeIndex <= 0 ? 0 : activeIndex - 1);
    return;
  }

  if (event.key === "Enter" && buttons.length && activeIndex >= 0) {
    event.preventDefault();
    selectOrderItemSuggestion(input, buttons[activeIndex].dataset.orderItemOption || "");
    return;
  }

  if (event.key === "Escape") {
    event.preventDefault();
    closeOrderItemSuggestions();
    input.focus();
  }
}

function addOrderItemRow(selectedValue = "", selectedUnit = "") {
  if (!orderItemsList) {
    return;
  }

  const row = document.createElement("div");
  row.className = "order-item-row";
  row.setAttribute("data-order-item-row", "true");
  row.innerHTML = `
    <div class="order-item-index" data-field="index">1</div>
    <div class="order-item-name-wrap">
      <input
        class="order-item-name"
        data-field="name"
        type="text"
        autocomplete="off"
        spellcheck="false"
        placeholder="Tìm hoặc nhập hàng hóa"
        value="${escapeHtml(selectedValue)}"
        required
      />
      <button type="button" class="order-item-name-toggle" data-order-item-toggle aria-label="Mở danh sách hàng hóa">
        <span></span>
      </button>
      <div class="order-item-suggestions hidden"></div>
    </div>
    <input
      class="order-item-quantity"
      data-field="quantity"
      type="number"
      min="0"
      step="1"
      value="1"
      placeholder="SL"
      required
    />
    <select class="order-item-unit" data-field="unit" required>
      ${buildOrderItemUnitOptions(selectedValue, selectedUnit)}
    </select>
    <input
      class="order-item-price"
      data-field="price"
      type="text"
      inputmode="numeric"
      placeholder="Đơn giá"
      required
    />
    <div class="order-item-line-summary">
      <span class="order-item-line-label">Thành tiền</span>
      <div class="order-item-total" data-field="line-total">0 ₫</div>
    </div>
  `;
  orderItemsList.appendChild(row);
  syncOrderDraftTotals();
  row.querySelector('[data-field="name"]')?.focus();
}

function addProductionOrderItemRow(values = {}) {
  if (!productionOrderItemsList) {
    return;
  }

  const quantityValue = String(values.quantity || "").trim();
  const doneValue = String(values.done || "").trim();
  const missingValue = String(values.missing || "").trim();
  const extraValue = String(values.extra || "").trim();
  const quantityNumber = Math.max(0, Number(quantityValue || 0));
  const doneNumber = Math.max(0, Number(doneValue || 0));
  const missingNumber = Math.max(0, Number(missingValue || 0));
  const extraNumber = Math.max(0, Number(extraValue || 0));
  const shouldDefaultMissingToQuantity =
    quantityNumber > 0 && doneNumber === 0 && missingNumber === 0 && extraNumber === 0;
  const hasProgressValues = (doneValue || missingValue || extraValue) && !shouldDefaultMissingToQuantity;
  const normalizedDoneValue = doneValue || "0";
  const normalizedMissingValue = hasProgressValues ? missingValue || "0" : quantityValue || "0";
  const normalizedExtraValue = extraValue || "0";

  const row = document.createElement("div");
  row.className = "production-order-row";
  row.setAttribute("data-order-item-row", "true");
  row.innerHTML = `
    <div class="production-order-index" data-field="index">
      <input type="checkbox" class="production-order-index-claim" data-production-claim-checkbox aria-label="Chọn dòng sản xuất" />
      <span data-field="index-label">1</span>
    </div>
    <input class="production-order-code" data-field="code" type="text" value="${escapeHtml(values.code || "")}" />
    <div class="order-item-name-wrap production-order-name-wrap">
      <span class="production-order-claimed-badge hidden" data-production-claimed-badge></span>
      <textarea
        class="order-item-name"
        data-field="name"
        autocomplete="off"
        spellcheck="false"
        value="${escapeHtml(values.name || "")}"
        required
        rows="1"
      ></textarea>
      <button type="button" class="order-item-name-toggle" data-order-item-toggle aria-label="Mở danh sách hàng hóa">
        <span></span>
      </button>
      <div class="order-item-suggestions hidden"></div>
    </div>
    <input class="production-order-norm" data-field="norm" type="text" value="${escapeHtml(values.norm || "")}" />
    <select class="production-order-unit" data-field="unit" required>
      <option value=""></option>
      ${buildOrderItemUnitOptions(values.name || "", values.unit || "")}
    </select>
    <input class="production-order-quantity" data-field="quantity" type="number" min="0" step="1" value="${escapeHtml(quantityValue)}" required />
    <input class="production-order-done" data-field="done" data-production-progress-field="done" type="number" min="0" step="1" value="${escapeHtml(normalizedDoneValue)}" />
    <input class="production-order-missing" data-field="missing" type="number" min="0" step="1" value="${escapeHtml(normalizedMissingValue)}" />
    <input class="production-order-extra" data-field="extra" type="number" min="0" step="1" value="${escapeHtml(normalizedExtraValue)}" />
    <input class="production-order-team" data-field="team" data-production-progress-field="team" type="number" min="0" step="1" value="${escapeHtml(values.team || "")}" />
  `;
  row.dataset.claimSelected = "false";
  row.querySelector(".production-order-index")?.addEventListener("click", (event) => {
    event.preventDefault();
    event.stopPropagation();
    toggleProductionClaimSelected(row);
  });
  row.querySelector("[data-production-claim-checkbox]")?.addEventListener("click", (event) => {
    event.stopPropagation();
  });
  row.querySelector("[data-production-claim-checkbox]")?.addEventListener("change", (event) => {
    setProductionClaimSelected(row, event.target.checked);
  });
  row.querySelectorAll("[data-production-progress-field]").forEach((input) => {
    input.addEventListener("input", () => {
      updateProductionProgressDerivedFields(row);
      row.dataset.progressDirty = row.dataset.claimEditable === "true" ? "true" : "false";
      syncProductionOrderDraft();
      updateProductionProgressButtons();
    });
    input.addEventListener("change", () => {
      updateProductionProgressDerivedFields(row);
      row.dataset.progressDirty = row.dataset.claimEditable === "true" ? "true" : "false";
      syncProductionOrderDraft();
      updateProductionProgressButtons();
    });
  });
  productionOrderItemsList.appendChild(row);
  const nameField = row.querySelector('[data-field="name"]');
  if (nameField) {
    nameField.value = String(values.name || "");
  }
  syncProductionOrderDraft();
  updateProductionProgressDerivedFields(row);
  autoResizeProductionOrderName(nameField);
  nameField?.focus();
}

function autoResizeProductionOrderName(input) {
  if (!input || input.tagName !== "TEXTAREA") {
    return;
  }

  input.style.height = "0px";
  input.style.height = `${Math.max(72, input.scrollHeight)}px`;
}

function syncProductionOrderDraft() {
  if (!productionOrderItemsList) {
    return;
  }

  const rows = Array.from(productionOrderItemsList.querySelectorAll("[data-order-item-row]"));
  rows.forEach((row, index) => {
    const indexNode = row.querySelector('[data-field="index-label"]');
    const unitSelect = row.querySelector('[data-field="unit"]');
    const name = row.querySelector('[data-field="name"]')?.value?.trim() || "";
    syncProductionOrderRowProductDetails(row);
    const availableUnits = buildOrderItemUnitOptions(name, unitSelect?.value || "");
    if (unitSelect && unitSelect.innerHTML !== availableUnits) {
      unitSelect.innerHTML = availableUnits;
    }
    autoResizeProductionOrderName(row.querySelector('[data-field="name"]'));
    if (indexNode) {
      indexNode.textContent = String(index + 1);
    }
  });

  if (productionOrderItemsEmpty) {
    productionOrderItemsEmpty.classList.toggle("hidden", rows.length > 0);
  }
}

function buildProductionOrderItemsSummary() {
  const rows = Array.from(productionOrderItemsList?.querySelectorAll("[data-order-item-row]") || []);
  return rows
    .map((row, index) => {
      const code = row.querySelector('[data-field="code"]')?.value?.trim() || "-";
      const name = row.querySelector('[data-field="name"]')?.value?.trim() || "";
      const norm = row.querySelector('[data-field="norm"]')?.value?.trim() || "-";
      const unit = row.querySelector('[data-field="unit"]')?.value?.trim() || "-";
      const quantity = row.querySelector('[data-field="quantity"]')?.value?.trim() || "0";
      const done = row.querySelector('[data-field="done"]')?.value?.trim() || "0";
      const missing = row.querySelector('[data-field="missing"]')?.value?.trim() || "0";
      const extra = row.querySelector('[data-field="extra"]')?.value?.trim() || "0";
      const team = row.querySelector('[data-field="team"]')?.value?.trim() || "-";
      return `${index + 1}. ${code} | ${name} | ĐM ${norm} | ${unit} | SL ${quantity} | Đã SX ${done} | Thiếu ${missing} | Dư ${extra} | Tổ ${team}`;
    })
    .filter(Boolean)
    .join("\n");
}

function collectProductionClaimItems(selectedOnly = false) {
  const rows = Array.from(productionOrderItemsList?.querySelectorAll("[data-order-item-row]") || []);
  return rows
    .map((row, index) => {
      if (row.dataset.claimed === "true") {
        return null;
      }
      const checkbox = row.querySelector("[data-production-claim-checkbox]");
      const isSelected = row.dataset.claimSelected === "true" || checkbox?.checked;
      if (selectedOnly && !isSelected) {
        return null;
      }

      return {
        line_number: index + 1,
        code: row.querySelector('[data-field="code"]')?.value?.trim() || "",
        name: row.querySelector('[data-field="name"]')?.value?.trim() || "",
        quantity: row.querySelector('[data-field="quantity"]')?.value?.trim() || "0",
        unit: row.querySelector('[data-field="unit"]')?.value?.trim() || "",
      };
    })
    .filter(Boolean);
}

function formatProductionClaimItem(item) {
  const quantity = String(item?.quantity || "").trim();
  const unit = String(item?.unit || "").trim();
  const name = String(item?.name || item?.code || "hạng mục").trim();
  return [quantity, unit, name].filter(Boolean).join(" ");
}

async function submitProductionProgress(action = "confirm") {
  const dirtyRows = collectPendingProductionProgressRows();
  if (!dirtyRows.length) {
    return;
  }
  const hasMissingRequiredField = dirtyRows.some((row) => {
    const doneValue = String(row.querySelector('[data-field="done"]')?.value || "").trim();
    const teamValue = String(row.querySelector('[data-field="team"]')?.value || "").trim();
    return !doneValue || !teamValue;
  });
  if (hasMissingRequiredField) {
    showToast("Cần nhập đủ Đã SX và Tổ máy trước khi xác nhận.", "error");
    return;
  }

  const orderId = normalizeOrderIdValue(productionOrderIdInput?.value || activeOrderDetails?.order_id || "");
  if (!orderId) {
    return;
  }

  const isCompleteAction = action === "complete";
  if (
    isCompleteAction &&
    dirtyRows.some((row) => {
      const quantity = Math.max(0, parseProductionNumber(row.querySelector('[data-field="quantity"]')?.value || 0));
      const done = Math.max(0, parseProductionNumber(row.querySelector('[data-field="done"]')?.value || 0));
      return done < quantity;
    })
  ) {
    showToast("Chỉ dùng xác nhận hoàn thành khi Đã SX lớn hơn hoặc bằng SL.", "error");
    return;
  }

  confirmProductionProgressButton && (confirmProductionProgressButton.disabled = true);
  completeProductionProgressButton && (completeProductionProgressButton.disabled = true);
  try {
    const progressUpdates = dirtyRows.map((row) => ({
      line_number: Number(row.dataset.lineNumber || 0),
      done: row.querySelector('[data-field="done"]')?.value?.trim() || "0",
      team: row.querySelector('[data-field="team"]')?.value?.trim() || "",
    }));
    let payload;
    try {
      payload = await postJson(orderProductionProgressUrl, {
        order_id: orderId,
        progress_updates: progressUpdates,
        order_items: buildProductionOrderItemsSummary(),
        mark_completed: isCompleteAction,
      });
    } catch (error) {
      if (!String(error?.message || "").toLowerCase().includes("method")) {
        throw error;
      }
      payload = await postJson(orderUpdateUrl, {
        order_id: orderId,
        progress_updates: progressUpdates,
        order_items: buildProductionOrderItemsSummary(),
        mark_completed: isCompleteAction,
      });
    }
    applyWorkspacePayload(payload);
    if (Array.isArray(payload.orders)) {
      activeOrderDetails =
        payload.orders.find((item) => String(item?.order_id || "").trim() === String(orderId || "").trim()) || activeOrderDetails;
      if (activeOrderDetails) {
        populateProductionOrderForm(activeOrderDetails);
        setOrderFormReadOnly(true);
        updateProductionClaimButtons(activeOrderDetails);
      }
    }
    showToast(isCompleteAction ? "Đã xác nhận hoàn thành sản xuất." : "Đã xác nhận tiến độ sản xuất.", "success");
  } catch (error) {
    showToast(error.message || "Không thể lưu tiến độ sản xuất.", "error");
  } finally {
    confirmProductionProgressButton && (confirmProductionProgressButton.disabled = false);
    completeProductionProgressButton && (completeProductionProgressButton.disabled = false);
    updateProductionProgressButtons();
  }
}

function isProductionPackagingCompleted(order = activeOrderDetails) {
  const packagingRecord = parseProductionPackagingRecord(order?.production_packaged_by_json);
  return Boolean(String(packagingRecord?.user_name || packagingRecord?.user_id || "").trim());
}

function isProductionReceiptCompleted(order = activeOrderDetails) {
  const receiptRecord = parseProductionPackagingRecord(order?.production_received_by_json);
  return Boolean(String(receiptRecord?.user_name || receiptRecord?.user_id || "").trim());
}

function isProductionOrderLockedForEditing(order) {
  return Boolean(isProductionOrder(order) && isProductionReceiptCompleted(order));
}

function isCurrentUserAdmin() {
  return String(currentUser?.role || "").trim().toLowerCase() === "admin";
}

function shouldShowCompletedProductionAdminEditButton(order) {
  return Boolean(
    activeOrderCreateKind === "production" &&
      isOrderModalReadOnly &&
      isCurrentUserAdmin() &&
      isProductionOrderLockedForEditing(order),
  );
}

function areAllProductionMissingValuesZero() {
  const rows = Array.from(productionOrderItemsList?.querySelectorAll(".production-order-row") || []);
  if (!rows.length) {
    return false;
  }
  return rows.every((row) => {
    const missing = Math.max(0, parseProductionNumber(row.querySelector('[data-field="missing"]')?.value || 0));
    return missing === 0;
  });
}

function updateProductionPackagingButton(order = activeOrderDetails) {
  if (!completeProductionPackagingButton) {
    return;
  }
  if (activeOrderCreateKind !== "production") {
    completeProductionPackagingButton.classList.add("hidden");
    completeProductionPackagingButton.disabled = true;
    return;
  }
  const hasDirtyRows = collectPendingProductionProgressRows().length > 0;
  const canShowButton =
    isOrderModalReadOnly &&
    activeOrderCreateKind === "production" &&
    !hasDirtyRows &&
    !isProductionPackagingCompleted(order) &&
    areAllProductionMissingValuesZero();
  completeProductionPackagingButton.classList.toggle("hidden", !canShowButton);
  completeProductionPackagingButton.disabled = !canShowButton;
}

function updateProductionReceiptButton(order = activeOrderDetails) {
  if (!completeProductionReceiptButton) {
    return;
  }
  if (activeOrderCreateKind !== "production") {
    completeProductionReceiptButton.classList.add("hidden");
    completeProductionReceiptButton.disabled = true;
    return;
  }
  const hasDirtyRows = collectPendingProductionProgressRows().length > 0;
  const canShowButton =
    isOrderModalReadOnly &&
    activeOrderCreateKind === "production" &&
    !hasDirtyRows &&
    isProductionPackagingCompleted(order) &&
    !isProductionReceiptCompleted(order) &&
    isOrderCreatedByCurrentUser(order);
  completeProductionReceiptButton.classList.toggle("hidden", !canShowButton);
  completeProductionReceiptButton.disabled = !canShowButton;
}

function syncProductionCompletionStamp(order = activeOrderDetails) {
  if (!productionCompletionStamp) {
    return;
  }
  const isCompleted = activeOrderCreateKind === "production" && isProductionReceiptCompleted(order);
  productionCompletionStamp.classList.toggle("hidden", !isCompleted);
  productionCompletionStamp.setAttribute("aria-hidden", isCompleted ? "false" : "true");
}

function openProductionPackagingConfirmModal() {
  productionPackagingConfirmModal?.classList.remove("hidden");
  productionPackagingConfirmBackdrop?.classList.remove("hidden");
  productionPackagingConfirmModal?.setAttribute("aria-hidden", "false");
}

function closeProductionPackagingConfirmModal() {
  productionPackagingConfirmModal?.classList.add("hidden");
  productionPackagingConfirmBackdrop?.classList.add("hidden");
  productionPackagingConfirmModal?.setAttribute("aria-hidden", "true");
}

function openProductionReceiptConfirmModal() {
  productionReceiptConfirmModal?.classList.remove("hidden");
  productionReceiptConfirmBackdrop?.classList.remove("hidden");
  productionReceiptConfirmModal?.setAttribute("aria-hidden", "false");
}

function closeProductionReceiptConfirmModal() {
  productionReceiptConfirmModal?.classList.add("hidden");
  productionReceiptConfirmBackdrop?.classList.add("hidden");
  productionReceiptConfirmModal?.setAttribute("aria-hidden", "true");
}

async function submitProductionPackagingComplete() {
  const orderId = normalizeOrderIdValue(productionOrderIdInput?.value || activeOrderDetails?.order_id || "");
  if (!orderId) {
    return;
  }
  if (!areAllProductionMissingValuesZero()) {
    showToast("Chỉ hoàn tất đóng gói khi tất cả giá trị cột Thiếu đã về 0.", "error");
    return;
  }
  if (confirmProductionPackagingCompleteButton) {
    confirmProductionPackagingCompleteButton.disabled = true;
  }
  try {
    const payload = await postJson(orderProductionPackagingCompleteUrl, {
      order_id: orderId,
    });
    applyWorkspacePayload(payload);
    if (Array.isArray(payload.orders)) {
      activeOrderDetails =
        payload.orders.find((item) => String(item?.order_id || "").trim() === String(orderId || "").trim()) || activeOrderDetails;
      if (activeOrderDetails) {
        populateProductionOrderForm(activeOrderDetails);
        setOrderFormReadOnly(true);
        updateProductionClaimButtons(activeOrderDetails);
      }
    }
    closeProductionPackagingConfirmModal();
    showToast("Đã hoàn tất đóng gói phiếu sản xuất.", "success");
  } catch (error) {
    showToast(error.message || "Không thể hoàn tất đóng gói.", "error");
  } finally {
    if (confirmProductionPackagingCompleteButton) {
      confirmProductionPackagingCompleteButton.disabled = false;
    }
    updateProductionPackagingButton();
    updateProductionReceiptButton();
  }
}

async function submitProductionReceiptComplete() {
  const orderId = normalizeOrderIdValue(productionOrderIdInput?.value || activeOrderDetails?.order_id || "");
  if (!orderId) {
    return;
  }
  if (confirmProductionReceiptCompleteButton) {
    confirmProductionReceiptCompleteButton.disabled = true;
  }
  try {
    const payload = await postJson(orderProductionReceiptCompleteUrl, {
      order_id: orderId,
    });
    applyWorkspacePayload(payload);
    if (Array.isArray(payload.orders)) {
      activeOrderDetails =
        payload.orders.find((item) => String(item?.order_id || "").trim() === String(orderId || "").trim()) || activeOrderDetails;
      if (activeOrderDetails) {
        populateProductionOrderForm(activeOrderDetails);
        setOrderFormReadOnly(true);
        updateProductionClaimButtons(activeOrderDetails);
      }
    }
    closeProductionReceiptConfirmModal();
    showToast("Đã xác nhận nhận đủ hàng.", "success");
  } catch (error) {
    showToast(error.message || "Không thể xác nhận nhận đủ hàng.", "error");
  } finally {
    if (confirmProductionReceiptCompleteButton) {
      confirmProductionReceiptCompleteButton.disabled = false;
    }
    updateProductionReceiptButton();
  }
}

async function submitProductionClaim(mode = "all") {
  const fallbackOrderId = normalizeOrderIdValue(productionOrderIdInput?.value || "");
  const order =
    activeOrderDetails ||
    orders.find((item) => String(item?.order_id || "").trim() === fallbackOrderId);
  if (!order) {
    showToast("Không tìm thấy phiếu sản xuất đang mở để nhận.", "error");
    return;
  }

  const claimMode = mode === "partial" ? "partial" : "all";
  const claimedItems = claimMode === "partial" ? collectProductionClaimItems(true) : collectProductionClaimItems(false);
  if (claimMode === "all" && !claimedItems.length) {
    showToast("Phiếu này không còn phần nào chưa được nhận.", "error");
    return;
  }
  if (claimMode === "partial" && !claimedItems.length) {
    showToast("Hãy chọn ít nhất một dòng để nhận sản xuất.", "error");
    return;
  }

  const payload = await postJson(orderProductionClaimUrl, {
    order_id: order.order_id,
    mode: claimMode,
    claimed_items: claimedItems,
  });
  applyWorkspacePayload(payload);
  if (Array.isArray(payload.orders)) {
    activeOrderDetails =
      payload.orders.find((item) => String(item?.order_id || "").trim() === String(order.order_id || "").trim()) || activeOrderDetails;
    if (activeOrderDetails) {
      await openOrderPreviewModal(activeOrderDetails);
    }
  }
  const successMessage =
    claimMode === "partial" ? "Đã nhận một phần phiếu sản xuất và gửi thông báo." : "Đã nhận toàn bộ phiếu sản xuất và gửi thông báo.";
  showToast(successMessage, "success");
}

function canChooseProductionRequester() {
  return String(currentUser?.role || "").trim().toLowerCase() === "admin";
}

function populateProductionRequesterOptions() {
  if (!productionRequesterSelect) {
    return;
  }

  const currentDepartment = String(currentUser?.department || "").trim().toLowerCase();
  const salesUsers = [...allUsers]
    .filter((user) => String(user?.department || "").trim().toLowerCase() === "sales")
    .sort((left, right) => String(left.name || "").localeCompare(String(right.name || ""), "vi"));

  const defaultRequester = canChooseProductionRequester()
    ? String(currentUser?.id || "")
    : currentDepartment === "sales"
      ? String(currentUser?.id || "")
      : "";
  const availableUsers = canChooseProductionRequester()
    ? salesUsers
    : salesUsers.filter((user) => String(user?.id || "") === defaultRequester);

  productionRequesterSelect.innerHTML = availableUsers.length
    ? availableUsers
        .map((user) => {
          const code = user.employee_code || user.username || user.id;
          const selected = String(user.id || "") === defaultRequester ? "selected" : "";
          return `<option value="${escapeHtml(user.id)}" ${selected}>${escapeHtml(`${user.name} • ${code}`)}</option>`;
        })
        .join("")
    : `<option value="">Chưa có người yêu cầu phù hợp</option>`;
  productionRequesterSelect.disabled = !canChooseProductionRequester();
  productionRequesterField?.classList.toggle("hidden", !canChooseProductionRequester());
  syncProductionRequesterSignature();
}

function syncProductionRequesterSignature() {
  if (!productionRequesterSignature) {
    return;
  }

  const selectedId = String(productionRequesterSelect?.value || "").trim();
  const selectedUser = allUsers.find((user) => String(user?.id || "").trim() === selectedId);
  productionRequesterSignature.textContent = selectedUser?.name || "-";
}

function parseProductionPackagingRecord(value) {
  try {
    const parsed = JSON.parse(value || "{}");
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch {
    return {};
  }
}

function syncProductionPerformerSignature(order = activeOrderDetails) {
  if (!productionPerformerSignature) {
    return;
  }
  const completions = parseProductionCompletions(order?.production_completed_by_json);
  if (!completions.length) {
    productionPerformerSignature.textContent = "-";
    return;
  }
  productionPerformerSignature.innerHTML = completions
    .map((entry) => `<div class="production-signature-entry">${escapeHtml(String(entry?.user_name || "").trim())}</div>`)
    .join("");
}

function syncProductionPackagingSignature(order = activeOrderDetails) {
  if (!productionPackagingSignature) {
    return;
  }
  const packagingRecord = parseProductionPackagingRecord(order?.production_packaged_by_json);
  const packagedByName = String(packagingRecord?.user_name || "").trim();
  productionPackagingSignature.innerHTML = packagedByName
    ? `<div class="production-signature-entry">${escapeHtml(packagedByName)}</div>`
    : "-";
}

function syncProductionReceiptSignature(order = activeOrderDetails) {
  if (!productionReceiptSignature) {
    return;
  }
  const receiptRecord = parseProductionPackagingRecord(order?.production_received_by_json);
  const receiptByName = String(receiptRecord?.user_name || "").trim();
  productionReceiptSignature.innerHTML = receiptByName
    ? `<div class="production-signature-entry">${escapeHtml(receiptByName)}</div>`
    : "-";
}

function syncOrderDraftTotals() {
  if (!orderItemsList) {
    return;
  }

  let totalValue = 0;
  const itemSummaries = [];

  const rows = Array.from(orderItemsList.querySelectorAll("[data-order-item-row]"));

  rows.forEach((row, index) => {
    const indexNode = row.querySelector('[data-field="index"]');
    const unitSelect = row.querySelector('[data-field="unit"]');
    const name = row.querySelector('[data-field="name"]')?.value?.trim() || "";
    const quantity = Math.max(0, Number(row.querySelector('[data-field="quantity"]')?.value || 0));
    const availableUnits = buildOrderItemUnitOptions(name, unitSelect?.value || "");
    if (unitSelect && unitSelect.innerHTML !== availableUnits) {
      unitSelect.innerHTML = availableUnits;
    }
    const unit = unitSelect?.value?.trim() || "";
    const priceInput = row.querySelector('[data-field="price"]');
    const price = Math.max(0, parseOrderDraftMoneyInput(priceInput?.value || 0));
    const lineTotal = Math.round(quantity * price);
    const lineTotalNode = row.querySelector('[data-field="line-total"]');

    if (indexNode) {
      indexNode.textContent = String(index + 1);
    }

    if (priceInput) {
      const formattedPrice = formatOrderDraftMoneyInput(priceInput.value);
      if (priceInput.value !== formattedPrice) {
        priceInput.value = formattedPrice;
      }
    }

    if (lineTotalNode) {
      lineTotalNode.textContent = formatOrderDraftMoney(lineTotal);
    }

    if (name && price > 0) {
      itemSummaries.push(`${index + 1}. ${name} x ${quantity} ${unit} • ${Math.round(price).toLocaleString("vi-VN")} đ = ${Math.round(lineTotal).toLocaleString("vi-VN")} đ`);
    }

    totalValue += lineTotal;
  });

  renderOrderItemEmptyState();

  if (orderItems) {
    orderItems.value = itemSummaries.join("\n");
  }

  if (orderValue) {
    orderValue.value = totalValue > 0 ? String(totalValue) : "";
  }

  if (orderTotalDisplay) {
    orderTotalDisplay.textContent = formatOrderDraftMoney(totalValue);
  }
}

function setOrderFormStep(step = "1") {
  const normalizedStep = String(step || "1");
  orderStepButtons.forEach((button) => {
    const isActive = button.dataset.orderStep === normalizedStep;
    button.classList.toggle("active", isActive);
    button.setAttribute("aria-selected", isActive ? "true" : "false");
  });

  orderStepPanels.forEach((panel) => {
    const isActive = panel.dataset.orderStepPanel === normalizedStep;
    panel.classList.toggle("hidden", !isActive);
    panel.setAttribute("aria-hidden", isActive ? "false" : "true");
  });
}

function focusFirstInvalidField(container) {
  const invalidField = container?.querySelector?.(":invalid");
  invalidField?.focus();
  invalidField?.reportValidity?.();
  return Boolean(invalidField);
}

function validateOrderItemsStep() {
  if (activeOrderCreateKind === "production") {
    const rows = Array.from(productionOrderItemsList?.querySelectorAll("[data-order-item-row]") || []);
    if (!rows.length) {
      showToast("Cần thêm ít nhất 1 dòng hàng hóa sản xuất.", "error");
      productionAddItemButton?.focus();
      return false;
    }

    for (const row of rows) {
      const nameInput = row.querySelector('[data-field="name"]');
      const quantityInput = row.querySelector('[data-field="quantity"]');
      const unitSelect = row.querySelector('[data-field="unit"]');
      if (!String(nameInput?.value || "").trim()) {
        showToast("Tên hàng hóa là bắt buộc.", "error");
        nameInput?.focus();
        return false;
      }
      if (Number(quantityInput?.value || 0) <= 0) {
        showToast("Số lượng sản xuất phải lớn hơn 0.", "error");
        quantityInput?.focus();
        return false;
      }
      if (!String(unitSelect?.value || "").trim()) {
        showToast("Hãy chọn đơn vị tính cho dòng sản xuất.", "error");
        unitSelect?.focus();
        return false;
      }
    }
    return true;
  }

  const rows = Array.from(orderItemsList?.querySelectorAll("[data-order-item-row]") || []);
  if (!rows.length) {
    setOrderFormStep("2");
    showToast("Cần thêm ít nhất 1 hàng hóa cho đơn hàng.", "error");
    orderAddItemButton?.focus();
    return false;
  }

  for (const row of rows) {
    const nameInput = row.querySelector('[data-field="name"]');
    const quantityInput = row.querySelector('[data-field="quantity"]');
    const unitSelect = row.querySelector('[data-field="unit"]');
    const priceInput = row.querySelector('[data-field="price"]');

    if (!String(nameInput?.value || "").trim()) {
      setOrderFormStep("2");
      showToast("Tên hàng hóa là bắt buộc.", "error");
      nameInput?.focus();
      return false;
    }

    if (Number(quantityInput?.value || 0) <= 0) {
      setOrderFormStep("2");
      showToast("Số lượng phải lớn hơn 0.", "error");
      quantityInput?.focus();
      return false;
    }

    if (!String(unitSelect?.value || "").trim()) {
      setOrderFormStep("2");
      showToast("Đơn vị hàng hóa là bắt buộc.", "error");
      unitSelect?.focus();
      return false;
    }

    if (parseOrderDraftMoneyInput(priceInput?.value || "") <= 0) {
      setOrderFormStep("2");
      showToast("Đơn giá phải lớn hơn 0.", "error");
      priceInput?.focus();
      return false;
    }
  }

  if (!String(orderItems?.value || "").trim() || Number(orderValue?.value || 0) <= 0) {
    setOrderFormStep("2");
    showToast("Đơn hàng chưa có tổng giá trị hợp lệ.", "error");
    orderAddItemButton?.focus();
    return false;
  }

  return true;
}

function validateOrderFormBeforeSubmit() {
  if (activeOrderCreateKind === "production") {
    if (!String(productionOrderIdInput?.value || "").trim()) {
      productionOrderIdInput?.focus();
      productionOrderIdInput?.reportValidity?.();
      return false;
    }
    if (!String(productionCustomerNameInput?.value || "").trim()) {
      productionCustomerNameInput?.focus();
      productionCustomerNameInput?.reportValidity?.();
      return false;
    }
    if (!String(productionRequesterSelect?.value || "").trim()) {
      productionRequesterSelect?.focus();
      productionRequesterSelect?.reportValidity?.();
      return false;
    }
    return validateOrderItemsStep();
  }

  const stepOnePanel = orderStepPanels.find((panel) => panel.dataset.orderStepPanel === "1");
  if (focusFirstInvalidField(stepOnePanel)) {
    setOrderFormStep("1");
    return false;
  }

  if (!validateOrderItemsStep()) {
    return false;
  }

  return checkTransportOrderInventory({ showSuccessToast: false }).ok;
}

function resetOrderFormMode() {
  activeOrderDraftEditId = "";
  activeOrderDraftEditRecord = null;
  isOrderModalReadOnly = false;
  resetProductionClaimMode();
  const titleByKind = {
    production: "Tạo đơn sản xuất mới",
    transport: "Tạo đơn vận chuyển mới",
  };
  if (orderModalTitle) {
    orderModalTitle.textContent = titleByKind[activeOrderCreateKind] || "Tạo đơn hàng mới";
  }
  if (confirmOrderButton) {
    confirmOrderButton.textContent = activeOrderCreateKind === "transport" ? "Tạo đơn vận chuyển" : "Tạo đơn sản xuất";
    confirmOrderButton.classList.remove("hidden");
  }
  checkOrderInventoryButton?.classList.toggle("hidden", activeOrderCreateKind !== "transport");
  confirmProductionProgressButton?.classList.add("hidden");
  completeProductionProgressButton?.classList.add("hidden");
  completeProductionPackagingButton?.classList.add("hidden");
  productionCompleteChip?.classList.add("hidden");
  editCompletedProductionOrderButton?.classList.add("hidden");
  productionCompletionStamp?.classList.add("hidden");
  productionCompletionStamp?.setAttribute("aria-hidden", "true");
  closeProductionPackagingConfirmModal();
  closeProductionReceiptConfirmModal();
  if (productionPerformerSignature) {
    productionPerformerSignature.textContent = "-";
  }
  if (productionPackagingSignature) {
    productionPackagingSignature.textContent = "-";
  }
  if (productionReceiptSignature) {
    productionReceiptSignature.textContent = "-";
  }
  productionAddItemButton?.classList.toggle("hidden", activeOrderCreateKind !== "production");
  orderAddItemButton?.classList.remove("hidden");
  if (cancelOrderModalButton) {
    cancelOrderModalButton.textContent = "Hủy";
  }
  if (orderIdInput) {
    orderIdInput.readOnly = false;
  }
  if (productionOrderIdInput) {
    productionOrderIdInput.readOnly = false;
  }
  setOrderFormReadOnly(false);
  orderModal?.classList.toggle("production-order-mode", activeOrderCreateKind === "production");
  productionOrderSheet?.classList.toggle("hidden", activeOrderCreateKind !== "production");
  orderStepButtons.forEach((button) => button.classList.toggle("hidden", activeOrderCreateKind === "production"));
  orderStepPanels.forEach((panel) => panel.classList.toggle("hidden", activeOrderCreateKind === "production"));
}

function setOrderFormReadOnly(isReadOnly) {
  isOrderModalReadOnly = Boolean(isReadOnly);
  orderModal?.classList.toggle("order-modal-readonly", isOrderModalReadOnly);
  const activeProductionOrder = activeOrderDetails || activeOrderDraftEditRecord;
  const showCompletedAdminEditButton = shouldShowCompletedProductionAdminEditButton(activeProductionOrder);
  const isProductionReceiptLocked =
    activeOrderCreateKind === "production" &&
    isProductionOrderLockedForEditing(activeProductionOrder) &&
    !isCurrentUserAdmin();

  const fields = Array.from(orderForm?.querySelectorAll("input, textarea, select") || []);
  fields.forEach((field) => {
    const tagName = String(field.tagName || "").toLowerCase();
    const type = String(field.getAttribute("type") || "").toLowerCase();
    if (isProductionReceiptLocked) {
      field.disabled = true;
      field.setAttribute("disabled", "disabled");
      if (tagName !== "select" && type !== "radio" && type !== "checkbox") {
        field.readOnly = true;
        field.setAttribute("readonly", "readonly");
      }
      return;
    }
    if (field.hasAttribute("data-production-claim-checkbox")) {
      const canUseCheckbox = activeProductionClaimMode === "partial";
      if (canUseCheckbox) {
        field.disabled = false;
        field.removeAttribute("disabled");
      } else {
        field.disabled = true;
        field.setAttribute("disabled", "disabled");
      }
      return;
    }
    if (field.hasAttribute("data-production-progress-field")) {
      const row = field.closest(".production-order-row");
      const canEditProgress = Boolean(
        isOrderModalReadOnly &&
        row?.dataset.claimEditable === "true" &&
        (!isProductionOrderLockedForEditing(activeProductionOrder) || (isCurrentUserAdmin() && !showCompletedAdminEditButton)),
      );
      field.disabled = false;
      if (canEditProgress) {
        field.readOnly = false;
        field.removeAttribute("readonly");
      } else {
        field.readOnly = isOrderModalReadOnly;
        if (isOrderModalReadOnly) {
          field.setAttribute("readonly", "readonly");
        } else {
          field.removeAttribute("readonly");
        }
      }
      return;
    }
    if (tagName === "select" || type === "radio" || type === "checkbox") {
      field.disabled = isOrderModalReadOnly;
      return;
    }
    field.readOnly = isOrderModalReadOnly;
  });

  const actionButtons = Array.from(orderForm?.querySelectorAll("button") || []);
  actionButtons.forEach((button) => {
    if (
      button === cancelOrderModalButton ||
      button === printProductionOrderButton ||
      button.dataset.orderStep
    ) {
      button.disabled = false;
      return;
    }
    if (isProductionReceiptLocked) {
      button.disabled = true;
      return;
    }
    button.disabled = isOrderModalReadOnly;
  });

  confirmOrderButton?.classList.toggle("hidden", isOrderModalReadOnly);
  productionAddItemButton?.classList.toggle("hidden", isOrderModalReadOnly || isProductionReceiptLocked || activeOrderCreateKind !== "production");
  orderAddItemButton?.classList.toggle("hidden", isOrderModalReadOnly);
  receiveAllProductionOrderButton?.classList.toggle("hidden", isProductionReceiptLocked);
  receivePartialProductionOrderButton?.classList.toggle("hidden", isProductionReceiptLocked);
  confirmPartialProductionOrderButton?.classList.toggle("hidden", isProductionReceiptLocked);
  confirmProductionProgressButton?.classList.toggle("hidden", !isOrderModalReadOnly || isProductionReceiptLocked || showCompletedAdminEditButton);
  completeProductionProgressButton?.classList.toggle("hidden", !isOrderModalReadOnly || isProductionReceiptLocked || showCompletedAdminEditButton);
  completeProductionPackagingButton?.classList.toggle("hidden", isProductionReceiptLocked);
  completeProductionReceiptButton?.classList.toggle("hidden", isProductionReceiptLocked);
  editCompletedProductionOrderButton?.classList.toggle("hidden", !showCompletedAdminEditButton);
  if (editCompletedProductionOrderButton) {
    editCompletedProductionOrderButton.disabled = false;
  }
  updateProductionClaimButtons();
  updateProductionProgressFieldAccess();
  updateProductionPackagingButton();
  updateProductionReceiptButton();
  syncProductionCompletionStamp(activeProductionOrder);
  if (cancelOrderModalButton) {
    cancelOrderModalButton.textContent = isOrderModalReadOnly ? "Đóng" : "Hủy";
  }
}

function populateTransportOrderForm(order) {
  if (orderIdInput) {
    orderIdInput.value = String(order?.order_id || "").replace(/^DH-/i, "");
  }
  if (orderCustomerName) {
    orderCustomerName.value = order?.customer_name || "";
  }
  if (orderSalesUser) {
    orderSalesUser.value = order?.sales_user_id || "";
  }
  if (orderDeliveryUser) {
    orderDeliveryUser.value = order?.delivery_user_id || "";
  }
  if (orderPlannedAt) {
    orderPlannedAt.value = order?.planned_delivery_at ? formatNotificationTime(order.planned_delivery_at) : "";
  }
  if (orderCreatedAt) {
    orderCreatedAt.value = order?.created_at ? formatNotificationTime(order.created_at) : formatNotificationTime(new Date().toISOString());
  }
  if (orderAddress) {
    orderAddress.value = order?.delivery_address || "";
  }
  if (orderNote) {
    orderNote.value = order?.note || order?.completion_note || "";
  }
  populateOrderDraftItems(order?.order_items || "");
}

function populateProductionOrderForm(order) {
  productionRequesterField?.classList.remove("hidden");
  if (productionOrderIdInput) {
    productionOrderIdInput.value = String(order?.order_id || "").replace(/^DH-/i, "");
  }
  if (productionCustomerNameInput) {
    productionCustomerNameInput.value = order?.customer_name || "";
  }
  if (productionRequesterSelect) {
    const requesterId = String(order?.sales_user_id || "").trim();
    const hasExistingOption = Array.from(productionRequesterSelect.options || []).some(
      (option) => String(option.value || "").trim() === requesterId,
    );
    if (requesterId && !hasExistingOption) {
      const requesterUser = allUsers.find((user) => String(user?.id || "").trim() === requesterId);
      if (requesterUser) {
        const option = document.createElement("option");
        option.value = requesterUser.id;
        const code = requesterUser.employee_code || requesterUser.username || requesterUser.id;
        option.textContent = `${requesterUser.name} • ${code}`;
        productionRequesterSelect.append(option);
      }
    }
    productionRequesterSelect.value = order?.sales_user_id || "";
  }
  syncProductionRequesterSignature();
  syncProductionPerformerSignature(order);
  syncProductionPackagingSignature(order);
  syncProductionReceiptSignature(order);
  syncProductionCompletionStamp(order);
  if (productionCreatedAtInput) {
    productionCreatedAtInput.value = order?.created_at ? formatNotificationTime(order.created_at) : formatNotificationTime(new Date().toISOString());
  }
  applyProductionTurnaroundFromNote(order?.note || "");
  populateProductionOrderDraftItems(order?.order_items || "");
  syncProductionClaimRowStates(order);
}

function finishProductionOrderPrintMode() {
  document.body.classList.remove("production-order-print-mode");
}

function printProductionOrderSheet() {
  if (activeOrderCreateKind !== "production" || !orderModal || orderModal.classList.contains("hidden")) {
    return;
  }

  closeOrderItemSuggestions();
  document.body.classList.add("production-order-print-mode");
  const handleAfterPrint = () => {
    finishProductionOrderPrintMode();
    window.removeEventListener("afterprint", handleAfterPrint);
  };
  window.addEventListener("afterprint", handleAfterPrint);
  window.print();
  window.setTimeout(finishProductionOrderPrintMode, 1000);
}

function setOrderCtaMenuOpen(isOpen) {
  orderCtaMenu?.classList.toggle("is-open", Boolean(isOpen));
  openOrderMenuButton?.setAttribute("aria-expanded", isOpen ? "true" : "false");
  if (isOpen) {
    positionOrderCtaDropdown();
  } else if (orderCtaDropdown) {
    orderCtaDropdown.style.top = "";
    orderCtaDropdown.style.left = "";
  }
}

function positionOrderCtaDropdown() {
  if (!openOrderMenuButton || !orderCtaDropdown) {
    return;
  }

  const rect = openOrderMenuButton.getBoundingClientRect();
  const gap = 12;
  const dropdownWidth = Math.max(220, orderCtaDropdown.offsetWidth || 0);
  const dropdownHeight = Math.max(rect.height, orderCtaDropdown.offsetHeight || 0);
  const maxLeft = Math.max(16, window.innerWidth - dropdownWidth - 16);
  const maxTop = Math.max(16, window.innerHeight - dropdownHeight - 16);
  const left = Math.min(rect.right + gap, maxLeft);
  const top = Math.min(Math.max(16, rect.top), maxTop);

  orderCtaDropdown.style.left = `${left}px`;
  orderCtaDropdown.style.top = `${top}px`;
}

function setOrdersBoardMenuOpen(isOpen) {
  ordersBoardMenu?.classList.toggle("is-open", Boolean(isOpen));
  openOrdersBoardMenuButton?.setAttribute("aria-expanded", isOpen ? "true" : "false");
  if (isOpen) {
    positionOrdersBoardDropdown();
  } else if (ordersBoardDropdown) {
    ordersBoardDropdown.style.top = "";
    ordersBoardDropdown.style.left = "";
  }
}

function positionOrdersBoardDropdown() {
  if (!openOrdersBoardMenuButton || !ordersBoardDropdown) {
    return;
  }

  const rect = openOrdersBoardMenuButton.getBoundingClientRect();
  const gap = 12;
  const dropdownWidth = Math.max(220, ordersBoardDropdown.offsetWidth || 0);
  const dropdownHeight = Math.max(rect.height, ordersBoardDropdown.offsetHeight || 0);
  const maxLeft = Math.max(16, window.innerWidth - dropdownWidth - 16);
  const maxTop = Math.max(16, window.innerHeight - dropdownHeight - 16);
  const left = Math.min(rect.right + gap, maxLeft);
  const top = Math.min(Math.max(16, rect.top), maxTop);

  ordersBoardDropdown.style.left = `${left}px`;
  ordersBoardDropdown.style.top = `${top}px`;
}

function enterOrderDraftEditMode(order) {
  activeOrderDraftEditId = String(order?.order_id || "").trim();
  activeOrderDraftEditRecord = order ? { ...order } : null;
  if (orderModalTitle) {
    orderModalTitle.textContent = "Sửa đơn hàng";
  }
  if (confirmOrderButton) {
    confirmOrderButton.textContent = "Xác Nhận Sửa Đơn";
  }
  if (orderIdInput) {
    orderIdInput.readOnly = true;
  }
}

function enterProductionOrderEditMode(order) {
  activeOrderDraftEditId = String(order?.order_id || "").trim();
  activeOrderDraftEditRecord = order ? { ...order } : null;
  if (orderModalTitle) {
    orderModalTitle.textContent = "Sửa phiếu sản xuất";
  }
  if (confirmOrderButton) {
    confirmOrderButton.textContent = "Xác Nhận Sửa Đơn";
  }
  if (productionOrderIdInput) {
    productionOrderIdInput.readOnly = true;
  }
}

async function openOrderModal(kind = "production") {
  clearPersistedOpenOrderPreview();
  if (!canCreateOrder()) {
    showToast("Chỉ kinh doanh, quản lý hoặc quản trị mới được tạo đơn hàng.", "error");
    return;
  }

  activeOrderCreateKind = kind === "transport" ? "transport" : "production";
  if (
    activeOrderCreateKind === "production" &&
    !canChooseProductionRequester() &&
    String(currentUser?.department || "").trim().toLowerCase() !== "sales"
  ) {
    showToast("Phiếu sản xuất mặc định lấy người yêu cầu là nhân viên kinh doanh. Chỉ admin mới được chọn người khác.", "error");
    return;
  }

  await loadOrderProducts();
  if (canViewOrders()) {
    await loadOrders();
    if (activeOrderCreateKind === "transport") {
      await loadSalesInventory();
    }
  }

  populateOrderSalesOptions();
  populateOrderDeliveryOptions();
  if (!orderSalesUser?.value) {
    showToast("Chưa có nhân viên kinh doanh phù hợp để gán đơn.", "error");
    return;
  }
  if (!orderDeliveryUser?.value) {
    showToast("Chưa có nhân viên giao hàng phù hợp để điều phối.", "error");
    return;
  }

  resetOrderFormMode();
  orderForm?.reset();
  populateOrderSalesOptions();
  populateProductionRequesterOptions();
  populateOrderDeliveryOptions();
  if (orderItemsList) {
    orderItemsList.innerHTML = "";
  }
  if (productionOrderItemsList) {
    productionOrderItemsList.innerHTML = "";
  }
  syncOrderDraftTotals();
  syncProductionOrderDraft();
  if (orderPlannedAt) {
    orderPlannedAt.value = "";
  }
  if (orderCreatedAt) {
    orderCreatedAt.value = formatNotificationTime(new Date().toISOString());
  }
  if (productionCreatedAtInput) {
    productionCreatedAtInput.value = formatNotificationTime(new Date().toISOString());
  }
  if (productionOrderIdInput) {
    productionOrderIdInput.value = activeOrderCreateKind === "production" ? getNextProductionOrderSequence() : "";
  }
  if (orderIdInput) {
    orderIdInput.value = activeOrderCreateKind === "transport" ? getNextTransportOrderSequence() : "";
  }
  if (productionCustomerNameInput) {
    productionCustomerNameInput.value = "";
  }
  if (productionNoteInput) {
    productionNoteInput.value = "";
  }
  if (productionTurnaroundOtherInput) {
    productionTurnaroundOtherInput.value = "";
  }
  const defaultTurnaroundRadio = orderForm?.querySelector('input[name="production-turnaround"][value="72H"]');
  if (defaultTurnaroundRadio) {
    defaultTurnaroundRadio.checked = true;
  }
  setOrderFormStep("1");
  orderModal?.classList.remove("hidden");
  orderBackdrop?.classList.remove("hidden");
  orderModal?.setAttribute("aria-hidden", "false");
  setOrderCtaMenuOpen(false);
  if (activeOrderCreateKind === "production") {
    productionCustomerNameInput?.focus();
  } else {
    orderIdInput?.focus();
  }
}

function openOrderDraftEditModal(order) {
  if (!canEditOrder(order) || !order) {
    return;
  }

  activeOrderCreateKind = "transport";
  populateOrderSalesOptions();
  populateOrderDeliveryOptions();
  resetOrderFormMode();
  orderForm?.reset();
  if (orderItemsList) {
    orderItemsList.innerHTML = "";
  }

  enterOrderDraftEditMode(order);
  populateTransportOrderForm(order);
  setOrderFormStep("1");
  orderModal?.classList.remove("hidden");
  orderBackdrop?.classList.remove("hidden");
  orderModal?.setAttribute("aria-hidden", "false");
  orderCustomerName?.focus();
}

async function openProductionOrderEditModal(order) {
  if (!canEditOrder(order) || !order) {
    if (isProductionOrderLockedForEditing(order)) {
      await openOrderPreviewModal(order);
    }
    return;
  }

  activeOrderCreateKind = "production";
  await loadOrderProducts();
  if (canViewOrders()) {
    await loadOrders();
  }

  resetOrderFormMode();
  orderForm?.reset();
  populateProductionRequesterOptions();
  if (productionOrderItemsList) {
    productionOrderItemsList.innerHTML = "";
  }

  enterProductionOrderEditMode(order);
  populateProductionOrderForm(order);
  setOrderFormStep("1");
  orderModal?.classList.remove("hidden");
  orderBackdrop?.classList.remove("hidden");
  orderModal?.setAttribute("aria-hidden", "false");
  productionCustomerNameInput?.focus();
}

async function openOrderPreviewModal(order) {
  if (!order) {
    return;
  }

  activeOrderDetails = order;
  activeOrderCreateKind = isProductionOrder(order) ? "production" : "transport";
  await loadOrderProducts();
  resetOrderFormMode();
  orderForm?.reset();
  populateOrderSalesOptions();
  populateProductionRequesterOptions();
  populateOrderDeliveryOptions();
  if (orderItemsList) {
    orderItemsList.innerHTML = "";
  }
  if (productionOrderItemsList) {
    productionOrderItemsList.innerHTML = "";
  }

  if (activeOrderCreateKind === "production") {
    populateProductionOrderForm(order);
    if (orderModalTitle) {
      orderModalTitle.textContent = "Xem phiếu sản xuất";
    }
  } else {
    populateTransportOrderForm(order);
    if (orderModalTitle) {
      orderModalTitle.textContent = "Xem đơn vận chuyển";
    }
  }

  setOrderFormStep("1");
  resetProductionClaimMode();
  setOrderFormReadOnly(true);
  orderModal?.classList.remove("hidden");
  orderBackdrop?.classList.remove("hidden");
  orderModal?.setAttribute("aria-hidden", "false");
  updateProductionClaimButtons(order);
  persistOpenOrderPreview(order);
}

function closeOrderModal() {
  clearPersistedOpenOrderPreview();
  closeProductionPackagingConfirmModal();
  orderModal?.classList.add("hidden");
  orderBackdrop?.classList.add("hidden");
  orderModal?.setAttribute("aria-hidden", "true");
  orderForm?.reset();
  if (orderItemsList) {
    orderItemsList.innerHTML = "";
  }
  syncOrderDraftTotals();
  setOrderFormStep("1");
  resetOrderFormMode();
}

function closePasswordModal() {
  passwordModal.classList.add("hidden");
  passwordBackdrop.classList.add("hidden");
  passwordModal.setAttribute("aria-hidden", "true");
  passwordForm.reset();
}

function getSortedPermissionUsers(users = allUsers) {
  return [...users].sort((left, right) => {
    return (
      labelDepartment(left.department).localeCompare(labelDepartment(right.department), "vi") ||
      String(left.name || "").localeCompare(String(right.name || ""), "vi")
    );
  });
}

function matchPermissionUserSearch(user, query = "") {
  const normalizedQuery = String(query || "").trim().toLowerCase();
  if (!normalizedQuery) {
    return true;
  }

  const departmentLabel = labelDepartment(user?.department).toLowerCase();
  const haystacks = [
    String(user?.name || "").toLowerCase(),
    String(user?.username || "").toLowerCase(),
    String(user?.employee_code || "").toLowerCase(),
    String(user?.title || "").toLowerCase(),
    departmentLabel,
  ];

  return haystacks.some((value) => value.includes(normalizedQuery));
}

function getVisiblePermissionUsers() {
  return getSortedPermissionUsers(allUsers).filter((user) => matchPermissionUserSearch(user, permissionsUserSearchValue));
}

function renderPermissionDirectory(users = []) {
  const card = document.createElement("section");
  card.className = "permission-card permission-directory-card";

  const cardHead = document.createElement("div");
  cardHead.className = "permission-card-head";
  cardHead.innerHTML = `
    <div class="permission-card-title">
      <p class="doc-title">Danh sách tài khoản</p>
      <small>${escapeHtml(users.length ? `Đang hiển thị ${users.length} tài khoản` : "Không có tài khoản phù hợp")}</small>
    </div>
  `;
  card.append(cardHead);

  const body = document.createElement("div");
  body.className = "permission-card-body";

  const tableWrap = document.createElement("div");
  tableWrap.className = "permission-user-table-wrap";

  if (!users.length) {
    tableWrap.innerHTML = `<div class="permission-empty">Không tìm thấy tài khoản nào khớp bộ lọc hiện tại.</div>`;
    body.append(tableWrap);
    card.append(body);
    return card;
  }

  const rows = users
    .map((user, index) => {
      const isActive = String(user?.id || "") === String(selectedPermissionUserId || "");
      return `
        <tr class="${isActive ? "is-active" : ""}" data-permission-user-row="${escapeHtml(user.id)}" tabindex="0">
          <td>${index + 1}</td>
          <td>${escapeHtml(user.name || "-")}</td>
          <td>${escapeHtml(user.employee_code || user.username || "-")}</td>
          <td>${escapeHtml(labelDepartment(user.department))}</td>
          <td>${escapeHtml(user.title || labelRole(user.role))}</td>
          <td><span class="role-mini-badge ${escapeHtml(roleBadgeClass(user.role))}">${escapeHtml(labelRole(user.role))}</span></td>
        </tr>
      `;
    })
    .join("");

  tableWrap.innerHTML = `
    <table class="permission-user-table">
      <thead>
        <tr>
          <th>STT</th>
          <th>Họ tên</th>
          <th>Mã NV</th>
          <th>Bộ phận</th>
          <th>Chức vụ</th>
          <th>Vai trò</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;

  body.append(tableWrap);
  card.append(body);
  return card;
}

function renderPermissionsPanel() {
  permissionsList.innerHTML = "";
  if (permissionsUserSearch) {
    permissionsUserSearch.value = permissionsUserSearchValue;
    permissionsUserSearch.placeholder = "Ten, tai khoan, ma NV, bo phan...";
    const searchLabel = permissionsUserSearch.closest(".permission-search")?.querySelector("span");
    if (searchLabel) {
      searchLabel.textContent = "Tim nhanh";
    }
  }

  const sortedUsers = getSortedPermissionUsers(allUsers);
  const visibleUsers = getVisiblePermissionUsers();
  let selectedUser =
    visibleUsers.find((user) => user.id === selectedPermissionUserId) ||
    sortedUsers.find((user) => user.id === selectedPermissionUserId) ||
    null;

  if (selectedUser) {
    selectedPermissionUserId = selectedUser.id;
  }

  permissionsList.append(renderPermissionDirectory(visibleUsers));

  if (String(currentUser?.role || "").trim().toLowerCase() === "admin" && isCreatingManagedUser) {
    permissionsList.append(renderCreateUserCard());
  }

  if (selectedUser) {
    permissionsList.append(renderPermissionCard(selectedUser));
    return;
  }

  if (!visibleUsers.length) {
    permissionsList.append(document.createRange().createContextualFragment(`<div class="permission-empty">Chưa có nhân viên nào trong danh sách.</div>`));
    return;
  }

  if (!isCreatingManagedUser) {
    permissionsList.append(
      document.createRange().createContextualFragment(
        `<div class="permission-empty">Bấm <strong>Thêm tài khoản</strong> để tạo mới, hoặc chọn một nhân viên trong bảng để xem và chỉnh sửa.</div>`,
      ),
    );
  }
}

function renderCreateUserCard() {
  const card = document.createElement("section");
  card.className = "permission-card";
  const defaultTitle = getSuggestedUserTitle("sales", "employee");
  card.innerHTML = `
    <div class="permission-card-head">
      <div class="permission-card-title">
        <p class="doc-title">Tạo tài khoản mới</p>
        <small>Tài khoản đăng nhập sẽ được lưu trực tiếp trong hệ thống nội bộ</small>
      </div>
    </div>
    <div class="permission-card-body">
      <div class="account-meta-grid">
        <label class="field">
          <span>Họ tên</span>
          <input type="text" id="create-user-name" placeholder="Nguyễn Văn A" />
        </label>
        <label class="field">
          <span>Tên đăng nhập</span>
          <input type="text" id="create-user-username" placeholder="nt001" />
        </label>
        <label class="field">
          <span>Mật khẩu</span>
          <input type="password" id="create-user-password" placeholder="Ít nhất 6 ký tự" />
        </label>
        <label class="field">
          <span>Mã NV</span>
          <input type="text" id="create-user-code" placeholder="NT001" />
        </label>
        <label class="field">
          <span>Bộ phận</span>
          <select id="create-user-department">
            <option value="sales">Kinh doanh</option>
            <option value="production">Sản xuất</option>
            <option value="operations">Vận chuyển</option>
            <option value="finance">Tài chính</option>
            <option value="hr">Nhân sự</option>
            <option value="executive">Ban giám đốc</option>
            <option value="general">Khác</option>
          </select>
        </label>
        <label class="field">
          <span>Vai trò</span>
          <select id="create-user-role">
            <option value="employee">Nhân viên</option>
            <option value="manager">Quản lý</option>
            <option value="director">Giám đốc</option>
            <option value="admin">Quản trị viên</option>
          </select>
        </label>
        <label class="field">
          <span>Email</span>
          <input type="email" id="create-user-email" placeholder="email@company.com" />
        </label>
        <label class="field">
          <span>Chức vụ</span>
          <select id="create-user-title">
            ${buildCreateUserTitleOptions("sales", "employee", defaultTitle)}
          </select>
        </label>
      </div>
    </div>
  `;

  const body = card.querySelector(".permission-card-body");
  if (body) {
    const actions = document.createElement("div");
    actions.className = "import-actions permission-card-actions";

    const cancelButton = document.createElement("button");
    cancelButton.type = "button";
    cancelButton.className = "ghost-button";
    cancelButton.textContent = "Dong";
    cancelButton.addEventListener("click", () => {
      isCreatingManagedUser = false;
      renderPermissionsPanel();
    });

    const submitButton = document.createElement("button");
    submitButton.type = "button";
    submitButton.className = "accent-button";
    submitButton.textContent = "Luu tai khoan moi";
    submitButton.addEventListener("click", createManagedUser);

    actions.append(cancelButton, submitButton);
    body.append(actions);
  }

  return card;
}

function getSuggestedUserTitle(department = "general", role = "employee") {
  const normalizedDepartment = String(department || "").trim().toLowerCase();
  const normalizedRole = String(role || "").trim().toLowerCase();

  if (normalizedRole === "admin") return "Quản trị hệ thống";
  if (normalizedRole === "director") return "Giám đốc";
  if (normalizedRole === "manager") {
    if (normalizedDepartment === "sales") return "Trưởng phòng kinh doanh";
    if (normalizedDepartment === "production") return "Quản lý sản xuất";
    if (normalizedDepartment === "operations") return "Quản lý vận hành";
    if (normalizedDepartment === "finance") return "Kế toán trưởng";
    if (normalizedDepartment === "hr") return "Trưởng phòng nhân sự";
    return "Quản lý";
  }

  if (normalizedDepartment === "sales") return "Nhân viên kinh doanh";
  if (normalizedDepartment === "production") return "Nhân viên sản xuất";
  if (normalizedDepartment === "operations") return "Nhân viên vận chuyển";
  if (normalizedDepartment === "finance") return "Nhân viên kế toán";
  if (normalizedDepartment === "hr") return "Nhân viên nhân sự";
  if (normalizedDepartment === "executive") return "Điều phối";
  return "Nhân viên";
}

function buildCreateUserTitleOptions(department = "general", role = "employee", selectedTitle = "") {
  const suggested = getSuggestedUserTitle(department, role);
  const titleOptions = [
    suggested,
    "Nhân viên kinh doanh",
    "Nhân viên sản xuất",
    "Nhân viên vận chuyển",
    "Nhân viên kế toán",
    "Nhân viên nhân sự",
    "Trưởng phòng kinh doanh",
    "Quản lý sản xuất",
    "Quản lý vận hành",
    "Kế toán trưởng",
    "Trưởng phòng nhân sự",
    "Giám đốc",
    "Quản trị hệ thống",
    "Điều phối",
    "Khác",
  ];

  const uniqueTitles = [...new Set(titleOptions.filter(Boolean))];
  return uniqueTitles
    .map((title) => `<option value="${escapeHtml(title)}" ${title === selectedTitle ? "selected" : ""}>${escapeHtml(title)}</option>`)
    .join("");
}

function handleCreateUserPresetChange(event) {
  const target = event.target;
  if (!target || (target.id !== "create-user-department" && target.id !== "create-user-role")) {
    return;
  }

  const department = document.querySelector("#create-user-department")?.value || "general";
  const role = document.querySelector("#create-user-role")?.value || "employee";
  const titleSelect = document.querySelector("#create-user-title");
  if (!titleSelect) {
    return;
  }

  const suggested = getSuggestedUserTitle(department, role);
  titleSelect.innerHTML = buildCreateUserTitleOptions(department, role, suggested);
  titleSelect.value = suggested;
}

function handleEditableUserProfilePresetChange(event) {
  const departmentSelect = event.target?.closest?.('[data-field="profile_department"]');
  if (!departmentSelect) {
    return;
  }

  const userId = departmentSelect.dataset.userId || "";
  const titleSelect = permissionsList?.querySelector(
    `[data-user-id="${userId}"][data-field="profile_title"]`,
  );
  if (!titleSelect) {
    return;
  }

  const department = departmentSelect.value || "general";
  const currentTitle = titleSelect.value || getSuggestedUserTitle(department, "employee");
  titleSelect.innerHTML = buildCreateUserTitleOptions(department, "employee", currentTitle);
  titleSelect.value = currentTitle;
}

function renderPermissionCard(user) {
  const card = document.createElement("section");
  card.className = "permission-card";
  card.dataset.userId = user.id;
  card.innerHTML = `
    <div class="permission-card-head">
      <div class="permission-card-title">
        <p class="doc-title">${escapeHtml(user.name)}</p>
        <small>${escapeHtml(`${user.username || "-"} • ${labelDepartment(user.department)}`)}</small>
      </div>
      <div class="permission-card-summary">
        <span class="role-mini-badge ${escapeHtml(roleBadgeClass(user.role))}">${escapeHtml(labelRole(user.role))}</span>
        <span class="mini-pill">${escapeHtml(labelPermissionAccessLevel(user.policy?.max_access_level || "basic"))}</span>
      </div>
    </div>
  `;

  const body = document.createElement("div");
  body.className = "permission-card-body";

  const accountMeta = document.createElement("div");
  accountMeta.className = "account-meta-grid editable-user-meta-grid";
  accountMeta.innerHTML = `
    <label class="field">
      <span>Họ tên</span>
      <input type="text" value="${escapeHtml(user.name || "")}" data-user-id="${escapeHtml(user.id)}" data-field="profile_name" />
    </label>
    <label class="field">
      <span>Tài khoản</span>
      <input type="text" value="${escapeHtml(user.username || "")}" data-user-id="${escapeHtml(user.id)}" data-field="profile_username" />
    </label>
    <label class="field">
      <span>Mã NV</span>
      <input type="text" value="${escapeHtml(user.employee_code || "")}" data-user-id="${escapeHtml(user.id)}" data-field="profile_employee_code" />
    </label>
    <label class="field">
      <span>Bộ phận</span>
      <select data-user-id="${escapeHtml(user.id)}" data-field="profile_department">
        <option value="sales" ${String(user.department || "") === "sales" ? "selected" : ""}>Kinh doanh</option>
        <option value="production" ${String(user.department || "") === "production" ? "selected" : ""}>Sản xuất</option>
        <option value="operations" ${String(user.department || "") === "operations" ? "selected" : ""}>Vận chuyển</option>
        <option value="finance" ${String(user.department || "") === "finance" ? "selected" : ""}>Tài chính</option>
        <option value="hr" ${String(user.department || "") === "hr" ? "selected" : ""}>Nhân sự</option>
        <option value="executive" ${String(user.department || "") === "executive" ? "selected" : ""}>Ban giám đốc</option>
        <option value="general" ${String(user.department || "") === "general" ? "selected" : ""}>Khác</option>
      </select>
    </label>
    <label class="field">
      <span>Chức vụ</span>
      <select data-user-id="${escapeHtml(user.id)}" data-field="profile_title">
        ${buildCreateUserTitleOptions(user.department || "general", "employee", user.title || getSuggestedUserTitle(user.department || "general", "employee"))}
      </select>
    </label>
    <label class="field">
      <span>Email</span>
      <input type="email" value="${escapeHtml(user.email || "")}" data-user-id="${escapeHtml(user.id)}" data-field="profile_email" />
    </label>
  `;
  body.append(accountMeta);

  if (String(currentUser?.role || "").trim().toLowerCase() === "admin") {
    const saveProfileButton = document.createElement("button");
    saveProfileButton.type = "button";
    saveProfileButton.className = "accent-button";
    saveProfileButton.textContent = "Lưu thông tin";
    saveProfileButton.addEventListener("click", () => saveUserProfile(user.id));
    body.append(saveProfileButton);
  }

  if (currentUser?.policy?.can_manage_permissions) {
    const accessField = document.createElement("label");
    accessField.className = "field";
    accessField.innerHTML = `
      <span>Mức truy cập</span>
      <select data-user-id="${escapeHtml(user.id)}" data-field="max_access_level">
        ${["basic", "advanced", "sensitive"]
          .map(
            (level) =>
              `<option value="${level}" ${user.policy?.max_access_level === level ? "selected" : ""}>${escapeHtml(labelPermissionAccessLevel(level))}</option>`,
          )
          .join("")}
      </select>
    `;
    body.append(accessField);

    const scopeField = document.createElement("label");
    scopeField.className = "field";
    scopeField.innerHTML = `
      <span>Phạm vi phòng ban</span>
      <select data-user-id="${escapeHtml(user.id)}" data-field="department_scope">
        <option value="own" ${user.policy?.department_scope !== "all" ? "selected" : ""}>${escapeHtml(labelDepartmentScope("own"))}</option>
        <option value="all" ${user.policy?.department_scope === "all" ? "selected" : ""}>${escapeHtml(labelDepartmentScope("all"))}</option>
      </select>
    `;
    body.append(scopeField);

    const departmentField = document.createElement("label");
    departmentField.className = "field";
    departmentField.innerHTML = `
      <span>Phòng ban bổ sung</span>
      <input
        type="text"
        value="${escapeHtml((user.policy?.allowed_departments || []).join(","))}"
        data-user-id="${escapeHtml(user.id)}"
        data-field="allowed_departments"
        placeholder="Ví dụ: finance,hr"
      />
    `;
    body.append(departmentField);

    const saveButton = document.createElement("button");
    saveButton.type = "button";
    saveButton.className = "accent-button";
    saveButton.textContent = "Lưu quyền";
    saveButton.addEventListener("click", () => saveUserPermissions(user.id));
    body.append(saveButton);

    const passwordField = document.createElement("label");
    passwordField.className = "field";
    passwordField.innerHTML = `
      <span>Đặt lại mật khẩu</span>
      <input
        type="password"
        data-user-id="${escapeHtml(user.id)}"
        data-field="reset_password"
        placeholder="Ít nhất 6 ký tự"
      />
    `;
    body.append(passwordField);

    const resetButton = document.createElement("button");
    resetButton.type = "button";
    resetButton.className = "ghost-button";
    resetButton.textContent = "Lưu mật khẩu mới";
    resetButton.addEventListener("click", () => resetUserPassword(user.id));
    body.append(resetButton);

    if (String(user.id || "") !== String(currentUser?.id || "")) {
      const deleteButton = document.createElement("button");
      deleteButton.type = "button";
      deleteButton.className = "ghost-button";
      deleteButton.textContent = "Xóa người dùng";
      deleteButton.addEventListener("click", () => deleteManagedUser(user.id, user.name));
      body.append(deleteButton);
    }
  }

  card.append(body);
  return card;
}

function roleBadgeClass(role) {
  const value = String(role || "").toLowerCase();
  if (value === "director") return "director";
  if (value === "admin") return "admin";
  if (value === "manager") return "manager";
  return "employee";
}

function selectPermissionUser(userId = "") {
  selectedPermissionUserId = String(userId || "").trim();
  isCreatingManagedUser = false;
  renderPermissionsPanel();
}

function handlePermissionDirectoryClick(event) {
  const row = event.target?.closest?.("[data-permission-user-row]");
  if (!row) {
    return;
  }

  selectPermissionUser(row.dataset.permissionUserRow || "");
}

function handlePermissionDirectoryKeydown(event) {
  const row = event.target?.closest?.("[data-permission-user-row]");
  if (!row) {
    return;
  }

  if (event.key !== "Enter" && event.key !== " ") {
    return;
  }

  event.preventDefault();
  selectPermissionUser(row.dataset.permissionUserRow || "");
}

async function saveUserProfile(userId) {
  if (String(currentUser?.role || "").trim().toLowerCase() !== "admin") {
    showToast("Chỉ admin mới được sửa thông tin người dùng.", "error");
    return;
  }

  const name = permissionsList.querySelector(`[data-user-id="${userId}"][data-field="profile_name"]`)?.value || "";
  const username = permissionsList.querySelector(`[data-user-id="${userId}"][data-field="profile_username"]`)?.value || "";
  const employeeCode =
    permissionsList.querySelector(`[data-user-id="${userId}"][data-field="profile_employee_code"]`)?.value || "";
  const department =
    permissionsList.querySelector(`[data-user-id="${userId}"][data-field="profile_department"]`)?.value || "general";
  const title = permissionsList.querySelector(`[data-user-id="${userId}"][data-field="profile_title"]`)?.value || "";
  const email = permissionsList.querySelector(`[data-user-id="${userId}"][data-field="profile_email"]`)?.value || "";

  try {
    const payload = await postJson(userProfileUrl, {
      user_id: userId,
      name,
      username,
      employee_code: employeeCode,
      department,
      title,
      email,
    });
    allUsers = payload.users || allUsers;
    currentUser = payload.currentUser || currentUser;
    selectedPermissionUserId = userId;
    renderCurrentUser();
    renderPermissionsPanel();
    showToast("Đã cập nhật thông tin người dùng.", "success");
  } catch (error) {
    showToast(error.message || "Không thể cập nhật thông tin người dùng.", "error");
  }
}

async function saveUserPermissions(userId) {
  const maxAccessLevel = permissionsList.querySelector(`[data-user-id="${userId}"][data-field="max_access_level"]`)?.value || "basic";
  const departmentScope = permissionsList.querySelector(`[data-user-id="${userId}"][data-field="department_scope"]`)?.value || "own";
  const allowedDepartments =
    permissionsList.querySelector(`[data-user-id="${userId}"][data-field="allowed_departments"]`)?.value || "";

  const payload = await postJson(userAccessUrl, {
    user_id: userId,
    max_access_level: maxAccessLevel,
    department_scope: departmentScope,
    allowed_departments: allowedDepartments,
  });

  allUsers = payload.users || [];
  currentUser = payload.currentUser || currentUser;
  renderCurrentUser();
  renderPermissionsPanel();
  showToast("Đã cập nhật quyền người dùng.", "success");
}

async function createManagedUser() {
  if (String(currentUser?.role || "").trim().toLowerCase() !== "admin") {
    showToast("Chỉ admin mới được tạo người dùng.", "error");
    return;
  }

  const name = document.querySelector("#create-user-name")?.value || "";
  const username = document.querySelector("#create-user-username")?.value || "";
  const password = document.querySelector("#create-user-password")?.value || "";
  const employeeCode = document.querySelector("#create-user-code")?.value || "";
  const department = document.querySelector("#create-user-department")?.value || "general";
  const role = document.querySelector("#create-user-role")?.value || "employee";
  const email = document.querySelector("#create-user-email")?.value || "";
  const title = document.querySelector("#create-user-title")?.value || "";

  try {
    const payload = await postJson(userCreateUrl, {
      name,
      username,
      password,
      employee_code: employeeCode,
      department,
      role,
      email,
      title,
    });
    allUsers = payload.users || allUsers;
    currentUser = payload.currentUser || currentUser;
    selectedPermissionUserId =
      String(payload?.users?.find?.((user) => String(user?.username || "").trim().toLowerCase() === String(username || "").trim().toLowerCase())?.id || "").trim() ||
      "";
    isCreatingManagedUser = false;
    renderCurrentUser();
    renderPermissionsPanel();
    showToast("Đã tạo người dùng mới.", "success");
  } catch (error) {
    showToast(error.message || "Không thể tạo người dùng.", "error");
  }
}

async function deleteManagedUser(userId, name = "") {
  if (String(currentUser?.role || "").trim().toLowerCase() !== "admin") {
    showToast("Chỉ admin mới được xóa người dùng.", "error");
    return;
  }
  if (!window.confirm(`Xóa người dùng "${name || userId}"?`)) {
    return;
  }

  try {
    const payload = await postJson(userDeleteUrl, { user_id: userId });
    allUsers = payload.users || allUsers;
    currentUser = payload.currentUser || currentUser;
    if (selectedPermissionUserId === userId) {
      selectedPermissionUserId = "";
    }
    if (isCreatingManagedUser) {
      isCreatingManagedUser = false;
    }
    renderCurrentUser();
    renderPermissionsPanel();
    showToast("Đã xóa người dùng.", "success");
  } catch (error) {
    showToast(error.message || "Không thể xóa người dùng.", "error");
  }
}

async function handlePasswordSubmit(event) {
  event.preventDefault();

  if (!newPasswordInput.value || newPasswordInput.value.length < 6) {
    showToast("Mật khẩu mới phải có ít nhất 6 ký tự.", "error");
    return;
  }

  if (newPasswordInput.value !== confirmPasswordInput.value) {
    showToast("Mật khẩu xác nhận không khớp.", "error");
    return;
  }

  confirmPasswordButton.disabled = true;
  try {
    const payload = await postJson(userPasswordUrl, {
      current_password: currentPasswordInput.value,
      new_password: newPasswordInput.value,
    });
    allUsers = payload.users || allUsers;
    currentUser = payload.currentUser || currentUser;
    renderCurrentUser();
    renderPermissionsPanel();
    closePasswordModal();
    showToast("Đã cập nhật mật khẩu.", "success");
  } catch (error) {
    showToast(error.message || "Không thể đổi mật khẩu.", "error");
  } finally {
    confirmPasswordButton.disabled = false;
  }
}

async function submitDeliveryCompletion(event) {
  event.preventDefault();

  if (!canCompleteDelivery()) {
    showToast("Bạn không có quyền xác nhận giao hàng.", "error");
    return;
  }

  confirmDeliveryButton.disabled = true;
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
    applyWorkspacePayload(payload);
    closeDeliveryModal();
    const deliveredTo = payload.salesUser?.name || "nhân viên kinh doanh";
    showToast(`Đã cập nhật đơn hàng và thông báo cho ${deliveredTo}.`, "success");
  } catch (error) {
    showToast(error.message || "Không thể xác nhận giao hàng.", "error");
  } finally {
    confirmDeliveryButton.disabled = false;
  }
}

async function submitOrderCreate(event) {
  event.preventDefault();

  if (!canCreateOrder()) {
    showToast("Bạn không có quyền tạo đơn hàng.", "error");
    return;
  }

  syncOrderDraftTotals();
  if (!validateOrderFormBeforeSubmit()) {
    return;
  }

  confirmOrderButton.disabled = true;
  try {
    const isEditingDraftOrder = Boolean(activeOrderDraftEditId);
    if (
      isEditingDraftOrder &&
      activeOrderCreateKind === "production" &&
      isProductionOrderLockedForEditing(activeOrderDraftEditRecord)
    ) {
      throw new Error("Phiếu sản xuất đã có người nhận hàng ký xác nhận, không thể sửa nữa.");
    }
    const productionTurnaroundValue = orderForm?.querySelector('input[name="production-turnaround"]:checked')?.value || "";
    const productionTurnaround =
      productionTurnaroundValue === "Khác"
        ? `Khác: ${String(productionTurnaroundOtherInput?.value || "").trim()}`
        : productionTurnaroundValue;
    const payloadBody = {
      order_id:
        activeOrderCreateKind === "production"
          ? normalizeOrderIdValue(productionOrderIdInput?.value || "")
          : activeOrderDraftEditId || normalizeOrderIdValue(orderIdInput?.value || ""),
      customer_name:
        activeOrderCreateKind === "production"
          ? productionCustomerNameInput?.value || ""
          : orderCustomerName?.value || "",
      sales_user_id:
        activeOrderCreateKind === "production"
          ? productionRequesterSelect?.value || ""
          : orderSalesUser?.value || "",
      delivery_user_id: orderDeliveryUser?.value || "",
      planned_delivery_at: activeOrderCreateKind === "production" ? "" : normalizePlannedDeliveryInput(orderPlannedAt?.value || ""),
      delivery_address: activeOrderCreateKind === "production" ? "" : orderAddress?.value || "",
      order_items: activeOrderCreateKind === "production" ? buildProductionOrderItemsSummary() : orderItems?.value || "",
      order_value: activeOrderCreateKind === "production" ? "" : orderValue?.value || "",
      note:
        activeOrderCreateKind === "production"
          ? [String(productionNoteInput?.value || "").trim(), productionTurnaround ? `Đề nghị trả hàng trong: ${productionTurnaround}` : ""]
              .filter(Boolean)
              .join("\n")
          : orderNote?.value || "",
      order_kind: activeOrderCreateKind,
      payment_status: activeOrderDraftEditRecord?.payment_status || "unpaid",
      payment_method: activeOrderDraftEditRecord?.payment_method || "",
    };
    const payload = await postJson(isEditingDraftOrder ? orderUpdateUrl : orderCreateUrl, payloadBody);
    applyWorkspacePayload(payload);
    closeOrderModal();
    if (isEditingDraftOrder) {
      showToast("Đã cập nhật đơn hàng chưa giao.", "success");
    } else {
      if (activeOrderCreateKind === "production") {
        const recipientCount = Number(payload.notification?.recipients || payload.productionRecipients?.length || 0);
        showToast(`Đã tạo phiếu sản xuất và gửi thông báo cho ${recipientCount} nhân viên sản xuất.`, "success");
      } else {
        const assignedTo = payload.deliveryUser?.name || "nhân viên giao hàng";
        showToast(`Đã tạo đơn và giao việc cho ${assignedTo}.`, "success");
      }
    }
  } catch (error) {
    showToast(error.message || "Không thể tạo đơn hàng.", "error");
  } finally {
    confirmOrderButton.disabled = false;
  }
}

function openOrderEditModal(order) {
  if (!canEditOrder(order) || !order) {
    if (isProductionOrderLockedForEditing(order)) {
      openOrderPreviewModal(order)
        .then(() => {
          closeOrdersBoardModal();
        })
        .catch((error) => {
          showToast(error.message || "Không thể mở phiếu sản xuất.", "error");
        });
    }
    return;
  }

  if (isProductionOrder(order)) {
    openProductionOrderEditModal(order).catch((error) => {
      showToast(error.message || "Không thể mở phiếu sản xuất để sửa.", "error");
    });
    return;
  }

  if (String(order.status || "").trim().toLowerCase() !== "completed") {
    openOrderDraftEditModal(order);
    return;
  }

  activeEditableOrderId = String(order.order_id || "").trim();
  activeEditableOrderRecord = order ? { ...order } : null;
  if (orderEditId) orderEditId.value = activeEditableOrderId;
  if (orderEditCustomerName) orderEditCustomerName.value = order.customer_name || "";
  if (orderEditAddress) orderEditAddress.value = order.delivery_address || "";
  if (orderEditItems) orderEditItems.value = order.order_items || "";
  if (orderEditValue) orderEditValue.value = order.order_value || "";
  if (orderEditPaymentStatus) orderEditPaymentStatus.value = order.payment_status || "unpaid";
  if (orderEditPaymentMethod) orderEditPaymentMethod.value = order.payment_method || "";
  if (orderEditNote) orderEditNote.value = order.completion_note || order.note || "";
  orderEditModal?.classList.remove("hidden");
  orderEditBackdrop?.classList.remove("hidden");
  orderEditModal?.setAttribute("aria-hidden", "false");
  orderEditCustomerName?.focus();
}

function closeOrderEditModal() {
  activeEditableOrderId = "";
  activeEditableOrderRecord = null;
  orderEditModal?.classList.add("hidden");
  orderEditBackdrop?.classList.add("hidden");
  orderEditModal?.setAttribute("aria-hidden", "true");
  orderEditForm?.reset();
}

async function submitOrderUpdate(event) {
  event.preventDefault();

  if (!canEditOrder(activeEditableOrderRecord) || !activeEditableOrderId) {
    showToast("Bạn không có quyền sửa đơn hàng.", "error");
    return;
  }

  confirmOrderEditButton.disabled = true;
  try {
    const payload = await postJson(orderUpdateUrl, {
      order_id: activeEditableOrderId,
      customer_name: orderEditCustomerName?.value || "",
      delivery_address: orderEditAddress?.value || "",
      order_items: orderEditItems?.value || "",
      order_value: orderEditValue?.value || "",
      payment_status: orderEditPaymentStatus?.value || "unpaid",
      payment_method: orderEditPaymentMethod?.value || "",
      note: orderEditNote?.value || "",
    });
    applyWorkspacePayload(payload);
    closeOrderEditModal();
    showToast("Đã cập nhật đơn hàng.", "success");
  } catch (error) {
    showToast(error.message || "Không thể cập nhật đơn hàng.", "error");
  } finally {
    confirmOrderEditButton.disabled = false;
  }
}

async function deleteOrder(orderOrId) {
  const targetOrderId =
    typeof orderOrId === "string"
      ? String(orderOrId || "").trim()
      : String(orderOrId?.order_id || "").trim();

  if (!canManageOrders()) {
    showToast("Bạn không có quyền xóa đơn hàng.", "error");
    return;
  }

  if (!targetOrderId) {
    showToast("Không xác định được mã đơn hàng để xóa.", "error");
    return;
  }

  if (!window.confirm(`Xóa đơn hàng "${targetOrderId}"?`)) {
    return;
  }

  try {
    const payload = await postJson(orderDeleteUrl, {
      order_id: targetOrderId,
    });
    applyWorkspacePayload(payload);
    closeOrderDetailsModal();
    closeOrderEditModal();
    showToast("Đã xóa đơn hàng.", "success");
  } catch (error) {
    showToast(error.message || "Không thể xóa đơn hàng.", "error");
  }
}

async function resetUserPassword(userId) {
  const input = permissionsList.querySelector(`[data-user-id="${userId}"][data-field="reset_password"]`);
  const newPassword = input?.value || "";

  if (!newPassword || newPassword.length < 6) {
    showToast("Mật khẩu mới phải có ít nhất 6 ký tự.", "error");
    return;
  }

  try {
    const payload = await postJson(userPasswordUrl, {
      user_id: userId,
      new_password: newPassword,
    });
    allUsers = payload.users || allUsers;
    currentUser = payload.currentUser || currentUser;
    renderCurrentUser();
    renderPermissionsPanel();
    showToast("Đã đặt lại mật khẩu người dùng.", "success");
  } catch (error) {
    showToast(error.message || "Không thể đặt lại mật khẩu.", "error");
  }
}

async function syncSheetDocumentsNow() {
  syncSheetButton.disabled = true;
  syncSheetButton.classList.add("is-syncing");

  try {
    const payload = await postJson(syncDocsUrl, {});
    renderSources(payload.sources || []);
    renderDocuments(payload.documents || [], payload.folders || []);
    await loadHealth();

    const sync = payload.sync || {};
    if (Array.isArray(sync.errors) && sync.errors.length > 0) {
      const firstError = sync.errors[0]?.message || "Không thể đồng bộ dữ liệu từ Google Sheet.";
      showToast(firstError, "error");
    } else if (sync.updated > 0) {
      showToast(`Đã đồng bộ ${sync.updated} nguồn từ Google Sheet.`, "success");
    } else if (sync.checked > 0) {
      showToast("Tất cả nguồn đồng bộ đã là bản mới nhất.", "success");
    } else {
      showToast("Không có nguồn đồng bộ nào để cập nhật.", "success");
    }
  } catch (error) {
    showToast(error.message || "Không thể đồng bộ dữ liệu từ Google Sheet.", "error");
  } finally {
    syncSheetButton.disabled = false;
    syncSheetButton.classList.remove("is-syncing");
  }
}

async function loadSources() {
  try {
    const payload = await fetchJson(sourcesUrl);
    sourceLibrary = payload.sources || [];
    updateDocumentFilterOptions(documentLibrary, [], sourceLibrary);
    renderSources(payload.sources || []);
  } catch {
    sourceLibrary = [];
    updateDocumentFilterOptions(documentLibrary, [], sourceLibrary);
    renderSources([]);
  }
}

async function loadDocuments() {
  try {
    const payload = await fetchJson(docsUrl);
    documentLibrary = payload.documents || [];
    updateDocumentFilterOptions(payload.documents || [], payload.folders || [], sourceLibrary);
    renderDocuments(payload.documents || [], payload.folders || []);
  } catch {
    documentLibrary = [];
    updateDocumentFilterOptions([], [], sourceLibrary);
    renderDocuments([], []);
  }
}

function updateDocumentFilterOptions(documents, folders = [], sources = []) {
  const folderMap = buildFolderOptionsByAccess(documents, folders, sources);
  applyAccessFilterPolicy();
  syncFolderFilterOptions(folderMap);
}

function buildFolderOptionsByAccess(documents, folders = [], sources = []) {
  const folderMap = {
    all: new Set(),
    basic: new Set(),
    advanced: new Set(),
    sensitive: new Set(),
  };

  for (const section of DOCUMENT_SECTION_TEMPLATE) {
    for (const folderName of section.folders) {
      folderMap.all.add(folderName);
      if (folderMap[section.level]) {
        folderMap[section.level].add(folderName);
      }
    }
  }

  for (const folder of folders) {
    const level = String(folder?.level || "").trim().toLowerCase();
    const name = String(folder?.name || "").trim();
    if (!name) {
      continue;
    }
    folderMap.all.add(name);
    if (folderMap[level]) {
      folderMap[level].add(name);
    }
  }

  for (const document of documents) {
    const level = String(document?.metadata?.access_level || "").trim().toLowerCase() || "basic";
    const name = resolveDocumentFolderName(document);
    if (!name) {
      continue;
    }
    folderMap.all.add(name);
    if (folderMap[level]) {
      folderMap[level].add(name);
    }
  }

  for (const source of sources) {
    const level = String(source?.access_level || "").trim().toLowerCase() || "basic";
    const name = resolveSourceFolderName(source);
    if (!name) {
      continue;
    }
    folderMap.all.add(name);
    if (folderMap[level]) {
      folderMap[level].add(name);
    }
  }

  return folderMap;
}

function syncFolderFilterOptions(folderMap = buildFolderOptionsByAccess(documentLibrary, [], sourceLibrary)) {
  updateSingleFolderFilter(libraryFolderFilter, libraryAccessFilter?.value, folderMap);
  updateSingleFolderFilter(askFolderFilter, askAccessFilter?.value, folderMap);
}

function applyAccessFilterPolicy() {
  const maxLevel = getCurrentMaxAccessLevel();
  applySingleAccessFilterPolicy(askAccessFilter, maxLevel);
  applySingleAccessFilterPolicy(libraryAccessFilter, maxLevel);
  renderAccessFilterHints(maxLevel);
}

function applySingleAccessFilterPolicy(selectNode, maxLevel) {
  if (!selectNode) {
    return;
  }

  const allowedLevels = getAllowedAccessLevels(maxLevel);
  const allowAllOption = allowedLevels.length > 1;

  for (const option of selectNode.options) {
    const value = String(option.value || "").trim().toLowerCase();
    if (!value) {
      option.disabled = !allowAllOption;
      option.hidden = !allowAllOption;
      continue;
    }
    const allowed = allowedLevels.includes(value);
    option.disabled = !allowed;
    option.hidden = !allowed;
  }

  const currentValue = String(selectNode.value || "").trim().toLowerCase();
  if (currentValue && allowedLevels.includes(currentValue)) {
    return;
  }
  selectNode.value = allowAllOption ? "" : allowedLevels[0] || "basic";
}

function renderAccessFilterHints(maxLevel = getCurrentMaxAccessLevel()) {
  const maxLabel = labelAccessLevel(maxLevel);
  const basicOnly = maxLevel === "basic";
  const personalSuffix = currentUser ? " Bạn vẫn có thể dùng mục Cá nhân cho tài liệu riêng của mình." : "";
  if (askFilterHint) {
    askFilterHint.textContent = basicOnly
      ? `Tài khoản hiện tại chỉ dùng mức ${maxLabel} cho dữ liệu dùng chung.${personalSuffix}`
      : `Chọn mức trước để chỉ tìm trong đúng nhóm tài liệu. Tài khoản hiện tại xem tối đa mức ${maxLabel}.`;
  }
  if (libraryFilterHint) {
    libraryFilterHint.textContent = basicOnly
      ? `Tài khoản hiện tại chỉ xem được tài liệu mức ${maxLabel} trong kho chung.${personalSuffix}`
      : `Chọn mức truy cập để thư mục và kết quả chỉ hiển thị trong đúng phạm vi được phép dùng. Tài khoản hiện tại xem tối đa mức ${maxLabel}.`;
  }
}

function getCurrentMaxAccessLevel() {
  const current = String(currentUser?.policy?.max_access_level || "basic").trim().toLowerCase();
  return ACCESS_LEVELS.includes(current) ? current : "basic";
}

function getAllowedAccessLevels(maxLevel = getCurrentMaxAccessLevel()) {
  const index = ACCESS_LEVELS.indexOf(maxLevel);
  const levels = index >= 0 ? ACCESS_LEVELS.slice(0, index + 1) : ["basic"];
  if (currentUser && !levels.includes("sensitive")) {
    levels.push("sensitive");
  }
  return levels;
}

function updateSingleFolderFilter(selectNode, accessLevel, folderMap) {
  if (!selectNode) {
    return;
  }

  const normalizedLevel = String(accessLevel || "").trim().toLowerCase();
  const allowed = normalizedLevel && folderMap[normalizedLevel]
    ? folderMap[normalizedLevel]
    : collectAccessibleFolders(folderMap);
  const options = [...allowed].sort((left, right) => left.localeCompare(right, "vi"));
  const previous = selectNode.value;
  selectNode.innerHTML = [`<option value="">Trong quyền của tôi</option>`]
    .concat(options.map((folder) => `<option value="${escapeHtml(folder)}">${escapeHtml(folder)}</option>`))
    .join("");

  if (options.includes(previous)) {
    selectNode.value = previous;
  } else {
    selectNode.value = "";
  }
}

function collectAccessibleFolders(folderMap) {
  const allowed = new Set();
  for (const level of getAllowedAccessLevels()) {
    for (const folder of folderMap[level] || []) {
      allowed.add(folder);
    }
  }
  return allowed;
}

function renderSources(sources) {
  sourceCountPill.textContent = String(sources.length);
  sourceList.innerHTML = "";
  const sections = groupSourcesByAccessLevel(sources);

  for (const section of sections) {
    const sectionNode = document.createElement("section");
    sectionNode.className = "doc-access-section";
    const canCreateSource = canCreateManagedContent() && ["basic", "advanced"].includes(String(section.level || "").toLowerCase());
    sectionNode.innerHTML = `
      <div class="doc-access-header">
        <span class="doc-access-title-wrap">
          <span class="doc-access-title">${escapeHtml(labelSectionAccessLevel(section.level))}</span>
          ${
            canCreateSource
              ? `<button type="button" class="folder-create-button source-create-button" data-level="${escapeHtml(section.level)}" aria-label="Tạo nguồn đồng bộ"></button>`
              : ""
          }
        </span>
        <span class="doc-access-count">${section.items.length}</span>
      </div>
    `;

    if (canCreateSource) {
      sectionNode.querySelector(".source-create-button")?.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();
        openImportModalWithPreset({
          mode: "sheet",
          level: section.level,
        });
      });
    }

    const list = document.createElement("div");
    list.className = "doc-folder-list";

    for (const source of section.items) {
      const item = document.createElement("details");
      item.className = "doc-folder compact-folder source-folder";
      const syncStatus = buildSourceSyncStatus(source);
      const sourceBadge = buildSourceStatusBadge(source);
      const sourceKindBadge = buildSourceKindBadge(source);
      item.innerHTML = `
        <summary class="doc-folder-summary">
          <span class="doc-folder-name source-folder-name">${escapeHtml(source.title || "Nguồn đồng bộ")}</span>
          <span class="doc-folder-summary-right">
            ${sourceKindBadge}
            ${sourceBadge}
            <span class="doc-folder-count">1</span>
          </span>
        </summary>
        <div class="doc-folder-list">
          <article class="doc-leaf source-leaf">
            <div class="doc-leaf-head">
              <span class="doc-file-icon icon-source" aria-hidden="true">${sourceLinkIconSvg()}</span>
              <div class="doc-leaf-meta">
                <div class="doc-title-row">
                  <p class="doc-title">${escapeHtml(source.title || "Nguồn đồng bộ")}</p>
                  ${sourceKindBadge}
                  ${sourceBadge}
                </div>
                <p class="doc-leaf-file">${escapeHtml(source.source_url || "")}</p>
                ${syncStatus ? `<p class="doc-sync-status">${escapeHtml(syncStatus)}</p>` : ""}
              </div>
            </div>
            <p class="doc-meta">Chủ đề: ${escapeHtml(source.topic_key || "")}</p>
          </article>
        </div>
      `;
      list.append(item);
    }

    sectionNode.append(list);
    sourceList.append(sectionNode);
  }
}

function renderDocuments(documents, folders = []) {
  documents = documents.filter(
    (document) => String(document?.metadata?.source_type || "").toLowerCase() !== "sheet",
  );
  updateDynamicFolderOptions(documents, folders);
  docCountPill.textContent = String(documents.length);
  docList.innerHTML = "";

  const sections = groupDocumentsByAccessLevel(documents, folders);

  for (const section of sections) {
    const sectionNode = document.createElement("section");
    sectionNode.className = "doc-access-section";
    const canCreateFolder = canCreateFolderInSection(section.level);
    sectionNode.innerHTML = `
      <div class="doc-access-header">
        <span class="doc-access-title-wrap">
          <span class="doc-access-title">${escapeHtml(labelSectionAccessLevel(section.level))}</span>
          ${
            canCreateFolder
              ? `<button type="button" class="folder-create-button" data-level="${escapeHtml(section.level)}" aria-label="Tạo thư mục"></button>`
              : ""
          }
        </span>
        <span class="doc-access-count">${section.count}</span>
      </div>
    `;

    if (canCreateFolder) {
      sectionNode.querySelector(".folder-create-button")?.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();
        createDocumentFolder(section.level).catch((error) => {
          showToast(error.message || "Khong the tao thu muc.", "error");
        });
      });
    }

    for (const group of section.groups) {
      const folder = document.createElement("details");
      folder.className = "doc-folder compact-folder";

      const summary = document.createElement("summary");
      summary.className = "doc-folder-summary";
      const canManageFolder = canManageCustomFolder(section.level, group.folder);
      summary.innerHTML = `
        <span class="doc-folder-name">${escapeHtml(group.folder)}</span>
        <span class="doc-folder-summary-right">
          ${
            canManageFolder
              ? `
                <span class="doc-folder-actions">
                  <button type="button" class="folder-action-button" data-action="rename">Sua</button>
                  <button type="button" class="folder-action-button danger" data-action="delete">Xoa</button>
                </span>
              `
              : ""
          }
          <span class="doc-folder-count">${group.items.length}</span>
        </span>
      `;
      folder.append(summary);

      if (canManageFolder) {
        const renameButton = summary.querySelector('[data-action="rename"]');
        const deleteButton = summary.querySelector('[data-action="delete"]');
        renameButton?.addEventListener("click", (event) => {
          event.preventDefault();
          event.stopPropagation();
          renameDocumentFolder(group.folder).catch((error) => {
            showToast(error.message || "Khong the doi ten thu muc.", "error");
          });
        });
        deleteButton?.addEventListener("click", (event) => {
          event.preventDefault();
          event.stopPropagation();
          deleteDocumentFolder(group.folder).catch((error) => {
            showToast(error.message || "Khong the xoa thu muc.", "error");
          });
        });
      }

      const list = document.createElement("div");
      list.className = "doc-folder-list";

      for (const documentItem of group.items) {
        const item = document.createElement("article");
        item.className = "doc-leaf";
        const canDelete = canDeleteDocument(documentItem);
        item.innerHTML = `
          <div class="doc-leaf-head">
            <span class="doc-file-icon ${fileIconClass(getFileExtension(documentItem.fileName))}" aria-hidden="true">${fileIconSvg()}</span>
            <div class="doc-leaf-meta">
              <p class="doc-title">${escapeHtml(documentItem.title)}</p>
            </div>
            <div class="doc-leaf-actions">
              <button type="button" class="leaf-action-button" data-action="view">Xem</button>
              <button type="button" class="leaf-action-button" data-action="download">Tải</button>
              ${
                canDelete
                  ? `<button type="button" class="leaf-action-button danger" data-action="delete">Xóa</button>`
                  : ""
              }
            </div>
          </div>
        `;
        item.querySelector('[data-action="view"]')?.addEventListener("click", () => {
          viewDocument(documentItem.path).catch((error) => {
            showToast(error.message || "Khong the mo tai lieu.", "error");
          });
        });
        item.querySelector('[data-action="download"]')?.addEventListener("click", () => {
          downloadDocument(documentItem.path);
        });
        item.querySelector('[data-action="delete"]')?.addEventListener("click", () => {
          deleteDocument(documentItem.path, documentItem.title).catch((error) => {
            showToast(error.message || "Khong the xoa tai lieu.", "error");
          });
        });
        list.append(item);
      }

      folder.append(list);
      sectionNode.append(folder);
    }

    docList.append(sectionNode);
  }
}

function appendUserMessage(question, image) {
  emptyState.classList.add("hidden");
  const article = document.createElement("article");
  article.className = "message-row user";

  const bubble = document.createElement("div");
  bubble.className = "message-bubble user-bubble";

  if (question) {
    const textNode = document.createElement("p");
    textNode.textContent = question;
    bubble.append(textNode);
  }

  if (image?.previewUrl) {
    const preview = document.createElement("img");
    preview.className = "user-image-preview";
    preview.src = image.previewUrl;
    preview.alt = "Ảnh đã gửi";
    bubble.append(preview);
  }

  article.append(bubble);
  chatThread.append(article);
  scrollToLatestMessage(article);
}

function appendAssistantLoadingMessage() {
  emptyState.classList.add("hidden");
  const article = document.createElement("article");
  article.className = "message-row assistant";
  article.innerHTML = `
    <div class="assistant-avatar">AI</div>
    <div class="message-card loading-card">
      <p class="message-label">Đang xử lý</p>
      <div class="typing-dots">
        <span></span>
        <span></span>
        <span></span>
      </div>
    </div>
  `;
  chatThread.append(article);
  scrollToLatestMessage(article);
  return article;
}

function hydrateAssistantMessage(container, result) {
  const route = routeName(result.route);
  container.innerHTML = "";
  container.className = "message-row assistant";

  const avatar = document.createElement("div");
  avatar.className = "assistant-avatar";
  avatar.textContent = route.short;
  container.append(avatar);

  const card = document.createElement("div");
  card.className = "message-card";

  const header = document.createElement("div");
  header.className = "message-topline";
  const label = document.createElement("p");
  label.className = "message-label";
  label.textContent = route.title;
  header.append(label, createMiniPill(route.pill, route.className));
  card.append(header);

  const answerBlock = document.createElement("div");
  answerBlock.className = "answer-block";
  for (const paragraph of toParagraphs(result.answer)) {
    const node = document.createElement("p");
    node.textContent = paragraph;
    answerBlock.append(node);
  }
  card.append(answerBlock);

  const sourcesWrap = document.createElement("div");
  sourcesWrap.className = "meta-section";
  sourcesWrap.append(createSectionTitle("Nguồn tham khảo"));

  const sourceGrid = document.createElement("div");
  sourceGrid.className = "source-grid";

  if (!result.sources?.length) {
    const empty = document.createElement("article");
    empty.className = "source-card";
    empty.innerHTML = `
      <p class="source-title">Chưa có nguồn hiển thị</p>
      <p class="source-snippet">Nguồn sẽ xuất hiện tại đây khi hệ thống tìm thấy tài liệu hoặc có dữ liệu web phù hợp.</p>
    `;
    sourceGrid.append(empty);
  } else {
    for (const source of result.sources) {
      const cardNode = document.createElement("article");
      cardNode.className = "source-card";

      const title = document.createElement("p");
      title.className = "source-title";
      title.textContent = source.title || "Nguồn";
      cardNode.append(title);

      const metaRow = document.createElement("div");
      metaRow.className = "source-meta-row";
      let hasMeta = false;

      if (source.url) {
        const link = document.createElement("a");
        link.href = source.url;
        link.target = "_blank";
        link.rel = "noreferrer";
        link.textContent = source.url;
        metaRow.append(link);
        hasMeta = true;
      } else if (source.path) {
        const pathText = document.createElement("p");
        pathText.className = "doc-meta";
        pathText.textContent = source.path;
        metaRow.append(pathText);
        hasMeta = true;
      }

      if (typeof source.score === "number") {
        const score = document.createElement("p");
        score.className = "doc-meta";
        score.textContent = `Điểm ${source.score.toFixed(2)}`;
        metaRow.append(score);
        hasMeta = true;
      }

      if (source.effectiveDate) {
        const effectiveDate = document.createElement("p");
        effectiveDate.className = "doc-meta";
        effectiveDate.textContent = `Hiệu lực ${source.effectiveDate}`;
        metaRow.append(effectiveDate);
        hasMeta = true;
      }

      if (source.owner) {
        const owner = document.createElement("p");
        owner.className = "doc-meta";
        owner.textContent = `Bộ phận ${labelDepartment(source.owner)}`;
        metaRow.append(owner);
        hasMeta = true;
      }

      if (hasMeta) {
        cardNode.append(metaRow);
      }

      if (source.snippet) {
        const snippet = document.createElement("p");
        snippet.className = "source-snippet";
        snippet.textContent = clampText(normalizeDisplayText(source.snippet), 140);
        cardNode.append(snippet);
      }

      sourceGrid.append(cardNode);
    }
  }

  sourcesWrap.append(sourceGrid);
  card.append(sourcesWrap);
  container.append(card);
  scrollToLatestMessage(container);
}

function renderHistory() {
  historyList.innerHTML = "";
  historyCount.textContent = String(conversationHistory.length);

  if (conversationHistory.length === 0) {
    const empty = document.createElement("div");
    empty.className = "history-item muted";
    empty.textContent = "Chưa có đoạn chat nào";
    historyList.append(empty);
    return;
  }

  for (const item of conversationHistory) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "history-item";
    if (item.question === selectedHistoryQuestion) {
      button.classList.add("is-active");
    }
    button.innerHTML = `
      <span>${escapeHtml(item.question)}</span>
      <small>${escapeHtml(routeName(item.route).pill)}</small>
    `;
    button.addEventListener("click", () => {
      selectedHistoryQuestion = item.question;
      questionInput.value = item.question;
      autoResizeTextarea();
      questionInput.focus();
      renderHistory();
    });
    historyList.append(button);
  }
}

async function openLibraryModal(preset = "general") {
  activeLibraryPreset = preset;
  libraryModal?.classList.remove("hidden");
  libraryBackdrop?.classList.remove("hidden");
  libraryModal?.setAttribute("aria-hidden", "false");
  setRepoTabActive(
    activeLibraryPreset === "production"
      ? "production-library"
      : activeLibraryPreset === "sales"
        ? "sales-library"
        : "library",
  );

  try {
    if (!documentLibrary.length) {
      await loadDocuments().catch(() => {});
    }
    if ((activeLibraryPreset === "production" || activeLibraryPreset === "sales") && canViewOrders()) {
      await loadOrders();
    }
    if (activeLibraryPreset === "sales") {
      await loadSalesInventory();
    }
    applyLibraryPreset();
    renderDocumentLibrary();
  } catch (error) {
    showToast(error.message || "Khong the mo kho tai lieu.", "error");
  }

  if (activeLibraryPreset === "production") {
    productionStockViewMode?.focus();
  } else if (activeLibraryPreset === "sales") {
    (canManageSalesInventory() ? salesStockNameInput : salesStockViewMode)?.focus?.();
  } else {
    librarySearch?.focus();
  }
}

function closeLibraryModal() {
  libraryModal?.classList.add("hidden");
  libraryBackdrop?.classList.add("hidden");
  libraryModal?.setAttribute("aria-hidden", "true");
  setRepoTabActive(null);
}

function openAuditModal() {
  loadAuditLogs()
    .then(() => {
      auditModal.classList.remove("hidden");
      auditBackdrop.classList.remove("hidden");
      auditModal.setAttribute("aria-hidden", "false");
      setRepoTabActive("audit");
    })
    .catch((error) => {
      showToast(error.message || "Khong the tai nhat ky.", "error");
    });
}

function closeAuditModal() {
  auditModal.classList.add("hidden");
  auditBackdrop.classList.add("hidden");
  auditModal.setAttribute("aria-hidden", "true");
  setRepoTabActive(null);
}

function openDashboardModal() {
  loadDashboard()
    .then(() => {
      dashboardModal.classList.remove("hidden");
      dashboardBackdrop.classList.remove("hidden");
      dashboardModal.setAttribute("aria-hidden", "false");
      setRepoTabActive("dashboard");
    })
    .catch((error) => {
      showToast(error.message || "Khong the tai dashboard.", "error");
    });
}

function closeDashboardModal() {
  dashboardModal.classList.add("hidden");
  dashboardBackdrop.classList.add("hidden");
  dashboardModal.setAttribute("aria-hidden", "true");
  setRepoTabActive(null);
}

function applyLibraryPreset() {
  if (libraryModalKicker) {
    libraryModalKicker.textContent =
      activeLibraryPreset === "production"
        ? "Kho của bộ phận sản xuất"
        : activeLibraryPreset === "sales"
          ? "Kho của bộ phận kinh doanh"
          : "Kho tài liệu";
  }
  if (libraryModalTitle) {
    libraryModalTitle.textContent =
      activeLibraryPreset === "production"
        ? "Nơi quản lí sản phẩm đã sản xuất"
        : activeLibraryPreset === "sales"
          ? "Nơi quản lí hàng hóa nhập ngoài"
          : "Quản lý tài liệu nội bộ";
  }

  if (!libraryFolderFilter) {
    return;
  }

  if (activeLibraryPreset === "general") {
    libraryFolderFilter.value = "";
    return;
  }

  if (activeLibraryPreset === "production") {
    libraryFolderFilter.value = Array.from(libraryFolderFilter.options || []).some((option) => option.value === "phòng Sản Xuất")
      ? "phòng Sản Xuất"
      : "";
    return;
  }

  if (activeLibraryPreset === "sales") {
    libraryFolderFilter.value = Array.from(libraryFolderFilter.options || []).some((option) => option.value === "phòng kinh doanh")
      ? "phòng kinh doanh"
      : "";
  }
}

function setRepoTabActive(activeTab) {
  openDashboardModalButton?.classList.toggle("is-active", activeTab === "dashboard");
  openLibraryModalButton?.classList.toggle("is-active", activeTab === "library");
  openProductionLibraryModalButton?.classList.toggle("is-active", activeTab === "production-library");
  openSalesLibraryModalButton?.classList.toggle("is-active", activeTab === "sales-library");
  openSalesProductsModalButton?.classList.toggle("is-active", activeTab === "sales-products");
  openAuditModalButton?.classList.toggle("is-active", activeTab === "audit");
}

async function closeDocumentViewerModal() {
  documentViewerModal.classList.add("hidden");
  documentViewerBackdrop.classList.add("hidden");
  documentViewerModal.setAttribute("aria-hidden", "true");

  if (activeNotificationOnViewerClose) {
    const notificationToDelete = activeNotificationOnViewerClose;
    activeNotificationOnViewerClose = null;
    await deleteNotification(notificationToDelete, { silent: true });
  }
}

function renderDocumentLibrary() {
  if (!librarySearch || !libraryAccessFilter || !libraryFolderFilter || !libraryStatusFilter || !libraryTableBody) {
    return;
  }
  const stockOnlyMode = activeLibraryPreset === "production" || activeLibraryPreset === "sales";
  libraryToolbar?.classList.toggle("hidden", stockOnlyMode);
  libraryFilterHint?.classList.toggle("hidden", stockOnlyMode);
  libraryDocumentsWrap?.classList.toggle("hidden", stockOnlyMode);
  renderProductionStockPanel();
  renderSalesStockPanel();
  if (stockOnlyMode) {
    return;
  }
  const query = librarySearch.value.trim().toLowerCase();
  const accessFilter = libraryAccessFilter.value.trim().toLowerCase();
  const folderFilter = libraryFolderFilter.value.trim();
  const statusFilter = libraryStatusFilter.value.trim().toLowerCase();
  const entries = [
    ...documentLibrary.map((document) => ({ kind: "document", item: document })),
    ...sourceLibrary.map((source) => ({ kind: "source", item: source })),
  ];
  const visibleEntries = entries.filter((entry) => {
    const folder = entry.kind === "source"
      ? resolveSourceFolderName(entry.item)
      : resolveDocumentFolderName(entry.item);
    const owner = entry.kind === "source"
      ? resolveSourceOwnerName(entry.item)
      : resolveDocumentOwnerName(entry.item);
    const haystack = [
      entry.item.title,
      folder,
      owner,
      entry.item.path,
      entry.kind === "source" ? entry.item.source_url : "",
    ]
      .join(" ")
      .toLowerCase();
    if (query && !haystack.includes(query)) {
      return false;
    }
    const accessLevel =
      entry.kind === "source"
        ? String(entry.item?.access_level || "").toLowerCase()
        : String(entry.item?.metadata?.access_level || "").toLowerCase();
    if (accessFilter && accessLevel !== accessFilter) {
      return false;
    }
    if (folderFilter && folder !== folderFilter) {
      return false;
    }
    const reviewStatus = entry.kind === "source"
      ? String(entry.item?.status || "").trim().toLowerCase()
      : String(entry.item?.metadata?.status || "").trim().toLowerCase();
    if (statusFilter && reviewStatus !== statusFilter) {
      return false;
    }
    return true;
  });

  libraryTableBody.innerHTML = "";
  if (!visibleEntries.length) {
    libraryTableBody.innerHTML = `
      <tr>
        <td colspan="6" class="library-empty">Chưa có tài liệu phù hợp.</td>
      </tr>
    `;
    return;
  }

  for (const entry of visibleEntries) {
    const isSource = entry.kind === "source";
    const record = entry.item;
    const row = document.createElement("tr");
    const canDelete = isSource ? false : canDeleteDocument(record);
    const canManagePrivate = isSource ? false : canManagePrivateDocument(record);
    const canManageSource = isSource && Boolean(currentUser?.policy?.can_manage_permissions);
    const canReviewRecord = canReviewLibraryRecord(entry);
    const statusBadge = isSource
      ? `${buildSourceKindBadge(record)}${buildSourceStatusBadge(record)}${buildApprovalStatusBadge(record.status || "")}`
      : `<span class="record-kind-badge document">Tài liệu</span>${buildApprovalStatusBadge(record?.metadata?.status || "")}`;
    row.innerHTML = `
      <td>
        <div class="table-title-row">
          <strong>${escapeHtml(record.title)}</strong>
          ${statusBadge}
        </div>
        <div class="table-subtle">${escapeHtml(isSource ? (record.source_url || record.path || "") : (record.path || ""))}</div>
      </td>
      <td>${escapeHtml(isSource ? resolveSourceFolderName(record) : resolveDocumentFolderName(record))}</td>
      <td>${escapeHtml(String(isSource ? (record?.added_at || record?.last_synced_at) : record?.metadata?.added_at || "").trim() || "-")}</td>
      <td>${escapeHtml(labelAccessLevel(isSource ? record?.access_level || "basic" : record?.metadata?.access_level || "basic"))}</td>
      <td>${escapeHtml(isSource ? resolveSourceOwnerName(record) : resolveDocumentOwnerName(record))}</td>
      <td>
        <div class="table-actions">
          <button type="button" class="leaf-action-button" data-action="view">${isSource ? "Mở" : "Xem"}</button>
          ${
            isSource
              ? `
                ${canManageSource ? `<button type="button" class="leaf-action-button" data-action="sync-source">Đồng bộ</button>` : ""}
                ${canManageSource ? `<button type="button" class="leaf-action-button" data-action="toggle-source">${record.syncable === false ? "Bật lại" : "Tạm dừng"}</button>` : ""}
              `
              : `<button type="button" class="leaf-action-button" data-action="download">Tải</button>`
          }
          ${
            canManagePrivate
              ? `<button type="button" class="leaf-action-button" data-action="share">Chia sẻ</button>`
              : ""
          }
          ${
            canReviewRecord
              ? `
                <button type="button" class="leaf-action-button accent" data-action="approve">Duyệt</button>
                <button type="button" class="leaf-action-button warning" data-action="reject">Từ chối</button>
              `
              : ""
          }
          ${
            canDelete || canManageSource
              ? `<button type="button" class="leaf-action-button danger" data-action="delete">Xóa</button>`
              : ""
          }
        </div>
      </td>
    `;
    row.querySelector('[data-action="view"]')?.addEventListener("click", () => {
      if (isSource) {
        openSource(record);
      } else {
        viewDocument(record.path).catch((error) => {
          showToast(error.message || "Khong the mo tai lieu.", "error");
        });
      }
    });
    row.querySelector('[data-action="download"]')?.addEventListener("click", () => {
      downloadDocument(record.path);
    });
    row.querySelector('[data-action="sync-source"]')?.addEventListener("click", () => {
      syncSourceNow(record).catch((error) => {
        showToast(error.message || "Khong the dong bo nguon.", "error");
      });
    });
    row.querySelector('[data-action="toggle-source"]')?.addEventListener("click", () => {
      toggleSourceSync(record).catch((error) => {
        showToast(error.message || "Khong the cap nhat nguon.", "error");
      });
    });
    row.querySelector('[data-action="delete"]')?.addEventListener("click", () => {
      if (isSource) {
        deleteSource(record).catch((error) => {
          showToast(error.message || "Khong the xoa nguon.", "error");
        });
      } else {
        deleteDocument(record.path, record.title).catch((error) => {
          showToast(error.message || "Khong the xoa tai lieu.", "error");
        });
      }
    });
    row.querySelector('[data-action="share"]')?.addEventListener("click", () => {
      openShareModal(record).catch((error) => {
        showToast(error.message || "Khong the cap nhat chia se.", "error");
      });
    });
    row.querySelector('[data-action="approve"]')?.addEventListener("click", () => {
      reviewLibraryRecord(entry, "approve").catch((error) => {
        showToast(error.message || "Khong the duyet muc nay.", "error");
      });
    });
    row.querySelector('[data-action="reject"]')?.addEventListener("click", () => {
      reviewLibraryRecord(entry, "reject").catch((error) => {
        showToast(error.message || "Khong the tu choi muc nay.", "error");
      });
    });
    libraryTableBody.append(row);
  }
}

function buildProductionStockMovements(orderList = orders, transfers = salesInventoryTransfers) {
  const movements = [];

  (Array.isArray(orderList) ? orderList : []).forEach((order) => {
    if (isProductionOrder(order)) {
      if (!isProductionReceiptCompleted(order)) {
        return;
      }
      const receiptRecord = parseProductionPackagingRecord(order?.production_received_by_json);
      const movementAt = String(receiptRecord?.completed_at || order?.updated_at || order?.created_at || "").trim();
      const parsedItems = parseProductionOrderItemsSummary(order?.order_items || "").filter((item) => item?.name || item?.code);
      parsedItems.forEach((item) => {
        const code = String(item?.code || "").trim();
        const name = String(item?.name || "").trim();
        const unit = String(item?.unit || "").trim() || "-";
        const quantity = Math.max(0, parseProductionNumber(item?.done || 0));
        if (!quantity || (!code && !name)) {
          return;
        }
        movements.push({
          movementType: "in",
          orderId: String(order?.order_id || "").trim(),
          movementAt,
          code,
          name,
          unit,
          quantity,
        });
      });
      return;
    }

    const movementAt = String(order?.created_at || order?.updated_at || "").trim();
    const parsedItems = parseOrderItemsSummary(order?.order_items || "").filter((item) => item?.name);
    parsedItems.forEach((item) => {
      const name = String(item?.name || "").trim();
      const unit = String(item?.unit || "").trim() || "-";
      const quantity = Math.max(0, parseProductionNumber(item?.quantity || 0));
      if (!quantity || !name) {
        return;
      }
      movements.push({
        movementType: "out",
        orderId: String(order?.order_id || "").trim(),
        movementAt,
        code: "",
        name,
        unit,
        quantity,
      });
    });
  });

  (Array.isArray(transfers) ? transfers : []).forEach((transfer) => {
    const code = String(transfer?.code || "").trim();
    const name = String(transfer?.name || "").trim();
    const unit = String(transfer?.unit || "").trim() || "-";
    const quantity = Math.max(0, parseProductionNumber(transfer?.quantity || 0));
    if (!quantity || (!code && !name)) {
      return;
    }
    movements.push({
      movementType: "out",
      orderId: String(transfer?.transfer_id || transfer?.id || "").trim(),
      movementAt: String(transfer?.transferred_at || transfer?.created_at || "").trim(),
      code,
      name,
      unit,
      quantity,
    });
  });

  return movements;
}

function formatProductionStockPeriod(value, mode = "day") {
  const date = new Date(value);
  if (!Number.isFinite(date.getTime())) {
    return "-";
  }
  if (mode === "month") {
    return new Intl.DateTimeFormat("vi-VN", {
      month: "2-digit",
      year: "numeric",
    }).format(date);
  }
  return new Intl.DateTimeFormat("vi-VN", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
  }).format(date);
}

function buildProductionStockSummary(orderList = orders, mode = productionStockView) {
  const movementMap = new Map();
  const normalizedMode = ["day", "month"].includes(String(mode || "").trim().toLowerCase()) ? String(mode).trim().toLowerCase() : "summary";

  buildProductionStockMovements(orderList, salesInventoryTransfers).forEach((movement) => {
    const movementDate = new Date(movement.movementAt || 0);
    const periodKey = normalizedMode === "summary"
      ? "summary"
      : normalizedMode === "month" && Number.isFinite(movementDate.getTime())
        ? `${movementDate.getFullYear()}-${String(movementDate.getMonth() + 1).padStart(2, "0")}`
        : Number.isFinite(movementDate.getTime())
          ? movementDate.toISOString().slice(0, 10)
          : "unknown";
    const itemKey = `${String(movement.name || "").trim().toLowerCase()}|${String(movement.unit || "").trim().toLowerCase()}`;
    const key = `${periodKey}|${itemKey}`;
    const existing = movementMap.get(key) || {
      periodKey,
      periodLabel: normalizedMode === "summary" ? "Tổng hợp" : formatProductionStockPeriod(movement.movementAt, normalizedMode),
      code: String(movement.code || "").trim(),
      name: String(movement.name || "").trim(),
      unit: String(movement.unit || "").trim() || "-",
      producedQuantity: 0,
      deliveredQuantity: 0,
      netQuantity: 0,
      completedOrders: new Set(),
      deliveryOrders: new Set(),
      lastUpdatedAt: "",
    };

    if (!existing.code && movement.code) {
      existing.code = movement.code;
    }
    if (movement.movementType === "in") {
      existing.producedQuantity += movement.quantity;
      existing.completedOrders.add(movement.orderId);
    } else {
      existing.deliveredQuantity += movement.quantity;
      existing.deliveryOrders.add(movement.orderId);
    }
    existing.netQuantity = existing.producedQuantity - existing.deliveredQuantity;
    if (movement.movementAt) {
      const currentTime = new Date(existing.lastUpdatedAt || 0).getTime();
      const nextTime = new Date(movement.movementAt).getTime();
      if (!existing.lastUpdatedAt || (Number.isFinite(nextTime) && nextTime > currentTime)) {
        existing.lastUpdatedAt = movement.movementAt;
      }
    }
    movementMap.set(key, existing);
  });

  return Array.from(movementMap.values())
    .map((item) => ({
      ...item,
      completedOrdersCount: item.completedOrders.size,
      deliveryOrdersCount: item.deliveryOrders.size,
    }))
    .sort((left, right) => {
      if (normalizedMode !== "summary" && left.periodKey !== right.periodKey) {
        return String(right.periodKey).localeCompare(String(left.periodKey), "vi");
      }
      const rightTime = new Date(right.lastUpdatedAt || 0).getTime();
      const leftTime = new Date(left.lastUpdatedAt || 0).getTime();
      if (Number.isFinite(rightTime) && Number.isFinite(leftTime) && rightTime !== leftTime) {
        return rightTime - leftTime;
      }
      return String(left.name || "").localeCompare(String(right.name || ""), "vi");
    });
}

function shouldShowProductionStockPanel() {
  const folderFilter = String(libraryFolderFilter?.value || "").trim();
  return activeLibraryPreset === "production" || folderFilter === "phòng Sản Xuất";
}

function renderProductionStockPanel() {
  if (!productionStockPanel || !productionStockBody || !productionStockCount || !productionStockHead) {
    return;
  }

  const shouldShow = shouldShowProductionStockPanel();
  productionStockPanel.classList.toggle("hidden", !shouldShow);
  productionStockPanel.setAttribute("aria-hidden", shouldShow ? "false" : "true");
  if (!shouldShow) {
    return;
  }

  if (productionStockViewMode) {
    productionStockViewMode.value = productionStockView;
  }
  productionStockSummary = canViewOrders() ? buildProductionStockSummary(orders, productionStockView) : [];
  productionStockCount.textContent = String(productionStockSummary.length);
  if (productionStockHint) {
    productionStockHint.textContent = canViewOrders()
      ? "Tồn được tính từ phiếu sản xuất đã hoàn tất nhận hàng và tự trừ khi tạo đơn giao hàng."
      : "Bạn không có quyền xem dữ liệu tồn sản xuất.";
  }
  productionStockHead.innerHTML = productionStockView === "summary"
    ? `
      <tr>
        <th>Mã hàng</th>
        <th>Tên hàng hóa</th>
        <th>ĐVT</th>
        <th>Nhập từ SX</th>
        <th>Xuất giao</th>
        <th>Tồn hiện tại</th>
        <th>Cập nhật gần nhất</th>
      </tr>
    `
    : `
      <tr>
        <th>${productionStockView === "month" ? "Tháng" : "Ngày"}</th>
        <th>Mã hàng</th>
        <th>Tên hàng hóa</th>
        <th>ĐVT</th>
        <th>Nhập SX</th>
        <th>Xuất giao</th>
        <th>Tồn ròng</th>
      </tr>
    `;

  productionStockBody.innerHTML = "";
  if (!canViewOrders() || !productionStockSummary.length) {
    productionStockBody.innerHTML = `
      <tr>
        <td colspan="7" class="library-empty">${canViewOrders() ? "Chưa có biến động kho sản xuất nào được ghi nhận." : "Không có quyền xem dữ liệu tồn sản xuất."}</td>
      </tr>
    `;
    return;
  }

  productionStockSummary.forEach((item) => {
    const row = document.createElement("tr");
    row.innerHTML = productionStockView === "summary"
      ? `
        <td>${escapeHtml(item.code || "-")}</td>
        <td><strong>${escapeHtml(item.name || "-")}</strong></td>
        <td>${escapeHtml(item.unit || "-")}</td>
        <td>${escapeHtml(formatProductionNumber(item.producedQuantity))}</td>
        <td>${escapeHtml(formatProductionNumber(item.deliveredQuantity))}</td>
        <td><strong>${escapeHtml(formatProductionNumber(item.netQuantity))}</strong></td>
        <td>${escapeHtml(item.lastUpdatedAt ? formatNotificationTime(item.lastUpdatedAt) : "-")}</td>
      `
      : `
        <td>${escapeHtml(item.periodLabel || "-")}</td>
        <td>${escapeHtml(item.code || "-")}</td>
        <td><strong>${escapeHtml(item.name || "-")}</strong></td>
        <td>${escapeHtml(item.unit || "-")}</td>
        <td>${escapeHtml(formatProductionNumber(item.producedQuantity))}</td>
        <td>${escapeHtml(formatProductionNumber(item.deliveredQuantity))}</td>
        <td><strong>${escapeHtml(formatProductionNumber(item.netQuantity))}</strong></td>
      `;
    productionStockBody.append(row);
  });
}

function canManageSalesInventory() {
  if (!currentUser) {
    return false;
  }
  const role = String(currentUser?.role || "").trim().toLowerCase();
  const department = String(currentUser?.department || "").trim().toLowerCase();
  return role === "admin" || role === "director" || role === "manager" || department === "sales";
}

function buildSalesInventoryMovements(receipts = salesInventoryReceipts, orderList = orders, transfers = salesInventoryTransfers) {
  const movements = [];

  (Array.isArray(receipts) ? receipts : []).forEach((receipt) => {
    const name = String(receipt?.name || "").trim();
    const code = String(receipt?.code || "").trim();
    const unit = String(receipt?.unit || "").trim() || "-";
    const quantity = Math.max(0, parseProductionNumber(receipt?.quantity || 0));
    if (!name || !quantity) {
      return;
    }
    movements.push({
      movementType: "in",
      movementAt: String(receipt?.received_at || receipt?.created_at || "").trim(),
      referenceId: String(receipt?.id || "").trim(),
      code,
      name,
      unit,
      quantity,
      supplier: String(receipt?.supplier || "").trim(),
      sourceDetail: String(receipt?.supplier || "").trim() || "Nhập ngoài",
      actorName: String(receipt?.created_by_user_name || "").trim(),
    });
  });

  (Array.isArray(transfers) ? transfers : []).forEach((transfer) => {
    const name = String(transfer?.name || "").trim();
    const code = String(transfer?.code || "").trim();
    const unit = String(transfer?.unit || "").trim() || "-";
    const quantity = Math.max(0, parseProductionNumber(transfer?.quantity || 0));
    if (!name || !quantity) {
      return;
    }
    movements.push({
      movementType: "in",
      movementAt: String(transfer?.transferred_at || transfer?.created_at || "").trim(),
      referenceId: String(transfer?.transfer_id || transfer?.id || "").trim(),
      code,
      name,
      unit,
      quantity,
      supplier: "Kho sản xuất",
      sourceDetail: "Kho sản xuất",
      actorName: String(transfer?.created_by_user_name || "").trim(),
    });
  });

  (Array.isArray(orderList) ? orderList : [])
    .filter((order) => !isProductionOrder(order))
    .forEach((order) => {
      const movementAt = String(order?.created_at || order?.updated_at || "").trim();
      const parsedItems = parseOrderItemsSummary(order?.order_items || "").filter((item) => item?.name);
      parsedItems.forEach((item) => {
        const name = String(item?.name || "").trim();
        const unit = String(item?.unit || "").trim() || "-";
        const quantity = Math.max(0, parseProductionNumber(item?.quantity || 0));
        if (!name || !quantity) {
          return;
        }
        movements.push({
          movementType: "out",
          movementAt,
          referenceId: String(order?.order_id || "").trim(),
          code: "",
          name,
          unit,
          quantity,
          supplier: "",
          sourceDetail: "",
          actorName: "",
        });
      });
    });

  return movements;
}

function buildSalesStockSummary(receipts = salesInventoryReceipts, orderList = orders, mode = salesStockView) {
  const movementMap = new Map();
  const normalizedMode = ["day", "month"].includes(String(mode || "").trim().toLowerCase()) ? String(mode).trim().toLowerCase() : "summary";

  buildSalesInventoryMovements(receipts, orderList, salesInventoryTransfers).forEach((movement) => {
    const movementDate = new Date(movement.movementAt || 0);
    const periodKey = normalizedMode === "summary"
      ? "summary"
      : normalizedMode === "month" && Number.isFinite(movementDate.getTime())
        ? `${movementDate.getFullYear()}-${String(movementDate.getMonth() + 1).padStart(2, "0")}`
        : Number.isFinite(movementDate.getTime())
          ? movementDate.toISOString().slice(0, 10)
          : "unknown";
    const itemKey = `${String(movement.name || "").trim().toLowerCase()}|${String(movement.unit || "").trim().toLowerCase()}`;
    const key = `${periodKey}|${itemKey}`;
    const existing = movementMap.get(key) || {
      periodKey,
      periodLabel: normalizedMode === "summary" ? "Tổng hợp" : formatProductionStockPeriod(movement.movementAt, normalizedMode),
      code: String(movement.code || "").trim(),
      name: String(movement.name || "").trim(),
      unit: String(movement.unit || "").trim() || "-",
      receivedQuantity: 0,
      deliveredQuantity: 0,
      netQuantity: 0,
      suppliers: new Set(),
      inboundRefs: new Set(),
      outboundRefs: new Set(),
      lastUpdatedAt: "",
      latestInboundAt: "",
      latestInboundBy: "",
      latestInboundSource: "",
    };

    if (!existing.code && movement.code) {
      existing.code = movement.code;
    }
    if (movement.movementType === "in") {
      existing.receivedQuantity += movement.quantity;
      existing.inboundRefs.add(movement.referenceId);
      if (movement.supplier) {
        existing.suppliers.add(movement.supplier);
      }
      const currentInboundTime = new Date(existing.latestInboundAt || 0).getTime();
      const nextInboundTime = new Date(movement.movementAt || 0).getTime();
      if (!existing.latestInboundAt || (Number.isFinite(nextInboundTime) && nextInboundTime >= currentInboundTime)) {
        existing.latestInboundAt = movement.movementAt || "";
        existing.latestInboundBy = String(movement.actorName || "").trim();
        existing.latestInboundSource = String(movement.sourceDetail || movement.supplier || "").trim();
      }
    } else {
      existing.deliveredQuantity += movement.quantity;
      existing.outboundRefs.add(movement.referenceId);
    }
    existing.netQuantity = existing.receivedQuantity - existing.deliveredQuantity;
    if (movement.movementAt) {
      const currentTime = new Date(existing.lastUpdatedAt || 0).getTime();
      const nextTime = new Date(movement.movementAt).getTime();
      if (!existing.lastUpdatedAt || (Number.isFinite(nextTime) && nextTime > currentTime)) {
        existing.lastUpdatedAt = movement.movementAt;
      }
    }
    movementMap.set(key, existing);
  });

  return Array.from(movementMap.values())
    .map((item) => ({
      ...item,
      supplierSummary: Array.from(item.suppliers).join(", "),
      inboundCount: item.inboundRefs.size,
      outboundCount: item.outboundRefs.size,
    }))
    .sort((left, right) => {
      if (normalizedMode !== "summary" && left.periodKey !== right.periodKey) {
        return String(right.periodKey).localeCompare(String(left.periodKey), "vi");
      }
      const rightTime = new Date(right.lastUpdatedAt || 0).getTime();
      const leftTime = new Date(left.lastUpdatedAt || 0).getTime();
      if (Number.isFinite(rightTime) && Number.isFinite(leftTime) && rightTime !== leftTime) {
        return rightTime - leftTime;
      }
      return String(left.name || "").localeCompare(String(right.name || ""), "vi");
    });
}

function buildCombinedInventoryAvailability(orderList = orders, receipts = salesInventoryReceipts) {
  const itemMap = new Map();

  (Array.isArray(orderList) ? orderList : []).forEach((order) => {
    if (isProductionOrder(order)) {
      if (!isProductionReceiptCompleted(order)) {
        return;
      }
      const parsedItems = parseProductionOrderItemsSummary(order?.order_items || "").filter((item) => item?.name || item?.code);
      parsedItems.forEach((item) => {
        const name = String(item?.name || "").trim();
        const unit = String(item?.unit || "").trim() || "-";
        const code = String(item?.code || "").trim();
        const quantity = Math.max(0, parseProductionNumber(item?.done || 0));
        if (!name || !quantity) {
          return;
        }
        const key = `${name.toLowerCase()}|${unit.toLowerCase()}`;
        const existing = itemMap.get(key) || { code, name, unit, quantity: 0 };
        if (!existing.code && code) {
          existing.code = code;
        }
        existing.quantity += quantity;
        itemMap.set(key, existing);
      });
      return;
    }

    const parsedItems = parseOrderItemsSummary(order?.order_items || "").filter((item) => item?.name);
    parsedItems.forEach((item) => {
      const name = String(item?.name || "").trim();
      const unit = String(item?.unit || "").trim() || "-";
      const quantity = Math.max(0, parseProductionNumber(item?.quantity || 0));
      if (!name || !quantity) {
        return;
      }
      const key = `${name.toLowerCase()}|${unit.toLowerCase()}`;
      const existing = itemMap.get(key) || { code: "", name, unit, quantity: 0 };
      existing.quantity -= quantity;
      itemMap.set(key, existing);
    });
  });

  (Array.isArray(receipts) ? receipts : []).forEach((receipt) => {
    const name = String(receipt?.name || "").trim();
    const unit = String(receipt?.unit || "").trim() || "-";
    const code = String(receipt?.code || "").trim();
    const quantity = Math.max(0, parseProductionNumber(receipt?.quantity || 0));
    if (!name || !quantity) {
      return;
    }
    const key = `${name.toLowerCase()}|${unit.toLowerCase()}`;
    const existing = itemMap.get(key) || { code, name, unit, quantity: 0 };
    if (!existing.code && code) {
      existing.code = code;
    }
    existing.quantity += quantity;
    itemMap.set(key, existing);
  });

  return itemMap;
}

function collectTransportOrderDemand() {
  return Array.from(orderItemsList?.querySelectorAll("[data-order-item-row]") || [])
    .map((row) => {
      const name = String(row.querySelector('[data-field="name"]')?.value || "").trim();
      const unit = String(row.querySelector('[data-field="unit"]')?.value || "").trim() || "-";
      const quantity = Math.max(0, parseProductionNumber(row.querySelector('[data-field="quantity"]')?.value || 0));
      return { name, unit, quantity };
    })
    .filter((item) => item.name && item.quantity > 0);
}

function checkTransportOrderInventory({ showSuccessToast = false } = {}) {
  if (activeOrderCreateKind !== "transport") {
    return { ok: true, shortages: [] };
  }

  const demandItems = collectTransportOrderDemand();
  if (!demandItems.length) {
    showToast("Hãy thêm hàng hóa vào đơn trước khi kiểm tra tồn kho.", "error");
    setOrderFormStep("2");
    orderAddItemButton?.focus();
    return { ok: false, shortages: [] };
  }

  const availabilityMap = buildCombinedInventoryAvailability(orders, salesInventoryReceipts);
  const shortages = [];

  demandItems.forEach((item) => {
    const key = `${String(item.name || "").trim().toLowerCase()}|${String(item.unit || "").trim().toLowerCase()}`;
    const available = availabilityMap.get(key)?.quantity || 0;
    if (available < item.quantity) {
      shortages.push({
        ...item,
        available,
        missing: item.quantity - available,
      });
    }
  });

  if (shortages.length) {
    const summary = shortages
      .map((item) => `${item.name} (${formatProductionNumber(item.available)}/${formatProductionNumber(item.quantity)} ${item.unit})`)
      .join("; ");
    showToast(`Không đủ tồn kho cho đơn này: ${summary}. Hãy tạo phiếu sản xuất hoặc nhập thêm hàng.`, "error");
    setOrderFormStep("2");
    checkOrderInventoryButton?.focus();
    return { ok: false, shortages };
  }

  if (showSuccessToast) {
    showToast("Tồn kho hiện tại đủ để tạo đơn giao hàng.", "success");
  }
  return { ok: true, shortages: [] };
}

function shouldShowSalesStockPanel() {
  const folderFilter = String(libraryFolderFilter?.value || "").trim();
  return activeLibraryPreset === "sales" || folderFilter === "phòng kinh doanh";
}

function renderSalesStockPanel() {
  if (!salesStockPanel || !salesStockBody || !salesStockCount || !salesStockHead) {
    return;
  }

  const shouldShow = shouldShowSalesStockPanel();
  salesStockPanel.classList.toggle("hidden", !shouldShow);
  salesStockPanel.setAttribute("aria-hidden", shouldShow ? "false" : "true");
  if (!shouldShow) {
    return;
  }

  if (salesStockViewMode) {
    salesStockViewMode.value = salesStockView;
  }
  if (salesStockReceiveForm) {
    salesStockReceiveForm.classList.toggle("hidden", !canManageSalesInventory());
  }
  if (salesStockTransferToggleButton) {
    salesStockTransferToggleButton.classList.toggle("hidden", !canManageSalesInventory());
  }
  if (salesStockReceivedAtInput && !salesStockReceivedAtInput.value) {
    salesStockReceivedAtInput.value = toLocalDateTimeInputValue(new Date());
  }
  salesStockSummary = canViewOrders() ? buildSalesStockSummary(salesInventoryReceipts, orders, salesStockView) : [];
  salesStockCount.textContent = String(salesStockSummary.length);
  if (salesStockHint) {
    salesStockHint.textContent = canViewOrders()
      ? "Hàng hóa được cộng từ các lần nhập kho ngoài và tự trừ khi tạo đơn giao hàng."
      : "Bạn không có quyền xem dữ liệu kho kinh doanh.";
  }
  salesStockHead.innerHTML = salesStockView === "summary"
    ? `
      <tr>
        <th>Mã hàng</th>
        <th>Tên hàng hóa</th>
        <th>ĐVT</th>
        <th>Nhập ngoài</th>
        <th>Xuất giao</th>
        <th>Tồn hiện tại</th>
        <th>Nguồn nhập</th>
        <th>Ngày giờ nhập</th>
        <th>Người nhập</th>
      </tr>
    `
    : `
      <tr>
        <th>${salesStockView === "month" ? "Tháng" : "Ngày"}</th>
        <th>Mã hàng</th>
        <th>Tên hàng hóa</th>
        <th>ĐVT</th>
        <th>Nhập ngoài</th>
        <th>Xuất giao</th>
        <th>Tồn ròng</th>
      </tr>
    `;

  salesStockBody.innerHTML = "";
  if (!canViewOrders() || !salesStockSummary.length) {
    salesStockBody.innerHTML = `
      <tr>
        <td colspan="7" class="library-empty">${canViewOrders() ? "Chưa có biến động kho kinh doanh nào được ghi nhận." : "Không có quyền xem dữ liệu kho kinh doanh."}</td>
      </tr>
    `;
    renderSalesStockTransferPanel();
    return;
  }

  salesStockSummary.forEach((item) => {
    const row = document.createElement("tr");
    row.innerHTML = salesStockView === "summary"
      ? `
        <td>${escapeHtml(item.code || "-")}</td>
        <td><strong>${escapeHtml(item.name || "-")}</strong></td>
        <td>${escapeHtml(item.unit || "-")}</td>
        <td>${escapeHtml(formatProductionNumber(item.receivedQuantity))}</td>
        <td>${escapeHtml(formatProductionNumber(item.deliveredQuantity))}</td>
        <td><strong>${escapeHtml(formatProductionNumber(item.netQuantity))}</strong></td>
        <td>${escapeHtml(item.latestInboundSource || item.supplierSummary || "-")}</td>
        <td>${escapeHtml(item.latestInboundAt ? formatNotificationTime(item.latestInboundAt) : "-")}</td>
        <td>${escapeHtml(item.latestInboundBy || "-")}</td>
      `
      : `
        <td>${escapeHtml(item.periodLabel || "-")}</td>
        <td>${escapeHtml(item.code || "-")}</td>
        <td><strong>${escapeHtml(item.name || "-")}</strong></td>
        <td>${escapeHtml(item.unit || "-")}</td>
        <td>${escapeHtml(formatProductionNumber(item.receivedQuantity))}</td>
        <td>${escapeHtml(formatProductionNumber(item.deliveredQuantity))}</td>
        <td><strong>${escapeHtml(formatProductionNumber(item.netQuantity))}</strong></td>
      `;
    salesStockBody.append(row);
  });
  renderSalesStockTransferPanel();
}

async function submitSalesStockReceive(event) {
  event.preventDefault();
  if (!canManageSalesInventory()) {
    showToast("Bạn không có quyền nhập kho kinh doanh.", "error");
    return;
  }

  salesStockReceiveButton.disabled = true;
  try {
    const payload = await postJson(salesInventoryReceiveUrl, {
      code: salesStockCodeInput?.value || "",
      name: salesStockNameInput?.value || "",
      unit: salesStockUnitInput?.value || "",
      quantity: salesStockQuantityInput?.value || "",
      supplier: salesStockSupplierInput?.value || "",
      received_at: salesStockReceivedAtInput?.value || "",
      note: salesStockNoteInput?.value || "",
    });
    applyWorkspacePayload(payload);
    salesStockReceiveForm?.reset();
    if (salesStockReceivedAtInput) {
      salesStockReceivedAtInput.value = toLocalDateTimeInputValue(new Date());
    }
    renderSalesStockPanel();
    showToast("Đã nhập hàng vào kho kinh doanh.", "success");
  } catch (error) {
    showToast(error.message || "Không thể nhập kho kinh doanh.", "error");
  } finally {
    salesStockReceiveButton.disabled = false;
  }
}

function getTransferableProductionStockItems() {
  const summaryItems = canViewOrders() ? buildProductionStockSummary(orders, "summary") : [];
  return (summaryItems || [])
    .filter((item) => Math.max(0, Number(item?.netQuantity || 0)) > 0)
    .map((item) => ({
      code: String(item?.code || "").trim(),
      name: String(item?.name || "").trim(),
      unit: String(item?.unit || "").trim() || "-",
      availableQuantity: Math.max(0, Number(item?.netQuantity || 0)),
    }))
    .sort((left, right) => String(left.name || "").localeCompare(String(right.name || ""), "vi"));
}

function renderSalesStockTransferPanel() {
  if (!salesStockTransferPanel || !salesStockTransferList) {
    return;
  }

  const canManage = canManageSalesInventory();
  salesStockTransferToggleButton?.classList.toggle("hidden", !canManage);
  if (salesStockTransferToggleButton) {
    salesStockTransferToggleButton.textContent = isSalesStockTransferPanelOpen ? "Đóng nhập từ kho sản xuất" : "Nhập hàng từ kho sản xuất";
  }
  salesStockTransferPanel.classList.toggle("hidden", !canManage || !isSalesStockTransferPanelOpen);
  salesStockTransferPanel.setAttribute("aria-hidden", !canManage || !isSalesStockTransferPanelOpen ? "true" : "false");
  if (!canManage || !isSalesStockTransferPanelOpen) {
    return;
  }

  if (salesStockTransferAtInput && !salesStockTransferAtInput.value) {
    salesStockTransferAtInput.value = toLocalDateTimeInputValue(new Date());
  }

  const transferableItems = getTransferableProductionStockItems();
  if (!transferableItems.length) {
    salesStockTransferList.innerHTML = `<div class="library-empty">Kho sản xuất hiện chưa có mặt hàng khả dụng để chuyển.</div>`;
    return;
  }

  salesStockTransferList.innerHTML = transferableItems
    .map(
      (item, index) => `
        <label class="sales-stock-transfer-item">
          <input type="checkbox" data-transfer-select="${index}" />
          <span class="sales-stock-transfer-meta">
            <strong>${escapeHtml(item.code || item.name || "-")}</strong>
            <span>${escapeHtml(item.name || "-")} • ${escapeHtml(item.unit || "-")} • Có thể chuyển: ${escapeHtml(formatProductionNumber(item.availableQuantity))}</span>
          </span>
          <input type="number" min="0" step="1" max="${escapeHtml(String(item.availableQuantity))}" placeholder="0" data-transfer-quantity="${index}" />
        </label>
      `,
    )
    .join("");
}

function toggleSalesStockTransferPanel() {
  if (!canManageSalesInventory()) {
    return;
  }
  isSalesStockTransferPanelOpen = !isSalesStockTransferPanelOpen;
  renderSalesStockTransferPanel();
}

async function submitSalesStockTransferFromProduction() {
  if (!canManageSalesInventory()) {
    showToast("Bạn không có quyền chuyển hàng từ kho sản xuất.", "error");
    return;
  }

  const transferableItems = getTransferableProductionStockItems();
  const items = transferableItems
    .map((item, index) => {
      const checked = salesStockTransferList?.querySelector(`[data-transfer-select="${index}"]`)?.checked;
      const quantityValue = salesStockTransferList?.querySelector(`[data-transfer-quantity="${index}"]`)?.value || "";
      const quantity = Math.max(0, parseProductionNumber(quantityValue));
      if (!checked || quantity <= 0) {
        return null;
      }
      return {
        code: item.code,
        name: item.name,
        unit: item.unit,
        quantity: Math.min(quantity, item.availableQuantity),
      };
    })
    .filter(Boolean);

  if (!items.length) {
    showToast("Hãy chọn ít nhất một mặt hàng và nhập số lượng cần chuyển.", "error");
    return;
  }

  salesStockTransferConfirmButton.disabled = true;
  try {
    const payload = await postJson(salesInventoryTransferUrl, {
      items,
      transferred_at: salesStockTransferAtInput?.value || "",
      note: salesStockTransferNoteInput?.value || "",
    });
    applyWorkspacePayload(payload);
    isSalesStockTransferPanelOpen = false;
    if (salesStockTransferAtInput) {
      salesStockTransferAtInput.value = toLocalDateTimeInputValue(new Date());
    }
    if (salesStockTransferNoteInput) {
      salesStockTransferNoteInput.value = "";
    }
    renderProductionStockPanel();
    renderSalesStockPanel();
    showToast("Đã nhập hàng từ kho sản xuất sang kho kinh doanh.", "success");
  } catch (error) {
    showToast(error.message || "Không thể chuyển hàng từ kho sản xuất.", "error");
  } finally {
    salesStockTransferConfirmButton.disabled = false;
  }
}

async function loadAuditLogs() {
  const payload = await fetchJson(auditLogsUrl);
  auditEntries = payload.logs || [];
  renderAuditLogs();
}

async function loadDashboard() {
  const payload = await fetchJson(dashboardUrl);
  dashboardStats = payload.stats || null;
  renderDashboard();
}

function renderDashboard() {
  if (!dashboardStats) {
    dashboardStatsGrid.innerHTML = "";
    dashboardFolders.innerHTML = "";
    dashboardUploaders.innerHTML = "";
    dashboardActions.innerHTML = "";
    return;
  }

  const cards = [
    { label: "Tài liệu", value: dashboardStats.documentCount || 0 },
    { label: "Nguồn đồng bộ", value: dashboardStats.sourceCount || 0 },
    { label: "Riêng tư", value: dashboardStats.privateCount || 0 },
    { label: "Chờ duyệt", value: dashboardStats.pendingApprovalCount || 0 },
    { label: "Nguồn tạm dừng", value: dashboardStats.sourcePausedCount || 0 },
    { label: "Nguồn lỗi", value: dashboardStats.sourceErrorCount || 0 },
  ];
  dashboardStatsGrid.innerHTML = cards
    .map(
      (card) => `
        <article class="dashboard-stat-card">
          <span class="dashboard-stat-label">${escapeHtml(card.label)}</span>
          <strong class="dashboard-stat-value">${escapeHtml(String(card.value))}</strong>
        </article>
      `,
    )
    .join("");

  renderDashboardSummaryList(dashboardFolders, dashboardStats.topFolders || []);
  renderDashboardSummaryList(dashboardUploaders, dashboardStats.topUploaders || []);
  renderDashboardActions(dashboardActions, dashboardStats.recentActions || []);
}

function renderDashboardSummaryList(container, items) {
  container.innerHTML = "";
  if (!items.length) {
    container.innerHTML = `<div class="history-item muted">Chưa có dữ liệu.</div>`;
    return;
  }

  for (const item of items) {
    const node = document.createElement("article");
    node.className = "dashboard-list-item";
    node.innerHTML = `
      <span>${escapeHtml(item.label || "-")}</span>
      <strong>${escapeHtml(String(item.count || 0))}</strong>
    `;
    container.append(node);
  }
}

function renderDashboardActions(container, actions) {
  container.innerHTML = "";
  if (!actions.length) {
    container.innerHTML = `<div class="history-item muted">Chưa có hoạt động.</div>`;
    return;
  }

  for (const action of actions) {
    const node = document.createElement("article");
    node.className = "dashboard-list-item wide";
    node.innerHTML = `
      <div>
        <strong>${escapeHtml(labelAuditAction(action.action || ""))}</strong>
        <div class="table-subtle">${escapeHtml(action.actor?.name || action.actor?.username || "Hệ thống")}</div>
      </div>
      <span>${escapeHtml(formatAuditTime(action.created_at || ""))}</span>
    `;
    container.append(node);
  }
}

function renderAuditLogs() {
  auditLogList.innerHTML = "";
  if (!auditEntries.length) {
    auditLogList.innerHTML = `<div class="history-item muted">Chưa có nhật ký nào.</div>`;
    return;
  }

  for (const entry of auditEntries) {
    const item = document.createElement("article");
    item.className = "audit-item";
    item.innerHTML = `
      <div class="audit-item-head">
        <strong>${escapeHtml(labelAuditAction(entry.action || ""))}</strong>
        <span class="mini-pill neutral">${escapeHtml(formatAuditTime(entry.created_at || ""))}</span>
      </div>
      <p class="table-subtle">${escapeHtml(entry.actor?.name || entry.actor?.username || "Hệ thống")}</p>
      <pre class="audit-item-body">${escapeHtml(formatAuditDetail(entry.detail || {}))}</pre>
    `;
    auditLogList.append(item);
  }
}

function resolveDocumentFolderName(document) {
  const pathValue = String(document?.path || "");
  const parts = pathValue.split("/");
  parts.pop();
  return formatFolderLabel(parts) || "internal";
}

function resolveDocumentOwnerName(document) {
  return (
    String(document?.metadata?.owner || "").trim() ||
    String(document?.metadata?.owner_user_id || "").trim() ||
    "-"
  );
}

function resolveSourceFolderName(source) {
  const storageFolder = String(source?.storage_folder || "").trim();
  if (!storageFolder) {
    return "internal";
  }
  const parts = storageFolder.split("/");
  if (parts[0] === "dữ liệu cá nhân") {
    return "dữ liệu cá nhân";
  }
  return parts[0] || "internal";
}

function resolveSourceOwnerName(source) {
  return String(source?.owner || "").trim() || String(source?.owner_user_id || "").trim() || "-";
}

function canDeleteDocument(document) {
  if (currentUser?.policy?.can_manage_permissions) {
    return true;
  }
  return String(document?.metadata?.owner_user_id || "").trim() === String(currentUser?.id || "").trim();
}

function canManagePrivateDocument(document) {
  return (
    String(document?.metadata?.access_level || "").trim().toLowerCase() === "sensitive" &&
    String(document?.metadata?.owner_user_id || "").trim() === String(currentUser?.id || "").trim()
  );
}

function canReviewLibraryRecord(entry) {
  if (!currentUser?.policy?.can_manage_permissions) {
    return false;
  }
  const status = entry.kind === "source"
    ? String(entry.item?.status || "").trim().toLowerCase()
    : String(entry.item?.metadata?.status || "").trim().toLowerCase();
  return status === "draft";
}

function openSource(source) {
  const url = String(source?.source_url || "").trim();
  if (!url) {
    showToast("Nguon dong bo nay chua co link mo.", "error");
    return;
  }
  window.open(url, "_blank", "noopener,noreferrer");
}

async function viewDocument(documentPath) {
  const payload = await fetchJson(`${docContentUrl}?path=${encodeURIComponent(documentPath)}`);
  const document = payload.document;
  documentViewerTitle.textContent = document.title || "Tài liệu";
  documentViewerMeta.innerHTML = `
    <span class="mini-pill neutral">${escapeHtml(labelDocumentStatus(document?.metadata?.status || "active"))}</span>
    <span class="mini-pill neutral">${escapeHtml(labelAccessLevel(document?.metadata?.access_level || "basic"))}</span>
    <span class="mini-pill neutral">${escapeHtml(resolveDocumentFolderName(document))}</span>
    <span class="mini-pill neutral">${escapeHtml(String(document?.metadata?.added_at || "").trim() || "-")}</span>
  `;
  renderDocumentPreview(documentViewerBody, document);
  documentViewerModal.classList.remove("hidden");
  documentViewerBackdrop.classList.remove("hidden");
  documentViewerModal.setAttribute("aria-hidden", "false");
}

async function downloadDocument(documentPath) {
  const response = await fetch(`${docDownloadUrl}?path=${encodeURIComponent(documentPath)}`, {
    headers: {
      ...(authToken ? { Authorization: `Bearer ${authToken}` } : {}),
    },
  });
  if (!response.ok) {
    const payload = await response.json().catch(() => ({}));
    throw new Error(payload.error || "Khong the tai tai lieu.");
  }

  const blob = await response.blob();
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = documentPath.split("/").pop() || "tai-lieu.md";
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

async function deleteDocument(documentPath, title) {
  if (!window.confirm(`Xóa tài liệu "${title}"?`)) {
    return;
  }
  const payload = await postJson(docDeleteUrl, { path: documentPath });
  applyWorkspacePayload(payload);
  showToast("Đã xóa tài liệu.", "success");
}

async function reviewLibraryRecord(entry, decision) {
  const isApprove = decision === "approve";
  const targetTitle = entry.item?.title || "mục này";
  if (!window.confirm(`${isApprove ? "Duyệt" : "Từ chối"} "${targetTitle}"?`)) {
    return;
  }

  const payload = entry.kind === "source"
    ? await postJson(reviewSourceUrl, { source_id: entry.item.id, decision })
    : await postJson(docReviewUrl, { path: entry.item.path, decision });
  applyWorkspacePayload(payload);
  showToast(isApprove ? "Đã duyệt." : "Đã từ chối.", "success");
}

function applyWorkspacePayload(payload) {
  if (Array.isArray(payload.orders)) {
    orders = payload.orders;
    renderOrdersBoard();
  }
  if (Array.isArray(payload.receipts)) {
    salesInventoryReceipts = payload.receipts;
  }
  if (Array.isArray(payload.transfers)) {
    salesInventoryTransfers = payload.transfers;
  }
  if (payload.sources) {
    sourceLibrary = payload.sources || [];
    renderSources(payload.sources || []);
  }
  documentLibrary = payload.documents || documentLibrary;
  updateDocumentFilterOptions(documentLibrary, payload.folders || [], sourceLibrary);
  renderDocuments(documentLibrary, payload.folders || []);
  renderDocumentLibrary();
  loadHealth().catch(() => {});
}

function labelAuditAction(action) {
  const labels = {
    "folder.create": "Tạo thư mục",
    "folder.rename": "Đổi tên thư mục",
    "folder.delete": "Xóa thư mục",
    "document.create": "Tạo tài liệu",
    "document.approve": "Duyệt tài liệu",
    "document.delete": "Xóa tài liệu",
    "document.access.update": "Cập nhật quyền riêng tư",
    "document.reject": "Từ chối tài liệu",
    "source.approve": "Duyệt nguồn đồng bộ",
    "source.create": "Tạo nguồn đồng bộ",
    "source.reject": "Từ chối nguồn đồng bộ",
    "source.sync": "Đồng bộ nguồn",
    "source.syncable": "Bật/tắt nguồn",
    "source.delete": "Xóa nguồn đồng bộ",
    "order.create": "Tạo đơn hàng",
    "order.update": "Sửa đơn hàng",
    "production.claim": "Nhận sản xuất",
    "delivery.complete": "Hoàn thành giao hàng",
    "sales-inventory.receive": "Nhập kho kinh doanh",
    "sales-inventory.transfer": "Chuyển từ kho sản xuất",
  };
  return labels[action] || action || "Hoạt động";
}

function formatAuditDetail(detail) {
  return Object.entries(detail)
    .map(([key, value]) => `${key}: ${String(value || "-")}`)
    .join("\n");
}

function formatAuditTime(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return "-";
  }
  return new Intl.DateTimeFormat("vi-VN", {
    dateStyle: "short",
    timeStyle: "short",
  }).format(date);
}

function buildUserDirectoryHint() {
  return allUsers
    .slice(0, 8)
    .map((user) => `${user.id}: ${user.name || user.username}`)
    .join("\n");
}

window.__openLibraryModal = () => {
  openLibraryModal();
};

function buildSourceKindBadge(source) {
  const sourceType = String(source?.type || source?.source_type || "sheet").trim().toLowerCase();
  const label = sourceType === "web" ? "Web" : "Sheet";
  return `<span class="record-kind-badge source-kind">${escapeHtml(label)}</span>`;
}

function buildApprovalStatusBadge(status) {
  const normalized = String(status || "").trim().toLowerCase();
  if (normalized === "draft") {
    return `<span class="record-kind-badge pending">Chờ duyệt</span>`;
  }
  if (normalized === "rejected") {
    return `<span class="record-kind-badge rejected">Từ chối</span>`;
  }
  if (normalized === "active") {
    return `<span class="record-kind-badge approved">Đã duyệt</span>`;
  }
  return "";
}

function labelDocumentStatus(status) {
  const normalized = String(status || "").trim().toLowerCase();
  if (normalized === "draft") {
    return "Chờ duyệt";
  }
  if (normalized === "rejected") {
    return "Từ chối";
  }
  if (normalized === "superseded") {
    return "Đã thay thế";
  }
  return "Đã duyệt";
}

function renderDocumentPreview(container, document) {
  container.innerHTML = "";
  const content = String(document?.content || "");
  const fileName = String(document?.file_name || document?.path || "").toLowerCase();

  if (fileName.endsWith(".json")) {
    renderJsonPreview(container, content);
    return;
  }

  if (fileName.endsWith(".csv") || fileName.endsWith(".tsv")) {
    renderDelimitedPreview(container, content, fileName.endsWith(".tsv") ? "\t" : ",");
    return;
  }

  if (fileName.endsWith(".md")) {
    renderMarkdownPreview(container, content);
    return;
  }

  const pre = document.createElement("pre");
  pre.className = "document-preview-plain";
  pre.textContent = content;
  container.append(pre);
}

function sourceLinkIconSvg() {
  return `
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
      <path d="M10 13a5 5 0 0 0 7.07 0l2.12-2.12a5 5 0 1 0-7.07-7.07L10.8 5.13" />
      <path d="M14 11a5 5 0 0 0-7.07 0L4.8 13.13a5 5 0 0 0 7.07 7.07L13.2 18.87" />
    </svg>
  `;
}

function renderJsonPreview(container, content) {
  const pre = document.createElement("pre");
  pre.className = "document-preview-plain";
  try {
    pre.textContent = JSON.stringify(JSON.parse(content), null, 2);
  } catch {
    pre.textContent = content;
  }
  container.append(pre);
}

function renderDelimitedPreview(container, content, separator) {
  const rows = String(content || "")
    .split(/\r?\n/)
    .filter(Boolean)
    .slice(0, 50)
    .map((line) => line.split(separator).map((cell) => cell.trim()));

  if (!rows.length) {
    const pre = document.createElement("pre");
    pre.className = "document-preview-plain";
    pre.textContent = content;
    container.append(pre);
    return;
  }

  const table = document.createElement("table");
  table.className = "document-preview-table";
  const headRow = rows[0];
  const thead = document.createElement("thead");
  const headerTr = document.createElement("tr");
  for (const cell of headRow) {
    const th = document.createElement("th");
    th.textContent = cell;
    headerTr.append(th);
  }
  thead.append(headerTr);
  table.append(thead);

  const tbody = document.createElement("tbody");
  for (const row of rows.slice(1)) {
    const tr = document.createElement("tr");
    for (const cell of row) {
      const td = document.createElement("td");
      td.textContent = cell;
      tr.append(td);
    }
    tbody.append(tr);
  }
  table.append(tbody);
  container.append(table);
}

function renderMarkdownPreview(container, content) {
  const blocks = String(content || "").replace(/\r/g, "").split(/\n{2,}/);
  for (const block of blocks) {
    const text = block.trim();
    if (!text) {
      continue;
    }

    if (text.startsWith("### ")) {
      const h3 = document.createElement("h3");
      h3.className = "document-preview-heading h3";
      h3.textContent = text.slice(4).trim();
      container.append(h3);
      continue;
    }

    if (text.startsWith("## ")) {
      const h2 = document.createElement("h2");
      h2.className = "document-preview-heading h2";
      h2.textContent = text.slice(3).trim();
      container.append(h2);
      continue;
    }

    if (text.startsWith("# ")) {
      const h1 = document.createElement("h2");
      h1.className = "document-preview-heading h1";
      h1.textContent = text.slice(2).trim();
      container.append(h1);
      continue;
    }

    const lines = text.split("\n").map((line) => line.trim());
    const isBulletList = lines.every((line) => line.startsWith("- ") || line.startsWith("* "));
    if (isBulletList) {
      const ul = document.createElement("ul");
      ul.className = "document-preview-list";
      for (const line of lines) {
        const li = document.createElement("li");
        li.textContent = line.slice(2).trim();
        ul.append(li);
      }
      container.append(ul);
      continue;
    }

    const paragraph = document.createElement("p");
    paragraph.className = "document-preview-paragraph";
    paragraph.textContent = lines.join(" ");
    container.append(paragraph);
  }
}

async function openShareModal(document) {
  activePrivateDocument = document;
  shareModalTitle.textContent = document.title || "Cập nhật quyền xem tài liệu";
  const ownerId = String(document?.metadata?.owner_user_id || "").trim();
  shareOwnerSelect.innerHTML = allUsers
    .map(
      (user) =>
        `<option value="${escapeHtml(user.id)}" ${user.id === ownerId ? "selected" : ""}>${escapeHtml(user.name || user.username)} (${escapeHtml(user.id)})</option>`,
    )
    .join("");
  shareUserSearch.value = "";
  renderShareModalUsers();
  shareModal.classList.remove("hidden");
  shareBackdrop.classList.remove("hidden");
  shareModal.setAttribute("aria-hidden", "false");
  shareUserSearch.focus();
}

function closeShareModal() {
  activePrivateDocument = null;
  shareModal.classList.add("hidden");
  shareBackdrop.classList.add("hidden");
  shareModal.setAttribute("aria-hidden", "true");
}

function renderShareModalUsers() {
  if (!activePrivateDocument) {
    shareModalUserList.innerHTML = "";
    return;
  }
  const query = shareUserSearch.value.trim().toLowerCase();
  const ownerId = String(activePrivateDocument?.metadata?.owner_user_id || "").trim();
  const selectedIds = new Set(
    String(activePrivateDocument?.metadata?.shared_with_users || "")
      .split(",")
      .map((item) => item.trim())
      .filter(Boolean),
  );
  shareModalUserList.innerHTML = "";
  for (const user of allUsers) {
    if (user.id === ownerId) {
      continue;
    }
    const haystack = [user.id, user.name, user.username, user.employee_code].join(" ").toLowerCase();
    if (query && !haystack.includes(query)) {
      continue;
    }
    const item = document.createElement("label");
    item.className = "share-user-option";
    item.innerHTML = `
      <input type="checkbox" value="${escapeHtml(user.id)}" ${selectedIds.has(user.id) ? "checked" : ""} />
      <span>${escapeHtml(user.name || user.username)} (${escapeHtml(user.id)})</span>
    `;
    shareModalUserList.append(item);
  }
}

async function submitShareModal() {
  if (!activePrivateDocument) {
    return;
  }
  const selectedUsers = [...shareModalUserList.querySelectorAll('input[type="checkbox"]:checked')]
    .map((input) => input.value.trim())
    .filter(Boolean)
    .join(",");
  const payload = await postJson(docAccessUrl, {
    path: activePrivateDocument.path,
    owner_user_id: shareOwnerSelect.value.trim(),
    shared_with_users: selectedUsers,
  });
  applyWorkspacePayload(payload);
  closeShareModal();
  showToast("Đã cập nhật chia sẻ.", "success");
}

async function syncSourceNow(source) {
  const payload = await postJson(syncOneSourceUrl, {
    source_id: source.id,
  });
  applyWorkspacePayload(payload);
  showToast("Đã đồng bộ nguồn.", "success");
}

async function toggleSourceSync(source) {
  const payload = await postJson(sourceSyncableUrl, {
    source_id: source.id,
    syncable: source.syncable === false,
  });
  applyWorkspacePayload(payload);
  showToast(source.syncable === false ? "Đã bật lại nguồn." : "Đã tạm dừng nguồn.", "success");
}

async function deleteSource(source) {
  if (!window.confirm(`Xóa nguồn đồng bộ "${source.title}"?`)) {
    return;
  }
  const payload = await postJson(deleteSourceUrl, {
    source_id: source.id,
  });
  applyWorkspacePayload(payload);
  showToast("Đã xóa nguồn đồng bộ.", "success");
}

function routeName(route) {
  if (route === "chat") {
    return {
      title: "Trò chuyện",
      pill: "Chat",
      className: "neutral",
      short: "CH",
    };
  }

  if (route === "internal") {
    return {
      title: "Trả lời từ kho nội bộ",
      pill: "Nội bộ",
      className: "internal",
      short: "NB",
    };
  }

  if (route === "web") {
    return {
      title: "Không có dữ liệu nội bộ, chuyển sang tìm trên web",
      pill: "Internet",
      className: "web",
      short: "WB",
    };
  }

  return {
    title: "Hệ thống gặp lỗi",
    pill: "Lỗi",
    className: "neutral",
    short: "LO",
  };
}

function openImportModal() {
  if (!currentUser) {
    showToast("Bạn cần đăng nhập để tải tài liệu.", "error");
    return;
  }
  importOwner.readOnly = true;
  syncImportPolicy();
  updateFolderOptions();
  renderPrivateShareOptions();
  importModal.classList.remove("hidden");
  importBackdrop.classList.remove("hidden");
  importModal.setAttribute("aria-hidden", "false");
  importTitle.focus();
}

function openImportModalWithPreset({ mode = "file", level = "basic" } = {}) {
  importStorageLevel.value = canCreateManagedContent()
    ? String(level || "basic").trim().toLowerCase()
    : "sensitive";
  setImportMode(mode);
  openImportModal();
}

function closeImportModal() {
  importModal.classList.add("hidden");
  importBackdrop.classList.add("hidden");
  importModal.setAttribute("aria-hidden", "true");
  importForm.reset();
  invalidateSheetValidation();
  setImportMode("file");
  syncImportPolicy();
  updateFolderOptions();
  renderPrivateShareOptions();
}

function setImportMode(mode) {
  const requestedMode = mode === "sheet" ? "sheet" : "file";
  importMode = !canCreateManagedContent() ? "file" : requestedMode;
  fileImportPanel.classList.toggle("hidden", importMode !== "file");
  sheetImportPanel.classList.toggle("hidden", importMode !== "sheet");
  importFile.disabled = importMode !== "file";
  importSheetUrl.disabled = importMode !== "sheet";
  checkSheetButton.disabled = importMode !== "sheet";
  updateImportSubmitState();

  for (const button of document.querySelectorAll(".tab-button")) {
    button.classList.toggle("active", button.dataset.mode === importMode);
    button.classList.toggle("hidden", !canCreateManagedContent() && button.dataset.mode === "sheet");
  }
}

function updateFolderOptions() {
  const level = String(importStorageLevel.value || "basic").trim().toLowerCase();
  const options = getStorageFolderOptions(level);
  const previous = importFolder.value;
  importFolder.innerHTML = options
    .map(
      (option) =>
        `<option value="${escapeHtml(option.value)}" ${option.value === previous ? "selected" : ""}>${escapeHtml(option.label)}</option>`,
    )
    .join("");

  if (![...importFolder.options].some((option) => option.value === previous)) {
    importFolder.value = options[0]?.value || "";
  }

  importOwner.value = resolveFolderOwnerLabel(importFolder.value);
  toggleCustomFolderInput();
}

function renderPrivateShareOptions() {
  const isPrivate = String(importStorageLevel.value || "").toLowerCase() === "sensitive";
  privateSharePanel.classList.toggle("hidden", !isPrivate);
  importOwner.value = resolveFolderOwnerLabel(importFolder.value);

  if (!isPrivate) {
    privateShareList.innerHTML = "";
    return;
  }

  const users = [...allUsers]
    .filter((user) => user.id !== currentUser?.id)
    .sort((left, right) => (left.name || left.username || "").localeCompare(right.name || right.username || "", "vi"));

  if (users.length === 0) {
    privateShareList.innerHTML = `<p class="helper-text">Chưa có người dùng nào khác để chia sẻ.</p>`;
    return;
  }

  privateShareList.innerHTML = users
    .map(
      (user) => `
        <label class="share-user-option">
          <input type="checkbox" value="${escapeHtml(user.id)}" />
          <span>${escapeHtml(user.name || user.username || user.id)} (${escapeHtml(labelDepartment(user.department))})</span>
        </label>
      `,
    )
    .join("");

  for (const input of privateShareList.querySelectorAll("input[type='checkbox']")) {
    input.addEventListener("change", invalidateSheetValidation);
  }
}

function getSelectedSharedUsers() {
  if (String(importStorageLevel.value || "").toLowerCase() !== "sensitive") {
    return [];
  }

  return [...privateShareList.querySelectorAll("input[type='checkbox']:checked")].map(
    (input) => input.value,
  );
}

function resolveFolderOwnerLabel(folderKey) {
  const labels = {
    internal: "internal",
    "van-chuyen": "Vận Chuyển",
    "ke-toan": "Kế Toán",
    "phong-kinh-doanh": "phòng kinh doanh",
    "phong-san-xuat": "phòng Sản Xuất",
    "du-lieu-ca-nhan": currentUser?.name || "dữ liệu cá nhân",
  };

  if (folderKey === "__new__" || isExistingCustomFolderOption(folderKey)) {
    return currentUser?.name || currentUser?.username || "thư mục mới";
  }

  return labels[String(folderKey || "").trim()] || "";
}

function getStorageFolderOptions(level) {
  if (level === "sensitive") {
    return [
      {
        value: "du-lieu-ca-nhan",
        label: `Cá nhân / ${resolvePrivateFolderLabel()}`,
      },
    ];
  }

  const baseOptions = STORAGE_FOLDER_OPTIONS[level] || STORAGE_FOLDER_OPTIONS.basic;
  const seen = new Set(
    baseOptions.flatMap((option) => [String(option.value || "").trim().toLowerCase(), String(option.label || "").trim().toLowerCase()]),
  );
  const extraOptions = (dynamicFolderOptions[level] || [])
    .filter((label) => {
      const normalized = String(label || "").trim().toLowerCase();
      if (!normalized || seen.has(normalized)) {
        return false;
      }
      seen.add(normalized);
      return true;
    })
    .map((label) => ({
      value: label,
      label,
    }));
  const options = [...baseOptions, ...extraOptions];

  if (canCreateManagedContent()) {
    options.push({
      value: "__new__",
      label: "Tạo thư mục mới...",
    });
  }

  return options;
}

function resolvePrivateFolderLabel() {
  return (
    String(currentUser?.employee_code || "").trim() ||
    String(currentUser?.username || "").trim() ||
    "tai-khoan-cua-ban"
  );
}

function buildTopicKey(title) {
  return String(title || "")
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function isCustomFolderOption(value) {
  return value === "__new__" || isExistingCustomFolderOption(value);
}

function toggleCustomFolderInput() {
  const isNewCustom = String(importFolder.value || "").trim() === "__new__";
  importCustomFolderField?.classList.toggle("hidden", !isNewCustom);
  if (!isNewCustom && importCustomFolder) {
    importCustomFolder.value = "";
  }
}

function isExistingCustomFolderOption(value) {
  const normalized = String(value || "").trim();
  return (
    normalized !== "" &&
    normalized !== "__new__" &&
    !Object.values(STORAGE_FOLDER_OPTIONS)
      .flat()
      .some((option) => option.value === normalized)
  );
}

async function validateSheetLink() {
  const title = importTitle.value.trim();
  const sheetUrl = importSheetUrl.value.trim();
  const effective_date = getTodayDateValue();
  const status = "active";
  const topic_key = buildTopicKey(title);
  const storage_level = String(importStorageLevel.value || "basic").trim().toLowerCase();
  const storage_folder_key = importFolder.value.trim();
  const custom_folder_name = importCustomFolder.value.trim();
  const shared_with_users = getSelectedSharedUsers().join(",");

  if (!title) {
    importTitle.focus();
    return;
  }

  if (!storage_folder_key || (isCustomFolderOption(storage_folder_key) && !custom_folder_name)) {
    showToast("Hãy nhập đủ metadata trước khi kiểm tra link.", "error");
    return;
  }

  if (!sheetUrl) {
    importSheetUrl.focus();
    return;
  }

  checkSheetButton.disabled = true;
  sheetCheckResult.className = "sheet-check-result";
  sheetCheckResult.textContent = "Đang kiểm tra khả năng đọc dữ liệu từ link...";
  sheetCheckResult.classList.remove("hidden");

  try {
    const payload = await postJson(checkDocsUrl, {
      title,
      sheetUrl,
      effective_date,
      status,
      topic_key,
      storage_level,
      storage_folder_key,
      custom_folder_name,
      shared_with_users,
    });
    isSheetValidated = true;
    sheetCheckResult.className = "sheet-check-result success";
    sheetCheckResult.textContent = `Đọc được dữ liệu. ${payload.result?.lineCount || 0} dòng. ${payload.result?.preview || ""}`;
    updateImportSubmitState();
  } catch (error) {
    isSheetValidated = false;
    sheetCheckResult.className = "sheet-check-result error";
    renderSheetCheckError(error.message || "Không đọc được dữ liệu từ link này.");
    updateImportSubmitState();
  } finally {
    checkSheetButton.disabled = false;
  }
}

function invalidateSheetValidation() {
  isSheetValidated = false;
  sheetCheckResult.textContent = "";
  sheetCheckResult.className = "sheet-check-result hidden";
  updateImportSubmitState();
}

function updateImportSubmitState() {
  const disabled = importMode === "sheet" ? !isSheetValidated : false;
  confirmImportButton.disabled = disabled;
}

function syncImportPolicy() {
  const canManage = canCreateManagedContent();
  for (const option of importStorageLevel.options) {
    const value = String(option.value || "").trim().toLowerCase();
    const allowed = canManage ? true : value === "sensitive";
    option.disabled = !allowed;
    option.hidden = !allowed;
  }
  if (!canManage) {
    importStorageLevel.value = "sensitive";
  }
}

function getTodayDateValue() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function renderSheetCheckError(message) {
  const duplicate = parseDuplicateSheetMessage(message);
  if (!duplicate) {
    sheetCheckResult.textContent = message;
    return;
  }

  sheetCheckResult.innerHTML = `
    <p class="sheet-check-title">Link này đã được thêm trước đó</p>
    <div class="sheet-check-grid">
      <p><strong>Link:</strong> ${escapeHtml(duplicate.link)}</p>
      <p><strong>Tiêu đề:</strong> ${escapeHtml(duplicate.title)}</p>
      <p><strong>Đường dẫn:</strong> ${escapeHtml(duplicate.path)}</p>
      <p><strong>Thời điểm thêm:</strong> ${escapeHtml(duplicate.addedAt)}</p>
    </div>
  `;
}

function handleComposerPaste(event) {
  const items = Array.from(event.clipboardData?.items || []);
  const imageItem = items.find((item) => item.type.startsWith("image/"));
  if (!imageItem) {
    return;
  }

  event.preventDefault();
  const file = imageItem.getAsFile();
  if (!file) {
    return;
  }

  setPastedImage(file).catch((error) => {
    showToast(error.message || "Không thể đọc ảnh đã dán.", "error");
  });
}

async function setPastedImage(file) {
  const previewUrl = URL.createObjectURL(file);
  const dataBase64 = await fileToBase64(file);

  if (pastedImage?.previewUrl) {
    URL.revokeObjectURL(pastedImage.previewUrl);
  }

  pastedImage = {
    name: file.name || "anh-da-dan.png",
    size: file.size,
    mimeType: file.type || "image/png",
    dataBase64,
    previewUrl,
    extractedText: "",
    ocrSkipped: false,
  };

  pastedImagePreview.src = previewUrl;
  pastedImageName.textContent = pastedImage.name;
  pastedImageSize.textContent = formatFileSize(pastedImage.size);
  composerAttachments.classList.remove("hidden");
  ocrPreview.classList.add("hidden");
  ocrDecision.classList.remove("hidden");
  showToast("Đã dán ảnh vào ô hỏi.", "success");
}

function clearPastedImage() {
  if (pastedImage?.previewUrl) {
    URL.revokeObjectURL(pastedImage.previewUrl);
  }

  pastedImage = null;
  pastedImagePreview.removeAttribute("src");
  pastedImageName.textContent = "Ảnh đã dán";
  pastedImageSize.textContent = "";
  ocrPreviewBody.textContent = "";
  ocrPreview.classList.add("hidden");
  ocrDecision.classList.add("hidden");
  composerAttachments.classList.add("hidden");
}

async function populateOcrPreview(forceRefresh = false) {
  if (!pastedImage) {
    return;
  }

  if (pastedImage.extractedText && !forceRefresh) {
    ocrPreviewBody.textContent = pastedImage.extractedText;
    ocrDecision.classList.add("hidden");
    ocrPreview.classList.remove("hidden");
    return;
  }

  refreshOcrButton.disabled = true;
  confirmOcrButton.disabled = true;
  skipOcrButton.disabled = true;
  ocrDecision.classList.add("hidden");
  ocrPreviewBody.textContent = "Đang đọc nội dung ảnh...";
  ocrPreview.classList.remove("hidden");

  try {
    const payload = await postJson(askUrl, {
      question: "",
      image: {
        mimeType: pastedImage.mimeType,
        dataBase64: pastedImage.dataBase64,
      },
      previewOnly: true,
    });

    pastedImage.extractedText = payload.extractedText || "Không trích được nội dung từ ảnh.";
    pastedImage.ocrSkipped = false;
    ocrPreviewBody.textContent = pastedImage.extractedText;
  } finally {
    refreshOcrButton.disabled = false;
    confirmOcrButton.disabled = false;
    skipOcrButton.disabled = false;
  }
}

function handleComposerDragEnter(event) {
  if (!hasDraggedImage(event)) {
    return;
  }

  event.preventDefault();
  dragDepth += 1;
  dropHint.classList.remove("hidden");
}

function handleComposerDragOver(event) {
  if (!hasDraggedImage(event)) {
    return;
  }

  event.preventDefault();
  dropHint.classList.remove("hidden");
}

function handleComposerDragLeave(event) {
  if (!hasDraggedImage(event)) {
    return;
  }

  event.preventDefault();
  dragDepth = Math.max(0, dragDepth - 1);
  if (dragDepth === 0) {
    dropHint.classList.add("hidden");
  }
}

function handleComposerDrop(event) {
  if (!hasDraggedImage(event)) {
    return;
  }

  event.preventDefault();
  dragDepth = 0;
  dropHint.classList.add("hidden");

  const file = Array.from(event.dataTransfer?.files || []).find((item) =>
    item.type.startsWith("image/"),
  );

  if (!file) {
    showToast("Chỉ hỗ trợ kéo thả file ảnh vào ô hỏi.", "error");
    return;
  }

  setPastedImage(file).catch((error) => {
    showToast(error.message || "Không thể đọc ảnh đã thả.", "error");
  });
}

function hasDraggedImage(event) {
  const types = Array.from(event.dataTransfer?.types || []);
  return types.includes("Files");
}

function showToast(message, type = "success") {
  clearTimeout(toastTimer);
  toast.textContent = message;
  toast.className = `toast ${type}`;
  toast.classList.remove("hidden");

  toastTimer = setTimeout(() => {
    toast.classList.add("hidden");
  }, 2600);
}

function startNotificationPolling() {
  stopNotificationPolling();
  if (!currentUser) {
    return;
  }
  notificationPollTimer = window.setInterval(() => {
    loadNotifications().catch(() => {});
  }, 10000);
}

function stopNotificationPolling() {
  if (notificationPollTimer) {
    window.clearInterval(notificationPollTimer);
    notificationPollTimer = null;
  }
}

function startNotificationStream() {
  stopNotificationStream();
  if (!authToken || !currentUser) {
    return;
  }

  notificationStream = new EventSource(`${notificationsStreamUrl}?token=${encodeURIComponent(authToken)}`);
  notificationStream.addEventListener("notification", (event) => {
    try {
      const item = JSON.parse(event.data || "{}");
      handleRealtimeNotification(item);
    } catch {}
  });
  notificationStream.onerror = () => {
    stopNotificationStream();
  };
}

function stopNotificationStream() {
  if (notificationStream) {
    notificationStream.close();
    notificationStream = null;
  }
}

function announceIncomingNotifications(nextNotifications = []) {
  if (!Array.isArray(nextNotifications)) {
    return;
  }

  const freshUnread = nextNotifications.filter((item) => {
    const id = String(item?.id || "").trim();
    return id && !seenNotificationIds.has(id) && !String(item?.read_at || "").trim();
  });

  for (const item of nextNotifications) {
    const id = String(item?.id || "").trim();
    if (id) {
      seenNotificationIds.add(id);
    }
  }

  if (!hasBootstrappedNotifications) {
    hasBootstrappedNotifications = true;
    return;
  }

  for (const item of freshUnread.slice(0, 2)) {
    showToast(item.title || item.message || "Có thông báo mới.", "success");
  }
}

function handleRealtimeNotification(item) {
  ensureNotificationUiStateLoaded();
  const id = String(item?.id || "").trim();
  if (!id || seenNotificationIds.has(id)) {
    return;
  }

  seenNotificationIds.add(id);
  hasBootstrappedNotifications = true;
  notifications = [item, ...notifications.filter((existing) => String(existing?.id || "").trim() !== id)].slice(0, 30);
  mergeNotificationArchiveItems([item]);
  renderNotifications();
  showToast(item.title || item.message || "Có thông báo mới.", "success");
}

function createSectionTitle(text) {
  const title = document.createElement("p");
  title.className = "section-title";
  title.textContent = text;
  return title;
}

function toParagraphs(text) {
  return normalizeDisplayText(text)
    .split(/\n+/)
    .map((part) => part.trim())
    .filter(Boolean);
}

function createMiniPill(text, className = "neutral") {
  const span = document.createElement("span");
  span.className = `mini-pill ${className}`;
  span.textContent = text;
  return span;
}

function setLoadingState(isLoading) {
  submitButton.disabled = isLoading;
  submitButton.textContent = isLoading ? "Đang gửi" : "Gửi";
  questionInput.disabled = isLoading;
}

function autoResizeTextarea() {
  questionInput.style.height = "auto";
  questionInput.style.height = `${Math.min(questionInput.scrollHeight, 220)}px`;
}

function scrollToLatestMessage(target) {
  if (!target) {
    return;
  }

  const composer = document.querySelector(".composer-wrap");
  const composerHeight = composer?.offsetHeight || 0;
  const top = window.scrollY + target.getBoundingClientRect().top;
  const offset = Math.max(140, composerHeight + 56);

  window.scrollTo({
    top: Math.max(0, top - offset),
    behavior: "smooth",
  });
}

async function fetchJson(url, options = {}) {
  const hadAuthToken = Boolean(authToken);
  const response = await fetch(url, {
    cache: "no-store",
    headers: {
      Accept: "application/json",
      ...(authToken && !options.skipAuth ? { Authorization: `Bearer ${authToken}` } : {}),
    },
  });

  const payload = await response.json();
  if (!response.ok) {
    if (response.status === 401 && !options.skipAuth) {
      authToken = "";
      localStorage.removeItem("authToken");
      if (hadAuthToken && !hasShownSessionExpiredToast) {
        hasShownSessionExpiredToast = true;
        showToast("Phiên đăng nhập đã hết, vui lòng đăng nhập lại.", "error");
      }
      showLoginModal();
    }
    throw new Error(payload.error || "Yêu cầu thất bại.");
  }

  if (hadAuthToken) {
    hasShownSessionExpiredToast = false;
  }

  return payload;
}

async function postJson(url, data, options = {}) {
  const hadAuthToken = Boolean(authToken);
  const response = await fetch(url, {
    method: "POST",
    cache: "no-store",
    headers: {
      "Content-Type": "application/json",
      Accept: "application/json",
      ...(authToken && !options.skipAuth ? { Authorization: `Bearer ${authToken}` } : {}),
    },
    body: JSON.stringify(data),
  });

  const payload = await response.json();
  if (!response.ok) {
    if (response.status === 401 && !options.skipAuth) {
      authToken = "";
      localStorage.removeItem("authToken");
      if (hadAuthToken && !hasShownSessionExpiredToast) {
        hasShownSessionExpiredToast = true;
        showToast("Phiên đăng nhập đã hết, vui lòng đăng nhập lại.", "error");
      }
      showLoginModal();
    }
    throw new Error(payload.error || "Yêu cầu thất bại.");
  }

  if (hadAuthToken) {
    hasShownSessionExpiredToast = false;
  }

  return payload;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function clampText(value, limit) {
  const text = String(value || "").trim();
  if (text.length <= limit) {
    return text;
  }

  return `${text.slice(0, limit).trimEnd()}...`;
}

function normalizeDisplayText(value) {
  return String(value || "")
    .replace(/\r/g, "")
    .split("\n")
    .map(normalizeDisplayLine)
    .join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function groupDocumentsByFolder(documents) {
  const map = new Map();

  for (const document of documents) {
    const pathValue = String(document.path || "");
    const parts = pathValue.split("/");
    const fileName = parts.pop() || pathValue;
    const folder = formatFolderLabel(parts) || "internal";

    if (!map.has(folder)) {
      map.set(folder, []);
    }

    map.get(folder).push({
      ...document,
      fileName,
    });
  }

  return [...map.entries()]
    .sort((left, right) => left[0].localeCompare(right[0], "vi"))
    .map(([folder, items]) => ({
      folder,
      items: items.sort((left, right) => left.title.localeCompare(right.title, "vi")),
    }));
}

function groupSourcesByAccessLevel(sources) {
  const sourceMap = new Map([
    ["basic", []],
    ["advanced", []],
  ]);

  for (const source of sources) {
    const level = normalizeSourceAccessLevel(source);
    if (!sourceMap.has(level)) {
      sourceMap.set(level, []);
    }
    sourceMap.get(level).push(source);
  }

  const sections = SOURCE_SECTION_TEMPLATE.map((section) => ({
    level: section.level,
    items: [...(sourceMap.get(section.level) || [])].sort((left, right) =>
      String(left.title || "").localeCompare(String(right.title || ""), "vi"),
    ),
  }));

  if ((sourceMap.get("sensitive") || []).length > 0) {
    sections.push({
      level: "sensitive",
      items: [...sourceMap.get("sensitive")].sort((left, right) =>
        String(left.title || "").localeCompare(String(right.title || ""), "vi"),
      ),
    });
  }

  return sections;
}

function normalizeSourceAccessLevel(source) {
  const level = String(source?.access_level || "basic").toLowerCase();
  if (level === "advanced") {
    return "advanced";
  }
  if (level === "sensitive") {
    return "sensitive";
  }
  return "basic";
}

function groupDocumentsByAccessLevel(documents, folders = []) {
  const grouped = groupDocumentsByFolder(documents);
  const groupedMap = new Map(grouped.map((group) => [group.folder, group]));

  return DOCUMENT_SECTION_TEMPLATE.map((section) => {
    const explicitFolders = folders
      .filter((folder) => String(folder.level || "").toLowerCase() === section.level)
      .map((folder) => folder.name)
      .filter(Boolean);
    const extraFolders = grouped
      .filter(
        (group) =>
          normalizeAccessLevel(group.items) === section.level &&
          !section.folders.includes(group.folder),
      )
      .map((group) => group.folder)
      .sort((left, right) => left.localeCompare(right, "vi"));
    const folderNames = [...new Set([...section.folders, ...explicitFolders, ...extraFolders])];
    const groups = folderNames.map((folder) => {
      const existing = groupedMap.get(folder);
      if (existing) {
        return existing;
      }

      return {
        folder,
        items: [],
      };
    });

    return {
      level: section.level,
      groups,
      count: countDocuments(groups),
    };
  });
}

function updateDynamicFolderOptions(documents, folders = []) {
  dynamicFolderOptions.basic = [];
  dynamicFolderOptions.advanced = [];
  customDocumentFolders.basic = [];
  customDocumentFolders.advanced = [];

  for (const folder of folders) {
    const level = String(folder?.level || "").toLowerCase();
    const name = String(folder?.name || "").trim();
    if (!["basic", "advanced"].includes(level) || !name) {
      continue;
    }
    if (!customDocumentFolders[level].includes(name)) {
      customDocumentFolders[level].push(name);
    }
  }

  const grouped = groupDocumentsByFolder(documents);
  for (const group of grouped) {
    const level = normalizeAccessLevel(group.items);
    if (level === "sensitive") {
      continue;
    }

    const defaults = DOCUMENT_SECTION_TEMPLATE.find((section) => section.level === level)?.folders || [];
    if (!defaults.includes(group.folder) && !customDocumentFolders[level].includes(group.folder)) {
      customDocumentFolders[level].push(group.folder);
    }
  }

  dynamicFolderOptions.basic.push(...customDocumentFolders.basic);
  dynamicFolderOptions.advanced.push(...customDocumentFolders.advanced);
}

function canManageCustomFolder(level, folderName) {
  if (!currentUser?.policy?.can_manage_permissions) {
    return false;
  }

  if (!["basic", "advanced"].includes(String(level || "").toLowerCase())) {
    return false;
  }

  const defaults = DOCUMENT_SECTION_TEMPLATE.find((section) => section.level === level)?.folders || [];
  return !defaults.includes(folderName);
}

function canCreateFolderInSection(level) {
  return Boolean(
    canCreateManagedContent() &&
      ["basic", "advanced"].includes(String(level || "").toLowerCase()),
  );
}

function canCreateManagedContent() {
  const role = String(currentUser?.role || "").trim().toLowerCase();
  return role === "admin" || role === "director";
}

function canCreateOrder() {
  const role = String(currentUser?.role || "").trim().toLowerCase();
  const department = String(currentUser?.department || "").trim().toLowerCase();
  return (
    role === "admin" ||
    role === "director" ||
    role === "manager" ||
    department === "sales"
  );
}

function canChooseSalesAssignee() {
  const role = String(currentUser?.role || "").trim().toLowerCase();
  return role === "admin" || role === "director" || role === "manager";
}

function canViewOrders() {
  if (!currentUser) {
    return false;
  }
  const role = String(currentUser?.role || "").trim().toLowerCase();
  const department = String(currentUser?.department || "").trim().toLowerCase();
  return (
    role === "admin" ||
    role === "director" ||
    role === "manager" ||
    department === "sales" ||
    department === "production" ||
    ["operations", "transport", "logistics", "delivery"].includes(department)
  );
}

function canManageOrders() {
  const role = String(currentUser?.role || "").trim().toLowerCase();
  return role === "admin";
}

function canClaimProductionOrder(order = activeOrderDetails) {
  if (!currentUser || !order || !isProductionOrder(order)) {
    return false;
  }
  const department = String(currentUser?.department || "").trim().toLowerCase();
  if (department === "sales") {
    return false;
  }
  return true;
}

function parseProductionClaimedByNames(value) {
  return String(value || "")
    .split(",")
    .map((item) => String(item || "").trim())
    .filter(Boolean);
}

function parseProductionClaims(value) {
  try {
    const parsed = JSON.parse(String(value || "[]"));
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function parseProductionCompletions(value) {
  try {
    const parsed = JSON.parse(String(value || "[]"));
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function getClaimedProductionLineNumbers(order) {
  const claimedLineNumbers = new Set();
  parseProductionClaims(order?.production_claims_json).forEach((claim) => {
    (Array.isArray(claim?.items) ? claim.items : []).forEach((item, index) => {
      const lineNumber = Number(item?.line_number || index + 1);
      if (Number.isFinite(lineNumber) && lineNumber > 0) {
        claimedLineNumbers.add(Math.trunc(lineNumber));
      }
    });
  });
  return claimedLineNumbers;
}

function getRemainingProductionLineNumbers(order) {
  const totalItems = parseProductionOrderItemsSummary(order?.order_items || "").filter((item) => item?.name || item?.code);
  const claimedLineNumbers = getClaimedProductionLineNumbers(order);
  return totalItems
    .map((item) => Number(item.index || 0))
    .filter((lineNumber) => Number.isFinite(lineNumber) && lineNumber > 0 && !claimedLineNumbers.has(lineNumber));
}

function getProductionClaimOwnersByLine(order) {
  const ownersByLine = new Map();
  parseProductionClaims(order?.production_claims_json).forEach((claim) => {
    const userName = String(claim?.user_name || "").trim() || "Nhân viên sản xuất";
    (Array.isArray(claim?.items) ? claim.items : []).forEach((item, index) => {
      const lineNumber = Number(item?.line_number || index + 1);
      if (!Number.isFinite(lineNumber) || lineNumber <= 0) {
        return;
      }
      const normalizedLine = Math.trunc(lineNumber);
      const currentOwners = ownersByLine.get(normalizedLine) || [];
      if (!currentOwners.includes(userName)) {
        currentOwners.push(userName);
      }
      ownersByLine.set(normalizedLine, currentOwners);
    });
  });
  return ownersByLine;
}

function getProductionClaimUserIdsByLine(order) {
  const userIdsByLine = new Map();
  parseProductionClaims(order?.production_claims_json).forEach((claim) => {
    const userId = String(claim?.user_id || "").trim();
    (Array.isArray(claim?.items) ? claim.items : []).forEach((item, index) => {
      const lineNumber = Number(item?.line_number || index + 1);
      if (!Number.isFinite(lineNumber) || lineNumber <= 0) {
        return;
      }
      const normalizedLine = Math.trunc(lineNumber);
      const currentUserIds = userIdsByLine.get(normalizedLine) || [];
      if (userId && !currentUserIds.includes(userId)) {
        currentUserIds.push(userId);
      }
      userIdsByLine.set(normalizedLine, currentUserIds);
    });
  });
  return userIdsByLine;
}

function updateProductionProgressFieldAccess() {
  const activeProductionOrder = activeOrderDetails || activeOrderDraftEditRecord;
  const showCompletedAdminEditButton = shouldShowCompletedProductionAdminEditButton(activeProductionOrder);
  Array.from(productionOrderItemsList?.querySelectorAll("[data-production-progress-field]") || []).forEach((field) => {
    const row = field.closest(".production-order-row");
    const canEdit = Boolean(
      isOrderModalReadOnly &&
      row?.dataset.claimEditable === "true" &&
      (!isProductionOrderLockedForEditing(activeProductionOrder) || (isCurrentUserAdmin() && !showCompletedAdminEditButton)),
    );
    field.disabled = false;
    if (canEdit) {
      field.readOnly = false;
      field.removeAttribute("readonly");
      field.classList.remove("production-progress-readonly");
    } else {
      field.readOnly = true;
      field.setAttribute("readonly", "readonly");
      field.classList.add("production-progress-readonly");
    }
  });
}

function parseProductionNumber(value) {
  const normalized = Number(String(value || "").replace(",", "."));
  return Number.isFinite(normalized) ? normalized : 0;
}

function formatProductionNumber(value) {
  if (!Number.isFinite(value)) {
    return "0";
  }
  if (Math.abs(value - Math.round(value)) < 0.000001) {
    return String(Math.round(value));
  }
  return String(value);
}

function updateProductionProgressDerivedFields(row) {
  if (!row) {
    return;
  }
  const quantity = Math.max(0, parseProductionNumber(row.querySelector('[data-field="quantity"]')?.value || 0));
  const done = Math.max(0, parseProductionNumber(row.querySelector('[data-field="done"]')?.value || 0));
  const missingField = row.querySelector('[data-field="missing"]');
  const extraField = row.querySelector('[data-field="extra"]');
  const missing = Math.max(0, quantity - done);
  const extra = Math.max(0, done - quantity);
  if (missingField) {
    missingField.value = formatProductionNumber(missing);
  }
  if (extraField) {
    extraField.value = formatProductionNumber(extra);
  }
}

function collectPendingProductionProgressRows() {
  return Array.from(productionOrderItemsList?.querySelectorAll(".production-order-row[data-progress-dirty='true']") || []).filter(
    (row) => row.dataset.claimEditable === "true",
  );
}

function updateProductionProgressButtons() {
  if (activeOrderCreateKind !== "production" || !isOrderModalReadOnly) {
    confirmProductionProgressButton?.classList.add("hidden");
    completeProductionProgressButton?.classList.add("hidden");
    if (confirmProductionProgressButton) {
      confirmProductionProgressButton.disabled = true;
    }
    if (completeProductionProgressButton) {
      completeProductionProgressButton.disabled = true;
    }
    productionCompleteChip?.classList.add("hidden");
    return;
  }
  if (shouldShowCompletedProductionAdminEditButton(activeOrderDetails || activeOrderDraftEditRecord)) {
    confirmProductionProgressButton?.classList.add("hidden");
    completeProductionProgressButton?.classList.add("hidden");
    if (confirmProductionProgressButton) {
      confirmProductionProgressButton.disabled = true;
    }
    if (completeProductionProgressButton) {
      completeProductionProgressButton.disabled = true;
    }
    return;
  }
  const dirtyRows = collectPendingProductionProgressRows();
  const hasDirtyRows = dirtyRows.length > 0;
  const hasMissingRequiredField =
    hasDirtyRows &&
    dirtyRows.some((row) => {
      const doneValue = String(row.querySelector('[data-field="done"]')?.value || "").trim();
      const teamValue = String(row.querySelector('[data-field="team"]')?.value || "").trim();
      return !doneValue || !teamValue;
    });
  const allComplete =
    hasDirtyRows &&
    !hasMissingRequiredField &&
    dirtyRows.every((row) => {
      const quantity = Math.max(0, parseProductionNumber(row.querySelector('[data-field="quantity"]')?.value || 0));
      const done = Math.max(0, parseProductionNumber(row.querySelector('[data-field="done"]')?.value || 0));
      return done >= quantity;
    });
  confirmProductionProgressButton?.classList.toggle("hidden", !hasDirtyRows || allComplete);
  completeProductionProgressButton?.classList.toggle("hidden", !hasDirtyRows || !allComplete);
  if (confirmProductionProgressButton) {
    confirmProductionProgressButton.disabled = hasMissingRequiredField;
  }
  if (completeProductionProgressButton) {
    completeProductionProgressButton.disabled = hasMissingRequiredField;
  }
  if (hasDirtyRows) {
    productionCompleteChip?.classList.add("hidden");
  }
  updateProductionPackagingButton();
  updateProductionReceiptButton();
}

function getProductionClaimStatus(order) {
  const productionItems = parseProductionOrderItemsSummary(order?.order_items || "").filter((item) => item?.name || item?.code);
  const claims = parseProductionClaims(order?.production_claims_json);
  const claimedLineNumbers = getClaimedProductionLineNumbers(order);
  const claimedByNames = new Set(parseProductionClaimedByNames(order?.production_claimed_by_names));
  let latestClaimAt = "";
  let latestClaimTime = 0;

  claims.forEach((claim) => {
    const userName = String(claim?.user_name || "").trim();
    if (userName) {
      claimedByNames.add(userName);
    }
    const claimedAt = String(claim?.claimed_at || "").trim();
    const claimedTime = Date.parse(claimedAt);
    if (claimedAt && Number.isFinite(claimedTime) && claimedTime >= latestClaimTime) {
      latestClaimAt = claimedAt;
      latestClaimTime = claimedTime;
    }
  });

  const totalLines = productionItems.length;
  const claimedLines = claimedLineNumbers.size;
  let status = "unclaimed";
  if (claimedLines > 0 && totalLines > 0 && claimedLines >= totalLines) {
    status = "full";
  } else if (claimedLines > 0 || claimedByNames.size > 0) {
    status = "partial";
  }

  return {
    status,
    totalLines,
    claimedLines,
    claimedByNames: Array.from(claimedByNames),
    latestClaimAt,
  };
}

function labelProductionClaimStatus(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "full") {
    return "Nhận toàn bộ";
  }
  if (normalized === "partial") {
    return "Nhận một phần";
  }
  return "Chưa nhận";
}

function getProductionClaimBadgeClass(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "full") {
    return "approved";
  }
  if (normalized === "partial") {
    return "pending";
  }
  return "paused";
}

function summarizeProductionAssignees(names = []) {
  const normalized = Array.isArray(names) ? names.filter(Boolean) : [];
  if (!normalized.length) {
    return "Chưa có";
  }
  if (normalized.length === 1) {
    return normalized[0];
  }
  return `${normalized[0]} +${normalized.length - 1}`;
}

function buildProductionItemPreview(order) {
  const items = parseProductionOrderItemsSummary(order?.order_items || "").filter((item) => item?.name || item?.code);
  const previewLines = items.slice(0, 2).map((item) => `${item.index}. ${item.name || item.code || ""}`);
  return {
    totalItems: items.length,
    previewText: previewLines.join("\n"),
    fullText: String(order?.order_items || "").trim(),
  };
}

function buildProductionQuantitySummary(order) {
  const items = parseProductionOrderItemsSummary(order?.order_items || "").filter((item) => item?.name || item?.code);
  const totalQuantity = items.reduce((sum, item) => {
    const quantity = Number.parseFloat(String(item?.quantity || "").replace(",", "."));
    return Number.isFinite(quantity) ? sum + quantity : sum;
  }, 0);
  const units = [...new Set(items.map((item) => String(item?.unit || "").trim()).filter(Boolean))];
  const unitLabel = units.length === 1 ? units[0] : "";
  const quantityLabel = Number.isFinite(totalQuantity) && totalQuantity > 0 ? `${totalQuantity}${unitLabel ? ` ${unitLabel}` : ""}` : "";
  return {
    totalItems: items.length,
    quantityLabel,
  };
}

function formatProductionLineClaimText(item) {
  if (!item) {
    return "";
  }
  const quantity = String(item.quantity || "").trim();
  const unit = String(item.unit || "").trim();
  const name = String(item.name || item.code || "").trim();
  const numericLine = Number(item.lineNumber || item.index || 0);
  const lineLabel = Number.isFinite(numericLine) && numericLine > 0 ? `Dòng ${Math.trunc(numericLine)}` : "";
  const content = [quantity, unit, name].filter(Boolean).join(" ");
  return [lineLabel, content].filter(Boolean).join(": ");
}

function buildProductionClaimBreakdown(order) {
  const allItems = parseProductionOrderItemsSummary(order?.order_items || "").filter((item) => item?.name || item?.code);
  const claims = parseProductionClaims(order?.production_claims_json);
  const claimedLineNumbers = new Set();
  const groupedClaims = new Map();

  claims.forEach((claim) => {
    const userName = String(claim?.user_name || "").trim() || "Nhân viên sản xuất";
    const existingItems = groupedClaims.get(userName) || new Map();
    (Array.isArray(claim?.items) ? claim.items : []).forEach((item, index) => {
      const lineNumber = Math.max(1, Number(item?.line_number || index + 1));
      if (!Number.isFinite(lineNumber)) {
        return;
      }
      claimedLineNumbers.add(Math.trunc(lineNumber));
      existingItems.set(Math.trunc(lineNumber), {
        lineNumber: Math.trunc(lineNumber),
        code: String(item?.code || "").trim(),
        name: String(item?.name || "").trim(),
        quantity: String(item?.quantity || "").trim(),
        unit: String(item?.unit || "").trim(),
      });
    });
    groupedClaims.set(userName, existingItems);
  });

  const assigneeDetails = Array.from(groupedClaims.entries()).map(([userName, itemsMap]) => {
    const items = Array.from(itemsMap.values())
      .sort((left, right) => left.lineNumber - right.lineNumber)
      .map((item) => formatProductionLineClaimText(item))
      .filter(Boolean);
    return {
      userName,
      text: items.join(", "),
    };
  });

  const remainingItems = allItems
    .filter((item) => !claimedLineNumbers.has(Number(item.index || 0)))
    .map((item) =>
      formatProductionLineClaimText({
        lineNumber: item.index,
        code: item.code,
        name: item.name,
        quantity: item.quantity,
        unit: item.unit,
      }),
    )
    .filter(Boolean);

  return {
    assigneeDetails,
    remainingText: remainingItems.join(", "),
  };
}

function renderProductionStatusSummary(statusInfo) {
  const totalLines = Math.max(0, Number(statusInfo?.totalLines || 0));
  const claimedLines = Math.max(0, Number(statusInfo?.claimedLines || 0));
  if (!totalLines) {
    return "Chưa có dòng sản xuất";
  }
  if (statusInfo?.status === "full") {
    return `Đã nhận toàn bộ ${claimedLines}/${totalLines} dòng`;
  }
  if (statusInfo?.status === "partial") {
    return `Đã nhận ${claimedLines}/${totalLines} dòng`;
  }
  return `Đã nhận 0/${totalLines} dòng`;
}

function getProductionClaimProgress(statusInfo) {
  const totalLines = Math.max(0, Number(statusInfo?.totalLines || 0));
  const claimedLines = Math.max(0, Number(statusInfo?.claimedLines || 0));
  if (!totalLines) {
    return 0;
  }
  return Math.max(0, Math.min(100, (claimedLines / totalLines) * 100));
}

function syncOrdersBoardUI() {
  const isProductionBoard = activeOrdersBoardKind === "production";
  if (ordersBoardTitle) {
    ordersBoardTitle.textContent = isProductionBoard ? "Theo dõi đơn sản xuất" : "Theo dõi đơn vận chuyển";
  }
  if (ordersAssigneeHeading) {
    ordersAssigneeHeading.textContent = "NV giao";
    ordersAssigneeHeading.classList.toggle("hidden", isProductionBoard);
  }
  if (ordersProductionStatusHeading) {
    ordersProductionStatusHeading.classList.toggle("hidden", !isProductionBoard);
  }
  ordersDeliveryHeading?.classList.toggle("hidden", isProductionBoard);
  ordersPaymentHeading?.classList.toggle("hidden", isProductionBoard);
  if (ordersStatusFilterLabel) {
    ordersStatusFilterLabel.textContent = isProductionBoard ? "Trạng thái SX" : "Trạng thái giao";
  }
  if (ordersSearch) {
    ordersSearch.placeholder = isProductionBoard ? "Mã đơn, khách hàng, NVKD, NVSX" : "Mã đơn, khách hàng, NVKD, NVGH";
  }
  if (ordersStatusFilter) {
    const nextMarkup = isProductionBoard
      ? `
        <option value="">Tất cả</option>
        <option value="unclaimed">Chưa nhận</option>
        <option value="partial">Nhận một phần</option>
        <option value="full">Nhận toàn bộ</option>
      `
      : `
        <option value="">Tất cả</option>
        <option value="assigned">Chưa giao</option>
        <option value="completed">Đã giao</option>
      `;
    if (ordersStatusFilter.dataset.mode !== (isProductionBoard ? "production" : "transport")) {
      ordersStatusFilter.innerHTML = nextMarkup;
      ordersStatusFilter.value = "";
      ordersStatusFilter.dataset.mode = isProductionBoard ? "production" : "transport";
    }
  }
  ordersPaymentFilterField?.classList.toggle("hidden", isProductionBoard);
  if (isProductionBoard && ordersPaymentFilter) {
    ordersPaymentFilter.value = "";
  }
  ordersToolbar?.classList.toggle("production-mode", isProductionBoard);
  ordersToolbar?.classList.toggle("transport-mode", !isProductionBoard);
  ordersTable?.classList.toggle("production-board", isProductionBoard);
  ordersTable?.classList.toggle("transport-board", !isProductionBoard);
}

function syncProductionClaimRowStates(order = activeOrderDetails) {
  const claimedLineNumbers = getClaimedProductionLineNumbers(order);
  const claimOwnersByLine = getProductionClaimOwnersByLine(order);
  const claimUserIdsByLine = getProductionClaimUserIdsByLine(order);
  const parsedItems = parseProductionOrderItemsSummary(order?.order_items || "");
  const currentUserId = String(currentUser?.id || "").trim();
  Array.from(productionOrderItemsList?.querySelectorAll("[data-order-item-row]") || []).forEach((row, index) => {
    const lineNumber = index + 1;
    const isClaimed = claimedLineNumbers.has(lineNumber);
    const ownerNames = claimOwnersByLine.get(lineNumber) || [];
    const ownerUserIds = claimUserIdsByLine.get(lineNumber) || [];
    const parsedItem = parsedItems[index] || null;
    const quantity = Math.max(0, parseProductionNumber(parsedItem?.quantity || row.querySelector('[data-field="quantity"]')?.value || 0));
    const done = Math.max(0, parseProductionNumber(parsedItem?.done || row.querySelector('[data-field="done"]')?.value || 0));
    const isCompleted = isClaimed && quantity > 0 && done >= quantity;
    row.dataset.claimed = isClaimed ? "true" : "false";
    row.dataset.lineNumber = String(lineNumber);
    row.dataset.claimedBy = ownerNames.join(", ");
    row.dataset.claimEditable = isClaimed && currentUserId && ownerUserIds.includes(currentUserId) ? "true" : "false";
    row.dataset.claimCompleted = isCompleted ? "true" : "false";
    if (row.dataset.claimEditable !== "true") {
      row.dataset.progressDirty = "false";
    }
    if (isClaimed) {
      row.dataset.claimSelected = "false";
    }
    const checkbox = row.querySelector("[data-production-claim-checkbox]");
    if (checkbox) {
      checkbox.checked = false;
      checkbox.disabled = isClaimed || activeProductionClaimMode !== "partial";
      checkbox.setAttribute("aria-hidden", isClaimed ? "true" : "false");
    }
    const badge = row.querySelector("[data-production-claimed-badge]");
    if (badge) {
      if (isClaimed && ownerNames.length) {
        badge.innerHTML = isCompleted
          ? `
          <span class="production-order-claimed-label">Đã nhận: ${escapeHtml(ownerNames.join(", "))}</span>
          <span class="production-order-claimed-status production-order-claimed-status-complete">✓ Hoàn thành</span>
        `
          : `
          <span class="production-order-claimed-label">Đã nhận: ${escapeHtml(ownerNames.join(", "))}</span>
          <span class="production-order-claimed-status">
            đang sản xuất<span class="production-order-claimed-dots" aria-hidden="true"><span>.</span><span>.</span><span>.</span></span>
          </span>
        `;
        badge.classList.remove("hidden");
      } else {
        badge.innerHTML = "";
        badge.classList.add("hidden");
      }
    }
  });
  updateProductionProgressFieldAccess();
  updateProductionProgressButtons();
}

function updateProductionClaimButtons(order = activeOrderDetails) {
  const canClaim = isOrderModalReadOnly && activeOrderCreateKind === "production" && canClaimProductionOrder(order);
  const totalLineCount = parseProductionOrderItemsSummary(order?.order_items || "").filter((item) => item?.name || item?.code).length;
  const remainingLineNumbers = canClaim ? getRemainingProductionLineNumbers(order) : [];
  const remainingCount = remainingLineNumbers.length;
  const canClaimRemaining = canClaim && remainingCount > 0;
  receiveAllProductionOrderButton?.classList.toggle("hidden", !canClaimRemaining);
  receivePartialProductionOrderButton?.classList.toggle("hidden", !canClaimRemaining || remainingCount <= 1);
  confirmPartialProductionOrderButton?.classList.toggle("hidden", activeProductionClaimMode !== "partial");
  if (receiveAllProductionOrderButton) {
    receiveAllProductionOrderButton.textContent = remainingCount > 0 && remainingCount < totalLineCount ? "Nhận tất cả phần còn lại" : "Nhận Tất Cả Đơn";
  }
  if (receivePartialProductionOrderButton) {
    receivePartialProductionOrderButton.textContent = "Nhận 1 Phần Đơn";
  }
}

function resetProductionClaimMode() {
  activeProductionClaimMode = "";
  orderModal?.classList.remove("production-claim-partial-mode");
  Array.from(productionOrderItemsList?.querySelectorAll(".production-order-row") || []).forEach((row) => {
    row.dataset.claimSelected = "false";
    const input = row.querySelector("[data-production-claim-checkbox]");
    if (input) {
      input.checked = false;
    }
  });
  syncProductionClaimRowStates(activeOrderDetails);
  updateProductionClaimButtons();
  updateProductionProgressButtons();
  updateProductionPackagingButton();
  updateProductionReceiptButton();
}

function setProductionClaimMode(mode = "") {
  activeProductionClaimMode = mode === "partial" ? "partial" : "";
  orderModal?.classList.toggle("production-claim-partial-mode", activeProductionClaimMode === "partial");
  Array.from(productionOrderItemsList?.querySelectorAll("[data-production-claim-checkbox]") || []).forEach((input) => {
    const row = input.closest(".production-order-row");
    const isClaimed = row?.dataset.claimed === "true";
    const canUseCheckbox = activeProductionClaimMode === "partial";
    if (canUseCheckbox && !isClaimed) {
      input.disabled = false;
      input.removeAttribute("disabled");
    } else {
      input.disabled = true;
      input.setAttribute("disabled", "disabled");
    }
  });
  if (activeProductionClaimMode !== "partial") {
    Array.from(productionOrderItemsList?.querySelectorAll(".production-order-row") || []).forEach((row) => {
      row.dataset.claimSelected = "false";
      const input = row.querySelector("[data-production-claim-checkbox]");
      if (input) {
        input.checked = false;
      }
    });
  }
  syncProductionClaimRowStates(activeOrderDetails);
  updateProductionClaimButtons();
  updateProductionPackagingButton();
  updateProductionReceiptButton();
}

function canEditOrder(order) {
  if (!currentUser || !order) {
    return false;
  }
  if (isCurrentUserAdmin()) {
    return true;
  }
  if (isProductionOrderLockedForEditing(order)) {
    return false;
  }
  return String(order?.created_by_user_id || "").trim() === String(currentUser?.id || "").trim();
}

function isOrderCreatedByCurrentUser(order) {
  if (!currentUser || !order) {
    return false;
  }
  return String(order?.created_by_user_id || "").trim() === String(currentUser?.id || "").trim();
}

function normalizeOrderIdValue(value) {
  const normalized = String(value || "")
    .trim()
    .replace(/^dh[-\s]*/i, "")
    .replace(/[^\dA-Za-z-]/g, "")
    .trim();
  return normalized ? `DH-${normalized.toUpperCase()}` : "";
}

function extractOrderSequenceValue(orderId) {
  const normalized = normalizeOrderIdValue(orderId);
  const match = normalized.match(/^DH-(\d+)$/i);
  return match ? Number.parseInt(match[1], 10) : Number.NaN;
}

function formatOrderSequenceValue(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric) || numeric < 0) {
    return "003000";
  }
  return String(Math.trunc(numeric)).padStart(6, "0");
}

function getNextProductionOrderSequence() {
  const numericValues = orders
    .map((item) => extractOrderSequenceValue(item?.order_id || ""))
    .filter((value) => Number.isFinite(value));
  const maxValue = numericValues.length ? Math.max(...numericValues) : 2999;
  return formatOrderSequenceValue(Math.max(3000, maxValue + 1));
}

function formatTransportOrderSequence(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric) || numeric < 0) {
    return "0000";
  }
  return String(Math.trunc(numeric)).padStart(4, "0");
}

function getNextTransportOrderSequence() {
  const numericValues = orders
    .filter((item) => !isProductionOrder(item))
    .map((item) => extractOrderSequenceValue(item?.order_id || ""))
    .filter((value) => Number.isFinite(value));
  const maxValue = numericValues.length ? Math.max(...numericValues) : -1;
  return formatTransportOrderSequence(Math.max(0, maxValue + 1));
}

function normalizePlannedDeliveryInput(value) {
  const raw = String(value || "").trim();
  if (!raw) {
    return "";
  }

  const slashMatch = raw.match(/^(\d{1,2})\s*\/\s*(\d{1,2})$/);
  if (slashMatch) {
    const day = Number(slashMatch[1]);
    const month = Number(slashMatch[2]);
    const year = new Date().getFullYear();
    if (day >= 1 && day <= 31 && month >= 1 && month <= 12) {
      return `${year}-${String(month).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
    }
  }

  const parsed = new Date(raw);
  if (Number.isFinite(parsed.getTime())) {
    return parsed.toISOString().slice(0, 10);
  }

  return raw;
}

function canCompleteDelivery() {
  const role = String(currentUser?.role || "").trim().toLowerCase();
  const department = String(currentUser?.department || "").trim().toLowerCase();
  return (
    role === "admin" ||
    role === "director" ||
    ["operations", "transport", "logistics", "delivery"].includes(department)
  );
}

function populateOrderSalesOptions() {
  if (!orderSalesUser) {
    return;
  }

  const allSalesUsers = [...allUsers]
    .filter((user) => String(user?.department || "").trim().toLowerCase() === "sales")
    .sort((left, right) => String(left.name || "").localeCompare(String(right.name || ""), "vi"));
  const salesUsers = canChooseSalesAssignee()
    ? allSalesUsers
    : allSalesUsers.filter((user) => String(user?.id || "") === String(currentUser?.id || ""));

  orderSalesUser.innerHTML = salesUsers.length
    ? salesUsers
        .map((user) => {
          const code = user.employee_code || user.username || user.id;
          const isCurrentSales = String(currentUser?.id || "") === String(user.id || "");
          return `<option value="${escapeHtml(user.id)}" ${isCurrentSales ? "selected" : ""}>${escapeHtml(`${user.name} • ${code}`)}</option>`;
        })
        .join("")
    : `<option value="">Chưa có nhân viên kinh doanh</option>`;
  orderSalesUser.disabled = !canChooseSalesAssignee();
}

function populateOrderDeliveryOptions() {
  if (!orderDeliveryUser) {
    return;
  }

  const deliveryUsers = [...allUsers]
    .filter((user) => {
      const department = String(user?.department || "").trim().toLowerCase();
      return ["operations", "transport", "logistics", "delivery"].includes(department);
    })
    .sort((left, right) => String(left.name || "").localeCompare(String(right.name || ""), "vi"));

  orderDeliveryUser.innerHTML = deliveryUsers.length
    ? deliveryUsers
        .map((user) => {
          const code = user.employee_code || user.username || user.id;
          return `<option value="${escapeHtml(user.id)}">${escapeHtml(`${user.name} • ${code}`)}</option>`;
        })
        .join("")
    : `<option value="">Chưa có nhân viên giao hàng</option>`;
}

async function loadOrders() {
  if (!canViewOrders()) {
    orders = [];
    renderOrdersBoard();
    return;
  }

  try {
    const payload = await fetchJson(ordersUrl);
    orders = payload.orders || [];
    populateOrdersDateFilters();
    renderOrdersBoard();
  } catch (error) {
    showToast(error.message || "Không thể tải danh sách đơn hàng.", "error");
  }
}

async function loadSalesInventory() {
  if (!canViewOrders()) {
    salesInventoryReceipts = [];
    salesInventoryTransfers = [];
    renderSalesStockPanel();
    return;
  }

  try {
    const payload = await fetchJson(salesInventoryUrl);
    salesInventoryReceipts = Array.isArray(payload?.receipts) ? payload.receipts : [];
    salesInventoryTransfers = Array.isArray(payload?.transfers) ? payload.transfers : [];
    renderSalesStockPanel();
  } catch (error) {
    showToast(error.message || "Không thể tải kho kinh doanh.", "error");
  }
}

async function loadOrderProducts() {
  try {
    const payload = await fetchJson(orderProductsUrl);
    orderProducts = getSortedOrderProducts(Array.isArray(payload.products) && payload.products.length ? payload.products : getDefaultOrderProducts());
  } catch {
    orderProducts = getSortedOrderProducts(getDefaultOrderProducts());
  }
}

function renderOrdersBoard() {
  if (!ordersTableBody) {
    return;
  }

  syncOrdersBoardUI();

  const keyword = String(ordersSearch?.value || "").trim().toLowerCase();
  const dateFromFilter = String(ordersDateFromFilter?.value || "").trim();
  const dateToFilter = String(ordersDateToFilter?.value || "").trim();
  const statusFilter = String(ordersStatusFilter?.value || "").trim().toLowerCase();
  const paymentFilter = String(ordersPaymentFilter?.value || "").trim().toLowerCase();

  const visibleOrders = orders.filter((item) => {
    const isProduction = isProductionOrder(item);
    const productionClaimStatus = isProduction ? getProductionClaimStatus(item) : null;
    if (activeOrdersBoardKind === "production" && !isProduction) {
      return false;
    }
    if (activeOrdersBoardKind === "transport" && isProduction) {
      return false;
    }

    const createdParts = extractOrderCreatedParts(item);
    const haystack = [
      item.order_id,
      item.customer_name,
      item.sales_user_name,
      item.delivery_user_name,
      item.production_claimed_by_names,
      item.delivery_address,
    ]
      .join(" ")
      .toLowerCase();
    if (keyword && !haystack.includes(keyword)) {
      return false;
    }
    if (dateFromFilter && (!createdParts.date || createdParts.date < dateFromFilter)) {
      return false;
    }
    if (dateToFilter && (!createdParts.date || createdParts.date > dateToFilter)) {
      return false;
    }
    if (statusFilter) {
      if (activeOrdersBoardKind === "production") {
        if (String(productionClaimStatus?.status || "").trim().toLowerCase() !== statusFilter) {
          return false;
        }
      } else if (String(item.status || "").trim().toLowerCase() !== statusFilter) {
        return false;
      }
    }
    if (activeOrdersBoardKind !== "production" && paymentFilter && String(item.payment_status || "").trim().toLowerCase() !== paymentFilter) {
      return false;
    }
    return true;
  });

  ordersTableBody.innerHTML = "";

  if (!visibleOrders.length) {
    const row = document.createElement("tr");
    row.innerHTML = `<td colspan="${activeOrdersBoardKind === "production" ? 5 : 6}"><span class="empty-copy">Không có đơn hàng phù hợp.</span></td>`;
    ordersTableBody.append(row);
    return;
  }

  for (const item of visibleOrders) {
    const isProduction = isProductionOrder(item);
    const productionClaimStatus = isProduction ? getProductionClaimStatus(item) : null;
    const productionAssigneeNames = productionClaimStatus?.claimedByNames || parseProductionClaimedByNames(item.production_claimed_by_names);
    const productionAssignees = productionAssigneeNames.join(", ");
    const productionAssigneeSummary = summarizeProductionAssignees(productionAssigneeNames);
    const productionItemPreview = isProduction ? buildProductionItemPreview(item) : { totalItems: 0, previewText: "", fullText: "" };
    const productionQuantitySummary = isProduction ? buildProductionQuantitySummary(item) : { totalItems: 0, quantityLabel: "" };
    const productionClaimBreakdown = isProduction ? buildProductionClaimBreakdown(item) : { assigneeDetails: [], remainingText: "" };
    const badgeClass = isProduction
      ? getProductionClaimBadgeClass(productionClaimStatus?.status)
      : String(item.status || "") === "completed"
        ? "approved"
        : "paused";
    const badgeLabel = isProduction ? labelProductionClaimStatus(productionClaimStatus?.status) : labelOrderStatus(item.status);
    const productionSummary = activeOrdersBoardKind === "production" && productionClaimStatus ? renderProductionStatusSummary(productionClaimStatus) : "";
    const progressValue = isProduction ? getProductionClaimProgress(productionClaimStatus) : 0;
    const latestClaimText =
      activeOrdersBoardKind === "production" && productionClaimStatus?.latestClaimAt
        ? `Nhận gần nhất: ${formatNotificationTime(productionClaimStatus.latestClaimAt)}`
        : "";
    const orderFacts = isProduction ? [productionQuantitySummary.totalItems ? `${productionQuantitySummary.totalItems} mặt hàng` : ""].filter(Boolean).join("") : "";
    const row = document.createElement("tr");
    row.classList.add("orders-table-row");
    if (activeOrdersBoardKind === "production") {
      row.classList.add(`production-status-${productionClaimStatus?.status || "unclaimed"}`);
    }
    row.innerHTML = `
      <td>
        <div class="table-title-row">
          <strong>${escapeHtml(item.order_id || "-")}</strong>
          ${activeOrdersBoardKind === "production" ? "" : `<span class="record-kind-badge ${escapeHtml(badgeClass)}">${escapeHtml(badgeLabel)}</span>`}
        </div>
        <p class="doc-leaf-file">Tên khách hàng: ${escapeHtml(item.customer_name || "-")}</p>
        ${orderFacts ? `<p class="doc-meta order-facts">${escapeHtml(orderFacts)}</p>` : ""}
        ${
          activeOrdersBoardKind === "production" && productionItemPreview.fullText
            ? `<div class="order-items-block">
                <p class="doc-meta order-items-preview is-collapsed">${escapeHtml(productionItemPreview.previewText || productionItemPreview.fullText)}</p>
                ${
                  productionItemPreview.totalItems > 2
                    ? `<button type="button" class="leaf-action-button order-items-toggle-button" data-order-items-toggle data-collapsed-text="Xem thêm" data-expanded-text="Thu gọn">Xem thêm</button>`
                    : ""
                }
                <pre class="order-items-full hidden">${escapeHtml(productionItemPreview.fullText)}</pre>
              </div>`
            : item.order_items
              ? `<p class="doc-meta">${escapeHtml(item.order_items)}</p>`
              : ""
        }
        ${item.order_value ? `<p class="doc-meta">Giá trị: ${escapeHtml(formatOrderValue(item.order_value))}</p>` : ""}
        <p class="doc-meta">${escapeHtml(item.delivery_address || item.note || "-")}</p>
        ${
          canEditOrder(item) || canManageOrders()
            ? `<div class="table-actions">
                ${canEditOrder(item) ? '<button type="button" class="leaf-action-button warning order-edit-button">Sửa đơn</button>' : ""}
                ${canManageOrders() ? '<button type="button" class="leaf-action-button danger order-delete-button">Xóa đơn</button>' : ""}
              </div>`
            : ""
        }
      </td>
      <td>${escapeHtml(item.sales_user_name || "-")}</td>
      ${
        activeOrdersBoardKind === "production"
          ? ""
          : `<td title="${escapeHtml(item.delivery_user_name || "-")}">${escapeHtml(item.delivery_user_name || "-")}</td>`
      }
      ${
        activeOrdersBoardKind === "production"
          ? `<td class="orders-production-status-cell">
              <div class="orders-production-status-stack">
                <span class="record-kind-badge ${escapeHtml(badgeClass)}">${escapeHtml(badgeLabel)}</span>
                <span class="doc-meta">${escapeHtml(productionSummary)}</span>
                <div class="orders-progress" aria-hidden="true">
                  <span class="orders-progress-bar" style="width:${escapeHtml(String(progressValue.toFixed(2)))}%"></span>
                </div>
                <span class="doc-meta">${escapeHtml(latestClaimText || "Chưa có thời điểm nhận")}</span>
              </div>
            </td>`
          : ""
      }
      ${
        activeOrdersBoardKind === "production"
          ? ""
          : `<td>${escapeHtml(item.completed_at ? formatNotificationTime(item.completed_at) : "Chưa giao")}</td>
      <td>${escapeHtml(labelPaymentStatus(item.payment_status || "unpaid"))}</td>`
      }
      <td>${escapeHtml(labelPaymentMethod(item.payment_method || ""))}</td>
    `;
    row.querySelector("[data-order-items-toggle]")?.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      const button = event.currentTarget;
      const block = button?.closest(".order-items-block");
      const preview = block?.querySelector(".order-items-preview");
      const full = block?.querySelector(".order-items-full");
      const expanded = !full?.classList.contains("hidden");
      full?.classList.toggle("hidden", expanded);
      preview?.classList.toggle("hidden", !expanded);
      button.textContent = expanded ? button.dataset.collapsedText || "Xem thêm" : button.dataset.expandedText || "Thu gọn";
    });
    row.addEventListener("click", async (event) => {
      event.preventDefault();
      event.stopPropagation();
      if (activeOrdersBoardKind === "production") {
        try {
          await openOrderPreviewModal(item);
          closeOrdersBoardModal();
        } catch (error) {
          showToast(error.message || "Không thể mở phiếu sản xuất.", "error");
        }
        return;
      }
      openOrderDetailsModal(item);
    });
    row.querySelector(".order-edit-button")?.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      openOrderEditModal(item);
    });
    row.querySelector(".order-delete-button")?.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      deleteOrder(item).catch((error) => {
        showToast(error.message || "Không thể xóa đơn hàng.", "error");
      });
    });
    ordersTableBody.append(row);
  }
}

function isProductionOrder(item) {
  const explicitKind = String(item?.order_kind || "").trim().toLowerCase();
  if (explicitKind === "production") {
    return true;
  }
  if (explicitKind === "transport") {
    return false;
  }

  const documentTitle = String(item?.document_title || "").trim().toLowerCase();
  if (documentTitle.includes("phieu de nghi san xuat") || documentTitle.includes("sản xuất")) {
    return true;
  }

  return !String(item?.delivery_user_id || "").trim();
}

function labelOrderStatus(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "completed") {
    return "Đã giao";
  }
  return "Chưa giao";
}

function labelPaymentStatus(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "paid") {
    return "Đã thanh toán";
  }
  return "Chưa thanh toán";
}

function labelPaymentMethod(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "cash") {
    return "Tiền mặt";
  }
  if (normalized === "bank_transfer") {
    return "Chuyển khoản";
  }
  return "Chưa rõ";
}

function formatOrderValue(value) {
  const amount = Number(String(value || "").replace(/[^\d.-]/g, ""));
  if (!Number.isFinite(amount) || amount <= 0) {
    return String(value || "").trim() || "0";
  }
  return new Intl.NumberFormat("vi-VN", {
    style: "currency",
    currency: "VND",
    maximumFractionDigits: 0,
  }).format(amount);
}

function populateDeliverySalesOptions() {
  if (!deliverySalesUser) {
    return;
  }

  const salesUsers = [...allUsers]
    .filter((user) => String(user?.department || "").trim().toLowerCase() === "sales")
    .sort((left, right) => String(left.name || "").localeCompare(String(right.name || ""), "vi"));

  deliverySalesUser.innerHTML = salesUsers.length
    ? salesUsers
        .map((user) => {
          const code = user.employee_code || user.username || user.id;
          return `<option value="${escapeHtml(user.id)}">${escapeHtml(`${user.name} • ${code}`)}</option>`;
        })
        .join("")
    : `<option value="">Chưa có nhân viên kinh doanh</option>`;
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

function formatNotificationTime(value) {
  const date = new Date(value);
  if (!Number.isFinite(date.getTime())) {
    return value || "-";
  }
  return new Intl.DateTimeFormat("vi-VN", {
    dateStyle: "medium",
    timeStyle: "short",
  }).format(date);
}

async function openNotification(item, options = {}) {
  if (!item) {
    return;
  }

  ensureNotificationUiStateLoaded();
  mergeNotificationArchiveItems([item]);

  try {
    const payload = await postJson(notificationsReadUrl, {
      notification_id: item.id,
    });
    notifications = payload.notifications || notifications;
    mergeNotificationArchiveItems(payload.notifications || [item]);
    renderNotifications();
  } catch {}

  if (!options.skipAutoHide) {
    scheduleNotificationAutoHide(item, 30000);
  }

  const orderId = String(item?.meta?.order_id || "").trim();
  if (orderId && canViewOrders()) {
    try {
      await loadOrders();
      const matchedOrder = orders.find((order) => String(order?.order_id || "").trim() === orderId);
      if (matchedOrder) {
        await openOrderPreviewModal(matchedOrder);
        return;
      }
    } catch {}
  }

  const documentPath = String(item?.meta?.document_path || "").trim();
  if (documentPath) {
    try {
      activeNotificationOnViewerClose = item;
      await viewDocument(documentPath);
      return;
    } catch (error) {
      activeNotificationOnViewerClose = null;
      showToast(error.message || "Không thể mở tài liệu từ thông báo.", "error");
      return;
    }
  }

  if (item.message) {
    showToast(item.message, "success");
  }
}

async function markAllNotificationsRead() {
  if (!notifications.some((item) => !String(item?.read_at || "").trim())) {
    showToast("Không còn thông báo chưa đọc.", "success");
    return;
  }

  try {
    const payload = await postJson(notificationsReadAllUrl, {});
    notifications = payload.notifications || notifications;
    mergeNotificationArchiveItems(payload.notifications || notifications);
    renderNotifications();
    showToast("Đã đánh dấu tất cả thông báo là đã đọc.", "success");
  } catch (error) {
    showToast(error.message || "Không thể cập nhật tất cả thông báo.", "error");
  }
}

async function deleteNotification(item, options = {}) {
  const notificationId = String(item?.id || "").trim();
  if (!notificationId) {
    return;
  }

  try {
    const payload = await postJson(notificationsDeleteUrl, {
      notification_id: notificationId,
    });
    notifications = payload.notifications || notifications;
    removeNotificationFromArchive(notificationId);
    hiddenNotificationIds.delete(notificationId);
    persistHiddenNotificationIds();
    const existingTimer = notificationAutoHideTimers.get(notificationId);
    if (existingTimer) {
      window.clearTimeout(existingTimer);
      notificationAutoHideTimers.delete(notificationId);
    }
    renderNotifications();
    if (!options.silent) {
      showToast("Đã xóa thông báo.", "success");
    }
  } catch (error) {
    showToast(error.message || "Không thể xóa thông báo.", "error");
  }
}

async function createDocumentFolder(level) {
  const folderName = window.prompt(`Tạo thư mục mới trong mục "${labelAccessLevel(level)}":`);
  if (folderName === null) {
    return;
  }

  const nextValue = folderName.trim();
  if (!nextValue) {
    return;
  }

  const payload = await postJson(createFolderUrl, {
    folder_name: nextValue,
    level,
  });
  applyWorkspacePayload(payload);
  showToast("Đã tạo thư mục mới.", "success");
}

async function renameDocumentFolder(folderName) {
  const nextFolderName = window.prompt(`Đổi tên thư mục "${folderName}" thành:`, folderName);
  if (nextFolderName === null) {
    return;
  }

  const nextValue = nextFolderName.trim();
  if (!nextValue || nextValue === folderName) {
    return;
  }

  const payload = await postJson(renameFolderUrl, {
    folder_name: folderName,
    next_folder_name: nextValue,
  });
  applyWorkspacePayload(payload);
  showToast("Đã đổi tên thư mục.", "success");
}

async function deleteDocumentFolder(folderName) {
  if (!window.confirm(`Xóa toàn bộ thư mục "${folderName}"?`)) {
    return;
  }

  const payload = await postJson(deleteFolderUrl, {
    folder_name: folderName,
  });
  applyWorkspacePayload(payload);
  showToast("Đã xóa thư mục.", "success");
}

function normalizeAccessLevel(items) {
  const level = String(items?.[0]?.metadata?.access_level || "basic").toLowerCase();
  if (level === "advanced") {
    return "advanced";
  }
  if (level === "sensitive") {
    return "sensitive";
  }
  return "basic";
}

function countDocuments(groups) {
  return groups.reduce((total, group) => total + group.items.length, 0);
}

function formatFolderLabel(parts) {
  if (!Array.isArray(parts) || parts.length === 0) {
    return "internal";
  }

  const internalIndex = parts.findIndex((part) => part === "internal");
  if (internalIndex >= 0) {
    const scopedParts = parts.slice(internalIndex + 1);
    const rootFolder = scopedParts[0] || "internal";
    if (rootFolder === "dữ liệu cá nhân") {
      return "dữ liệu cá nhân";
    }
    return rootFolder;
  }

  return parts[0] || "internal";
}

function getFileExtension(fileName) {
  const match = String(fileName || "").toLowerCase().match(/(\.[a-z0-9]+)$/);
  return match?.[1] || "";
}

function buildDocumentSyncStatus(documentItem) {
  const metadata = documentItem?.metadata || {};
  if (String(metadata.source_type || "").toLowerCase() !== "sheet") {
    return "";
  }

  if (metadata.synced_at) {
    return `Đã đồng bộ lúc ${metadata.synced_at}`;
  }

  return "Đang chờ đồng bộ";
}

function buildSourceSyncStatus(source) {
  if (source.last_error) {
    return `Lỗi đồng bộ: ${source.last_error}`;
  }

  if (source.syncable === false) {
    return "Nguồn đang tạm dừng đồng bộ.";
  }

  if (source.last_synced_at) {
    return `Đã đồng bộ lúc ${source.last_synced_at}`;
  }

  return "Đang chờ đồng bộ";
}

function buildSourceStatusBadge(source) {
  if (source.last_error) {
    return `<span class="record-kind-badge error">Lỗi</span>`;
  }
  if (source.syncable === false) {
    return `<span class="record-kind-badge paused">Tạm dừng</span>`;
  }
  return "";
}

function fileIconClass(extension) {
  if (extension === ".md") {
    return "icon-md";
  }

  if (extension === ".xlsx" || extension === ".csv") {
    return "icon-sheet";
  }

  if (extension === ".docx" || extension === ".txt") {
    return "icon-doc";
  }

  if (extension === ".pptx") {
    return "icon-slide";
  }

  if (extension === ".json") {
    return "icon-json";
  }

  return "icon-doc";
}

function fileIconSvg() {
  return `
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
      <path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"></path>
      <path d="M14 3v5h5"></path>
      <path d="M8 13h8"></path>
      <path d="M8 17h5"></path>
    </svg>
  `;
}

function normalizeDisplayLine(line) {
  const trimmed = line.trim();
  if (!trimmed) {
    return "";
  }

  const bulletLike = trimmed.match(/^\*{2,}\s*(.+)$/);
  if (bulletLike?.[1]) {
    return `• ${stripInlineMarkdown(bulletLike[1])}`;
  }

  return stripInlineMarkdown(trimmed);
}

function stripInlineMarkdown(text) {
  return String(text || "")
    .replace(/\*{2,}([^*]+)\*{2,}/g, "$1")
    .replace(/\*([^*]+)\*/g, "$1")
    .replace(/_{2,}([^_]+)_{2,}/g, "$1")
    .replace(/_([^_]+)_/g, "$1")
    .trim();
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const result = String(reader.result || "");
      const base64 = result.includes(",") ? result.split(",").pop() : "";
      if (!base64) {
        reject(new Error("Không đọc được file đã chọn."));
        return;
      }

      resolve(base64);
    };
    reader.onerror = () => reject(new Error("Không đọc được file đã chọn."));
    reader.readAsDataURL(file);
  });
}

function formatFileSize(size) {
  if (!Number.isFinite(size) || size <= 0) {
    return "";
  }

  if (size < 1024 * 1024) {
    return `${(size / 1024).toFixed(1)} KB`;
  }

  return `${(size / (1024 * 1024)).toFixed(2)} MB`;
}

function parseDuplicateSheetMessage(message) {
  const text = String(message || "");
  if (!text.includes("Link Google Sheet nay da duoc them truoc do.")) {
    return null;
  }

  const rows = text.split(/\n+/);
  const lookup = Object.fromEntries(
    rows
      .map((row) => {
        const index = row.indexOf(":");
        if (index === -1) {
          return null;
        }

        return [row.slice(0, index).trim(), row.slice(index + 1).trim()];
      })
      .filter(Boolean),
  );

  return {
    link: lookup.Link || "",
    title: lookup["Tieu de"] || "",
    path: lookup["Duong dan"] || "",
    addedAt: lookup["Thoi diem them"] || "",
  };
}

function labelRole(role) {
  const value = String(role || "").toLowerCase();
  if (value === "employee") return "Nhân viên";
  if (value === "manager") return "Quản lý";
  if (value === "director") return "Giám đốc";
  if (value === "admin") return "Quản trị viên";
  return role || "-";
}

function avatarIconForRole(role) {
  const value = String(role || "").toLowerCase();
  if (value === "director") {
    return `
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
        <path d="M12 4l2.1 4.3 4.7.7-3.4 3.3.8 4.7L12 14.8 7.8 17l.8-4.7L5.2 9l4.7-.7L12 4z"></path>
      </svg>
    `;
  }
  if (value === "admin") {
    return `
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
        <path d="M12 3l7 4v5c0 4.4-2.9 7.8-7 9-4.1-1.2-7-4.6-7-9V7l7-4z"></path>
        <path d="M9.5 12.3l1.7 1.7 3.3-4"></path>
      </svg>
    `;
  }
  if (value === "manager") {
    return `
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
        <rect x="4" y="6" width="16" height="12" rx="2"></rect>
        <path d="M9 6V4h6v2"></path>
        <path d="M4 11h16"></path>
      </svg>
    `;
  }
  if (value === "employee") {
    return `
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
        <path d="M18 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path>
        <circle cx="12" cy="7" r="4"></circle>
      </svg>
    `;
  }
  return `
    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
      <circle cx="12" cy="12" r="9"></circle>
      <path d="M12 8v4"></path>
      <path d="M12 16h.01"></path>
    </svg>
  `;
}


function labelDepartment(department) {
  const value = String(department || "").toLowerCase();
  if (value === "sales") return "Kinh doanh";
  if (value === "hr") return "Nhân sự";
  if (value === "finance") return "Tài chính";
  if (value === "operations" || value === "transport" || value === "logistics" || value === "delivery") return "Vận chuyển";
  if (value === "production") return "Sản xuất";
  if (value === "executive") return "Ban giám đốc";
  return department || "-";
}

function labelAccessLevel(level) {
  const value = String(level || "").toLowerCase();
  if (value === "basic") return "Cơ bản";
  if (value === "advanced") return "Nâng cao";
  if (value === "sensitive") return "Cá nhân";
  return level || "-";
}

function labelSectionAccessLevel(level) {
  return labelAccessLevel(level);
}

function labelPermissionAccessLevel(level) {
  const value = String(level || "").toLowerCase();
  if (value === "sensitive") {
    return "ALL";
  }
  return labelAccessLevel(level);
}

function labelDepartmentScope(scope) {
  const value = String(scope || "").toLowerCase();
  if (value === "all") return "Toàn bộ phòng ban";
  return "Chỉ phòng ban của người này";
}




