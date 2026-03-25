/**
 * 付款申請原型：主檔／明細頁籤、款項用途連動、付款名稱來自設定檔
 * 版本號請與 README.md 首段「版本」欄位同步維護。
 */
(function () {
  const APP_VERSION = "1.0.0";

  const STORAGE_KEY = "payment_vendor_config_v1";
  const APPLICATIONS_KEY = "payment_applications_v1";
  const PETTY_CASH_LIMIT_KEY = "payment_petty_cash_limit_v1";
  const DEFAULT_APPLICANT = "ruby.lee";
  const DEFAULT_PETTY_CASH_LIMIT = 5000;
  let pettyCashLimit = DEFAULT_PETTY_CASH_LIMIT;

  const PURPOSE_OPTIONS = [
    "進貨款項(PO單)",
    "進口相關費用(LCM)",
    "旅費(國內外出差費用)",
    "廣告費",
    "勞務費(KOL、律師、會計師)",
    "通路費用(通路後扣)",
    "URMART月結廠商",
    "倉庫相關費用",
    "其他費用",
  ];

  const VOUCHER_OPTIONS = [
    "發票",
    "收據",
    "合約",
    "報價單",
    "勞務報酬單",
    "簽呈(無憑證補貼款項)",
    "差旅費",
    "通路費用(通路後扣)",
  ];

  /** 款項用途 → 備註欄位名稱、說明、憑證樣式選項 */
  const PURPOSE_RULES = {
    "進貨款項(PO單)": {
      remarkLabel: "PO單",
      remarkPlaceholder: "請輸入 PO 單號",
      remarkHelp: "填寫採購單號",
      voucherStyles: ["發票"],
    },
    "進口相關費用(LCM)": {
      remarkLabel: "報關單號",
      remarkPlaceholder: "請輸入報關單號",
      remarkHelp: "填寫報關單號",
      voucherStyles: ["收據"],
    },
    "旅費(國內外出差費用)": {
      remarkLabel: "出差期間",
      remarkPlaceholder: "請輸入出差期間（起訖日期）",
      remarkHelp: "填寫起訖日期",
      voucherStyles: ["合約"],
      periodPicker: true,
    },
    廣告費: {
      remarkLabel: "結算月",
      remarkPlaceholder: "請輸入結算月（YYYY-MM）",
      remarkHelp: "填寫 YYYY-MM",
      voucherStyles: ["報價單"],
    },
    "勞務費(KOL、律師、會計師)": {
      remarkLabel: "勞務期間",
      remarkPlaceholder: "請輸入勞務期間",
      remarkHelp: "填寫服務期間",
      voucherStyles: ["勞務報酬單"],
      periodPicker: true,
    },
    "通路費用(通路後扣)": {
      remarkLabel: "結算月",
      remarkPlaceholder: "請輸入結算月（YYYY-MM）",
      remarkHelp: "填寫 YYYY-MM",
      voucherStyles: ["通路費用(通路後扣)", "簽呈(無憑證補貼款項)"],
    },
    "URMART月結廠商": {
      remarkLabel: "結算月",
      remarkPlaceholder: "請輸入結算月（YYYY-MM）",
      remarkHelp: "填寫 YYYY-MM",
      voucherStyles: ["差旅費"],
    },
    "倉庫相關費用": {
      remarkLabel: "結算月",
      remarkPlaceholder: "請輸入結算月（YYYY-MM）",
      remarkHelp: "填寫 YYYY-MM",
      voucherStyles: ["發票", "收據", "差旅費"],
    },
    其他費用: {
      remarkLabel: "使用期間",
      remarkPlaceholder: "請輸入使用期間",
      remarkHelp: "填寫期間",
      voucherStyles: ["發票", "收據", "合約", "報價單"],
      periodPicker: true,
    },
  };

  /** 憑證樣式 → 稅別規則（fixedUntax: 固定未稅；both: 可選） */
  const VOUCHER_TAX_RULES = {
    發票: "both",
    收據: "fixedUntax",
    合約: "both",
    報價單: "both",
    勞務報酬單: "fixedUntax",
    "簽呈(無憑證補貼款項)": "fixedUntax",
    差旅費: "both",
    "通路費用(通路後扣)": "both",
  };

  const PAYMENT_TYPE = {
    GENERAL: "GENERAL",
    PETTY_CASH: "PETTY_CASH",
    URGENT: "URGENT",
    PREPAY: "PREPAY",
  };

  const APPLICATION_STATUS = {
    DRAFT: "draft",
    PENDING_APPROVAL: "pending_approval",
    APPROVED: "approved",
    REJECTED: "rejected",
    PAID: "paid",
    // PREPAY 核銷/結算流程（母子單勾稽）
    WRITE_OFF_PENDING: "writeoff_pending", // 待核銷
    PARTIAL_WRITE_OFF: "partial_writeoff", // 部分核銷（相容舊資料）
    WRITE_OFF_AUDITING: "writeoff_auditing", // 核銷審核中（母單：子單待財務審）
    WRITE_OFF_REVIEWING: "writeoff_reviewing", // 結算審核中（相容舊資料）
    SETTLED: "settled", // 已結案
    PAYMENT_PENDING: "payment_pending", // 待付款（母單審核通過後，出納付款前）
    WRITE_OFF_REJECTED: "writeoff_rejected", // 被退件
  };

  function parseInputDate(s) {
    if (!s || typeof s !== "string") return null;
    const p = s.trim().split("-");
    if (p.length !== 3) return null;
    const y = parseInt(p[0], 10);
    const m = parseInt(p[1], 10);
    const d = parseInt(p[2], 10);
    if (!Number.isFinite(y) || !Number.isFinite(m) || !Number.isFinite(d)) return null;
    return new Date(y, m - 1, d);
  }

  function formatISODate(date) {
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${y}-${m}-${day}`;
  }

  function addDays(date, n) {
    const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
    d.setDate(d.getDate() + n);
    return d;
  }

  /** 零用金：非週四 → 下一個週四；週四 → 下週四（+7） */
  function nextPettyCashThursday(applyDate) {
    const d = new Date(applyDate.getFullYear(), applyDate.getMonth(), applyDate.getDate());
    const dow = d.getDay();
    if (dow === 4) {
      d.setDate(d.getDate() + 7);
      return d;
    }
    let add = (4 - dow + 7) % 7;
    if (add === 0) add = 7;
    d.setDate(d.getDate() + add);
    return d;
  }

  /**
   * 一般付款（GENERAL）：
   * 以「申請日所屬區間」決定付款日：
   * - 每月 21 日至次月 5 日（含）之申請，付款日為次月 15 日
   * - 每月 6 日至 20 日（含）之申請，付款日為當月 30 日（若當月天數不足 30，取當月最後一天）
   */
  function calcGeneralExpectedDate(applyDate) {
    const y = applyDate.getFullYear();
    const m = applyDate.getMonth();
    const day = applyDate.getDate();
    const daysInMonth = new Date(y, m + 1, 0).getDate();

    // day 1~5：落在「上個月21~本月5」區間，付款日為本月15
    if (day <= 5) return new Date(y, m, 15);
    // day 6~20：付款日為當月30
    if (day <= 20) return new Date(y, m, Math.min(30, daysInMonth));
    // day 21~31：付款日為次月15
    return new Date(y, m + 1, 15);
  }

  let vendors = [];
  let nextId = 1;
  let comboOpen = false;
  let comboHighlight = -1;
  let applications = [];
  let currentApplicationId = null;
  let isMasterCreated = false;
  let currentApplicationStatus = APPLICATION_STATUS.DRAFT;

  function $(sel, root) {
    return (root || document).querySelector(sel);
  }

  function nowIso() {
    return new Date().toISOString();
  }

  function formatDateTime(iso) {
    if (!iso) return "";
    const d = new Date(iso);
    if (Number.isNaN(d.getTime())) return "";
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    const hh = String(d.getHours()).padStart(2, "0");
    const mm = String(d.getMinutes()).padStart(2, "0");
    return `${y}/${m}/${day} ${hh}:${mm}`;
  }

  function generateApplicationId() {
    const d = new Date();
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    const seq = String(applications.length + 1).padStart(4, "0");
    return `AP${y}${m}${day}-${seq}`;
  }

  function loadVendorsFromStorage() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return;
      const data = JSON.parse(raw);
      if (Array.isArray(data.vendors)) {
        vendors = data.vendors;
      }
    } catch (e) {
      console.warn(e);
    }
  }

  function saveVendorsToStorage() {
    localStorage.setItem(
      STORAGE_KEY,
      JSON.stringify({ vendors, savedAt: new Date().toISOString() })
    );
  }

  function updateVendorBadge() {
    const el = $("#vendor-count");
    if (el) el.textContent = String(vendors.length);
  }

  function loadPettyCashLimitFromStorage() {
    try {
      const raw = localStorage.getItem(PETTY_CASH_LIMIT_KEY);
      if (raw == null) return;
      const n = Number(raw);
      if (Number.isFinite(n) && n >= 0) pettyCashLimit = n;
    } catch (e) {
      console.warn(e);
    }
  }

  function persistPettyCashLimitToStorage(value) {
    localStorage.setItem(PETTY_CASH_LIMIT_KEY, String(value));
  }

  function syncPettyCashLimitInput() {
    const el = $("#petty-cash-limit");
    if (!el) return;
    el.value = String(pettyCashLimit);
  }

  function initPettyCashLimitSetting() {
    const input = $("#petty-cash-limit");
    const btn = $("#btn-save-petty-cash-limit");
    const msgEl = $("#petty-cash-limit-msg");
    if (!input || !btn) return;

    const showMsg = (s) => {
      if (msgEl) msgEl.textContent = s;
    };

    btn.addEventListener("click", () => {
      const n = Number(input.value);
      if (!Number.isFinite(n) || n < 0) {
        alert("零用金上限請輸入有效數字（>= 0）。");
        showMsg("");
        return;
      }
      pettyCashLimit = n;
      persistPettyCashLimitToStorage(pettyCashLimit);
      showMsg("已更新零用金上限。");
      applyScenarioRules();
      recalcPettyCapHint();
    });
  }

  function loadApplications() {
    try {
      const raw = localStorage.getItem(APPLICATIONS_KEY);
      if (!raw) return;
      const data = JSON.parse(raw);
      if (Array.isArray(data)) applications = data;
    } catch (e) {
      console.warn(e);
    }
  }

  function saveApplications() {
    localStorage.setItem(APPLICATIONS_KEY, JSON.stringify(applications));
  }

  /** 只擷取 Customer Name（或相容：name / customerName）作為選項；其餘欄位忽略 */
  function nameFromRow(item) {
    if (item == null) return "";
    if (typeof item === "string") return item.trim();
    if (typeof item !== "object") return "";
    for (const k of Object.keys(item)) {
      if (k.trim().toLowerCase() === "customer name") {
        return String(item[k] != null ? item[k] : "").trim();
      }
    }
    if (Object.prototype.hasOwnProperty.call(item, "customerName")) {
      return String(item.customerName != null ? item.customerName : "").trim();
    }
    if (Object.prototype.hasOwnProperty.call(item, "name")) {
      return String(item.name != null ? item.name : "").trim();
    }
    return "";
  }

  /** 將任意陣列正規化為 { name }（內部仍用 name 存顯示文字；序號僅供內部鍵值） */
  function vendorsNameOnly(rows) {
    const out = [];
    for (let i = 0; i < rows.length; i++) {
      const name = nameFromRow(rows[i]);
      if (!name) continue;
      out.push({ name, code: String(i + 1) });
    }
    return out;
  }

  function jsonArrayFromRoot(data) {
    if (Array.isArray(data)) return data;
    if (data && typeof data === "object") {
      const prefer = ["vendors", "data", "items", "list", "records"];
      for (const k of prefer) {
        if (Array.isArray(data[k])) return data[k];
      }
      for (const k of Object.keys(data)) {
        if (Array.isArray(data[k])) return data[k];
      }
    }
    return null;
  }

  function parseVendorJson(text) {
    const data = JSON.parse(text);
    const arr = jsonArrayFromRoot(data);
    if (!arr || !arr.length) {
      throw new Error(
        "JSON 需為陣列，或物件內含一個陣列（例如 { \"vendors\": [...] }）；每筆請含 Customer Name 欄位"
      );
    }
    const list = vendorsNameOnly(arr);
    if (!list.length) {
      throw new Error(
        "找不到任何客戶名稱：請確認每筆有 Customer Name（或相容欄位 name）"
      );
    }
    return list;
  }

  /** CSV：第一列為標題，只讀取 Customer Name 欄（不分大小寫；相容 name），其餘欄位忽略 */
  function parseVendorCsv(text) {
    const lines = text.split(/\r?\n/).filter((l) => l.trim());
    if (lines.length < 2) throw new Error("CSV 至少需要標題與一筆資料");
    const header = lines[0].split(",").map((h) => h.trim().replace(/^\uFEFF/, ""));
    const nameIdx = header.findIndex((h) => {
      const x = h.trim().toLowerCase();
      return x === "customer name" || x === "name";
    });
    if (nameIdx < 0) {
      throw new Error(
        "CSV 標題需包含 Customer Name 欄位（其餘欄位可保留，僅擷取 Customer Name）"
      );
    }
    const rows = [];
    for (let r = 1; r < lines.length; r++) {
      const cells = lines[r].split(",");
      const name = (cells[nameIdx] || "").trim().replace(/^"|"$/g, "");
      if (name) rows.push({ name });
    }
    return vendorsNameOnly(rows);
  }

  function getDetailItemsFromTable() {
    const out = [];
    document.querySelectorAll("#detail-body tr[data-row-id]").forEach((tr) => {
      if (tr._detail) out.push({ ...tr._detail });
    });
    return out;
  }

  function clearDetailTable() {
    $("#detail-body").innerHTML = "";
    nextId = 1;
    updateDetailCount();
    updateMasterTotal();
  }

  function renderDetailItems(items) {
    clearDetailTable();
    (items || []).forEach((d) => {
      const id = nextId++;
      const tr = document.createElement("tr");
      tr.dataset.rowId = String(id);
      tr._detail = d;
      fillSummaryRow(tr, id, d);
      $("#detail-body").appendChild(tr);
      bindSummaryRowActions(tr);
    });
    updateDetailCount();
    updateMasterTotal();
  }

  function collectMasterForm() {
    return {
      applicant: ($("#applicant").value || "").trim(),
      applyDate: $("#applyDate").value,
      payCategory: $("#payCategory").value,
      vendorName: ($("#vendor-hidden").value || $("#vendor-search").value || "").trim(),
      currency: $("#currency").value,
      payMethod: $("#payMethod").value,
      transferFee: $("#transferFee").value,
      expectedDate: $("#expectedDate").value,
      totalPayment: parseNum($("#totalPayment").value),
    };
  }

  function fillMasterForm(master) {
    $("#applicant").value = master?.applicant || "";
    $("#applyDate").value = master?.applyDate || $("#applyDate").value;
    $("#payCategory").value = master?.payCategory || $("#payCategory").value;
    $("#vendor-search").value = master?.vendorName || "";
    $("#vendor-hidden").value = master?.vendorName || "";
    $("#currency").value = master?.currency || $("#currency").value;
    $("#payMethod").value = master?.payMethod || $("#payMethod").value;
    $("#transferFee").value = master?.transferFee || $("#transferFee").value;
    $("#expectedDate").value = master?.expectedDate || "";
    syncExpectedPaymentDate();
    initPaymentMethod();
    applyScenarioRules();
  }

  function setMasterCreated(created) {
    isMasterCreated = !!created;
    $("#btn-edit-detail").textContent = "建立";
    $("#btn-add-row").disabled = !isMasterCreated;
    const tabBtn = $("#tab-btn-detail");
    if (tabBtn) {
      if (!isMasterCreated) {
        tabBtn.classList.add("hidden");
        tabBtn.setAttribute("aria-hidden", "true");
      } else {
        tabBtn.classList.remove("hidden");
        tabBtn.setAttribute("aria-hidden", "false");
      }
    }
  }

  function statusToLabel(status) {
    if (!status) return "draft";
    if (status === "voided") return "作廢";
    if (status === "submitted") return "待審核";
    if (status === APPLICATION_STATUS.DRAFT) return "草稿";
    if (status === APPLICATION_STATUS.PENDING_APPROVAL) return "待審核";
    if (status === APPLICATION_STATUS.APPROVED) return "審核通過";
    if (status === APPLICATION_STATUS.REJECTED) return "審核不通過";
    if (status === APPLICATION_STATUS.PAID) return "已完成";
    if (status === APPLICATION_STATUS.WRITE_OFF_PENDING) return "待核銷";
    if (status === APPLICATION_STATUS.PARTIAL_WRITE_OFF) return "部分核銷";
    if (status === APPLICATION_STATUS.WRITE_OFF_AUDITING) return "核銷審核中";
    if (status === APPLICATION_STATUS.WRITE_OFF_REVIEWING) return "結算審核中";
    if (status === APPLICATION_STATUS.SETTLED) return "已結案";
    if (status === APPLICATION_STATUS.PAYMENT_PENDING) return "待付款";
    if (status === APPLICATION_STATUS.WRITE_OFF_REJECTED) return "被退件";
    return String(status);
  }

  /** 列表狀態標籤樣式 class（與 CSS .status-pill.st-* 對應） */
  function statusPillClass(statusKey) {
    if (!statusKey || statusKey === "voided") return "st-voided";
    if (statusKey === APPLICATION_STATUS.DRAFT) return "st-draft";
    if (statusKey === APPLICATION_STATUS.PENDING_APPROVAL || statusKey === "submitted") return "st-pending";
    if (statusKey === APPLICATION_STATUS.REJECTED) return "st-rejected";
    if (statusKey === APPLICATION_STATUS.APPROVED) return "st-approved";
    if (statusKey === APPLICATION_STATUS.PAID) return "st-paid";
    if (statusKey === APPLICATION_STATUS.WRITE_OFF_PENDING) return "st-writeoff-pending";
    if (statusKey === APPLICATION_STATUS.PARTIAL_WRITE_OFF) return "st-partial-writeoff";
    if (statusKey === APPLICATION_STATUS.WRITE_OFF_AUDITING) return "st-writeoff-auditing";
    if (statusKey === APPLICATION_STATUS.WRITE_OFF_REVIEWING) return "st-writeoff-reviewing";
    if (statusKey === APPLICATION_STATUS.SETTLED) return "st-settled";
    if (statusKey === APPLICATION_STATUS.PAYMENT_PENDING) return "st-payment-pending";
    if (statusKey === APPLICATION_STATUS.WRITE_OFF_REJECTED) return "st-writeoff-rejected";
    return "st-default";
  }

  /** 明細頁籤已出現且為目前分頁時，才顯示頁籤列「提交」 */
  function updateTabsSubmitVisibility() {
    const wrap = $("#tabs-submit-actions");
    const submitBtn = $("#btn-submit");
    if (!wrap) return;
    const tabDetail = $("#tab-btn-detail");
    const detailPanel = $("#panel-detail");
    const detailTabExists = tabDetail && !tabDetail.classList.contains("hidden");
    const onDetailTab = detailPanel && detailPanel.classList.contains("active");
    const app = getCurrentApplication();
    const locked = app ? isReviewLocked(app) : true;
    const show = isMasterCreated && detailTabExists && onDetailTab && !locked;
    wrap.classList.toggle("hidden", !show);
    if (submitBtn) submitBtn.disabled = !canEditApplication(app);
  }

  function payCategoryToLabel(payCategory) {
    if (!payCategory) return "";
    if (payCategory === PAYMENT_TYPE.PETTY_CASH) return "個人代墊報支";
    if (payCategory === PAYMENT_TYPE.GENERAL) return "廠商實報實銷";
    if (payCategory === PAYMENT_TYPE.PREPAY) return "廠商預付／訂金";
    return String(payCategory);
  }

  function setFormMeta(applicationId, status) {
    $("#meta-application-id").textContent = applicationId || "（尚未建立）";
    $("#meta-application-status").textContent = statusToLabel(status);
  }

  function getCurrentApplication() {
    if (!currentApplicationId) return null;
    return applications.find((x) => x.applicationId === currentApplicationId) || null;
  }

  function canEditApplication(app) {
    if (!app) return false;
    if (app.voided) return false;
    return app.status === APPLICATION_STATUS.DRAFT || app.status === APPLICATION_STATUS.REJECTED;
  }

  function isReviewLocked(app) {
    // 尚未建立／選取申請單（新增第二筆時 currentApplicationId 為 null）應可編輯，不可沿用上一筆鎖定狀態
    if (!app) return false;
    if (app.voided) return true;
    return (
      app.status === APPLICATION_STATUS.PENDING_APPROVAL ||
      app.status === "submitted" ||
      app.status === APPLICATION_STATUS.APPROVED ||
      app.status === APPLICATION_STATUS.PAID ||
      // PREPAY 核銷流程：只讓使用者做「核銷/結算」操作，不開啟主檔/明細編輯
      app.status === APPLICATION_STATUS.WRITE_OFF_PENDING ||
      app.status === APPLICATION_STATUS.PARTIAL_WRITE_OFF ||
      app.status === APPLICATION_STATUS.WRITE_OFF_AUDITING ||
      app.status === APPLICATION_STATUS.WRITE_OFF_REVIEWING ||
      app.status === APPLICATION_STATUS.SETTLED ||
      app.status === APPLICATION_STATUS.PAYMENT_PENDING ||
      app.status === APPLICATION_STATUS.WRITE_OFF_REJECTED
    );
  }

  function setReviewButtonsVisibility(app) {
    const wrap = $("#tabs-review-actions");
    const isPendingReview =
      app &&
      (app.status === APPLICATION_STATUS.PENDING_APPROVAL || app.status === "submitted") &&
      !app.voided;
    if (wrap) wrap.classList.toggle("hidden", !isPendingReview);
  }

  function lockMasterFields(locked) {
    const ids = [
      "#applicant",
      "#applyDate",
      "#vendor-search",
      "#currency",
      "#payMethod",
      "#transferFee",
      "#expectedDate",
    ];
    ids.forEach((sel) => {
      const el = $(sel);
      if (!el) return;
      if (locked) el.setAttribute("disabled", "true");
      else el.removeAttribute("disabled");
    });
    const payCat = $("#payCategory");
    if (payCat) {
      if (locked || isMasterCreated) payCat.setAttribute("disabled", "true");
      else payCat.removeAttribute("disabled");
    }
    const addRow = $("#btn-add-row");
    if (addRow) {
      addRow.disabled = locked || !isMasterCreated;
      addRow.classList.remove("hidden");
    }

  }

  function renderFormState() {
    const app = getCurrentApplication();
    currentApplicationStatus = app?.status || APPLICATION_STATUS.DRAFT;
    const locked = isReviewLocked(app);
    lockMasterFields(locked);
    setReviewButtonsVisibility(app);

    // 「建立」：僅在尚未建立主單時顯示；建立後隱藏（改以頁籤切換明細）
    const btnCreate = $("#btn-edit-detail");
    if (btnCreate) {
      btnCreate.classList.toggle("hidden", isMasterCreated);
    }

    // PREPAY 核銷/結算 UI（疊加式擴充）
    applyWriteoffUI(app);

    // PREPAY 核銷結算按鈕（在 writeoff tab 也要可用）
    applyWriteoffReviewButtonsVisibility(app);

    // PREPAY 子單：財務審核面板
    applyChildReviewPanelUI(app);

    updateTabsSubmitVisibility();
  }

  function getApplicationTotal(app) {
    if (!app) return 0;
    const itemsTotal = (app.items || []).reduce((s, x) => s + parseNum(String(x.payAmount || 0)), 0);
    if (itemsTotal > 0) return itemsTotal;
    return parseNum(String(app.master?.totalPayment ?? 0));
  }

  function getRemainingBalanceForApp(app) {
    const v = app?.remainingBalance;
    if (v != null && v !== "") return parseNum(String(v));
    // 若為 PREPAY 母單且尚未寫入 remainingBalance，先用總金額推算
    if (app?.master?.payCategory === PAYMENT_TYPE.PREPAY) return getApplicationTotal(app);
    return 0;
  }

  function isPrepayApplication(app) {
    return !!app?.master && app.master.payCategory === PAYMENT_TYPE.PREPAY;
  }

  function isPrepayMother(app) {
    return isPrepayApplication(app) && (app.parentId == null);
  }

  /** 預付母單：須財務審核通過後才顯示「核銷紀錄」頁籤（草稿／待審／退件不顯示） */
  function canShowPrepayWriteoffTab(app) {
    if (!app || !isPrepayMother(app) || app.voided) return false;
    const s = app.status;
    return (
      s === APPLICATION_STATUS.APPROVED ||
      s === APPLICATION_STATUS.PAID ||
      s === APPLICATION_STATUS.WRITE_OFF_PENDING ||
      s === APPLICATION_STATUS.PARTIAL_WRITE_OFF ||
      s === APPLICATION_STATUS.WRITE_OFF_AUDITING ||
      s === APPLICATION_STATUS.WRITE_OFF_REVIEWING ||
      s === APPLICATION_STATUS.SETTLED ||
      s === APPLICATION_STATUS.WRITE_OFF_REJECTED
    );
  }

  function getWriteoffChildren(parentId) {
    if (!parentId) return [];
    return applications
      .filter((a) => a.parentId === parentId)
      .sort((x, y) => String(y.createdAt || "").localeCompare(String(x.createdAt || "")));
  }

  function renderWriteoffHistoryTable(motherApp) {
    const body = $("#writeoff-history-body");
    if (!body) return;
    const children = getWriteoffChildren(motherApp?.applicationId);
    if (!children.length) {
      body.innerHTML =
        '<tr><td colspan="3" style="color:#8c8c8c">目前無核銷紀錄</td></tr>';
      return;
    }
    body.innerHTML = children
      .map((c) => {
        const childId = c.applicationId || "";
        const amount = getApplicationTotal(c);
        const statusText = escapeHtml(statusToLabel(c.status));
        return `<tr>
          <td>${escapeHtml(childId)}</td>
          <td class="cell-num">${formatMoney(amount)}</td>
          <td>${statusText}</td>
        </tr>`;
      })
      .join("");
  }

  function applyWriteoffUI(app) {
    const meta = $("#meta-remaining-balance");
    const tabBtn = $("#tab-btn-writeoff");
    const panel = $("#panel-writeoff");
    const addBtn = $("#btn-add-writeoff");
    const groupTitle = $("#writeoff-group-title");
    const groupParentId = $("#writeoff-group-parent-id");
    const groupRemaining = $("#writeoff-group-remaining");

    if (!meta || !tabBtn || !panel || !addBtn) return;

    if (app && isPrepayMother(app)) {
      const showWriteoffTab = canShowPrepayWriteoffTab(app);
      tabBtn.classList.toggle("hidden", !showWriteoffTab);
      if (!showWriteoffTab) {
        meta.classList.add("hidden");
        if (groupTitle) groupTitle.classList.add("hidden");
        addBtn.classList.add("hidden");
        const panelWriteoff = $("#panel-writeoff");
        if (panelWriteoff && panelWriteoff.classList.contains("active")) {
          document.querySelector('.tab[data-tab="master"]')?.click();
        }
      } else {
        const remaining = getRemainingBalanceForApp(app);
        meta.classList.remove("hidden");
        meta.querySelector("strong") && (meta.querySelector("strong").textContent = formatMoney(remaining));
        if (groupTitle) groupTitle.classList.remove("hidden");
        if (groupParentId) groupParentId.textContent = app.applicationId;
        if (groupRemaining) groupRemaining.textContent = formatMoney(remaining);
        renderWriteoffHistoryTable(app);
        const hideAdd =
          app.status === APPLICATION_STATUS.WRITE_OFF_REVIEWING ||
          app.status === APPLICATION_STATUS.SETTLED ||
          app.status === APPLICATION_STATUS.PAID;
        addBtn.classList.toggle("hidden", hideAdd);
      }
    } else {
      meta.classList.add("hidden");
      tabBtn.classList.add("hidden");
      panel.classList.add("hidden");
      addBtn.classList.add("hidden");
      if (groupTitle) groupTitle.classList.add("hidden");
    }
  }

  function applyWriteoffReviewButtonsVisibility(app) {
    const btn = $("#btn-review-writeoff-approve");
    if (!btn) return;
    btn.classList.add("hidden");
  }

  function getWriteoffAttributeLabel(childType) {
    if (!childType) return "";
    if (childType === "WRITE_OFF") return "一般核銷";
    if (childType === "GENERAL") return "補尾款";
    if (childType === "REFUND") return "退款單";
    return String(childType);
  }

  function applyChildReviewPanelUI(app) {
    const panel = $("#child-review-panel");
    if (!panel) return;

    // 子單：parentId 不為 null
    const isChild = !!app && app.parentId != null;
    const eligible =
      isChild &&
      (app.status === APPLICATION_STATUS.PENDING_APPROVAL || app.status === "submitted");

    if (!eligible) {
      panel.classList.add("hidden");
      return;
    }

    const mother = applications.find((x) => x.applicationId === app.parentId);
    $("#child-review-parent-id").textContent = mother?.applicationId || String(app.parentId || "");
    $("#child-review-attribute").textContent = getWriteoffAttributeLabel(app.type);

    panel.classList.remove("hidden");
  }

  function activateWriteoffTabIfNeeded(app) {
    if (!app || !isPrepayMother(app) || !canShowPrepayWriteoffTab(app)) return;
    const targetStatuses = [
      APPLICATION_STATUS.WRITE_OFF_PENDING,
      APPLICATION_STATUS.PARTIAL_WRITE_OFF,
      APPLICATION_STATUS.WRITE_OFF_AUDITING,
      APPLICATION_STATUS.WRITE_OFF_REVIEWING,
      APPLICATION_STATUS.SETTLED,
    ];
    if (!targetStatuses.includes(app.status)) return;

    const tabBtn = $("#tab-btn-writeoff");
    if (!tabBtn) return;
    if (tabBtn.classList.contains("hidden")) tabBtn.classList.remove("hidden");
    // 直接切換成 writeoff 面板（避免依賴使用者手動點擊）
    const panels = document.querySelectorAll(".tab-panel");
    const tabs = document.querySelectorAll(".tab");
    tabs.forEach((t) => {
      const isWriteoff = t.getAttribute("data-tab") === "writeoff";
      t.classList.toggle("active", isWriteoff);
      if (isWriteoff) t.classList.remove("hidden");
    });
    panels.forEach((p) => {
      const isWriteoff = p.getAttribute("data-panel") === "writeoff";
      p.classList.toggle("active", isWriteoff);
      p.classList.toggle("hidden", !isWriteoff);
    });
    updateTabsSubmitVisibility();
  }

  // ===== PREPAY：核銷浮窗（Writeoff Modal）=====
  /** 母單核銷來源明細：優先 details，否則 items；無資料時給一筆備用列 */
  function getMotherWriteoffSourceRows(mother) {
    if (!mother) return [];
    const fromDetails = mother.details;
    if (Array.isArray(fromDetails) && fromDetails.length > 0) return fromDetails;
    const fromItems = mother.items;
    if (Array.isArray(fromItems) && fromItems.length > 0) return fromItems;
    const total = getApplicationTotal(mother);
    return [
      {
        purpose: "其他費用",
        remarkNo: "",
        payAmount: total,
        untaxed: total,
        tax: 0,
      },
    ];
  }

  function getLineOriginalAmount(row) {
    if (!row) return 0;
    if (row.payAmount != null && row.payAmount !== "") return parseNum(String(row.payAmount));
    return parseNum(String(row.untaxed || 0)) + parseNum(String(row.tax || 0));
  }

  /** 該母單項目已被核銷／送審中的金額（待審核 + 已完成） */
  function getWrittenOffAmountForMotherLine(motherId, lineIndex) {
    if (!motherId) return 0;
    return applications
      .filter((a) => {
        if (a.parentId !== motherId || a.voided) return false;
        const idx = a.writeoffSourceLineIndex;
        if (idx == null || idx === "") return lineIndex === 0;
        return Number(idx) === lineIndex;
      })
      .filter(
        (a) =>
          a.status === APPLICATION_STATUS.PAID ||
          a.status === APPLICATION_STATUS.PENDING_APPROVAL ||
          a.status === "submitted"
      )
      .reduce((s, a) => s + parseNum(String(a.writeoffInvoiceAmount ?? getApplicationTotal(a))), 0);
  }

  /** 選定母單明細列之可核銷餘額（原申請金額 − 已送審／已核銷） */
  function getRemainingForMotherLine(mother, lineIndex) {
    const rows = getMotherWriteoffSourceRows(mother);
    const row = rows[lineIndex];
    if (!row) return 0;
    const orig = getLineOriginalAmount(row);
    const used = getWrittenOffAmountForMotherLine(mother.applicationId, lineIndex);
    return Math.max(0, orig - used);
  }

  function populateWriteoffSourceDetailSelect(mother) {
    const sel = $("#writeoff-source-detail");
    if (!sel) return;
    sel.innerHTML = "";
    const rows = getMotherWriteoffSourceRows(mother);
    rows.forEach((row, idx) => {
      const purpose = row.purpose || "其他費用";
      const remarkRaw = (row.remarkNo || "").trim();
      const remarkLabel = remarkRaw || "—";
      const orig =
        row.payAmount != null && row.payAmount !== ""
          ? parseNum(String(row.payAmount))
          : parseNum(String(row.untaxed || 0)) + parseNum(String(row.tax || 0));
      const opt = document.createElement("option");
      opt.value = String(idx);
      opt.textContent = `[${purpose}] 備註單號: ${remarkLabel} (原申請金額: ${formatMoney(orig)})`;
      sel.appendChild(opt);
    });
    if (rows.length) sel.selectedIndex = 0;
  }

  /** 子單核銷明細：供 details 與 items（表格）共用欄位 */
  function buildWriteoffChildDetailFromSelection(selectedLine, payAmountForLine, voucherNo, attachmentName) {
    const purpose = selectedLine.purpose || "其他費用";
    const remarkNo = selectedLine.remarkNo || "";
    const periodStart = selectedLine.periodStart || "";
    const periodEnd = selectedLine.periodEnd || "";
    const note = "(系統自動產生) 關聯母單核銷";
    const amt = parseNum(String(payAmountForLine));
    const detailRow = {
      purpose,
      remarkNo,
      voucherStyle: "發票",
      voucherNo: voucherNo,
      payAmount: amt,
      note,
    };
    const itemRow = {
      ...detailRow,
      periodStart,
      periodEnd,
      taxType: "未稅",
      untaxed: amt,
      tax: 0,
      fileName: attachmentName,
    };
    return { detailRow, itemRow };
  }

  function openWriteoffModal() {
    const app = getCurrentApplication();
    if (!app || !isPrepayMother(app)) return;
    const modal = $("#writeoff-modal");
    if (!modal) return;

    populateWriteoffSourceDetailSelect(app);
    $("#writeoff-actual-invoice").value = "";
    $("#writeoff-voucher-no").value = "";
    $("#writeoff-attachment").value = "";
    $("#writeoff-is-final").checked = false;
    recalcWriteoffDiffHint();

    modal.hidden = false;
    document.body.style.overflow = "hidden";
    $("#writeoff-actual-invoice")?.focus();
  }

  function closeWriteoffModal() {
    const modal = $("#writeoff-modal");
    if (modal) modal.hidden = true;
    // 若 detail modal 還在顯示，避免把底層鎖定解除
    if ($("#detail-modal") && !$("#detail-modal").hidden) {
      document.body.style.overflow = "hidden";
    } else {
      document.body.style.overflow = "";
    }
  }

  function recalcWriteoffDiffHint() {
    const hint = $("#writeoff-diff-hint");
    const app = getCurrentApplication();
    if (!hint || !app) return;

    const sel = $("#writeoff-source-detail");
    const idx = sel ? parseInt(String(sel.value || "0"), 10) : 0;
    const idxSafe = Number.isNaN(idx) ? 0 : idx;
    const rows = getMotherWriteoffSourceRows(app);
    const row = rows[idxSafe];
    const orig = row ? getLineOriginalAmount(row) : 0;
    const lineRemaining = getRemainingForMotherLine(app, idxSafe);
    const invoice = parseNum($("#writeoff-actual-invoice").value);
    if (!invoice) {
      hint.textContent = `本項目原申請金額：${formatMoney(orig)}；可核銷餘額：${formatMoney(lineRemaining)}（未輸入實際發票／核銷金額）`;
      return;
    }

    const displayDiff = orig - invoice;
    const diffLabel = displayDiff >= 0 ? `+${formatMoney(displayDiff)}` : formatMoney(displayDiff);
    hint.textContent = `本項目可核銷餘額：${formatMoney(lineRemaining)}｜差額（原申請金額 − 本次核銷金額）：${formatMoney(orig)} − ${formatMoney(invoice)} = ${diffLabel}`;
  }

  function submitWriteoff() {
    const mother = getCurrentApplication();
    if (!mother || !isPrepayMother(mother)) return;
    if (!canShowPrepayWriteoffTab(mother)) {
      alert("母單須先經財務審核通過後，才開放核銷。");
      return;
    }
    const canMotherWriteoff =
      mother.status === APPLICATION_STATUS.WRITE_OFF_PENDING ||
      mother.status === APPLICATION_STATUS.PARTIAL_WRITE_OFF ||
      mother.status === APPLICATION_STATUS.WRITE_OFF_AUDITING;
    if (!canMotherWriteoff) {
      alert("目前母單狀態不可新增核銷。");
      return;
    }
    const invoice = parseNum($("#writeoff-actual-invoice").value);
    if (!invoice || invoice <= 0) {
      alert("請輸入實際發票總金額（> 0）");
      $("#writeoff-actual-invoice")?.focus();
      return;
    }

    const writeoffVoucherNo = ($("#writeoff-voucher-no").value || "").trim();
    if (!writeoffVoucherNo) {
      alert("請填寫實際發票號碼（憑證號碼）");
      $("#writeoff-voucher-no")?.focus();
      return;
    }

    const fileInput = $("#writeoff-attachment");
    const attachmentName = fileInput?.files?.[0]?.name || "";
    if (!attachmentName) {
      alert("請上傳發票照片（附件）");
      fileInput?.focus();
      return;
    }

    const sourceRows = getMotherWriteoffSourceRows(mother);
    const detailSel = $("#writeoff-source-detail");
    const selVal = detailSel?.value;
    if (selVal === undefined || selVal === null || selVal === "") {
      alert("請選擇要核銷的母單項目");
      detailSel?.focus();
      return;
    }
    const srcIdx = parseInt(String(selVal), 10);
    if (Number.isNaN(srcIdx) || srcIdx < 0 || srcIdx >= sourceRows.length) {
      alert("母單項目選擇無效，請重新選擇");
      detailSel?.focus();
      return;
    }
    const selectedMotherLine = sourceRows[srcIdx];
    const lineRemaining = getRemainingForMotherLine(mother, srcIdx);

    const isFinal = !!$("#writeoff-is-final")?.checked;
    const diff = invoice - lineRemaining;
    const EPS = 0.000001;

    const createChild = ({ type, amount }) => {
      const childId = generateApplicationId();
      const childPayCategory =
        type === "GENERAL" || type === "REFUND" ? PAYMENT_TYPE.GENERAL : PAYMENT_TYPE.PREPAY;

      const { detailRow, itemRow } = buildWriteoffChildDetailFromSelection(
        selectedMotherLine,
        amount,
        writeoffVoucherNo,
        attachmentName
      );
      const childItems = [itemRow];
      const childDetails = [detailRow];

      applications.unshift({
        applicationId: childId,
        parentId: mother.applicationId,
        type,
        createdAt: nowIso(),
        updatedAt: nowIso(),
        status: APPLICATION_STATUS.PENDING_APPROVAL,
        voided: false,
        remainingBalance: 0,
        writeoffInvoiceAmount: invoice,
        writeoffIsFinal: isFinal,
        writeoffSourceLineIndex: srcIdx,
        master: {
          ...mother.master,
          payCategory: childPayCategory,
          totalPayment: amount,
          expectedDate: mother.master?.expectedDate || "",
          writeoffVoucherNo,
          writeoffAttachmentName: attachmentName,
        },
        items: childItems,
        details: childDetails,
      });
      return childId;
    };

    // 情境判斷（A/B/C/D）：差額以「所選母單項目之可核銷餘額」為基準
    if (!isFinal) {
      if (invoice > lineRemaining) {
        alert("未勾選結清時，實際發票金額不可大於此母單項目之可核銷餘額。");
        return;
      }
      createChild({ type: "WRITE_OFF", amount: invoice });
    } else if (Math.abs(diff) <= EPS) {
      createChild({
        type: "WRITE_OFF",
        amount: invoice,
      });
    } else if (diff > 0) {
      createChild({ type: "GENERAL", amount: diff });
    } else {
      createChild({
        type: "REFUND",
        amount: Math.abs(diff),
      });
    }

    const motherIdxAfter = applications.findIndex((x) => x.applicationId === mother.applicationId);
    if (motherIdxAfter < 0) return;
    applications[motherIdxAfter].status = APPLICATION_STATUS.WRITE_OFF_AUDITING;

    saveApplications();
    closeWriteoffModal();

    // 更新畫面（核銷紀錄/剩餘金額/按鈕顯示）
    const updatedMother = applications.find((x) => x.applicationId === mother.applicationId);
    if (!updatedMother) return;
    currentApplicationId = updatedMother.applicationId;
    fillMasterForm(updatedMother.master || {});
    renderDetailItems(updatedMother.items || []);
    setFormMeta(updatedMother.applicationId, updatedMother.voided ? "voided" : updatedMother.status);
    renderFormState();

    // 確保使用者立刻看到剛新增的核銷紀錄（包含結算審核中）
    activateWriteoffTabIfNeeded(updatedMother);
  }

  // ===== PREPAY：核銷結算審核（舊批次流程已停用，保留函數避免外部引用錯誤）=====
  function approveWriteoffSettlement() {
    return false;
  }

  function getMotherById(motherId) {
    if (!motherId) return null;
    return applications.find((x) => x.applicationId === motherId) || null;
  }

  /** 財務審核通過（母單：一般／預付首審；子單：核銷子單） */
  function approveTicket(id) {
    const app = applications.find((x) => x.applicationId === id);
    if (!app || app.voided) return false;
    if (app.parentId != null) return approveChildTicket(id);
    return approveMotherTicket(id);
  }

  function approveMotherTicket(id) {
    const app = applications.find((x) => x.applicationId === id);
    if (!app || app.voided || app.parentId != null) return false;
    if (app.status !== APPLICATION_STATUS.PENDING_APPROVAL && app.status !== "submitted") return false;

    const current = $("#expectedDate")?.value || "";
    const defaultDate = current || formatISODate(new Date());
    const financeExpected = prompt("請輸入財務預期付款日（YYYY-MM-DD）", defaultDate);
    if (!financeExpected) return false;
    const ok = /^\d{4}-\d{2}-\d{2}$/.test(financeExpected.trim());
    if (!ok) {
      alert("日期格式錯誤，請輸入 YYYY-MM-DD");
      return false;
    }
    $("#expectedDate").value = financeExpected.trim();
    clearExpectedDateError();

    // 審核通過後一律先進「待付款」，由財務於列表點「完成付款」並回填實際付款日後才轉為已完成（或預付則進入待核銷）
    const targetStatus = APPLICATION_STATUS.PAYMENT_PENDING;

    currentApplicationId = id;
    upsertCurrentApplication(targetStatus);

    const idx = applications.findIndex((x) => x.applicationId === id);
    if (idx >= 0) {
      applications[idx].reviewedAt = nowIso();
      applications[idx].reviewNote = applications[idx].reviewNote || "";
      saveApplications();
    }
    return true;
  }

  /** 財務審核通過（核銷子單）：此時才扣母單餘額 */
  function approveChildTicket(id) {
    const childIdx = applications.findIndex((x) => x.applicationId === id);
    if (childIdx < 0) return false;
    const child = applications[childIdx];
    if (child.voided || child.parentId == null) return false;
    if (child.status !== APPLICATION_STATUS.PENDING_APPROVAL && child.status !== "submitted") return false;

    const motherIdx = applications.findIndex((x) => x.applicationId === child.parentId);
    if (motherIdx < 0) return false;

    const inv = parseNum(String(child.writeoffInvoiceAmount ?? getApplicationTotal(child)));
    if (inv <= 0) return false;

    const mother = applications[motherIdx];
    const motherRem = getRemainingBalanceForApp(mother);
    const newRem = Math.max(0, motherRem - inv);
    applications[motherIdx].remainingBalance = newRem;

    applications[childIdx].status = APPLICATION_STATUS.PAID;
    applications[childIdx].reviewedAt = nowIso();

    if (newRem > 0) {
      applications[motherIdx].status = APPLICATION_STATUS.WRITE_OFF_PENDING;
    } else {
      // 母單未核銷餘額已歸零：核銷子單皆通過後應為「已完成」，不再卡在「待核銷」
      applications[motherIdx].status = APPLICATION_STATUS.PAID;
    }

    saveApplications();
    renderFormState();
    return true;
  }

  /** 財務審核不通過 */
  function rejectTicket(id) {
    const app = applications.find((x) => x.applicationId === id);
    if (!app || app.voided) return false;
    if (app.parentId != null) return rejectChildTicket(id);
    return rejectMotherTicket(id);
  }

  function rejectMotherTicket(id) {
    const app = applications.find((x) => x.applicationId === id);
    if (!app || app.voided || app.parentId != null) return false;
    if (app.status !== APPLICATION_STATUS.PENDING_APPROVAL && app.status !== "submitted") return false;

    const note = prompt("請輸入審核不通過原因（必填）");
    if (!note || !note.trim()) {
      alert("審核原因不可空白");
      return false;
    }

    currentApplicationId = id;
    upsertCurrentApplication(APPLICATION_STATUS.DRAFT);
    const idx = applications.findIndex((x) => x.applicationId === id);
    if (idx >= 0) {
      applications[idx].reviewNote = note.trim();
      applications[idx].reviewedAt = nowIso();
      saveApplications();
    }
    return true;
  }

  function rejectChildTicket(id) {
    const childIdx = applications.findIndex((x) => x.applicationId === id);
    if (childIdx < 0) return false;
    const child = applications[childIdx];
    if (child.voided || child.parentId == null) return false;
    if (child.status !== APPLICATION_STATUS.PENDING_APPROVAL && child.status !== "submitted") return false;

    const note = prompt("請輸入審核不通過原因（必填）");
    if (!note || !note.trim()) {
      alert("退件原因不可空白");
      return false;
    }

    applications[childIdx].status = APPLICATION_STATUS.WRITE_OFF_REJECTED;
    applications[childIdx].reviewNote = note.trim();
    applications[childIdx].reviewedAt = nowIso();

    const motherIdx = applications.findIndex((x) => x.applicationId === child.parentId);
    if (motherIdx >= 0 && applications[motherIdx].status === APPLICATION_STATUS.WRITE_OFF_AUDITING) {
      applications[motherIdx].status = APPLICATION_STATUS.WRITE_OFF_PENDING;
    }

    saveApplications();
    renderFormState();
    return true;
  }

  function openMasterCreateModal() {
    const today = new Date();
    $("#master-modal-applicant").value = $("#applicant").value || DEFAULT_APPLICANT;
    $("#master-modal-apply-date").value =
      $("#applyDate").value ||
      formatISODate(new Date(today.getFullYear(), today.getMonth(), today.getDate()));
    $("#master-create-modal").hidden = false;
    document.body.style.overflow = "hidden";
    $("#master-modal-applicant")?.focus();
  }

  function closeMasterCreateModal() {
    $("#master-create-modal").hidden = true;
    if ($("#detail-modal").hidden) document.body.style.overflow = "";
  }

  function createMasterFromModal() {
    const applicant = ($("#master-modal-applicant").value || "").trim();
    const applyDate = $("#master-modal-apply-date").value;
    if (!applicant) {
      alert("請先填寫申請人");
      $("#master-modal-applicant")?.focus();
      return;
    }
    if (!applyDate) {
      alert("請先填寫申請日");
      $("#master-modal-apply-date")?.focus();
      return;
    }
    $("#applicant").value = applicant;
    $("#applyDate").value = applyDate;
    syncExpectedPaymentDate();
    const vr = validateMasterFieldsComplete({ forSubmit: false });
    if (!vr.ok) {
      closeMasterCreateModal();
      applyMasterValidationResult(vr);
      return;
    }
    clearExpectedDateError();

    const appId = generateApplicationId();
    currentApplicationId = appId;
    const master = collectMasterForm();
    applications.unshift({
      applicationId: appId,
      createdAt: nowIso(),
      updatedAt: nowIso(),
      status: APPLICATION_STATUS.DRAFT,
      voided: false,
      parentId: null,
      type: null,
      remainingBalance: master.payCategory === PAYMENT_TYPE.PREPAY ? master.totalPayment : null,
      master,
      items: [],
    });
    saveApplications();
    setMasterCreated(true);
    setFormMeta(appId, APPLICATION_STATUS.DRAFT);
    renderFormState();
    closeMasterCreateModal();
    const tabBtn = $("#tab-btn-detail");
    tabBtn.classList.remove("hidden");
    tabBtn.setAttribute("aria-hidden", "false");
    tabBtn.click();
  }

  function initMasterCreateModal() {
    $("#master-create-modal-save").addEventListener("click", createMasterFromModal);
    $("#master-create-modal-cancel").addEventListener("click", closeMasterCreateModal);
    $("#master-create-modal-close").addEventListener("click", closeMasterCreateModal);
    $("#master-create-modal").addEventListener("click", (e) => {
      if (e.target.id === "master-create-modal") closeMasterCreateModal();
    });
    document.addEventListener("keydown", (e) => {
      const m = $("#master-create-modal");
      if (e.key === "Escape" && m && !m.hidden) closeMasterCreateModal();
    });
  }

  function showListPage() {
    $("#application-list-page").classList.remove("hidden");
    $("#application-form-page").classList.add("hidden");
    renderApplicationList();
  }

  function showFormPage() {
    $("#application-list-page").classList.add("hidden");
    $("#application-form-page").classList.remove("hidden");
  }

  /** 列表列：母單操作欄按鈕條件（集中業務規則，供 renderApplicationList 使用） */
  function renderActionButtonsForParent(parent) {
    const canEdit =
      !parent.voided &&
      (parent.status === APPLICATION_STATUS.DRAFT || parent.status === APPLICATION_STATUS.REJECTED);
    const canVoid =
      !parent.voided &&
      (parent.status === APPLICATION_STATUS.DRAFT || parent.status === APPLICATION_STATUS.REJECTED);
    const canView = !parent.voided;
    const canWriteoffGo =
      !parent.voided &&
      parent.master?.payCategory === PAYMENT_TYPE.PREPAY &&
      (parent.status === APPLICATION_STATUS.WRITE_OFF_PENDING ||
        parent.status === APPLICATION_STATUS.PARTIAL_WRITE_OFF);
    const canComplete =
      !parent.voided &&
      (parent.status === APPLICATION_STATUS.PAYMENT_PENDING ||
        parent.status === APPLICATION_STATUS.APPROVED);
    return { canEdit, canVoid, canView, canWriteoffGo, canComplete };
  }

  function getFilteredApplications() {
    const applicant = ($("#filter-applicant").value || "").trim().toLowerCase();
    const from = $("#filter-date-from").value;
    const to = $("#filter-date-to").value;
    const id = ($("#filter-id").value || "").trim().toLowerCase();
    return applications.filter((a) => {
      if (applicant && !String(a.master?.applicant || "").toLowerCase().includes(applicant)) return false;
      if (id && !String(a.applicationId || "").toLowerCase().includes(id)) return false;
      if (from && String(a.master?.applyDate || "") < from) return false;
      if (to && String(a.master?.applyDate || "") > to) return false;
      return true;
    });
  }

  function renderApplicationList() {
    const body = $("#application-list-body");
    const filtered = getFilteredApplications();
    if (!filtered.length) {
      body.innerHTML = '<tr><td colspan="11" style="color:#8c8c8c">目前無資料</td></tr>';
      return;
    }

    // ===== 先把 filtered 分組 =====
    const idSet = new Set(filtered.map((a) => String(a.applicationId)));
    const childrenByParent = {};
    filtered.forEach((a) => {
      if (a?.parentId == null) return;
      const pid = String(a.parentId).trim();
      if (!pid) return;
      childrenByParent[pid] = childrenByParent[pid] || [];
      childrenByParent[pid].push(a);
    });
    Object.keys(childrenByParent).forEach((pid) => {
      childrenByParent[pid].sort((x, y) => String(y.createdAt || "").localeCompare(String(x.createdAt || "")));
    });

    const isRoot = (a) => a?.parentId == null || !idSet.has(String(a.parentId));
    const roots = filtered.filter(isRoot);

    // 產生 parent/child HTML（child 預設收合）
    const html = [];
    roots.forEach((parent) => {
      const parentId = String(parent.applicationId);
      const children = childrenByParent[parentId] || [];
      const aHasChildren = children.length > 0;

      // parent row
      const itemsTotal = (parent.items || []).reduce((s, x) => s + parseNum(String(x.payAmount || 0)), 0);
      const masterTotal = parseNum(String(parent.master?.totalPayment ?? 0));
      const total = itemsTotal > 0 ? itemsTotal : masterTotal;
      const itemCount = (parent.items || []).length;

      const statusKey = parent.voided ? "voided" : parent.status;
      const statusClass = statusPillClass(statusKey);
      const statusText = statusToLabel(statusKey);
      const rowCls = parent.voided ? "row-voided" : "";
      const actualPaymentDate = parent.paidAt || parent.actualPaymentDate || "";

      const { canEdit, canVoid, canView, canWriteoffGo, canComplete } = renderActionButtonsForParent(parent);

      html.push(`<tr class="${rowCls} parent-row" data-row-type="parent" data-app-id="${escapeHtml(parent.applicationId)}" data-parent-id="${escapeHtml(parent.applicationId)}">
        <td>${aHasChildren ? '<span class="row-arrow" aria-hidden="true">▶</span>' : ''}${escapeHtml(parent.applicationId)}</td>
        <td>${escapeHtml(payCategoryToLabel(parent.master?.payCategory))}</td>
        <td>${escapeHtml(parent.master?.applicant || "")}</td>
        <td>${escapeHtml(parent.master?.applyDate || "")}</td>
        <td>${itemCount}</td>
        <td class="cell-num">${formatMoney(total)}</td>
        <td>${escapeHtml(formatDateTime(parent.createdAt))}</td>
        <td>${escapeHtml(parent.master?.expectedDate || "")}</td>
        <td>${escapeHtml(actualPaymentDate || "")}</td>
        <td><span class="status-pill ${statusClass}">${statusText}</span></td>
        <td class="col-actions">
          ${canEdit ? '<button type="button" class="link btn-row-edit">編輯</button>' : ''}
          ${canVoid ? '<button type="button" class="link btn-row-void">作廢</button>' : ''}
          ${canView ? '<button type="button" class="link btn-row-view">檢視</button>' : ''}
          ${canWriteoffGo ? '<button type="button" class="link btn-row-writeoff">進行核銷</button>' : ''}
          ${canComplete ? '<button type="button" class="link btn-row-complete">完成付款</button>' : ''}
        </td>
      </tr>`);

      // child rows
      children.forEach((child) => {
        const childItemsTotal = (child.items || []).reduce((s, x) => s + parseNum(String(x.payAmount || 0)), 0);
        const childMasterTotal = parseNum(String(child.master?.totalPayment ?? 0));
        const childTotal = childItemsTotal > 0 ? childItemsTotal : childMasterTotal;
        const childItemCount = (child.items || []).length;

        const childStatusKey = child.voided ? "voided" : child.status;
        const childStatusClass = statusPillClass(childStatusKey);
        const childStatusText = statusToLabel(childStatusKey);
        const childRowCls = child.voided ? "row-voided" : "";
        const childActualPaymentDate = child.paidAt || child.actualPaymentDate || "";
        const childViewBtn = !child.voided ? '<button type="button" class="link btn-row-view">檢視</button>' : "";
        html.push(`<tr class="${childRowCls} child-row is-collapsed" data-row-type="child" data-app-id="${escapeHtml(child.applicationId)}" data-parent-id="${escapeHtml(parent.applicationId)}">
          <td><span class="child-indent" aria-hidden="true">└</span>${escapeHtml(child.applicationId)}</td>
          <td>${escapeHtml(payCategoryToLabel(child.master?.payCategory))}</td>
          <td>${escapeHtml(child.master?.applicant || "")}</td>
          <td>${escapeHtml(child.master?.applyDate || "")}</td>
          <td>${childItemCount}</td>
          <td class="cell-num">${formatMoney(childTotal)}</td>
          <td>${escapeHtml(formatDateTime(child.createdAt))}</td>
          <td>${escapeHtml(child.master?.expectedDate || "")}</td>
          <td>${escapeHtml(childActualPaymentDate || "")}</td>
          <td><span class="status-pill ${childStatusClass}">${childStatusText}</span></td>
          <td class="col-actions">${childViewBtn}</td>
        </tr>`);
      });
    });

    body.innerHTML = html.join("");

    // ===== 事件綁定 =====
    // Parent：點擊切換子單顯示
    body.querySelectorAll("tr[data-row-type='parent']").forEach((tr) => {
      const parentId = tr.getAttribute("data-parent-id");
      const arrowEl = tr.querySelector(".row-arrow");

      tr.addEventListener("click", (e) => {
        if (e.target && e.target.closest && e.target.closest("button")) return;
        const safePid = String(parentId).replace(/"/g, '\\"');
        const childRows = body.querySelectorAll(`tr.child-row[data-parent-id="${safePid}"]`);
        const anyCollapsed = Array.from(childRows).some((cr) => cr.classList.contains("is-collapsed"));
        const nextCollapsed = !anyCollapsed;

        childRows.forEach((cr) => {
          cr.classList.toggle("is-collapsed", nextCollapsed);
        });
        if (arrowEl) arrowEl.textContent = nextCollapsed ? "▶" : "▽";
      });
    });

    // Buttons + child row click
    body.querySelectorAll("tr[data-app-id]").forEach((tr) => {
      const appId = tr.getAttribute("data-app-id");
      const rowType = tr.getAttribute("data-row-type");

      tr.querySelector(".btn-row-edit")?.addEventListener("click", (e) => {
        e.stopPropagation();
        openApplication(appId);
      });
      tr.querySelector(".btn-row-void")?.addEventListener("click", (e) => {
        e.stopPropagation();
        voidApplication(appId);
      });
      tr.querySelector(".btn-row-view")?.addEventListener("click", (e) => {
        e.stopPropagation();
        openApplication(appId);
      });
      tr.querySelector(".btn-row-writeoff")?.addEventListener("click", (e) => {
        e.stopPropagation();
        openApplicationForWriteoff(appId);
      });
      tr.querySelector(".btn-row-complete")?.addEventListener("click", (e) => {
        e.stopPropagation();
        completePaymentApplication(appId);
      });

      if (rowType === "child") {
        tr.addEventListener("click", (e) => {
          if (e.target && e.target.closest && e.target.closest("button")) return;
          openApplication(appId);
        });
      }
    });
  }

  function openApplication(appId) {
    const app = applications.find((x) => x.applicationId === appId);
    if (!app) return;
    currentApplicationId = app.applicationId;
    showFormPage();
    fillMasterForm(app.master || {});
    renderDetailItems(app.items || []);
    setMasterCreated(true);
    setFormMeta(app.applicationId, app.voided ? "voided" : app.status);
    renderFormState();

    // PREPAY：若處於待核銷/部分核銷，預設切到「核銷紀錄」面板
    activateWriteoffTabIfNeeded(app);
  }

  function switchToWriteoffTab() {
    const tabBtn = $("#tab-btn-writeoff");
    if (tabBtn && !tabBtn.classList.contains("hidden")) tabBtn.click();
  }

  /** 列表「進行核銷」：開啟明細並切到核銷紀錄分頁 */
  function openApplicationForWriteoff(appId) {
    openApplication(appId);
    const app = getCurrentApplication();
    if (app && isPrepayMother(app) && canShowPrepayWriteoffTab(app)) switchToWriteoffTab();
  }

  function voidApplication(appId) {
    const app = applications.find((x) => x.applicationId === appId);
    if (!app) return;
    if (app.status !== APPLICATION_STATUS.DRAFT && app.status !== APPLICATION_STATUS.REJECTED) {
      alert("已送出或進入審核流程不可作廢");
      return;
    }
    app.voided = true;
    saveApplications();
    renderApplicationList();
  }

  function completePaymentApplication(appId) {
    const app = applications.find((x) => x.applicationId === appId);
    if (!app) return;
    if (
      app.status !== APPLICATION_STATUS.PAYMENT_PENDING &&
      app.status !== APPLICATION_STATUS.APPROVED
    ) {
      alert("目前狀態不可完成付款");
      return;
    }
    const today = new Date();
    const defaultDate = formatISODate(today);
    const actualDate = prompt("請輸入實際付款日（YYYY-MM-DD）", defaultDate);
    if (!actualDate) return;
    // 基本檢核：只接受 YYYY-MM-DD
    const ok = /^\d{4}-\d{2}-\d{2}$/.test(actualDate.trim());
    if (!ok) {
      alert("日期格式錯誤，請輸入 YYYY-MM-DD");
      return;
    }

    app.paidAt = actualDate.trim();
    app.actualPaymentDate = actualDate.trim();
    // PREPAY 完成付款後，進入「待核銷」而非直接「已完成付款」
    if (app.master?.payCategory === PAYMENT_TYPE.PREPAY) {
      app.status = APPLICATION_STATUS.WRITE_OFF_PENDING;
      app.remainingBalance = getApplicationTotal(app);
    } else {
      app.status = APPLICATION_STATUS.PAID;
    }
    app.paidCompletedAt = nowIso();

    saveApplications();
    if (currentApplicationId === appId) {
      setMasterCreated(true);
      setFormMeta(app.applicationId, app.status);
      renderFormState();
    }
    showListPage();
  }

  function upsertCurrentApplication(statusOverride) {
    if (!currentApplicationId) return;
    const master = collectMasterForm();
    let items = getDetailItemsFromTable();
    const idx = applications.findIndex((x) => x.applicationId === currentApplicationId);
    if (idx < 0) return;

    const prevApp = applications[idx];
    const nextStatus = statusOverride || prevApp.status || APPLICATION_STATUS.DRAFT;
    const isPrepayMother = prevApp.parentId == null && master.payCategory === PAYMENT_TYPE.PREPAY;
    const prevStatus = prevApp.status;

    let totalPayment = items.reduce((s, x) => s + parseNum(String(x.payAmount || 0)), 0);
    const isWriteoffPhaseMother = isPrepayMother && [
      APPLICATION_STATUS.WRITE_OFF_PENDING,
      APPLICATION_STATUS.PARTIAL_WRITE_OFF,
      APPLICATION_STATUS.WRITE_OFF_AUDITING,
      APPLICATION_STATUS.WRITE_OFF_REVIEWING,
      APPLICATION_STATUS.SETTLED,
    ].includes(prevStatus);

    // PREPAY 核銷階段：避免因為 writeoff tab 切換後畫面明細為空，誤把母單 totalPayment 覆寫成 0
    if (isWriteoffPhaseMother && (!items || items.length === 0) && prevApp.items && prevApp.items.length > 0) {
      items = prevApp.items;
      totalPayment = parseNum(String(prevApp.master?.totalPayment ?? 0));
    }

    applications[idx] = {
      ...applications[idx],
      master: { ...master, totalPayment },
      items,
      status: nextStatus,
      updatedAt: nowIso(),
    };
    if (statusOverride === APPLICATION_STATUS.PENDING_APPROVAL) {
      applications[idx].submittedAt = nowIso();
    }
    // PREPAY 母單在送出前/修改期間，remainingBalance 以「總金額」初始化/更新
    if (
      isPrepayMother &&
      (prevStatus === APPLICATION_STATUS.DRAFT || prevStatus === APPLICATION_STATUS.REJECTED)
    ) {
      applications[idx].remainingBalance = totalPayment;
    }
    saveApplications();
    setFormMeta(currentApplicationId, applications[idx].status);
  }

  let modalEditingRowId = null;

  function applyTaxRuleByVoucher(voucherStyle, preferredTaxType) {
    const selTax = $("#modal-taxType");
    if (!selTax) return;
    const mode = VOUCHER_TAX_RULES[voucherStyle] || "both";
    selTax.disabled = false;
    if (mode === "fixedUntax") {
      selTax.innerHTML = '<option value="未稅">未稅</option>';
      selTax.value = "未稅";
      selTax.disabled = true;
    } else {
      selTax.innerHTML = '<option value="應稅">應稅</option><option value="未稅">未稅</option>';
      selTax.value = preferredTaxType === "未稅" ? "未稅" : "應稅";
    }
    syncAmountInputMode();
  }

  function applyPurposeToModal(purpose) {
    const rule = PURPOSE_RULES[purpose] || PURPOSE_RULES["其他費用"];
    const labelRemark = $("#modal-remark-label");
    const inpRemark = $("#modal-remarkNo");
    const periodBlock = $("#modal-period-range");
    const periodStart = $("#modal-period-start");
    const periodEnd = $("#modal-period-end");

    if (labelRemark) labelRemark.textContent = rule.remarkLabel || "備註單號";
    if (inpRemark) {
      inpRemark.placeholder = rule.remarkPlaceholder;
      inpRemark.setAttribute("aria-label", rule.remarkLabel);
    }
    if (periodBlock) {
      const usePeriod = !!rule.periodPicker;
      periodBlock.classList.toggle("hidden", !usePeriod);
      inpRemark.classList.toggle("hidden", usePeriod);
      inpRemark.readOnly = usePeriod;
      if (!usePeriod) {
        periodStart.value = "";
        periodEnd.value = "";
      } else {
        syncPeriodRangeToRemark();
      }
    }
    applyTaxRuleByVoucher($("#modal-voucherStyle").value);
    recalcModal();
  }

  function escapeHtml(s) {
    const div = document.createElement("div");
    div.textContent = s;
    return div.innerHTML;
  }

  function parseNum(v) {
    if (v == null || v === "") return 0;
    const n = parseFloat(String(v).replace(/,/g, ""));
    return Number.isFinite(n) ? n : 0;
  }

  function formatDateDisplay(isoDate) {
    return String(isoDate || "").replace(/-/g, "/");
  }

  function parsePeriodFromRemark(text) {
    const m = String(text || "").match(
      /(\d{4})[/-](\d{2})[/-](\d{2})\s*[~～-]\s*(\d{4})[/-](\d{2})[/-](\d{2})/
    );
    if (!m) return null;
    return {
      start: `${m[1]}-${m[2]}-${m[3]}`,
      end: `${m[4]}-${m[5]}-${m[6]}`,
    };
  }

  function syncPeriodRangeToRemark() {
    const block = $("#modal-period-range");
    if (!block || block.classList.contains("hidden")) return;
    const s = $("#modal-period-start").value;
    const e = $("#modal-period-end").value;
    if (s && e) {
      $("#modal-remarkNo").value = `${formatDateDisplay(s)} ~ ${formatDateDisplay(e)}`;
    } else if (!s && !e) {
      $("#modal-remarkNo").value = "";
    }
  }

  function formatMoney(n) {
    return n.toLocaleString("zh-TW", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    });
  }

  function syncAmountInputMode() {
    const taxType = $("#modal-taxType")?.value;
    const untaxedInput = $("#modal-untaxed");
    const taxInput = $("#modal-tax");
    const payInput = $("#modal-payAmount");
    if (!untaxedInput || !taxInput || !payInput) return;

    const isTaxed = taxType === "應稅";
    untaxedInput.readOnly = isTaxed;
    payInput.readOnly = !isTaxed;
    taxInput.readOnly = true;

    untaxedInput.classList.toggle("readonly-field", isTaxed);
    payInput.classList.toggle("readonly-field", !isTaxed);
    taxInput.classList.add("readonly-field");
  }

  function recalcModal() {
    const untaxedEl = $("#modal-untaxed");
    const taxEl = $("#modal-tax");
    const payEl = $("#modal-payAmount");
    const taxType = $("#modal-taxType")?.value;
    if (!untaxedEl || !taxEl || !payEl) return;

    let untaxed = parseNum(untaxedEl.value);
    let tax = 0;
    let pay = parseNum(payEl.value);

    if (taxType === "應稅") {
      pay = parseNum(payEl.value);
      untaxed = Math.round((pay / 1.05) * 100) / 100;
      tax = Math.round((pay - untaxed) * 100) / 100;
      untaxedEl.value = formatMoney(untaxed);
      taxEl.value = formatMoney(tax);
    } else {
      untaxed = parseNum(untaxedEl.value);
      tax = 0;
      pay = untaxed;
      taxEl.value = formatMoney(tax);
      payEl.value = formatMoney(pay);
    }
  }

  function updateMasterTotal() {
    let sum = 0;
    let count = 0;
    document.querySelectorAll("#detail-body tr[data-row-id]").forEach((tr) => {
      const d = tr._detail;
      if (d && d.payAmount != null) {
        sum += parseNum(String(d.payAmount));
        count += 1;
      }
    });
    const el = $("#totalPayment");
    if (el) el.value = count > 0 ? formatMoney(sum) : "";
    // 零用金上限提示需要跟著總金額更新
    recalcPettyCapHint();
  }

  function clearDetailModalForm() {
    $("#modal-purpose").value = "其他費用";
    $("#modal-remarkNo").value = "";
    $("#modal-period-start").value = "";
    $("#modal-period-end").value = "";
    $("#modal-voucherNo").value = "";
    $("#modal-untaxed").value = "";
    $("#modal-note").value = "";
    $("#modal-attachment").value = "";
    $("#modal-tax").value = "";
    applyPurposeToModal("其他費用");
    syncAmountInputMode();
  }

  function closeDetailModal() {
    const modal = $("#detail-modal");
    if (modal) modal.hidden = true;
    document.body.style.overflow = "";
  }

  function openDetailModal(editRowId) {
    const modal = $("#detail-modal");
    const title = $("#detail-modal-title");
    const isEdit = editRowId != null && editRowId !== "";
    const app = getCurrentApplication();
    if (!canEditApplication(app)) {
      alert("此單已進入審核流程，無法編輯明細");
      return;
    }

    if (isEdit) {
      const tr = document.querySelector(`#detail-body tr[data-row-id="${editRowId}"]`);
      if (!tr || !tr._detail) return;
      modalEditingRowId = Number(editRowId);
      if (title) title.textContent = "編輯明細";
      const d = tr._detail;
      $("#modal-purpose").value = d.purpose;
      applyPurposeToModal(d.purpose);
      $("#modal-remarkNo").value = d.remarkNo || "";
      if (d.periodStart || d.periodEnd) {
        $("#modal-period-start").value = d.periodStart || "";
        $("#modal-period-end").value = d.periodEnd || "";
        syncPeriodRangeToRemark();
      } else {
        const parsed = parsePeriodFromRemark(d.remarkNo);
        if (parsed) {
          $("#modal-period-start").value = parsed.start;
          $("#modal-period-end").value = parsed.end;
          syncPeriodRangeToRemark();
        }
      }
      const vs = $("#modal-voucherStyle");
      if (vs && d.voucherStyle && Array.from(vs.options).some((o) => o.value === d.voucherStyle)) {
        vs.value = d.voucherStyle;
      }
      $("#modal-voucherNo").value = d.voucherNo || "";
      applyTaxRuleByVoucher(vs ? vs.value : "", d.taxType || "應稅");
      $("#modal-untaxed").value =
        d.untaxed != null ? formatMoney(parseNum(String(d.untaxed))) : "";
      $("#modal-tax").value = d.tax != null ? formatMoney(parseNum(String(d.tax))) : "";
      $("#modal-note").value = d.note || "";
      $("#modal-attachment").value = "";
      recalcModal();
    } else {
      modalEditingRowId = null;
      if (title) title.textContent = "新增明細";
      clearDetailModalForm();
    }

    // PREPAY：套用憑證號碼顯示/隱藏與憑證樣式限制
    applyScenarioToDetailModalFields();

    if (modal) modal.hidden = false;
    document.body.style.overflow = "hidden";
    $("#modal-purpose")?.focus();
  }

  function collectDetailFromModal() {
    const purpose = $("#modal-purpose").value;
    const rule = PURPOSE_RULES[purpose] || PURPOSE_RULES["其他費用"];
    const periodStart = $("#modal-period-start").value;
    const periodEnd = $("#modal-period-end").value;
    if (rule.periodPicker) {
      syncPeriodRangeToRemark();
    }
    return {
      purpose,
      remarkNo: ($("#modal-remarkNo").value || "").trim(),
      periodStart: rule.periodPicker ? periodStart : "",
      periodEnd: rule.periodPicker ? periodEnd : "",
      voucherStyle: $("#modal-voucherStyle").value,
      voucherNo: ($("#modal-voucherNo").value || "").trim(),
      taxType: $("#modal-taxType").value,
      untaxed: parseNum($("#modal-untaxed").value),
      tax: parseNum($("#modal-tax").value),
      payAmount: parseNum($("#modal-payAmount").value),
      note: ($("#modal-note").value || "").trim(),
      fileName: $("#modal-attachment").files?.[0]?.name || "",
    };
  }

  function saveDetailFromModal() {
    const app = getCurrentApplication();
    if (!canEditApplication(app)) {
      alert("此單已進入審核流程，無法儲存明細");
      return;
    }
    const d = collectDetailFromModal();
    const rule = PURPOSE_RULES[d.purpose] || PURPOSE_RULES["其他費用"];
    if (rule.periodPicker) {
      if (!d.periodStart || !d.periodEnd) {
        alert("請選擇期間起訖日期");
        $("#modal-period-start")?.focus();
        return;
      }
      if (d.periodStart > d.periodEnd) {
        alert("期間起日不可晚於迄日");
        $("#modal-period-start")?.focus();
        return;
      }
    }
    // 預付（PREPAY）母單：憑證號碼不必填（核銷補單時會在 writeoff modal 收集）
    const isPrepay = app?.master?.payCategory === PAYMENT_TYPE.PREPAY;
    if (!d.voucherNo && !isPrepay) {
      const msg = d.voucherStyle === "發票" ? "請填寫發票號碼" : "請填寫憑證號碼";
      alert(msg);
      $("#modal-voucherNo")?.focus();
      return;
    }
    if (isPrepay) d.voucherNo = "";
    if (d.untaxed <= 0) {
      alert("請填寫未稅金額（須大於 0）");
      $("#modal-untaxed")?.focus();
      return;
    }

    // 零用金：總金額上限防呆（包含編輯情境時的精確加總）
    if (getScenario() === PAYMENT_TYPE.PETTY_CASH) {
      let sumOther = 0;
      document
        .querySelectorAll("#detail-body tr[data-row-id]")
        .forEach((tr) => {
          const rid = tr.dataset.rowId;
          if (modalEditingRowId != null && String(rid) === String(modalEditingRowId)) return;
          if (tr._detail && tr._detail.payAmount != null) sumOther += parseNum(String(tr._detail.payAmount));
        });
      const nextSum = sumOther + (d.payAmount || 0);
      if (nextSum > pettyCashLimit) {
        const msg = `超過零用金上限 ${pettyCashLimit.toLocaleString("zh-TW")} 元，請改走一般廠商付款流程，或先走專案簽呈。`;
        $("#petty-amount-error").textContent = msg;
        $("#petty-amount-error").classList.remove("hidden");
        alert(msg);
        return;
      }
    }

    if (modalEditingRowId != null) {
      const tr = document.querySelector(`#detail-body tr[data-row-id="${modalEditingRowId}"]`);
      if (tr) {
        fillSummaryRow(tr, modalEditingRowId, d);
        tr._detail = d;
        bindSummaryRowActions(tr);
      }
    } else {
      const id = nextId++;
      const tr = document.createElement("tr");
      tr.dataset.rowId = String(id);
      tr._detail = d;
      fillSummaryRow(tr, id, d);
      $("#detail-body").appendChild(tr);
      bindSummaryRowActions(tr);
    }

    closeDetailModal();
    updateDetailCount();
    updateMasterTotal();
    if (isMasterCreated && currentApplicationId) upsertCurrentApplication();
  }

  function fillSummaryRow(tr, id, d) {
    const att = d.fileName ? escapeHtml(d.fileName) : "—";
    tr.innerHTML = `
      <td>${id}</td>
      <td>${escapeHtml(d.purpose)}</td>
      <td>${escapeHtml(d.remarkNo || "")}</td>
      <td>${escapeHtml(d.voucherStyle)}</td>
      <td>${escapeHtml(d.voucherNo)}</td>
      <td>${escapeHtml(d.taxType)}</td>
      <td class="cell-num">${formatMoney(d.untaxed)}</td>
      <td class="cell-num">${formatMoney(d.tax)}</td>
      <td class="cell-num">${formatMoney(d.payAmount)}</td>
      <td>${escapeHtml(d.note || "")}</td>
      <td>${att}</td>
      <td class="col-actions">
        <button type="button" class="link btn-edit-row">編輯</button>
        <button type="button" class="link btn-del-row">刪除</button>
      </td>
    `;
  }

  function bindSummaryRowActions(tr) {
    const rid = tr.dataset.rowId;
    const editBtn = tr.querySelector(".btn-edit-row");
    const delBtn = tr.querySelector(".btn-del-row");
    const app = getCurrentApplication();
    const canEdit = canEditApplication(app);
    if (editBtn) {
      editBtn.classList.toggle("hidden", !canEdit);
      editBtn.addEventListener("click", () => openDetailModal(rid));
    }
    if (delBtn) {
      delBtn.classList.toggle("hidden", !canEdit);
      delBtn.addEventListener("click", () => {
        if (!canEdit) return;
        tr.remove();
        updateDetailCount();
        updateMasterTotal();
        if (isMasterCreated && currentApplicationId) upsertCurrentApplication();
      });
    }
  }

  function initDetailModal() {
    const selPurpose = $("#modal-purpose");
    selPurpose.innerHTML = PURPOSE_OPTIONS.map(
      (p) => `<option value="${escapeHtml(p)}">${escapeHtml(p)}</option>`
    ).join("");
    const selVoucher = $("#modal-voucherStyle");
    selVoucher.innerHTML = VOUCHER_OPTIONS.map(
      (v) => `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`
    ).join("");
    selPurpose.addEventListener("change", () => {
      applyPurposeToModal(selPurpose.value);
    });
    $("#modal-voucherStyle").addEventListener("change", () => {
      const prevTax = $("#modal-taxType").value;
      applyTaxRuleByVoucher($("#modal-voucherStyle").value, prevTax);
      recalcModal();
    });
    $("#modal-period-start").addEventListener("change", syncPeriodRangeToRemark);
    $("#modal-period-end").addEventListener("change", syncPeriodRangeToRemark);

    $("#modal-untaxed").addEventListener("input", () => recalcModal());
    $("#modal-payAmount").addEventListener("input", () => recalcModal());
    $("#modal-taxType").addEventListener("change", () => {
      syncAmountInputMode();
      recalcModal();
    });

    $("#detail-modal-save").addEventListener("click", saveDetailFromModal);
    $("#detail-modal-cancel").addEventListener("click", closeDetailModal);
    $("#detail-modal-close").addEventListener("click", closeDetailModal);
    $("#detail-modal").addEventListener("click", (e) => {
      if (e.target.id === "detail-modal") closeDetailModal();
    });
    document.addEventListener("keydown", (e) => {
      const m = $("#detail-modal");
      if (e.key === "Escape" && m && !m.hidden) {
        closeDetailModal();
      }
    });
  }

  function updateDetailCount() {
    const n = document.querySelectorAll("#detail-body tr[data-row-id]").length;
    const el = $("#detail-total");
    if (el) el.textContent = String(n);
  }

  function initTabs() {
    const tabs = document.querySelectorAll(".tab");
    const panels = document.querySelectorAll(".tab-panel");
    tabs.forEach((tab) => {
      tab.addEventListener("click", () => {
        const target = tab.getAttribute("data-tab");
        tabs.forEach((t) => t.classList.toggle("active", t === tab));
        panels.forEach((p) => {
          const isActive = p.getAttribute("data-panel") === target;
          p.classList.toggle("active", isActive);
          p.classList.toggle("hidden", !isActive);
        });
        updateTabsSubmitVisibility();
      });
    });
  }

  /** 建立主單後才顯示明細頁籤，並切換至該頁籤 */
  function initRevealDetailTab() {
    const btn = $("#btn-edit-detail");
    const tabBtn = $("#tab-btn-detail");
    if (!btn || !tabBtn) return;
    btn.addEventListener("click", () => {
      if (!isMasterCreated) {
        const applicant = ($("#applicant").value || DEFAULT_APPLICANT).trim();
        let applyDate = $("#applyDate").value;
        if (!applyDate) {
          const today = new Date();
          applyDate = formatISODate(
            new Date(today.getFullYear(), today.getMonth(), today.getDate())
          );
          $("#applyDate").value = applyDate;
        }
        $("#applicant").value = applicant || DEFAULT_APPLICANT;

        syncExpectedPaymentDate();
        const vr = validateMasterFieldsComplete({ forSubmit: false });
        if (!vr.ok) {
          applyMasterValidationResult(vr);
          return;
        }
        clearExpectedDateError();

        currentApplicationId = generateApplicationId();
        const master = collectMasterForm();
        applications.unshift({
          applicationId: currentApplicationId,
          createdAt: nowIso(),
          updatedAt: nowIso(),
          status: APPLICATION_STATUS.DRAFT,
          voided: false,
          parentId: null,
          type: null,
          remainingBalance: master.payCategory === PAYMENT_TYPE.PREPAY ? master.totalPayment : null,
          master,
          items: [],
        });
        saveApplications();
        setMasterCreated(true);
        setFormMeta(currentApplicationId, APPLICATION_STATUS.DRAFT);
        renderFormState();
      }

      if (tabBtn.classList.contains("hidden")) {
        tabBtn.classList.remove("hidden");
        tabBtn.setAttribute("aria-hidden", "false");
      }
      tabBtn.click();
    });
  }

  function filterCombo(q) {
    const qq = (q || "").trim().toLowerCase();
    if (!qq) return vendors.slice(0, 50);
    return vendors.filter((v) => v.name.toLowerCase().includes(qq)).slice(0, 50);
  }

  function renderComboList(items) {
    const ul = $("#combo-list");
    if (!items.length) {
      ul.innerHTML =
        '<li style="cursor:default;color:#8c8c8c">無符合項目，請上傳設定檔（需含 Customer Name）或調整關鍵字</li>';
      comboHighlight = -1;
      return;
    }
    ul.innerHTML = items
      .map(
        (v, i) =>
          `<li data-index="${i}" data-name="${escapeHtml(v.name)}">${escapeHtml(v.name)}</li>`
      )
      .join("");
    comboHighlight = items.length ? 0 : -1;
    updateHighlight();
  }

  function updateHighlight() {
    const lis = document.querySelectorAll("#combo-list li");
    lis.forEach((li, i) => li.classList.toggle("highlight", i === comboHighlight));
  }

  function openCombo() {
    comboOpen = true;
    $("#combo-list").classList.add("open");
    renderComboList(filterCombo($("#vendor-search").value));
  }

  function closeCombo() {
    comboOpen = false;
    $("#combo-list").classList.remove("open");
  }

  function selectVendor(name) {
    $("#vendor-search").value = name;
    $("#vendor-hidden").value = name;
    closeCombo();
  }

  function initVendorCombo() {
    const input = $("#vendor-search");
    const ul = $("#combo-list");

    input.addEventListener("focus", () => openCombo());
    input.addEventListener("input", () => {
      renderComboList(filterCombo(input.value));
      openCombo();
    });

    ul.addEventListener("mousedown", (e) => {
      const li = e.target.closest("li");
      if (!li) return;
      selectVendor(li.getAttribute("data-name"));
    });

    document.addEventListener("click", (e) => {
      if (!e.target.closest(".combo-wrap")) closeCombo();
    });

    input.addEventListener("keydown", (e) => {
      const items = document.querySelectorAll("#combo-list li");
      if (!items.length) return;
      if (e.key === "ArrowDown") {
        e.preventDefault();
        comboHighlight = Math.min(comboHighlight + 1, items.length - 1);
        updateHighlight();
      } else if (e.key === "ArrowUp") {
        e.preventDefault();
        comboHighlight = Math.max(comboHighlight - 1, 0);
        updateHighlight();
      } else if (e.key === "Enter" && comboHighlight >= 0) {
        e.preventDefault();
        const li = items[comboHighlight];
        if (li) selectVendor(li.getAttribute("data-name"));
      }
    });
  }

  function updateExpectedDateCurrentRule() {
    const el = $("#expected-date-current-rule");
    if (!el) return;
    const type = $("#payCategory").value;
    const lines = {
      [PAYMENT_TYPE.GENERAL]:
        "目前類別「一般」：依每月截止規則自動帶入（不可手動修改）。",
      [PAYMENT_TYPE.PREPAY]:
        "目前類別「預付款」：請在申請日起 5 日內（含）選擇預計付款日。",
      [PAYMENT_TYPE.PETTY_CASH]:
        "目前類別「零用金」：預計付款日已依「下一個週四」規則自動帶入（不可手動修改）。",
      [PAYMENT_TYPE.URGENT]:
        "目前類別「急件付款」：請在申請日起 5 日內（含）選擇預計付款日。",
    };
    el.textContent = lines[type] || "";
  }

  function clearExpectedDateError() {
    const err = $("#expected-date-error");
    const inp = $("#expectedDate");
    if (err) {
      err.textContent = "";
      err.classList.add("hidden");
    }
    if (inp) inp.classList.remove("input-error");
  }

  function showExpectedDateError(msg) {
    const err = $("#expected-date-error");
    const inp = $("#expectedDate");
    if (err) {
      err.textContent = msg;
      err.classList.remove("hidden");
    }
    if (inp) inp.classList.add("input-error");
  }

  function syncExpectedPaymentDate() {
    const type = $("#payCategory").value;
    const apply = parseInputDate($("#applyDate").value);
    const exp = $("#expectedDate");
    if (!exp) return;
    if (!apply) {
      exp.value = "";
      updateExpectedDateCurrentRule();
      return;
    }
    if (type === PAYMENT_TYPE.GENERAL) {
      exp.value = formatISODate(calcGeneralExpectedDate(apply));
      exp.readOnly = true;
      exp.classList.add("readonly-field");
      exp.removeAttribute("min");
      exp.removeAttribute("max");
    } else if (type === PAYMENT_TYPE.PETTY_CASH) {
      exp.value = formatISODate(nextPettyCashThursday(apply));
      exp.readOnly = true;
      exp.classList.add("readonly-field");
      exp.removeAttribute("min");
      exp.removeAttribute("max");
    } else if (type === PAYMENT_TYPE.PREPAY) {
      // 預付款：須在申請日起 5 日內（含）
      exp.readOnly = false;
      exp.classList.remove("readonly-field");
      const min = formatISODate(apply);
      const max = formatISODate(addDays(apply, 5));
      exp.min = min;
      exp.max = max;
      if (!exp.value || exp.value < min || exp.value > max) {
        exp.value = min;
      }
    } else if (type === PAYMENT_TYPE.URGENT) {
      exp.readOnly = false;
      exp.classList.remove("readonly-field");
      const min = formatISODate(apply);
      const max = formatISODate(addDays(apply, 5));
      exp.min = min;
      exp.max = max;
      if (!exp.value || exp.value < min || exp.value > max) {
        exp.value = min;
      }
    }
    clearExpectedDateError();
    updateExpectedDateCurrentRule();
  }

  function validateExpectedPaymentDate() {
    const type = $("#payCategory").value;
    const apply = parseInputDate($("#applyDate").value);
    const expStr = $("#expectedDate").value;
    const exp = parseInputDate(expStr);
    if (!apply) return { ok: false, msg: "請選擇申請日" };
    if (!expStr || !exp) return { ok: false, msg: "請填寫預計付款日" };
    if (
      type !== PAYMENT_TYPE.GENERAL &&
      type !== PAYMENT_TYPE.PETTY_CASH &&
      type !== PAYMENT_TYPE.PREPAY &&
      type !== PAYMENT_TYPE.URGENT
    ) {
      return { ok: false, msg: "無效的付款類別" };
    }
    if (type === PAYMENT_TYPE.URGENT || type === PAYMENT_TYPE.PREPAY) {
      const max = addDays(apply, 5);
      if (exp < apply || exp > max) {
        const label = type === PAYMENT_TYPE.PREPAY ? "預付款" : "急件付款";
        return {
          ok: false,
          msg: `${label}：預計付款日須在申請日起 5 日內（含）`,
        };
      }
    }
    if (type === PAYMENT_TYPE.GENERAL) {
      const want = calcGeneralExpectedDate(apply);
      if (formatISODate(exp) !== formatISODate(want)) {
        return {
          ok: false,
          msg:
            "一般付款：預計付款日依規則計算失敗（每月21~次月5=>次月15；每月6~20=>當月30）",
        };
      }
    }
    if (type === PAYMENT_TYPE.PETTY_CASH) {
      const want = nextPettyCashThursday(apply);
      if (formatISODate(exp) !== formatISODate(want)) {
        return {
          ok: false,
          msg: "零用金：預計付款日須為下一個週四（申請日若為週四則為下週四）",
        };
      }
    }
    return { ok: true };
  }

  /** 主單欄位必填檢核（建立前／送出前共用） */
  function validateMasterFieldsComplete({ forSubmit } = {}) {
    const applicant = ($("#applicant").value || "").trim();
    if (!applicant) return { ok: false, msg: "請填寫申請人", focus: "#applicant" };
    const applyDate = $("#applyDate").value;
    if (!applyDate) return { ok: false, msg: "請填寫申請日", focus: "#applyDate" };
    const payCategory = $("#payCategory").value;
    if (!payCategory) return { ok: false, msg: "請選擇申請款項情境", focus: "#payCategory" };
    const vendor = ($("#vendor-hidden").value || $("#vendor-search").value || "").trim();
    if (!vendor) return { ok: false, msg: "請選擇或輸入付款名稱（供應商）", focus: "#vendor-search" };
    const currency = $("#currency").value;
    if (!currency) return { ok: false, msg: "請選擇付款幣別", focus: "#currency" };
    const payMethod = $("#payMethod").value;
    if (!payMethod) return { ok: false, msg: "請選擇付款方式", focus: "#payMethod" };
    const transferFee = $("#transferFee").value;
    if (!transferFee) return { ok: false, msg: "請選擇匯款手續費", focus: "#transferFee" };

    const dt = validateExpectedPaymentDate();
    if (!dt.ok) return { ok: false, msg: dt.msg, focus: "#expectedDate" };

    if (forSubmit) {
      const total = parseNum($("#totalPayment").value || "");
      if (!total || total <= 0) {
        return { ok: false, msg: "總付款金額須大於 0，請先完成款項憑證明細", focus: "#totalPayment" };
      }
      const n = document.querySelectorAll("#detail-body tr[data-row-id]").length;
      if (n === 0) return { ok: false, msg: "請至少新增一筆款項憑證明細", focusDetail: true };
    }
    return { ok: true };
  }

  function applyMasterValidationResult(r) {
    if (r.ok) return true;
    alert(r.msg);
    if (r.focusDetail) {
      const tabBtn = $("#tab-btn-detail");
      if (tabBtn && !tabBtn.classList.contains("hidden")) tabBtn.click();
      $("#panel-detail")?.scrollIntoView({ behavior: "smooth", block: "start" });
      $("#btn-add-row")?.focus();
      return false;
    }
    $("#panel-master")?.scrollIntoView({ behavior: "smooth", block: "start" });
    if (r.focus === "#expectedDate") showExpectedDateError(r.msg);
    else clearExpectedDateError();
    if (r.focus) $(r.focus)?.focus();
    return false;
  }

  function onExpectedDateUserInput() {
    const type = $("#payCategory").value;
    if (type !== PAYMENT_TYPE.URGENT && type !== PAYMENT_TYPE.PREPAY) return;
    const r = validateExpectedPaymentDate();
    if (!r.ok) showExpectedDateError(r.msg);
    else clearExpectedDateError();
  }

  function initPaymentDateRules() {
    const sel = $("#payCategory");
    const apply = $("#applyDate");
    const exp = $("#expectedDate");
    const today = new Date();
    apply.value = formatISODate(
      new Date(today.getFullYear(), today.getMonth(), today.getDate())
    );
    sel.addEventListener("change", syncExpectedPaymentDate);
    sel.addEventListener("change", applyScenarioRules);
    apply.addEventListener("change", syncExpectedPaymentDate);
    exp.addEventListener("change", onExpectedDateUserInput);
    exp.addEventListener("blur", onExpectedDateUserInput);
    syncExpectedPaymentDate();
    applyScenarioRules();
  }

  function getScenario() {
    return $("#payCategory")?.value || PAYMENT_TYPE.GENERAL;
  }

  function getEmployeeName() {
    const v = $("#applicant")?.value || "";
    return v.trim() || DEFAULT_APPLICANT;
  }

  function setVendorLocked(name) {
    const input = $("#vendor-search");
    const hidden = $("#vendor-hidden");
    if (input) {
      input.value = name;
      input.setAttribute("disabled", "true");
    }
    if (hidden) hidden.value = name;
    closeCombo();
  }

  function setVendorUnlocked() {
    const input = $("#vendor-search");
    if (input) {
      input.removeAttribute("disabled");
    }
  }

  function setPayMethodLocked(methodValue) {
    const sel = $("#payMethod");
    if (!sel) return;
    sel.value = methodValue;
    sel.setAttribute("disabled", "true");
    // re-sync transfer fee row
    const row = $("#row-transfer-fee");
    const lab = $("#label-transfer-fee");
    if (row) row.classList.toggle("hidden", methodValue !== "匯款");
    if (lab) lab.classList.toggle("hidden", methodValue !== "匯款");
  }

  function setPayMethodUnlocked() {
    const sel = $("#payMethod");
    if (!sel) return;
    sel.removeAttribute("disabled");
    // sync transfer fee row based on current value
    const methodValue = sel.value;
    const row = $("#row-transfer-fee");
    const lab = $("#label-transfer-fee");
    if (row) row.classList.toggle("hidden", methodValue !== "匯款");
    if (lab) lab.classList.toggle("hidden", methodValue !== "匯款");
  }

  function applyScenarioInfo() {
    const box = $("#scenario-info");
    const v = getScenario();
    if (!box) return;
    if (v === PAYMENT_TYPE.PETTY_CASH) {
      box.querySelector(".rule-title").textContent = "個人代墊報支（零用金）";
      box.innerHTML = "";
      box.insertAdjacentHTML(
        "afterbegin",
        '<span class="rule-title">個人代墊報支（零用金）</span><div style="margin-top:4px">我已先用個人款項墊付餐費、車資、雜支等，請公司退款至我的薪資帳戶。</div>'
      );
    } else if (v === PAYMENT_TYPE.PREPAY) {
      box.querySelector(".rule-title").textContent = "廠商預付／訂金（預付款）";
      box.innerHTML = "";
      box.insertAdjacentHTML(
        "afterbegin",
        '<span class="rule-title">廠商預付／訂金（預付款）</span><div style="margin-top:4px">目前僅有合約或報價單，需先申請款項支付給廠商，事後再補發票核銷。</div>'
      );
    } else {
      // GENERAL
      box.innerHTML =
        '<span class="rule-title">廠商實報實銷（一般付款）</span><div style="margin-top:4px">已取得廠商開立的正式發票或收據，請公司直接匯款給該廠商。</div>';
    }
  }

  function applyScenarioRules() {
    // 1) 以情境進行畫面防呆（鎖定/解鎖）
    const v = getScenario();
    const applicantName = getEmployeeName();

    $("#petty-amount-error")?.classList.add("hidden");
    $("#petty-amount-error").textContent = "";

    if (v === PAYMENT_TYPE.PETTY_CASH) {
      // 強制鎖定收款人/付款名稱
      setVendorLocked(applicantName);
      // 強制鎖定付款方式：匯款
      setPayMethodLocked("匯款");
    } else {
      setVendorUnlocked();
      setPayMethodUnlocked();
    }

    if (v === PAYMENT_TYPE.PREPAY) {
      // 預付款目前未要求額外防呆，沿用一般付款的畫面自由度
      setVendorUnlocked();
      setPayMethodUnlocked();
    }

    applyScenarioInfo();
    recalcPettyCapHint();
    applyScenarioToDetailModalFields();
  }

  function applyScenarioToDetailModalFields() {
    const app = getCurrentApplication();
    const isPrepay = !!app?.master && app.master.payCategory === PAYMENT_TYPE.PREPAY;

    const voucherNoLabel = document.querySelector('label[for="modal-voucherNo"]');
    const voucherNoInput = $("#modal-voucherNo");
    const selVoucherStyle = $("#modal-voucherStyle");
    if (!voucherNoLabel || !voucherNoInput || !selVoucherStyle) return;

    // PREPAY：隱藏/灰化「憑證號碼」，且限制憑證樣式只能「合約/報價單」
    if (isPrepay) {
      if (!voucherNoLabel.dataset.originalText) voucherNoLabel.dataset.originalText = voucherNoLabel.textContent;
      voucherNoLabel.textContent = "憑證號碼（預付款無需填寫）";
      voucherNoLabel.style.opacity = "0.6";
      voucherNoInput.readOnly = true;
      voucherNoInput.classList.add("readonly-field");
      voucherNoInput.placeholder = "預付款無需填寫";
      voucherNoInput.value = "";

      const allowed = ["合約", "報價單"];
      selVoucherStyle.innerHTML = allowed
        .map((v) => `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`)
        .join("");
      if (!allowed.includes(selVoucherStyle.value)) {
        selVoucherStyle.value = allowed[0];
      }
      applyTaxRuleByVoucher(selVoucherStyle.value);
      recalcModal();
    } else {
      if (voucherNoLabel.dataset.originalText) voucherNoLabel.textContent = voucherNoLabel.dataset.originalText;
      voucherNoLabel.style.opacity = "";
      voucherNoInput.readOnly = false;
      voucherNoInput.classList.remove("readonly-field");
      voucherNoInput.placeholder = "";

      // 還原憑證樣式選項
      selVoucherStyle.innerHTML = VOUCHER_OPTIONS.map(
        (v) => `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`
      ).join("");
      recalcModal();
    }
  }

  function recalcPettyCapHint() {
    const v = getScenario();
    if (v !== PAYMENT_TYPE.PETTY_CASH) return;
    const sum = parseNum($("#totalPayment").value || "");
    if (sum > pettyCashLimit) {
      $("#petty-amount-error").textContent = `超過零用金上限 ${pettyCashLimit.toLocaleString("zh-TW")} 元，請改走一般廠商付款流程或專案簽呈。`;
      $("#petty-amount-error").classList.remove("hidden");
    } else {
      $("#petty-amount-error").classList.add("hidden");
      $("#petty-amount-error").textContent = "";
    }
  }

  function initPaymentMethod() {
    const sel = $("#payMethod");
    const row = $("#row-transfer-fee");
    const lab = $("#label-transfer-fee");
    function sync() {
      const isTransfer = sel.value === "匯款";
      row.classList.toggle("hidden", !isTransfer);
      if (lab) lab.classList.toggle("hidden", !isTransfer);
    }
    if (!sel.dataset.boundPaymentMethod) {
      sel.addEventListener("change", sync);
      sel.dataset.boundPaymentMethod = "1";
    }
    sync();
  }

  function initFileUpload() {
    $("#file-config").addEventListener("change", (e) => {
      const file = e.target.files && e.target.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = () => {
        try {
          const text = String(reader.result);
          let list;
          if (/\.json$/i.test(file.name)) {
            list = parseVendorJson(text);
          } else if (/\.csv$/i.test(file.name)) {
            list = parseVendorCsv(text);
          } else {
            try {
              list = parseVendorJson(text);
            } catch {
              list = parseVendorCsv(text);
            }
          }
          vendors = list;
          saveVendorsToStorage();
          updateVendorBadge();
          $("#config-msg").textContent = `已載入 ${vendors.length} 筆付款名稱。`;
        } catch (err) {
          $("#config-msg").textContent = "解析失敗：" + err.message;
        }
      };
      reader.readAsText(file, "UTF-8");
      e.target.value = "";
    });

    $("#btn-download-sample").addEventListener("click", async () => {
      const fallback = () => {
        const sample = {
          version: "1.0",
          vendors: [
            { code: "V001", "Customer Name": "台灣理光股份有限公司" },
            { "Customer Name": "Ur-09 範例廠商" },
          ],
        };
        const blob = new Blob([JSON.stringify(sample, null, 2)], {
          type: "application/json;charset=utf-8",
        });
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = "payment-vendors.example.json";
        a.click();
        URL.revokeObjectURL(a.href);
      };
      try {
        const res = await fetch("config/payment-vendors.example.json");
        if (!res.ok) throw new Error("fetch failed");
        const text = await res.text();
        const blob = new Blob([text], { type: "application/json;charset=utf-8" });
        const a = document.createElement("a");
        a.href = URL.createObjectURL(blob);
        a.download = "payment-vendors.example.json";
        a.click();
        URL.revokeObjectURL(a.href);
      } catch {
        fallback();
      }
    });
  }

  function resetFormForNew() {
    $("#applicant").value = DEFAULT_APPLICANT;
    $("#vendor-search").value = "";
    $("#vendor-hidden").value = "";
    $("#payCategory").value = PAYMENT_TYPE.GENERAL;
    $("#currency").value = "臺幣TWD";
    $("#payMethod").value = "匯款";
    $("#transferFee").value = "公司負擔";
    const today = new Date();
    $("#applyDate").value = formatISODate(
      new Date(today.getFullYear(), today.getMonth(), today.getDate())
    );
    syncExpectedPaymentDate();
    applyScenarioRules();
    clearExpectedDateError();
    clearDetailTable();
    setMasterCreated(false);
    setFormMeta("", "draft");
    // 新建/切換時先把 PREPAY 核銷 UI 隱藏，避免殘留
    applyWriteoffUI(null);
    initPaymentMethod();
    const masterTab = document.querySelector('.tab[data-tab="master"]');
    if (masterTab) masterTab.click();
    renderFormState();
  }

  function initMasterAutoSave() {
    const fields = [
      "#applicant",
      "#applyDate",
      "#payCategory",
      "#vendor-search",
      "#currency",
      "#payMethod",
      "#transferFee",
      "#expectedDate",
    ];
    fields.forEach((sel) => {
      $(sel).addEventListener("change", () => {
        if (isMasterCreated && currentApplicationId) upsertCurrentApplication();
      });
    });
  }

  function initActions() {
    $("#btn-add-row").addEventListener("click", () => {
      if (!isMasterCreated) {
        alert("請先建立主單，再新增明細");
        return;
      }
      const app = getCurrentApplication();
      if (!canEditApplication(app)) {
        alert("此單已進入審核流程，無法編輯明細");
        return;
      }
      openDetailModal();
    });
    $("#btn-submit").addEventListener("click", () => {
      if (!isMasterCreated || !currentApplicationId) {
        alert("請先建立主單");
        return;
      }
      const app = getCurrentApplication();
      if (!canEditApplication(app)) {
        alert("目前狀態不可送出審核");
        return;
      }

      const vr = validateMasterFieldsComplete({ forSubmit: true });
      if (!vr.ok) {
        applyMasterValidationResult(vr);
        return;
      }
      clearExpectedDateError();

      // 零用金上限再做一次保護（避免繞過明細儲存防呆）
      if (getScenario() === PAYMENT_TYPE.PETTY_CASH) {
        const sum = parseNum($("#totalPayment").value || "");
        if (sum > pettyCashLimit) {
          const msg = `超過零用金上限 ${pettyCashLimit.toLocaleString("zh-TW")} 元，請改走一般廠商付款流程，或專案簽呈。`;
          $("#petty-amount-error").textContent = msg;
          $("#petty-amount-error").classList.remove("hidden");
          alert(msg);
          return;
        }
      }
      upsertCurrentApplication(APPLICATION_STATUS.PENDING_APPROVAL);
      alert("已送出申請，待財務審核。\n總金額：" + ($("#totalPayment").value || ""));
      showListPage();
    });

    $("#btn-review-approve")?.addEventListener("click", () => {
      const app = getCurrentApplication();
      if (!app || app.voided) return;
      if (!approveTicket(app.applicationId)) return;
      alert("審核通過");
      showListPage();
    });

    $("#btn-review-reject")?.addEventListener("click", () => {
      const app = getCurrentApplication();
      if (!app || app.voided) return;
      if (!rejectTicket(app.applicationId)) return;
      alert(app.parentId != null ? "已退件" : "已退回草稿（可重新修改並送出）");
      showListPage();
    });


    $("#btn-back-list").addEventListener("click", () => {
      if (isMasterCreated && currentApplicationId) upsertCurrentApplication();
      showListPage();
    });
    $("#btn-new-application").addEventListener("click", () => {
      currentApplicationId = null;
      showFormPage();
      resetFormForNew();
    });
    $("#btn-filter-reset").addEventListener("click", () => {
      $("#filter-applicant").value = "";
      $("#filter-date-from").value = "";
      $("#filter-date-to").value = "";
      $("#filter-id").value = "";
      renderApplicationList();
    });
    ["#filter-applicant", "#filter-date-from", "#filter-date-to", "#filter-id"].forEach((sel) => {
      $(sel).addEventListener("input", renderApplicationList);
      $(sel).addEventListener("change", renderApplicationList);
    });

    // PREPAY：核銷浮窗事件綁定
    $("#btn-add-writeoff")?.addEventListener("click", () => openWriteoffModal());
    $("#writeoff-modal-submit")?.addEventListener("click", () => submitWriteoff());
    $("#writeoff-modal-cancel")?.addEventListener("click", () => closeWriteoffModal());
    $("#writeoff-modal-close")?.addEventListener("click", () => closeWriteoffModal());
    $("#writeoff-actual-invoice")?.addEventListener("input", () => recalcWriteoffDiffHint());
    $("#writeoff-is-final")?.addEventListener("change", () => recalcWriteoffDiffHint());
    $("#writeoff-source-detail")?.addEventListener("change", () => recalcWriteoffDiffHint());
    $("#btn-review-writeoff-approve")?.addEventListener("click", () => {
      showListPage();
    });
    document.addEventListener("keydown", (e) => {
      const modal = $("#writeoff-modal");
      if (e.key === "Escape" && modal && !modal.hidden) closeWriteoffModal();
    });
  }

  function init() {
    if (typeof console !== "undefined" && console.info) {
      console.info(`[付款申請模組] 版本 ${APP_VERSION}`);
    }
    loadVendorsFromStorage();
    loadPettyCashLimitFromStorage();
    loadApplications();
    updateVendorBadge();
    syncPettyCashLimitInput();
    if (!vendors.length) {
      $("#config-msg").textContent =
        "尚未上傳設定檔。請上傳 JSON／CSV，或使用範例檔格式準備資料。";
    } else {
      $("#config-msg").textContent = `已從瀏覽器讀取 ${vendors.length} 筆付款名稱（可重新上傳覆寫）。`;
    }

    initTabs();
    initRevealDetailTab();
    initVendorCombo();
    initPaymentDateRules();
    initPaymentMethod();
    initFileUpload();
    initDetailModal();
    initMasterCreateModal();
    initPettyCashLimitSetting();
    initMasterAutoSave();
    initActions();
    resetFormForNew();
    showListPage();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
