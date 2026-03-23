/**
 * 付款申請原型：主檔／明細頁籤、款項用途連動、付款名稱來自設定檔
 */
(function () {
  const STORAGE_KEY = "payment_vendor_config_v1";
  const APPLICATIONS_KEY = "payment_applications_v1";
  const DEFAULT_APPLICANT = "ruby.lee";

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
  };

  const APPLICATION_STATUS = {
    DRAFT: "draft",
    PENDING_APPROVAL: "pending_approval",
    APPROVED: "approved",
    REJECTED: "rejected",
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
    if (status === APPLICATION_STATUS.DRAFT) return "待申請";
    if (status === APPLICATION_STATUS.PENDING_APPROVAL) return "待審核";
    if (status === APPLICATION_STATUS.APPROVED) return "審核通過";
    if (status === APPLICATION_STATUS.REJECTED) return "審核不通過";
    return String(status);
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
    if (!app) return true;
    if (app.voided) return true;
    return (
      app.status === APPLICATION_STATUS.PENDING_APPROVAL ||
      app.status === "submitted" ||
      app.status === APPLICATION_STATUS.APPROVED
    );
  }

  function setReviewButtonsVisibility(app) {
    const approveBtn = $("#btn-review-approve");
    const rejectBtn = $("#btn-review-reject");
    if (!approveBtn || !rejectBtn) return;
    const show =
      app &&
      (app.status === APPLICATION_STATUS.PENDING_APPROVAL || app.status === "submitted") &&
      !app.voided;
    approveBtn.classList.toggle("hidden", !show);
    rejectBtn.classList.toggle("hidden", !show);
  }

  function lockMasterFields(locked) {
    const ids = [
      "#applicant",
      "#applyDate",
      "#payCategory",
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
    const addRow = $("#btn-add-row");
    if (addRow) {
      addRow.disabled = locked || !isMasterCreated;
      addRow.classList.toggle("hidden", locked);
    }

    const submitBtn = $("#btn-submit");
    if (submitBtn) {
      const app = getCurrentApplication();
      submitBtn.disabled = !canEditApplication(app);
      submitBtn.classList.toggle("hidden", locked);
    }
  }

  function renderFormState() {
    const app = getCurrentApplication();
    currentApplicationStatus = app?.status || APPLICATION_STATUS.DRAFT;
    const locked = isReviewLocked(app);
    lockMasterFields(locked);
    setReviewButtonsVisibility(app);

    // 「建立」只給 draft / rejected 編輯者使用；pending_approval / approved 只可檢視
    const btnCreate = $("#btn-edit-detail");
    if (btnCreate) {
      const canEdit = canEditApplication(app);
      btnCreate.classList.toggle("hidden", !canEdit);
    }
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
    const dt = validateExpectedPaymentDate();
    if (!dt.ok) {
      showExpectedDateError(dt.msg);
      closeMasterCreateModal();
      $("#expectedDate")?.focus();
      return;
    }

    const appId = generateApplicationId();
    currentApplicationId = appId;
    applications.unshift({
      applicationId: appId,
      createdAt: nowIso(),
      updatedAt: nowIso(),
      status: APPLICATION_STATUS.DRAFT,
      voided: false,
      master: collectMasterForm(),
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
    const rows = getFilteredApplications();
    if (!rows.length) {
      body.innerHTML = '<tr><td colspan="8" style="color:#8c8c8c">目前無資料</td></tr>';
      return;
    }
    body.innerHTML = rows
      .map((a) => {
        const total = (a.items || []).reduce((s, x) => s + parseNum(String(x.payAmount || 0)), 0);
        const itemCount = (a.items || []).length;
        const statusKey = a.voided ? "voided" : a.status;
        const statusClass =
          statusKey === "voided"
            ? "voided"
            : statusKey === APPLICATION_STATUS.DRAFT
              ? "draft"
              : statusKey === APPLICATION_STATUS.PENDING_APPROVAL || statusKey === "submitted"
                ? "pending"
                : statusKey === APPLICATION_STATUS.APPROVED
                  ? "approved"
                  : statusKey === APPLICATION_STATUS.REJECTED
                    ? "rejected"
                    : "";
        const statusText = statusToLabel(statusKey);
        const rowCls = a.voided ? "row-voided" : "";
        const canEdit =
          !a.voided &&
          (a.status === APPLICATION_STATUS.DRAFT ||
            a.status === APPLICATION_STATUS.REJECTED);
        const canVoid =
          !a.voided &&
          (a.status === APPLICATION_STATUS.DRAFT ||
            a.status === APPLICATION_STATUS.REJECTED);
        const canView =
          !a.voided &&
          (a.status === APPLICATION_STATUS.PENDING_APPROVAL ||
            a.status === "submitted" ||
            a.status === APPLICATION_STATUS.APPROVED);
        return `<tr class="${rowCls}" data-app-id="${escapeHtml(a.applicationId)}">
          <td>${escapeHtml(a.applicationId)}</td>
          <td>${escapeHtml(a.master?.applicant || "")}</td>
          <td>${escapeHtml(a.master?.applyDate || "")}</td>
          <td>${itemCount}</td>
          <td class="cell-num">${formatMoney(total)}</td>
          <td>${escapeHtml(formatDateTime(a.createdAt))}</td>
          <td><span class="status-pill ${statusClass}">${statusText}</span></td>
          <td class="col-actions">
            ${canEdit ? '<button type="button" class="link btn-row-edit">編輯</button>' : ''}
            ${canVoid ? '<button type="button" class="link btn-row-void">作廢</button>' : ''}
            ${canView ? '<button type="button" class="link btn-row-view">檢視</button>' : ''}
          </td>
        </tr>`;
      })
      .join("");

    body.querySelectorAll("tr[data-app-id]").forEach((tr) => {
      const appId = tr.getAttribute("data-app-id");
      tr.addEventListener("click", (e) => {
        if (e.target.closest("button")) return;
        openApplication(appId);
      });
      tr.querySelector(".btn-row-edit")?.addEventListener("click", () => openApplication(appId));
      tr.querySelector(".btn-row-void")?.addEventListener("click", () => voidApplication(appId));
      tr.querySelector(".btn-row-view")?.addEventListener("click", () => openApplication(appId));
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
  }

  function voidApplication(appId) {
    const app = applications.find((x) => x.applicationId === appId);
    if (!app) return;
    if (app.status !== APPLICATION_STATUS.DRAFT) {
      alert("已送出或進入審核流程不可作廢");
      return;
    }
    app.voided = true;
    saveApplications();
    renderApplicationList();
  }

  function upsertCurrentApplication(statusOverride) {
    if (!currentApplicationId) return;
    const master = collectMasterForm();
    const items = getDetailItemsFromTable();
    const idx = applications.findIndex((x) => x.applicationId === currentApplicationId);
    if (idx < 0) return;
    const nextStatus = statusOverride || applications[idx].status || APPLICATION_STATUS.DRAFT;
    applications[idx] = {
      ...applications[idx],
      master: { ...master, totalPayment: items.reduce((s, x) => s + parseNum(String(x.payAmount || 0)), 0) },
      items,
      status: nextStatus,
      updatedAt: nowIso(),
    };
    if (statusOverride === APPLICATION_STATUS.PENDING_APPROVAL) {
      applications[idx].submittedAt = nowIso();
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
    if (!d.voucherNo) {
      alert("請填寫憑證號碼");
      $("#modal-voucherNo")?.focus();
      return;
    }
    if (d.untaxed <= 0) {
      alert("請填寫未稅金額（須大於 0）");
      $("#modal-untaxed")?.focus();
      return;
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
        panels.forEach((p) =>
          p.classList.toggle("active", p.getAttribute("data-panel") === target)
        );
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
        const dt = validateExpectedPaymentDate();
        if (!dt.ok) {
          showExpectedDateError(dt.msg);
          $("#expectedDate")?.focus();
          return;
        }

        currentApplicationId = generateApplicationId();
        applications.unshift({
          applicationId: currentApplicationId,
          createdAt: nowIso(),
          updatedAt: nowIso(),
          status: APPLICATION_STATUS.DRAFT,
          voided: false,
          master: collectMasterForm(),
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
      type !== PAYMENT_TYPE.URGENT
    ) {
      return { ok: false, msg: "無效的付款類別" };
    }
    if (type === PAYMENT_TYPE.URGENT) {
      const max = addDays(apply, 5);
      if (exp < apply || exp > max) {
        return { ok: false, msg: "急件付款：預計付款日須在申請日起 5 日內（含）" };
      }
    }
    if (type === PAYMENT_TYPE.GENERAL) {
      const want = calcGeneralExpectedDate(apply);
      if (formatISODate(exp) !== formatISODate(want)) {
        return {
          ok: false,
          msg:
            "一般付款：預計付款日依規則計算失敗（21~次月5=>次月15；6~20=>當月30）",
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

  function onExpectedDateUserInput() {
    if ($("#payCategory").value !== PAYMENT_TYPE.URGENT) return;
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
    apply.addEventListener("change", syncExpectedPaymentDate);
    exp.addEventListener("change", onExpectedDateUserInput);
    exp.addEventListener("blur", onExpectedDateUserInput);
    syncExpectedPaymentDate();
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
    clearExpectedDateError();
    clearDetailTable();
    setMasterCreated(false);
    setFormMeta("", "draft");
    initPaymentMethod();
    const masterTab = document.querySelector('.tab[data-tab="master"]');
    if (masterTab) masterTab.click();
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
      const dt = validateExpectedPaymentDate();
      if (!dt.ok) {
        showExpectedDateError(dt.msg);
        const panel = $("#panel-master");
        if (panel) panel.scrollIntoView({ behavior: "smooth", block: "start" });
        $("#expectedDate")?.focus();
        return;
      }
      clearExpectedDateError();
      upsertCurrentApplication(APPLICATION_STATUS.PENDING_APPROVAL);
      alert("已送出申請，待財務審核。\n總金額：" + ($("#totalPayment").value || ""));
      showListPage();
    });

    $("#btn-review-approve")?.addEventListener("click", () => {
      const app = getCurrentApplication();
      if (!app || app.voided) return;
      if (app.status !== APPLICATION_STATUS.PENDING_APPROVAL && app.status !== "submitted") {
        alert("目前狀態不可審核通過");
        return;
      }
      upsertCurrentApplication(APPLICATION_STATUS.APPROVED);
      const idx = applications.findIndex((x) => x.applicationId === currentApplicationId);
      if (idx >= 0) applications[idx].reviewNote = applications[idx].reviewNote || "";
      applications[idx].reviewedAt = nowIso();
      saveApplications();
      alert("審核通過");
      showListPage();
    });

    $("#btn-review-reject")?.addEventListener("click", () => {
      const app = getCurrentApplication();
      if (!app || app.voided) return;
      if (app.status !== APPLICATION_STATUS.PENDING_APPROVAL && app.status !== "submitted") {
        alert("目前狀態不可審核不通過");
        return;
      }
      const note = prompt("請輸入審核不通過原因（必填）");
      if (!note || !note.trim()) {
        alert("審核原因不可空白");
        return;
      }
      upsertCurrentApplication(APPLICATION_STATUS.REJECTED);
      const idx = applications.findIndex((x) => x.applicationId === currentApplicationId);
      if (idx >= 0) {
        applications[idx].reviewNote = note.trim();
        applications[idx].reviewedAt = nowIso();
      }
      saveApplications();
      alert("已回覆：審核不通過（可重新修改並送出）");
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
  }

  function init() {
    loadVendorsFromStorage();
    loadApplications();
    updateVendorBadge();
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
