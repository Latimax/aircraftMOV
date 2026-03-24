const DB_NAME = "aircraft-movement-register";
const DB_VERSION = 1;
const STORE_RECORDS = "records";
const STORE_META = "meta";
const DEFAULT_WORKBOOK_NAME = "Aircraft_Movement_Register.xlsx";
const MOVEMENT_SHEET_PREFIX = "Aircraft Movement";
const EXCEL_MAX_ROWS = 1048576;
const THEME_KEY = "aircraft-register-theme";
const WATCH_OPTIONS = ["Red", "Blue", "Green", "White"];
const AIRCRAFT_OPTIONS = [
    "B737", "A320", "B777", "A330", "B787", "A380",
    "A319", "A321", "E190", "E175", "CRJ900", "ATR72",
    "B757", "B767", "MD80", "C130", "G550", "DH8D"
];
const EXPORT_HEADERS = [
    "S/N", "DATE", "WATCH ON DUTY", "AIRCRAFT TYPE", "REG. NO",
    "TIME OF ARRIVAL", "TIME OF DEPARTURE", "SOULS ON BOARD LANDING",
    "SOULS ON BOARD TAKE OFF", "CREW ON BOARD", "OPERATING COMPANY", "DESTINATION"
];

const state = {
    db: null,
    viewMode: "day",
    viewDate: "",
    page: 1,
    pageSize: 25,
    totalRecords: 0,
    totalDays: 0,
    currentViewCount: 0,
    totalPages: 1,
    availableDates: [],
    editingId: null,
    linkedWorkbookHandle: null,
    linkedWorkbookName: "",
    projectDirectoryHandle: null,
    projectDirectoryName: "",
    renderToken: 0
};

const modalState = { resolver: null };
const el = {};

try {
    document.documentElement.dataset.theme = localStorage.getItem(THEME_KEY) || "light";
} catch (error) {
    document.documentElement.dataset.theme = "light";
}

function toYmdLocal(date) {
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function todayYmd() {
    return toYmdLocal(new Date());
}

function fmtNumber(value) {
    return new Intl.NumberFormat().format(value || 0);
}

function fmtDate(dateKey, sep = "-") {
    if (!dateKey) return "-";
    const [year, month, day] = String(dateKey).split("-");
    return [day?.padStart(2, "0"), month?.padStart(2, "0"), year].join(sep);
}

function fmtDateWb(dateKey) {
    return fmtDate(dateKey, "/");
}

function esc(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

function pad(value, width) {
    const num = Number.parseInt(value, 10);
    if (!Number.isFinite(num) || num < 0) return "".padStart(width, "0");
    return String(num).padStart(width, "0");
}

function fmtSouls(value) {
    return pad(value, 3);
}

function fmtCrew(value) {
    return pad(value, 2);
}

function normTime(raw) {
    const value = String(raw || "").trim();
    if (!value || value === "-" || value === "--") return "-";
    const digits = value.replace(/\D/g, "");
    if (!digits) return "-";
    if (digits.length === 3) return `0${digits}`;
    if (digits.length === 4) return digits;
    return digits.slice(0, 4);
}

function validTime(value) {
    if (value === "-") return true;
    if (!/^\d{4}$/.test(value)) return false;
    const hour = Number.parseInt(value.slice(0, 2), 10);
    const minute = Number.parseInt(value.slice(2), 10);
    return hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59;
}

function normHeader(value) {
    return String(value || "").toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
}

function coerceCell(value) {
    if (value == null) return "";
    if (value instanceof Date) return value;
    if (typeof value === "object") {
        if ("text" in value) return value.text || "";
        if ("result" in value) return value.result || "";
        if ("richText" in value && Array.isArray(value.richText)) {
            return value.richText.map((part) => part.text || "").join("");
        }
    }
    return value;
}

function normImportedDate(value) {
    if (!value && value !== 0) return "";
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
        return toYmdLocal(value);
    }
    if (typeof value === "number" && Number.isFinite(value)) {
        const date = new Date(Math.round((value - 25569) * 86400 * 1000));
        if (!Number.isNaN(date.getTime())) return toYmdLocal(date);
    }
    const raw = String(value).trim();
    if (!raw) return "";
    if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
    const match = raw.replace(/[.-]/g, "/").match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (!match) return "";
    let [, day, month, year] = match;
    if (year.length === 2) year = `20${year}`;
    return `${year.padStart(4, "0")}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
}

function setBusy(label) {
    el.busyStatus.textContent = label;
}

function currentTheme() {
    return document.documentElement.dataset.theme === "dark" ? "dark" : "light";
}

function updateThemeControls() {
    if (!el.themeStateLabel || !el.themeToggleBtn) return;
    const theme = currentTheme();
    el.themeStateLabel.textContent = theme === "light" ? "Light Theme" : "Dark Theme";
    el.themeToggleBtn.textContent = theme === "light" ? "Switch to Dark Theme" : "Switch to Light Theme";
}

function applyTheme(theme) {
    const nextTheme = theme === "dark" ? "dark" : "light";
    document.documentElement.dataset.theme = nextTheme;
    try {
        localStorage.setItem(THEME_KEY, nextTheme);
    } catch (error) {
        console.warn("Theme preference could not be stored.", error);
    }
    updateThemeControls();
}

function appShell() {
    return `
        <header class="mb-6 glass-panel p-5 sm:p-6">
            <div class="flex flex-col gap-4 lg:flex-row lg:items-center lg:justify-between">
                <div class="flex items-start gap-4">
                    <div class="flex h-14 w-14 items-center justify-center rounded-2xl border border-copper/25 bg-copper/10 text-lg font-bold text-copper">AM</div>
                    <div>
                        <p class="section-label">Operations Header</p>
                        <h2 class="mt-2 text-2xl font-semibold text-white">Aircraft Movement Register</h2>
                        <p class="mt-2 max-w-3xl text-sm leading-7 text-mist/65">Light theme is the default and uses a green, yellow, black, and white palette. You can switch themes here whenever you need a darker console view.</p>
                    </div>
                </div>
                <div class="flex flex-col gap-3 sm:flex-row sm:items-center">
                    <div class="rounded-2xl border border-white/10 bg-ink/45 px-4 py-3 text-xs uppercase tracking-[0.24em] text-mist/55">
                        Active Theme
                        <span id="themeStateLabel" class="ml-2 font-semibold text-white">Light Theme</span>
                    </div>
                    <button id="themeToggleBtn" type="button" class="secondary-btn">Switch to Dark Theme</button>
                </div>
            </div>
        </header>

        <section class="glass-panel p-6 sm:p-8">
            <div class="flex flex-col gap-6 xl:flex-row xl:items-end xl:justify-between">
                <div class="max-w-4xl">
                    <p class="section-label">Aircraft Movement Control</p>
                    <h1 class="mt-3 text-3xl font-semibold tracking-tight text-white sm:text-5xl">
                        Professional aircraft movement logging with linked workbook updates.
                    </h1>
                    <p class="mt-4 max-w-3xl text-sm leading-7 text-mist/70 sm:text-base">
                        Tailwind UI, IndexedDB-backed storage, grouped operational days, workbook updates,
                        exports, backups, and row editing without rendering the entire register at once.
                    </p>
                </div>
                <div class="grid gap-3 sm:grid-cols-2 xl:min-w-[420px]">
                    <div class="stat-chip"><p class="text-[0.72rem] uppercase tracking-[0.24em] text-mist/55">Total Records</p><p id="totalRecordsValue" class="mt-2 text-2xl font-semibold text-white">0</p></div>
                    <div class="stat-chip"><p class="text-[0.72rem] uppercase tracking-[0.24em] text-mist/55">Tracked Days</p><p id="totalDaysValue" class="mt-2 text-2xl font-semibold text-white">0</p></div>
                    <div class="stat-chip"><p class="text-[0.72rem] uppercase tracking-[0.24em] text-mist/55">Current View</p><p id="currentViewValue" class="mt-2 text-2xl font-semibold text-white">0</p></div>
                    <div class="stat-chip"><p class="text-[0.72rem] uppercase tracking-[0.24em] text-mist/55">Storage Mode</p><p class="mt-2 text-sm font-semibold text-white">IndexedDB + Paged Render</p></div>
                </div>
            </div>
        </section>

        <section class="mt-6 grid gap-6 xl:grid-cols-[1.1fr,0.9fr]">
            <div class="glass-panel p-6 sm:p-7">
                <div class="flex flex-col gap-5">
                    <div class="flex flex-col gap-2 sm:flex-row sm:items-end sm:justify-between">
                        <div><p class="section-label">Workbook Operations</p><h2 class="mt-2 text-2xl font-semibold text-white">Linked file workflow</h2></div>
                        <div class="rounded-2xl border border-white/10 bg-ink/45 px-4 py-3 text-xs uppercase tracking-[0.24em] text-mist/55">Update in place when browser access is granted</div>
                    </div>
                    <div class="grid gap-3 lg:grid-cols-3">
                        <button id="connectFolderBtn" type="button" class="secondary-btn">Connect Project Folder</button>
                        <button id="importWorkbookBtn" type="button" class="secondary-btn">Import or Link Workbook</button>
                        <button id="updateWorkbookBtn" type="button" class="primary-btn">Update Linked Workbook</button>
                        <button id="exportViewBtn" type="button" class="secondary-btn">Export Current View</button>
                        <button id="backupWorkbookBtn" type="button" class="secondary-btn">Backup Snapshot</button>
                        <button id="clearRegisterBtn" type="button" class="ghost-btn border border-danger/25 text-danger hover:bg-danger/10 hover:text-white">Clear Register</button>
                    </div>
                    <input id="excelFileInput" type="file" accept=".xlsx,.xlsm,.xls" class="hidden">
                    <div class="grid gap-3 lg:grid-cols-2">
                        <div class="rounded-3xl border border-white/10 bg-ink/45 p-5">
                            <p class="text-[0.72rem] uppercase tracking-[0.24em] text-mist/55">Connected Folder</p>
                            <p id="projectFolderStatus" class="mt-2 text-sm font-medium text-white">No folder connected</p>
                            <p class="mt-2 text-sm text-mist/60">Used for direct workbook updates and automatic backups.</p>
                        </div>
                        <div class="rounded-3xl border border-white/10 bg-ink/45 p-5">
                            <p class="text-[0.72rem] uppercase tracking-[0.24em] text-mist/55">Linked Workbook</p>
                            <p id="linkedWorkbookStatus" class="mt-2 text-sm font-medium text-white">No workbook linked</p>
                            <p class="mt-2 text-sm text-mist/60">Import an existing workbook to overwrite the same file instead of saving a fresh one every time.</p>
                        </div>
                    </div>
                </div>
            </div>

            <div class="glass-panel p-6 sm:p-7">
                <div class="flex flex-col gap-5">
                    <div class="flex flex-col gap-2 sm:flex-row sm:items-end sm:justify-between">
                        <div><p class="section-label">View Controller</p><h2 class="mt-2 text-2xl font-semibold text-white">Navigate by day or full register</h2></div>
                        <div id="busyStatus" class="rounded-2xl border border-white/10 bg-ink/45 px-4 py-3 text-xs uppercase tracking-[0.24em] text-mist/55">Idle</div>
                    </div>
                    <div class="grid gap-4 md:grid-cols-2 xl:grid-cols-4">
                        <div><label class="control-label" for="viewModeSelect">View Scope</label><select id="viewModeSelect" class="field-shell-lite"><option value="day">Single Day</option><option value="all">All Days</option></select></div>
                        <div><label class="control-label" for="viewDateInput">View Date</label><input id="viewDateInput" type="date" class="field-shell-lite"></div>
                        <div><label class="control-label" for="pageSizeSelect">Rows Per Page</label><select id="pageSizeSelect" class="field-shell-lite"><option value="25">25 rows</option><option value="50">50 rows</option><option value="100">100 rows</option><option value="250">250 rows</option></select></div>
                        <div><label class="control-label" for="jumpDateBtn">Recorded Day Jump</label><button id="jumpDateBtn" type="button" class="secondary-btn w-full">Jump to Nearest Recorded Day</button></div>
                    </div>
                    <div class="grid gap-3 sm:grid-cols-3">
                        <button id="prevDateBtn" type="button" class="ghost-btn rounded-2xl border border-white/10">Previous Recorded Day</button>
                        <button id="todayDateBtn" type="button" class="ghost-btn rounded-2xl border border-white/10">Today</button>
                        <button id="nextDateBtn" type="button" class="ghost-btn rounded-2xl border border-white/10">Next Recorded Day</button>
                    </div>
                    <div class="rounded-3xl border border-white/10 bg-ink/45 p-5">
                        <div class="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
                            <p id="viewSummary" class="text-sm font-medium text-white">Viewing register for today.</p>
                            <p id="pageSummary" class="text-sm text-mist/65">Page 1 of 1</p>
                        </div>
                        <p class="mt-3 text-sm leading-7 text-mist/60">Rows stay outside the DOM and are rendered page by page so the interface remains usable when the register grows large.</p>
                    </div>
                </div>
            </div>
        </section>

        <section class="mt-6 grid gap-6 xl:grid-cols-[0.92fr,1.08fr]">
            <div class="glass-panel p-6 sm:p-7">
                <div class="flex flex-col gap-5">
                    <div class="flex flex-col gap-2 sm:flex-row sm:items-end sm:justify-between">
                        <div><p class="section-label">Record Form</p><h2 id="formTitle" class="mt-2 text-2xl font-semibold text-white">Create movement record</h2></div>
                        <div id="serialPreview" class="rounded-2xl border border-copper/25 bg-copper/10 px-4 py-3 text-sm font-semibold tracking-[0.14em] text-copper">Next S/N for selected day: 1</div>
                    </div>
                    <form id="entryForm" class="grid gap-4 lg:grid-cols-2 xl:grid-cols-3">
                        <div><label class="control-label" for="dateInput">Date</label><input id="dateInput" type="date" class="field-shell" required></div>
                        <div><label class="control-label" for="watchSelect">Watch On Duty</label><select id="watchSelect" class="field-shell" required></select></div>
                        <div><label class="control-label" for="aircraftSelect">Aircraft Type</label><select id="aircraftSelect" class="field-shell" required></select></div>
                        <div><label class="control-label" for="regNoInput">Registration No</label><input id="regNoInput" type="text" class="field-shell" placeholder="Example: 5NTON" required></div>
                        <div><label class="control-label" for="arrivalTimeInput">Time Of Arrival</label><input id="arrivalTimeInput" type="text" maxlength="4" class="field-shell" placeholder="HHMM or -" required></div>
                        <div><label class="control-label" for="departureTimeInput">Time Of Departure</label><input id="departureTimeInput" type="text" maxlength="4" class="field-shell" placeholder="HHMM or -" required></div>
                        <div><label class="control-label" for="soulsLandingInput">Souls On Board Landing</label><input id="soulsLandingInput" type="number" min="0" step="1" class="field-shell" value="0" required></div>
                        <div><label class="control-label" for="soulsTakeoffInput">Souls On Board Take Off</label><input id="soulsTakeoffInput" type="number" min="0" step="1" class="field-shell" value="0" required></div>
                        <div><label class="control-label" for="crewOnBoardInput">Crew On Board</label><input id="crewOnBoardInput" type="number" min="0" step="1" class="field-shell" value="0" required></div>
                        <div class="lg:col-span-2 xl:col-span-2"><label class="control-label" for="operatingCompanyInput">Operating Company</label><input id="operatingCompanyInput" type="text" class="field-shell" placeholder="Example: Air Peace" required></div>
                        <div><label class="control-label" for="destinationInput">Destination</label><input id="destinationInput" type="text" class="field-shell" placeholder="Example: LOS / ABV" required></div>
                    </form>
                    <div class="rounded-3xl border border-white/10 bg-ink/45 p-5 text-sm leading-7 text-mist/65">Arrival and departure accept either a four-digit time or a single dash for no time. S/N resets per operational day and updates automatically when a row is edited, moved, or deleted.</div>
                    <div class="flex flex-wrap gap-3">
                        <button id="submitBtn" type="button" class="primary-btn">Save Record</button>
                        <button id="resetFormBtn" type="button" class="secondary-btn">Reset Form</button>
                        <button id="cancelEditBtn" type="button" class="ghost-btn hidden rounded-2xl border border-copper/25 text-copper hover:bg-copper/10 hover:text-white">Cancel Edit</button>
                    </div>
                </div>
            </div>

            <div class="glass-panel p-6 sm:p-7">
                <div class="flex flex-col gap-5">
                    <div class="flex flex-col gap-2 sm:flex-row sm:items-end sm:justify-between">
                        <div><p class="section-label">Movement Register</p><h2 class="mt-2 text-2xl font-semibold text-white">Operational day groups with row actions</h2></div>
                        <div class="flex flex-wrap gap-2">
                            <button id="prevPageBtn" type="button" class="ghost-btn rounded-2xl border border-white/10">Previous Page</button>
                            <button id="nextPageBtn" type="button" class="ghost-btn rounded-2xl border border-white/10">Next Page</button>
                        </div>
                    </div>
                    <div class="overflow-hidden rounded-[26px] border border-white/10 bg-ink/45">
                        <div class="max-h-[70vh] overflow-auto scrollbar-thin">
                            <table class="min-w-[1450px] w-full border-separate border-spacing-0">
                                <thead><tr>
                                    <th class="table-head-cell">S/N</th><th class="table-head-cell">Date</th><th class="table-head-cell">Watch</th><th class="table-head-cell">Aircraft</th><th class="table-head-cell">Reg. No</th><th class="table-head-cell">Arrival</th><th class="table-head-cell">Departure</th><th class="table-head-cell">Souls Landing</th><th class="table-head-cell">Souls Take Off</th><th class="table-head-cell">Crew</th><th class="table-head-cell">Operating Company</th><th class="table-head-cell">Destination</th><th class="table-head-cell">Actions</th>
                                </tr></thead>
                                <tbody id="tableBody"></tbody>
                            </table>
                        </div>
                    </div>
                </div>
            </div>
        </section>
    `;
}

function modalShell() {
    return `
        <div class="modal-panel">
            <div class="flex items-start justify-between gap-4">
                <div>
                    <p id="modalTone" class="section-label">Notice</p>
                    <h3 id="modalTitle" class="mt-2 text-2xl font-semibold text-white">Message</h3>
                </div>
                <button id="modalCloseBtn" type="button" class="ghost-btn rounded-2xl border border-white/10">Close</button>
            </div>
            <p id="modalDescription" class="mt-4 text-sm leading-7 text-mist/72"></p>
            <div id="modalDetails" class="mt-4 hidden rounded-3xl border border-white/10 bg-ink/45 p-4 text-sm leading-7 text-mist/65"></div>
            <div id="modalActions" class="mt-6 flex flex-wrap gap-3"></div>
        </div>
    `;
}

function cacheDom() {
    document.getElementById("app").innerHTML = appShell();
    document.getElementById("messageModal").innerHTML = modalShell();
    [
        "themeStateLabel", "themeToggleBtn",
        "totalRecordsValue", "totalDaysValue", "currentViewValue", "busyStatus",
        "projectFolderStatus", "linkedWorkbookStatus", "viewModeSelect", "viewDateInput",
        "pageSizeSelect", "jumpDateBtn", "prevDateBtn", "todayDateBtn", "nextDateBtn",
        "viewSummary", "pageSummary", "entryForm", "formTitle", "serialPreview",
        "dateInput", "watchSelect", "aircraftSelect", "regNoInput", "arrivalTimeInput",
        "departureTimeInput", "soulsLandingInput", "soulsTakeoffInput", "crewOnBoardInput",
        "operatingCompanyInput", "destinationInput", "submitBtn", "resetFormBtn",
        "cancelEditBtn", "connectFolderBtn", "importWorkbookBtn", "updateWorkbookBtn",
        "exportViewBtn", "backupWorkbookBtn", "clearRegisterBtn", "excelFileInput",
        "tableBody", "prevPageBtn", "nextPageBtn", "toastContainer", "messageModal",
        "modalTone", "modalTitle", "modalDescription", "modalDetails", "modalActions",
        "modalCloseBtn"
    ].forEach((id) => { el[id] = document.getElementById(id); });
}

function loadSelects() {
    el.watchSelect.innerHTML = ['<option value="" disabled selected>Select watch</option>', ...WATCH_OPTIONS.map((v) => `<option value="${esc(v)}">${esc(v)}</option>`)].join("");
    el.aircraftSelect.innerHTML = ['<option value="" disabled selected>Select aircraft type</option>', ...AIRCRAFT_OPTIONS.map((v) => `<option value="${esc(v)}">${esc(v)}</option>`)].join("");
}

function reqToPromise(request) {
    return new Promise((resolve, reject) => {
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
    });
}

function txDone(tx) {
    return new Promise((resolve, reject) => {
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
        tx.onabort = () => reject(tx.error);
    });
}

function openDb() {
    return new Promise((resolve, reject) => {
        const request = indexedDB.open(DB_NAME, DB_VERSION);
        request.onupgradeneeded = () => {
            const db = request.result;
            if (!db.objectStoreNames.contains(STORE_RECORDS)) {
                const store = db.createObjectStore(STORE_RECORDS, { keyPath: "id" });
                store.createIndex("dateKey", "dateKey", { unique: false });
                store.createIndex("dateSerial", ["dateKey", "serial"], { unique: true });
            }
            if (!db.objectStoreNames.contains(STORE_META)) {
                db.createObjectStore(STORE_META, { keyPath: "key" });
            }
        };
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
    });
}

async function getMeta(key) {
    const tx = state.db.transaction(STORE_META, "readonly");
    const row = await reqToPromise(tx.objectStore(STORE_META).get(key));
    await txDone(tx);
    return row ? row.value : null;
}

async function setMeta(key, value) {
    const tx = state.db.transaction(STORE_META, "readwrite");
    tx.objectStore(STORE_META).put({ key, value });
    await txDone(tx);
}

async function putRecord(record) {
    const tx = state.db.transaction(STORE_RECORDS, "readwrite");
    tx.objectStore(STORE_RECORDS).put(record);
    await txDone(tx);
}

async function getRecord(id) {
    const tx = state.db.transaction(STORE_RECORDS, "readonly");
    const row = await reqToPromise(tx.objectStore(STORE_RECORDS).get(id));
    await txDone(tx);
    return row || null;
}

async function deleteRecordById(id) {
    const tx = state.db.transaction(STORE_RECORDS, "readwrite");
    tx.objectStore(STORE_RECORDS).delete(id);
    await txDone(tx);
}

async function clearRecords() {
    const tx = state.db.transaction(STORE_RECORDS, "readwrite");
    tx.objectStore(STORE_RECORDS).clear();
    await txDone(tx);
}

async function replaceAllRecords(records) {
    const tx = state.db.transaction(STORE_RECORDS, "readwrite");
    const store = tx.objectStore(STORE_RECORDS);
    store.clear();
    records.forEach((record) => store.put(record));
    await txDone(tx);
}

async function countAllRecords() {
    const tx = state.db.transaction(STORE_RECORDS, "readonly");
    const count = await reqToPromise(tx.objectStore(STORE_RECORDS).count());
    await txDone(tx);
    return count;
}

async function countByDate(dateKey) {
    const tx = state.db.transaction(STORE_RECORDS, "readonly");
    const count = await reqToPromise(tx.objectStore(STORE_RECORDS).index("dateKey").count(IDBKeyRange.only(dateKey)));
    await txDone(tx);
    return count;
}

async function countDistinctDates() {
    const tx = state.db.transaction(STORE_RECORDS, "readonly");
    const index = tx.objectStore(STORE_RECORDS).index("dateKey");
    let count = 0;
    await new Promise((resolve, reject) => {
        const req = index.openKeyCursor(null, "nextunique");
        req.onsuccess = () => {
            const cursor = req.result;
            if (!cursor) return resolve();
            count += 1;
            cursor.continue();
        };
        req.onerror = () => reject(req.error);
    });
    await txDone(tx);
    return count;
}

async function getAvailableDates() {
    const tx = state.db.transaction(STORE_RECORDS, "readonly");
    const index = tx.objectStore(STORE_RECORDS).index("dateKey");
    const dates = [];
    await new Promise((resolve, reject) => {
        const req = index.openKeyCursor(null, "prevunique");
        req.onsuccess = () => {
            const cursor = req.result;
            if (!cursor) return resolve();
            dates.push(cursor.key);
            cursor.continue();
        };
        req.onerror = () => reject(req.error);
    });
    await txDone(tx);
    return dates;
}

async function collectRecords({ mode, dateKey, offset, limit }) {
    const tx = state.db.transaction(STORE_RECORDS, "readonly");
    const index = tx.objectStore(STORE_RECORDS).index("dateSerial");
    const range = mode === "day" ? IDBKeyRange.bound([dateKey, 1], [dateKey, Number.MAX_SAFE_INTEGER]) : null;
    const rows = [];
    let skipped = offset;
    await new Promise((resolve, reject) => {
        const req = index.openCursor(range, "next");
        req.onsuccess = () => {
            const cursor = req.result;
            if (!cursor || rows.length >= limit) return resolve();
            if (skipped > 0) {
                skipped -= 1;
                cursor.continue();
                return;
            }
            rows.push(cursor.value);
            cursor.continue();
        };
        req.onerror = () => reject(req.error);
    });
    await txDone(tx);
    return rows;
}

async function getAllRecordsForExport(mode, dateKey) {
    const tx = state.db.transaction(STORE_RECORDS, "readonly");
    const index = tx.objectStore(STORE_RECORDS).index("dateSerial");
    const range = mode === "day" ? IDBKeyRange.bound([dateKey, 1], [dateKey, Number.MAX_SAFE_INTEGER]) : null;
    const rows = [];
    await new Promise((resolve, reject) => {
        const req = index.openCursor(range, "next");
        req.onsuccess = () => {
            const cursor = req.result;
            if (!cursor) return resolve();
            rows.push(cursor.value);
            cursor.continue();
        };
        req.onerror = () => reject(req.error);
    });
    await txDone(tx);
    return rows;
}

async function shiftSerials(dateKey, startSerial, delta) {
    if (!delta) return;
    const tx = state.db.transaction(STORE_RECORDS, "readwrite");
    const index = tx.objectStore(STORE_RECORDS).index("dateSerial");
    const range = IDBKeyRange.bound([dateKey, startSerial], [dateKey, Number.MAX_SAFE_INTEGER]);
    await new Promise((resolve, reject) => {
        const req = index.openCursor(range, "prev");
        req.onsuccess = () => {
            const cursor = req.result;
            if (!cursor) return resolve();
            cursor.update({ ...cursor.value, serial: cursor.value.serial + delta });
            cursor.continue();
        };
        req.onerror = () => reject(req.error);
    });
    await txDone(tx);
}

function pushToast(message, tone = "success") {
    const toneClass = tone === "error"
        ? "border-danger/35 bg-danger/10"
        : tone === "warning"
            ? "border-copper/35 bg-copper/10"
            : "border-success/30 bg-success/10";
    const node = document.createElement("div");
    node.className = `rounded-3xl border px-4 py-3 text-white shadow-soft ${toneClass}`;
    node.innerHTML = `<p class="text-sm leading-6">${esc(message)}</p>`;
    el.toastContainer.appendChild(node);
    window.setTimeout(() => {
        node.style.opacity = "0";
        node.style.transform = "translateX(12px)";
        window.setTimeout(() => node.remove(), 240);
    }, 3600);
}

function closeModal(result = false) {
    el.messageModal.dataset.open = "false";
    el.messageModal.setAttribute("aria-hidden", "true");
    if (modalState.resolver) {
        const resolve = modalState.resolver;
        modalState.resolver = null;
        resolve(result);
    }
}

function showModal({ tone = "Notice", title, description, details = "", confirmLabel = "Close", cancelLabel = "", confirmClass = "primary-btn", hideClose = false }) {
    return new Promise((resolve) => {
        modalState.resolver = resolve;
        el.modalTone.textContent = tone;
        el.modalTitle.textContent = title;
        el.modalDescription.textContent = description;
        el.modalDetails.textContent = details;
        el.modalDetails.classList.toggle("hidden", !details);
        el.modalCloseBtn.classList.toggle("hidden", hideClose);
        el.modalActions.innerHTML = "";

        const ok = document.createElement("button");
        ok.type = "button";
        ok.className = confirmClass;
        ok.textContent = confirmLabel;
        ok.addEventListener("click", () => closeModal(true));
        el.modalActions.appendChild(ok);

        if (cancelLabel) {
            const cancel = document.createElement("button");
            cancel.type = "button";
            cancel.className = "secondary-btn";
            cancel.textContent = cancelLabel;
            cancel.addEventListener("click", () => closeModal(false));
            el.modalActions.appendChild(cancel);
        }

        el.messageModal.dataset.open = "true";
        el.messageModal.setAttribute("aria-hidden", "false");
    });
}

async function refreshMetaState() {
    state.viewMode = (await getMeta("viewMode")) || "day";
    state.viewDate = (await getMeta("viewDate")) || todayYmd();
    state.pageSize = Number((await getMeta("pageSize")) || 25);
    state.linkedWorkbookHandle = (await getMeta("linkedWorkbookHandle")) || null;
    state.linkedWorkbookName = (await getMeta("linkedWorkbookName")) || "";
    state.projectDirectoryHandle = (await getMeta("projectDirectoryHandle")) || null;
    state.projectDirectoryName = (await getMeta("projectDirectoryName")) || "";
}

async function persistViewState() {
    await setMeta("viewMode", state.viewMode);
    await setMeta("viewDate", state.viewDate);
    await setMeta("pageSize", state.pageSize);
}

function updateViewSummary() {
    el.viewSummary.textContent = state.viewMode === "all"
        ? "Viewing all recorded days."
        : `Viewing records for ${fmtDate(state.viewDate)}.`;
    el.pageSummary.textContent = `Page ${state.page} of ${state.totalPages}`;
    el.viewDateInput.disabled = state.viewMode === "all";
    el.jumpDateBtn.disabled = !state.availableDates.length;
    el.prevDateBtn.disabled = !state.availableDates.length;
    el.nextDateBtn.disabled = !state.availableDates.length;
}

async function updateSerialPreview() {
    const dateKey = el.dateInput.value;
    if (!dateKey) {
        el.serialPreview.textContent = "Select a day to preview S/N";
        return;
    }
    if (state.editingId) {
        const current = await getRecord(state.editingId);
        if (current) {
            if (current.dateKey === dateKey) {
                el.serialPreview.textContent = `Editing S/N ${current.serial} for ${fmtDate(dateKey)}`;
                return;
            }
            el.serialPreview.textContent = `Moved entries land at S/N ${(await countByDate(dateKey)) + 1} for ${fmtDate(dateKey)}`;
            return;
        }
    }
    el.serialPreview.textContent = `Next S/N for selected day: ${(await countByDate(dateKey)) + 1}`;
}

async function refreshDashboard() {
    state.totalRecords = await countAllRecords();
    state.totalDays = await countDistinctDates();
    state.availableDates = await getAvailableDates();
    el.totalRecordsValue.textContent = fmtNumber(state.totalRecords);
    el.totalDaysValue.textContent = fmtNumber(state.totalDays);
    el.currentViewValue.textContent = fmtNumber(state.currentViewCount);
    el.projectFolderStatus.textContent = state.projectDirectoryName || "No folder connected";
    el.linkedWorkbookStatus.textContent = state.linkedWorkbookName || "No workbook linked";
    updateViewSummary();
    await updateSerialPreview();
}

function rowHtml(record) {
    return `
        <tr class="table-row" data-id="${esc(record.id)}">
            <td class="table-cell font-semibold text-white">${record.serial}</td>
            <td class="table-cell">${esc(fmtDateWb(record.dateKey))}</td>
            <td class="table-cell">${esc(record.watch)}</td>
            <td class="table-cell">${esc(record.aircraftType)}</td>
            <td class="table-cell">${esc(record.regNo)}</td>
            <td class="table-cell">${esc(record.arrivalTime)}</td>
            <td class="table-cell">${esc(record.departureTime)}</td>
            <td class="table-cell">${esc(record.soulsLanding)}</td>
            <td class="table-cell">${esc(record.soulsTakeoff)}</td>
            <td class="table-cell">${esc(record.crewOnBoard)}</td>
            <td class="table-cell">${esc(record.operatingCompany)}</td>
            <td class="table-cell">${esc(record.destination)}</td>
            <td class="table-cell">
                <div class="flex gap-2">
                    <button type="button" data-action="edit" class="ghost-btn rounded-xl border border-signal/20 text-signal hover:bg-signal/10 hover:text-white">Edit</button>
                    <button type="button" data-action="delete" class="ghost-btn rounded-xl border border-danger/20 text-danger hover:bg-danger/10 hover:text-white">Delete</button>
                </div>
            </td>
        </tr>
    `;
}

async function renderTable() {
    const token = ++state.renderToken;
    setBusy("Refreshing register");
    const count = state.viewMode === "all" ? state.totalRecords : await countByDate(state.viewDate);
    state.currentViewCount = count;
    state.totalPages = Math.max(1, Math.ceil(count / state.pageSize));
    if (state.page > state.totalPages) state.page = state.totalPages;

    const records = await collectRecords({
        mode: state.viewMode,
        dateKey: state.viewDate,
        offset: (state.page - 1) * state.pageSize,
        limit: state.pageSize
    });
    if (token !== state.renderToken) return;

    if (!records.length) {
        el.tableBody.innerHTML = `<tr><td colspan="13" class="table-cell py-8 text-center text-sm text-mist/60">No movements available for this view yet.</td></tr>`;
    } else {
        let activeDate = "";
        let html = "";
        records.forEach((record) => {
            if (record.dateKey !== activeDate) {
                activeDate = record.dateKey;
                html += `<tr><td colspan="13" class="px-4 py-4"><div class="day-banner">Operational Day - ${esc(fmtDate(record.dateKey))}</div></td></tr>`;
            }
            html += rowHtml(record);
        });
        el.tableBody.innerHTML = html;
    }

    el.currentViewValue.textContent = fmtNumber(state.currentViewCount);
    updateViewSummary();
    setBusy("Idle");
}

function resetForm({ preserveDate = true } = {}) {
    const dateValue = preserveDate ? (el.dateInput.value || state.viewDate || todayYmd()) : todayYmd();
    el.entryForm.reset();
    el.dateInput.value = dateValue;
    el.watchSelect.value = "";
    el.aircraftSelect.value = "";
    el.arrivalTimeInput.value = "-";
    el.departureTimeInput.value = "-";
    el.soulsLandingInput.value = "0";
    el.soulsTakeoffInput.value = "0";
    el.crewOnBoardInput.value = "0";
}

function exitEditMode() {
    state.editingId = null;
    el.formTitle.textContent = "Create movement record";
    el.submitBtn.textContent = "Save Record";
    el.cancelEditBtn.classList.add("hidden");
}

function readForm() {
    const dateKey = el.dateInput.value;
    const watch = el.watchSelect.value;
    const aircraftType = el.aircraftSelect.value;
    const regNo = el.regNoInput.value.trim().toUpperCase();
    const arrivalTime = normTime(el.arrivalTimeInput.value);
    const departureTime = normTime(el.departureTimeInput.value);
    const soulsLanding = Number.parseInt(el.soulsLandingInput.value, 10);
    const soulsTakeoff = Number.parseInt(el.soulsTakeoffInput.value, 10);
    const crewOnBoard = Number.parseInt(el.crewOnBoardInput.value, 10);
    const operatingCompany = el.operatingCompanyInput.value.trim();
    const destination = el.destinationInput.value.trim();

    if (!dateKey) throw new Error("Date is required.");
    if (!watch) throw new Error("Watch on duty is required.");
    if (!aircraftType) throw new Error("Aircraft type is required.");
    if (!regNo) throw new Error("Registration number is required.");
    if (!operatingCompany) throw new Error("Operating company is required.");
    if (!destination) throw new Error("Destination is required.");
    if (!validTime(arrivalTime)) throw new Error("Arrival time must be HHMM or -.");
    if (!validTime(departureTime)) throw new Error("Departure time must be HHMM or -.");
    if (!Number.isFinite(soulsLanding) || soulsLanding < 0) throw new Error("Souls on board landing must be zero or greater.");
    if (!Number.isFinite(soulsTakeoff) || soulsTakeoff < 0) throw new Error("Souls on board take off must be zero or greater.");
    if (!Number.isFinite(crewOnBoard) || crewOnBoard < 0) throw new Error("Crew on board must be zero or greater.");

    return {
        dateKey,
        watch,
        aircraftType,
        regNo,
        arrivalTime,
        departureTime,
        soulsLanding: fmtSouls(soulsLanding),
        soulsTakeoff: fmtSouls(soulsTakeoff),
        crewOnBoard: fmtCrew(crewOnBoard),
        operatingCompany,
        destination
    };
}

async function saveMovement(payload) {
    const now = new Date().toISOString();
    if (state.editingId) {
        const existing = await getRecord(state.editingId);
        if (!existing) throw new Error("The selected record no longer exists.");
        let serial = existing.serial;
        if (existing.dateKey !== payload.dateKey) {
            await shiftSerials(existing.dateKey, existing.serial + 1, -1);
            serial = (await countByDate(payload.dateKey)) + 1;
        }
        const updated = { ...existing, ...payload, serial, updatedAt: now };
        await putRecord(updated);
        return updated;
    }

    const record = {
        ...payload,
        id: crypto.randomUUID ? crypto.randomUUID() : `row-${Date.now()}-${Math.random().toString(16).slice(2)}`,
        serial: (await countByDate(payload.dateKey)) + 1,
        createdAt: now,
        updatedAt: now
    };
    await putRecord(record);
    return record;
}

async function removeMovement(id) {
    const existing = await getRecord(id);
    if (!existing) return;
    await deleteRecordById(id);
    await shiftSerials(existing.dateKey, existing.serial + 1, -1);
}

async function startEdit(id) {
    const record = await getRecord(id);
    if (!record) {
        pushToast("That record no longer exists.", "error");
        return;
    }
    state.editingId = id;
    el.formTitle.textContent = "Edit movement record";
    el.submitBtn.textContent = "Update Record";
    el.cancelEditBtn.classList.remove("hidden");
    el.dateInput.value = record.dateKey;
    el.watchSelect.value = record.watch;
    el.aircraftSelect.value = record.aircraftType;
    el.regNoInput.value = record.regNo;
    el.arrivalTimeInput.value = record.arrivalTime;
    el.departureTimeInput.value = record.departureTime;
    el.soulsLandingInput.value = Number.parseInt(record.soulsLanding, 10);
    el.soulsTakeoffInput.value = Number.parseInt(record.soulsTakeoff, 10);
    el.crewOnBoardInput.value = Number.parseInt(record.crewOnBoard, 10);
    el.operatingCompanyInput.value = record.operatingCompany;
    el.destinationInput.value = record.destination;
    await updateSerialPreview();
    el.dateInput.scrollIntoView({ behavior: "smooth", block: "center" });
    el.dateInput.focus();
}

async function saveFromForm() {
    try {
        setBusy(state.editingId ? "Updating record" : "Saving record");
        const saved = await saveMovement(readForm());
        if (state.viewMode === "day") {
            state.viewDate = saved.dateKey;
            el.viewDateInput.value = saved.dateKey;
            await setMeta("viewDate", saved.dateKey);
        }
        pushToast(`${state.editingId ? "Updated" : "Saved"} S/N ${saved.serial} for ${fmtDate(saved.dateKey)}.`, "success");
        exitEditMode();
        resetForm({ preserveDate: true });
        await refreshDashboard();
        await renderTable();
    } catch (error) {
        await showModal({
            tone: "Validation",
            title: "Unable to save record",
            description: error.message || "The form contains invalid data.",
            confirmLabel: "Understood",
            hideClose: true
        });
    } finally {
        setBusy("Idle");
    }
}

async function deleteFromTable(id) {
    const record = await getRecord(id);
    if (!record) {
        pushToast("That record is already gone.", "warning");
        return;
    }
    const confirmed = await showModal({
        tone: "Warning",
        title: "Delete this movement record?",
        description: `This removes S/N ${record.serial} on ${fmtDate(record.dateKey)} and renumbers the remaining rows for that day.`,
        confirmLabel: "Delete Record",
        cancelLabel: "Cancel",
        confirmClass: "ghost-btn rounded-2xl border border-danger/30 bg-danger/10 text-danger hover:bg-danger/15 hover:text-white"
    });
    if (!confirmed) return;
    setBusy("Deleting record");
    await removeMovement(id);
    if (state.editingId === id) {
        exitEditMode();
        resetForm({ preserveDate: true });
    }
    pushToast("Record deleted and daily numbering realigned.", "success");
    await refreshDashboard();
    await renderTable();
    setBusy("Idle");
}

async function navigateRecordedDay(direction) {
    if (!state.availableDates.length) return;
    const sorted = [...state.availableDates].sort();
    let index = sorted.indexOf(state.viewDate);
    if (index === -1) index = direction > 0 ? -1 : sorted.length;
    state.viewDate = sorted[Math.min(sorted.length - 1, Math.max(0, index + direction))];
    state.page = 1;
    el.viewDateInput.value = state.viewDate;
    el.dateInput.value = state.viewDate;
    await persistViewState();
    await updateSerialPreview();
    await renderTable();
}

async function jumpToNearestDay() {
    if (!state.availableDates.length) return;
    const current = el.viewDateInput.value || todayYmd();
    const sorted = [...state.availableDates].sort();
    let nearest = sorted[0];
    let delta = Number.POSITIVE_INFINITY;
    sorted.forEach((dateKey) => {
        const distance = Math.abs(new Date(dateKey).getTime() - new Date(current).getTime());
        if (distance < delta) {
            delta = distance;
            nearest = dateKey;
        }
    });
    state.viewDate = nearest;
    state.page = 1;
    el.viewDateInput.value = nearest;
    if (!el.dateInput.value) el.dateInput.value = nearest;
    await persistViewState();
    await updateSerialPreview();
    await renderTable();
}

function downloadBlob(blob, name) {
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = name;
    link.click();
    window.setTimeout(() => URL.revokeObjectURL(link.href), 1000);
}

async function ensurePermission(handle, mode = "readwrite") {
    if (!handle || typeof handle.queryPermission !== "function") return false;
    const options = { mode };
    if (await handle.queryPermission(options) === "granted") return true;
    return (await handle.requestPermission(options)) === "granted";
}

async function connectProjectFolder() {
    if (typeof window.showDirectoryPicker !== "function") {
        await showModal({
            tone: "Browser Support",
            title: "Folder linking is not available here",
            description: "This browser does not expose the File System Access API for directory linking. Export actions will fall back to normal downloads."
        });
        return;
    }
    try {
        const handle = await window.showDirectoryPicker({ mode: "readwrite" });
        state.projectDirectoryHandle = handle;
        state.projectDirectoryName = handle.name || "Project folder";
        await setMeta("projectDirectoryHandle", handle);
        await setMeta("projectDirectoryName", state.projectDirectoryName);
        pushToast(`Connected folder: ${state.projectDirectoryName}`, "success");
        await refreshDashboard();
    } catch (error) {
        if (error && error.name !== "AbortError") pushToast("Could not connect the project folder.", "error");
    }
}

async function getLinkedWorkbookHandleForWrite() {
    if (state.linkedWorkbookHandle && await ensurePermission(state.linkedWorkbookHandle, "readwrite")) {
        return state.linkedWorkbookHandle;
    }
    if (state.projectDirectoryHandle && await ensurePermission(state.projectDirectoryHandle, "readwrite")) {
        const handle = await state.projectDirectoryHandle.getFileHandle(DEFAULT_WORKBOOK_NAME, { create: true });
        state.linkedWorkbookHandle = handle;
        state.linkedWorkbookName = handle.name || DEFAULT_WORKBOOK_NAME;
        await setMeta("linkedWorkbookHandle", handle);
        await setMeta("linkedWorkbookName", state.linkedWorkbookName);
        return handle;
    }
    if (typeof window.showSaveFilePicker === "function") {
        const handle = await window.showSaveFilePicker({
            suggestedName: DEFAULT_WORKBOOK_NAME,
            types: [{ description: "Excel Workbook", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] } }]
        });
        state.linkedWorkbookHandle = handle;
        state.linkedWorkbookName = handle.name || DEFAULT_WORKBOOK_NAME;
        await setMeta("linkedWorkbookHandle", handle);
        await setMeta("linkedWorkbookName", state.linkedWorkbookName);
        return handle;
    }
    return null;
}

function createSheet(workbook, index) {
    const sheet = workbook.addWorksheet(index === 1 ? MOVEMENT_SHEET_PREFIX : `${MOVEMENT_SHEET_PREFIX} ${index}`);
    sheet.columns = [
        { width: 8 }, { width: 16 }, { width: 18 }, { width: 18 }, { width: 16 }, { width: 18 },
        { width: 18 }, { width: 22 }, { width: 22 }, { width: 16 }, { width: 24 }, { width: 20 }
    ];
    sheet.views = [{ state: "frozen", ySplit: 1 }];
    const row = sheet.getRow(1);
    row.values = EXPORT_HEADERS;
    row.height = 24;
    row.eachCell((cell) => {
        cell.font = { bold: true, color: { argb: "FFF8FAFC" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF11263A" } };
        cell.alignment = { vertical: "middle", horizontal: "center" };
        cell.border = {
            top: { style: "thin", color: { argb: "FF35526B" } },
            left: { style: "thin", color: { argb: "FF35526B" } },
            bottom: { style: "thin", color: { argb: "FF35526B" } },
            right: { style: "thin", color: { argb: "FF35526B" } }
        };
    });
    return sheet;
}

function addDayRow(sheet, rowNumber, dateKey) {
    sheet.mergeCells(`A${rowNumber}:L${rowNumber}`);
    const cell = sheet.getCell(`A${rowNumber}`);
    cell.value = `OPERATIONAL DAY - ${fmtDate(dateKey)}`;
    cell.font = { bold: true, color: { argb: "FFF3E8D6" }, size: 12 };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4D3417" } };
    cell.alignment = { vertical: "middle", horizontal: "left" };
    cell.border = {
        top: { style: "thin", color: { argb: "FFD6A86B" } },
        left: { style: "thin", color: { argb: "FFD6A86B" } },
        bottom: { style: "thin", color: { argb: "FFD6A86B" } },
        right: { style: "thin", color: { argb: "FFD6A86B" } }
    };
}

function addRecordSheetRow(sheet, rowNumber, record) {
    const row = sheet.getRow(rowNumber);
    row.values = [
        record.serial, fmtDateWb(record.dateKey), record.watch, record.aircraftType, record.regNo,
        record.arrivalTime, record.departureTime, record.soulsLanding, record.soulsTakeoff,
        record.crewOnBoard, record.operatingCompany, record.destination
    ];
    row.eachCell((cell, index) => {
        cell.alignment = { vertical: "middle", horizontal: index >= 11 ? "left" : "center" };
        cell.border = { bottom: { style: "thin", color: { argb: "FFDFE7EF" } } };
    });
}

async function buildWorkbook({ mode, dateKey, preserveLinkedSheets = false }) {
    const workbook = new ExcelJS.Workbook();
    if (preserveLinkedSheets && state.linkedWorkbookHandle && await ensurePermission(state.linkedWorkbookHandle, "read")) {
        try {
            const file = await state.linkedWorkbookHandle.getFile();
            await workbook.xlsx.load(await file.arrayBuffer());
        } catch (error) {
            console.warn("Could not reload linked workbook before update.", error);
        }
    }
    workbook.worksheets.filter((sheet) => sheet.name.startsWith(MOVEMENT_SHEET_PREFIX)).slice().forEach((sheet) => workbook.removeWorksheet(sheet.id));

    const records = await getAllRecordsForExport(mode, dateKey);
    if (!records.length) {
        const sheet = createSheet(workbook, 1);
        sheet.mergeCells("A2:L2");
        sheet.getCell("A2").value = "No aircraft movements available for the selected export scope.";
        sheet.getCell("A2").font = { italic: true, color: { argb: "FF4F6476" } };
        return workbook;
    }

    let sheetIndex = 1;
    let sheet = createSheet(workbook, sheetIndex);
    let rowNumber = 2;
    let activeDate = "";
    records.forEach((record) => {
        const needsBanner = activeDate !== record.dateKey;
        const requiredRows = needsBanner ? 2 : 1;
        if (rowNumber + requiredRows - 1 > EXCEL_MAX_ROWS) {
            sheetIndex += 1;
            sheet = createSheet(workbook, sheetIndex);
            rowNumber = 2;
            activeDate = "";
        }
        if (needsBanner) {
            addDayRow(sheet, rowNumber, record.dateKey);
            rowNumber += 1;
            activeDate = record.dateKey;
        }
        addRecordSheetRow(sheet, rowNumber, record);
        rowNumber += 1;
    });
    return workbook;
}

async function importWorkbookFromBuffer(buffer, linkedHandle = null, workbookName = "") {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const sheet = workbook.worksheets[0];
    if (!sheet) throw new Error("The workbook does not contain any worksheet.");

    const headerLookup = new Map();
    sheet.getRow(1).eachCell((cell, index) => headerLookup.set(normHeader(coerceCell(cell.value)), index));
    const aliases = {
        dateKey: ["DATE"],
        watch: ["WATCH ON DUTY", "WATCH"],
        aircraftType: ["AIRCRAFT TYPE"],
        regNo: ["REG NO", "REG. NO", "REGISTRATION NO", "REGISTRATION NUMBER"],
        arrivalTime: ["TIME OF ARRIVAL", "ARRIVAL TIME"],
        departureTime: ["TIME OF DEPARTURE", "TIME OF DEPARTMENT", "DEPARTURE TIME"],
        soulsLanding: ["SOULS ON BOARD LANDING", "SOULS LANDING"],
        soulsTakeoff: ["SOULS ON BOARD TAKE OFF", "SOULS TAKEOFF", "SOULS ON BOARD TAKEOFF"],
        crewOnBoard: ["CREW ON BOARD", "CREW"],
        operatingCompany: ["OPERATING COMPANY"],
        destination: ["DESTINATION"]
    };
    const columns = {};
    Object.entries(aliases).forEach(([key, values]) => {
        const hit = values.find((value) => headerLookup.has(normHeader(value)));
        if (hit) columns[key] = headerLookup.get(normHeader(hit));
    });
    if (!columns.dateKey || !columns.watch || !columns.aircraftType || !columns.regNo) {
        throw new Error("The workbook headers do not match the expected aircraft movement format.");
    }

    const imported = [];
    const serialMap = new Map();
    sheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;
        const firstCell = String(coerceCell(row.getCell(1).value) || "").trim().toUpperCase();
        if (firstCell.startsWith("OPERATIONAL DAY")) return;
        const dateKey = normImportedDate(coerceCell(row.getCell(columns.dateKey).value));
        if (!dateKey) return;
        const getText = (columnNumber) => {
            if (!columnNumber) return "";
            const value = coerceCell(row.getCell(columnNumber).value);
            return value instanceof Date ? toYmdLocal(value) : String(value || "").trim();
        };
        const serial = (serialMap.get(dateKey) || 0) + 1;
        serialMap.set(dateKey, serial);
        imported.push({
            id: crypto.randomUUID ? crypto.randomUUID() : `row-${Date.now()}-${Math.random().toString(16).slice(2)}`,
            serial,
            dateKey,
            watch: getText(columns.watch),
            aircraftType: getText(columns.aircraftType),
            regNo: getText(columns.regNo).toUpperCase(),
            arrivalTime: normTime(getText(columns.arrivalTime)),
            departureTime: normTime(getText(columns.departureTime)),
            soulsLanding: fmtSouls(getText(columns.soulsLanding)),
            soulsTakeoff: fmtSouls(getText(columns.soulsTakeoff)),
            crewOnBoard: fmtCrew(getText(columns.crewOnBoard)),
            operatingCompany: getText(columns.operatingCompany),
            destination: getText(columns.destination),
            createdAt: new Date().toISOString(),
            updatedAt: new Date().toISOString()
        });
    });

    await replaceAllRecords(imported);
    if (linkedHandle) {
        state.linkedWorkbookHandle = linkedHandle;
        state.linkedWorkbookName = workbookName || linkedHandle.name || "";
        await setMeta("linkedWorkbookHandle", linkedHandle);
        await setMeta("linkedWorkbookName", state.linkedWorkbookName);
    }
    exitEditMode();
    resetForm({ preserveDate: false });
    state.page = 1;
    await refreshDashboard();
    await renderTable();
    await showModal({
        tone: "Import Complete",
        title: "Workbook imported into the register",
        description: `Loaded ${fmtNumber(imported.length)} movement record${imported.length === 1 ? "" : "s"} from ${workbookName || "the workbook"}.`,
        details: imported.length ? "The register is now stored in IndexedDB, so edits, deletes, and filtered views no longer depend on rendering every row at once." : "No movement rows were found after the header row.",
        confirmLabel: "Continue"
    });
}

async function importWorkbook() {
    if (state.totalRecords > 0) {
        const ok = await showModal({
            tone: "Replace Register",
            title: "Import and replace the current register?",
            description: "Importing a workbook replaces the locally stored register with the workbook contents.",
            confirmLabel: "Replace Register",
            cancelLabel: "Cancel"
        });
        if (!ok) return;
    }
    if (typeof window.showOpenFilePicker === "function") {
        try {
            const [handle] = await window.showOpenFilePicker({
                multiple: false,
                types: [{ description: "Excel Workbook", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] } }]
            });
            if (!handle) return;
            const file = await handle.getFile();
            setBusy("Importing workbook");
            await importWorkbookFromBuffer(await file.arrayBuffer(), handle, file.name);
            pushToast(`Imported ${file.name}`, "success");
        } catch (error) {
            if (error && error.name !== "AbortError") {
                await showModal({ tone: "Import Error", title: "Workbook import failed", description: error.message || "The workbook could not be imported." });
            }
        } finally {
            setBusy("Idle");
        }
        return;
    }
    el.excelFileInput.click();
}

async function exportLinkedWorkbook() {
    try {
        setBusy("Preparing workbook");
        const workbook = await buildWorkbook({ mode: "all", dateKey: state.viewDate, preserveLinkedSheets: true });
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        const handle = await getLinkedWorkbookHandleForWrite();
        if (handle) {
            const writable = await handle.createWritable();
            await writable.write(buffer);
            await writable.close();
            pushToast(`Updated ${state.linkedWorkbookName || handle.name || DEFAULT_WORKBOOK_NAME}`, "success");
            await refreshDashboard();
            return;
        }
        downloadBlob(blob, DEFAULT_WORKBOOK_NAME);
        pushToast("Linked file access is unavailable here. Workbook downloaded instead.", "warning");
    } catch (error) {
        if (error && error.name === "AbortError") {
            setBusy("Idle");
            return;
        }
        await showModal({ tone: "Export Error", title: "Linked workbook update failed", description: error.message || "The workbook could not be updated." });
    } finally {
        setBusy("Idle");
    }
}

async function exportCurrentView() {
    try {
        setBusy("Exporting current view");
        const workbook = await buildWorkbook({ mode: state.viewMode, dateKey: state.viewDate, preserveLinkedSheets: false });
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        const name = state.viewMode === "all" ? "Aircraft_Movement_Full_Register.xlsx" : `Aircraft_Movement_${state.viewDate}.xlsx`;
        if (typeof window.showSaveFilePicker === "function") {
            try {
                const handle = await window.showSaveFilePicker({
                    suggestedName: name,
                    types: [{ description: "Excel Workbook", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] } }]
                });
                const writable = await handle.createWritable();
                await writable.write(buffer);
                await writable.close();
                pushToast(`Exported ${state.viewMode === "all" ? "full register" : fmtDate(state.viewDate)}.`, "success");
                return;
            } catch (error) {
                if (error && error.name === "AbortError") {
                    setBusy("Idle");
                    return;
                }
                throw error;
            }
        }
        downloadBlob(blob, name);
        pushToast(`Downloaded ${state.viewMode === "all" ? "full register" : fmtDate(state.viewDate)}.`, "success");
    } catch (error) {
        await showModal({ tone: "Export Error", title: "Current view export failed", description: error.message || "The selected view could not be exported." });
    } finally {
        setBusy("Idle");
    }
}

async function backupWorkbook() {
    try {
        setBusy("Creating backup");
        const workbook = await buildWorkbook({ mode: "all", dateKey: state.viewDate, preserveLinkedSheets: false });
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        const stamp = new Date().toISOString().replace(/[:.]/g, "-");
        const name = `Aircraft_Movement_Backup_${stamp}.xlsx`;
        if (state.projectDirectoryHandle && await ensurePermission(state.projectDirectoryHandle, "readwrite")) {
            const handle = await state.projectDirectoryHandle.getFileHandle(name, { create: true });
            const writable = await handle.createWritable();
            await writable.write(buffer);
            await writable.close();
            pushToast(`Backup saved to ${state.projectDirectoryName}`, "success");
            return;
        }
        if (typeof window.showSaveFilePicker === "function") {
            try {
                const handle = await window.showSaveFilePicker({
                    suggestedName: name,
                    types: [{ description: "Excel Workbook", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] } }]
                });
                const writable = await handle.createWritable();
                await writable.write(buffer);
                await writable.close();
                pushToast("Backup file saved.", "success");
                return;
            } catch (error) {
                if (error && error.name === "AbortError") {
                    setBusy("Idle");
                    return;
                }
                throw error;
            }
        }
        downloadBlob(blob, name);
        pushToast("Backup downloaded.", "success");
    } catch (error) {
        await showModal({ tone: "Backup Error", title: "Backup creation failed", description: error.message || "The backup workbook could not be created." });
    } finally {
        setBusy("Idle");
    }
}

async function clearRegister() {
    const ok = await showModal({
        tone: "Warning",
        title: "Clear the entire register?",
        description: "This removes every stored aircraft movement row from the local register. Linked folder and workbook settings stay in place.",
        confirmLabel: "Clear Register",
        cancelLabel: "Cancel",
        confirmClass: "ghost-btn rounded-2xl border border-danger/30 bg-danger/10 text-danger hover:bg-danger/15 hover:text-white"
    });
    if (!ok) return;
    setBusy("Clearing register");
    await clearRecords();
    exitEditMode();
    resetForm({ preserveDate: false });
    state.page = 1;
    await refreshDashboard();
    await renderTable();
    pushToast("Register cleared.", "success");
    setBusy("Idle");
}

async function requestPersistentStorage() {
    if (navigator.storage && navigator.storage.persist) {
        try {
            await navigator.storage.persist();
        } catch (error) {
            console.warn("Persistent storage request failed.", error);
        }
    }
}

function onTimeInput(event) {
    let value = event.target.value.trim();
    if (value === "-" || value === "--") {
        event.target.value = "-";
        return;
    }
    event.target.value = value.replace(/\D/g, "").slice(0, 4);
}

function bindEvents() {
    el.themeToggleBtn.addEventListener("click", () => {
        applyTheme(currentTheme() === "light" ? "dark" : "light");
    });

    el.submitBtn.addEventListener("click", saveFromForm);
    el.resetFormBtn.addEventListener("click", async () => {
        exitEditMode();
        resetForm({ preserveDate: true });
        await updateSerialPreview();
    });
    el.cancelEditBtn.addEventListener("click", async () => {
        exitEditMode();
        resetForm({ preserveDate: true });
        await updateSerialPreview();
    });

    el.dateInput.addEventListener("change", updateSerialPreview);
    el.arrivalTimeInput.addEventListener("input", onTimeInput);
    el.departureTimeInput.addEventListener("input", onTimeInput);

    el.viewModeSelect.addEventListener("change", async () => {
        state.viewMode = el.viewModeSelect.value;
        state.page = 1;
        await persistViewState();
        await renderTable();
    });

    el.viewDateInput.addEventListener("change", async () => {
        state.viewDate = el.viewDateInput.value || todayYmd();
        state.page = 1;
        await persistViewState();
        await renderTable();
    });

    el.pageSizeSelect.addEventListener("change", async () => {
        state.pageSize = Number(el.pageSizeSelect.value) || 25;
        state.page = 1;
        await persistViewState();
        await renderTable();
    });

    el.prevPageBtn.addEventListener("click", async () => {
        if (state.page <= 1) return;
        state.page -= 1;
        await renderTable();
    });

    el.nextPageBtn.addEventListener("click", async () => {
        if (state.page >= state.totalPages) return;
        state.page += 1;
        await renderTable();
    });

    el.prevDateBtn.addEventListener("click", async () => {
        state.viewMode = "day";
        el.viewModeSelect.value = "day";
        await navigateRecordedDay(-1);
        await persistViewState();
    });

    el.nextDateBtn.addEventListener("click", async () => {
        state.viewMode = "day";
        el.viewModeSelect.value = "day";
        await navigateRecordedDay(1);
        await persistViewState();
    });

    el.todayDateBtn.addEventListener("click", async () => {
        const today = todayYmd();
        state.viewMode = "day";
        state.viewDate = today;
        state.page = 1;
        el.viewModeSelect.value = "day";
        el.viewDateInput.value = today;
        el.dateInput.value = today;
        await persistViewState();
        await updateSerialPreview();
        await renderTable();
    });

    el.jumpDateBtn.addEventListener("click", async () => {
        state.viewMode = "day";
        el.viewModeSelect.value = "day";
        await jumpToNearestDay();
    });

    el.connectFolderBtn.addEventListener("click", connectProjectFolder);
    el.importWorkbookBtn.addEventListener("click", importWorkbook);
    el.updateWorkbookBtn.addEventListener("click", exportLinkedWorkbook);
    el.exportViewBtn.addEventListener("click", exportCurrentView);
    el.backupWorkbookBtn.addEventListener("click", backupWorkbook);
    el.clearRegisterBtn.addEventListener("click", clearRegister);

    el.excelFileInput.addEventListener("change", async (event) => {
        const file = event.target.files && event.target.files[0];
        if (!file) return;
        try {
            setBusy("Importing workbook");
            await importWorkbookFromBuffer(await file.arrayBuffer(), null, file.name);
            pushToast(`Imported ${file.name}`, "success");
        } catch (error) {
            await showModal({ tone: "Import Error", title: "Workbook import failed", description: error.message || "The workbook could not be imported." });
        } finally {
            el.excelFileInput.value = "";
            setBusy("Idle");
        }
    });

    el.tableBody.addEventListener("click", async (event) => {
        const button = event.target.closest("button[data-action]");
        if (!button) return;
        const row = event.target.closest("tr[data-id]");
        if (!row) return;
        if (button.dataset.action === "edit") {
            await startEdit(row.dataset.id);
            return;
        }
        if (button.dataset.action === "delete") {
            await deleteFromTable(row.dataset.id);
        }
    });

    el.modalCloseBtn.addEventListener("click", () => closeModal(false));
    el.messageModal.addEventListener("click", (event) => {
        if (event.target === el.messageModal) closeModal(false);
    });
}

async function init() {
    cacheDom();
    loadSelects();
    bindEvents();
    updateThemeControls();
    resetForm({ preserveDate: false });
    setBusy("Starting");

    state.db = await openDb();
    await requestPersistentStorage();
    await refreshMetaState();

    el.viewModeSelect.value = state.viewMode;
    el.viewDateInput.value = state.viewDate;
    el.pageSizeSelect.value = String(state.pageSize);
    el.dateInput.value = state.viewDate || todayYmd();
    if (!el.viewDateInput.value) {
        state.viewDate = todayYmd();
        el.viewDateInput.value = state.viewDate;
        el.dateInput.value = state.viewDate;
    }

    await refreshDashboard();
    await renderTable();
    setBusy("Idle");
}

init().catch(async (error) => {
    console.error(error);
    if (document.getElementById("messageModal")) {
        await showModal({
            tone: "Startup Error",
            title: "The app could not start correctly",
            description: error.message || "An unexpected startup error occurred."
        });
    } else {
        window.alert(error.message || "An unexpected startup error occurred.");
    }
});
