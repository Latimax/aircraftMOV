(function (global) {
    const app = global.AircraftMovementApp = global.AircraftMovementApp || {};
    const {
        DEFAULT_ENTRY_ROWS,
        ENTRY_ROW_INCREMENT,
        ENTRY_FIELD_NAMES,
        DEFAULT_EXPORT_NAME,
        WATCH_OPTIONS,
        AIRCRAFT_OPTIONS,
        AUTO_SAVE_DELAY_MS,
        WORKBOOK_SYNC_DELAY_MS
    } = app.constants;
    const utils = app.utils;
    const store = app.store;
    const workbook = app.workbook;

    const state = {
        records: [],
        entryRows: [],
        defaultDate: utils.todayYmd(),
        savedViewMode: "all",
        savedViewDate: utils.todayYmd(),
        importedRecordCount: 0,
        currentShiftOnly: false,
        shiftSessionsByDate: {},
        workbookSourceLabel: "Checking project workbook...",
        workbookSourceName: "",
        syncHandle: null,
        syncHandleName: "",
        rowSaveTimers: new Map(),
        draftPersistTimer: null,
        workbookSyncTimer: null,
        workbookSyncRunning: false
    };

    const el = {};

    function cacheDom() {
        [
            "workbookSourceStatus",
            "saveStatus",
            "importedRecordCount",
            "totalRegisterCount",
            "savedRecordCount",
            "currentShiftRecordCount",
            "entryRowCount",
            "defaultDateInput",
            "savedViewMode",
            "savedViewDateInput",
            "addRowsBtn",
            "startShiftBtn",
            "toggleShiftViewBtn",
            "reloadWorkbookBtn",
            "downloadWorkbookBtn",
            "connectSyncBtn",
            "shiftStatus",
            "syncStatus",
            "statsDateLabel",
            "aircraftStats",
            "watchStats",
            "entrySheetBody",
            "savedTableBody",
            "savedTableSummary",
            "aircraftSuggestionList",
            "toastContainer"
        ].forEach((id) => {
            el[id] = document.getElementById(id);
        });
    }

    function recordMap() {
        return new Map(state.records.map((record) => [record.id, record]));
    }

    function setSaveStatus(message, tone = "idle") {
        el.saveStatus.textContent = message;
        el.saveStatus.dataset.tone = tone;
    }

    function setSyncStatus(message, tone = "muted") {
        el.syncStatus.textContent = message;
        el.syncStatus.dataset.tone = tone;
    }

    function pushToast(message, tone = "success") {
        const toast = document.createElement("div");
        toast.className = "toast";
        toast.dataset.tone = tone;
        toast.textContent = message;
        el.toastContainer.appendChild(toast);
        global.setTimeout(() => {
            toast.style.opacity = "0";
            toast.style.transform = "translateY(-4px)";
            global.setTimeout(() => toast.remove(), 220);
        }, 2800);
    }

    function createDefaultEntryRow(overrides = {}) {
        return utils.createBlankEntryRow({ dateKey: state.defaultDate, ...overrides });
    }

    function createEntryRows(count) {
        return Array.from({ length: count }, () => createDefaultEntryRow());
    }

    function ensureMinimumEntryRows() {
        if (state.entryRows.length >= DEFAULT_ENTRY_ROWS) return;
        state.entryRows.push(...createEntryRows(DEFAULT_ENTRY_ROWS - state.entryRows.length));
    }

    function entryRowToRecordRow(record, rowId = "") {
        return {
            ...utils.createBlankEntryRow({ rowId: rowId || utils.createLocalId("entry") }),
            recordId: record.id,
            dateKey: record.dateKey,
            watch: record.watch,
            aircraftType: record.aircraftType,
            regNo: record.regNo,
            arrivalTime: record.arrivalTime,
            departureTime: record.departureTime,
            soulsLanding: String(Number.parseInt(record.soulsLanding, 10) || 0),
            soulsTakeoff: String(Number.parseInt(record.soulsTakeoff, 10) || 0),
            crewOnBoard: String(Number.parseInt(record.crewOnBoard, 10) || 0),
            operatingCompany: record.operatingCompany,
            destination: record.destination,
            lastSavedAt: record.updatedAt,
            statusText: `Saved ${utils.formatClock(record.updatedAt)}`,
            statusTone: "saved"
        };
    }

    function getShiftSession(dateKey = state.defaultDate) {
        return state.shiftSessionsByDate[dateKey] || null;
    }

    function getCurrentShiftRecords(records = state.records, dateKey = state.defaultDate) {
        const session = getShiftSession(dateKey);
        const dayRecords = records.filter((record) => record.dateKey === dateKey);
        if (!session) return dayRecords;
        return dayRecords.filter((record) => record.serial > session.startSerial);
    }

    function applyShiftFilter(records) {
        if (!state.currentShiftOnly) return records;
        const session = getShiftSession(state.defaultDate);
        if (!session) return records;
        return records.filter((record) => record.dateKey !== state.defaultDate || record.serial > session.startSerial);
    }

    function buildSavedViewRecords() {
        const baseRecords = state.savedViewMode !== "date"
            ? utils.sortRecordsForDisplay(state.records)
            : utils.sortRecordsForDisplay(state.records.filter((record) => record.dateKey === state.savedViewDate));
        return applyShiftFilter(baseRecords);
    }

    function currentSerialLabel(row) {
        const effectiveDate = row.dateKey || state.defaultDate;
        if (!utils.rowHasUserData(row, state.defaultDate)) return "Auto";
        const existing = row.recordId ? state.records.find((record) => record.id === row.recordId) : null;
        if (existing && existing.dateKey === effectiveDate) return String(existing.serial);
        const sameDateCount = state.records.filter((record) => record.dateKey === effectiveDate && record.id !== row.recordId).length;
        return String(sameDateCount + 1);
    }

    function rowHtml(row) {
        const watchOptions = ['<option value=""></option>', ...WATCH_OPTIONS.map((option) => {
            const selected = row.watch === option ? " selected" : "";
            return `<option value="${utils.escapeHtml(option)}"${selected}>${utils.escapeHtml(option)}</option>`;
        })].join("");

        return `
            <tr class="entry-row" data-row-id="${utils.escapeHtml(row.rowId)}" data-status-tone="${utils.escapeHtml(row.statusTone)}">
                <td class="sn-cell" data-role="serial">${utils.escapeHtml(currentSerialLabel(row))}</td>
                <td><input class="sheet-input" data-field="dateKey" type="date" value="${utils.escapeHtml(row.dateKey)}"></td>
                <td><select class="sheet-select sheet-select--watch" data-field="watch">${watchOptions}</select></td>
                <td><input class="sheet-input sheet-input--aircraft" data-field="aircraftType" type="text" list="aircraftSuggestionList" value="${utils.escapeHtml(row.aircraftType)}" placeholder="Type or select"></td>
                <td><input class="sheet-input sheet-input--reg" data-field="regNo" type="text" value="${utils.escapeHtml(row.regNo)}" placeholder="5NTON"></td>
                <td><input class="sheet-input sheet-input--time" data-field="arrivalTime" type="text" inputmode="numeric" maxlength="4" value="${utils.escapeHtml(row.arrivalTime)}" placeholder="HHMM or -"></td>
                <td><input class="sheet-input sheet-input--time" data-field="departureTime" type="text" inputmode="numeric" maxlength="4" value="${utils.escapeHtml(row.departureTime)}" placeholder="HHMM or -"></td>
                <td><input class="sheet-input sheet-input--count" data-field="soulsLanding" type="text" inputmode="numeric" value="${utils.escapeHtml(row.soulsLanding)}" placeholder="0"></td>
                <td><input class="sheet-input sheet-input--count" data-field="soulsTakeoff" type="text" inputmode="numeric" value="${utils.escapeHtml(row.soulsTakeoff)}" placeholder="0"></td>
                <td><input class="sheet-input sheet-input--crew" data-field="crewOnBoard" type="text" inputmode="numeric" value="${utils.escapeHtml(row.crewOnBoard)}" placeholder="0"></td>
                <td><input class="sheet-input sheet-input--company" data-field="operatingCompany" type="text" value="${utils.escapeHtml(row.operatingCompany)}" placeholder="Operating company"></td>
                <td><input class="sheet-input" data-field="destination" type="text" value="${utils.escapeHtml(row.destination)}" placeholder="Destination"></td>
                <td><span class="status-badge" data-role="status" data-tone="${utils.escapeHtml(row.statusTone)}">${utils.escapeHtml(row.statusText)}</span></td>
                <td><button type="button" class="mini-action" data-action="reset-entry-row">Reset</button></td>
            </tr>
        `;
    }

    function renderEntrySheet() {
        el.entrySheetBody.innerHTML = state.entryRows.map(rowHtml).join("");
        el.entryRowCount.textContent = String(state.entryRows.length);
    }

    function savedDateDividerHtml(dateKey) {
        return `
            <tr class="saved-date-divider">
                <td colspan="13">
                    <div class="saved-date-divider__line">
                        <span class="saved-date-divider__label">Operational Day - ${utils.escapeHtml(utils.formatDate(dateKey))}</span>
                    </div>
                </td>
            </tr>
        `;
    }

    function savedRecordHtml(record) {
        return `
            <tr data-record-id="${utils.escapeHtml(record.id)}">
                <td>${record.serial}</td>
                <td>${utils.escapeHtml(utils.formatWorkbookDate(record.dateKey))}</td>
                <td>${utils.escapeHtml(record.watch)}</td>
                <td>${utils.escapeHtml(record.aircraftType)}</td>
                <td>${utils.escapeHtml(record.regNo)}</td>
                <td>${utils.escapeHtml(record.arrivalTime)}</td>
                <td>${utils.escapeHtml(record.departureTime)}</td>
                <td>${utils.escapeHtml(record.soulsLanding)}</td>
                <td>${utils.escapeHtml(record.soulsTakeoff)}</td>
                <td>${utils.escapeHtml(record.crewOnBoard)}</td>
                <td>${utils.escapeHtml(record.operatingCompany)}</td>
                <td>${utils.escapeHtml(record.destination)}</td>
                <td>
                    <div class="action-group">
                        <button type="button" class="mini-action" data-action="edit-saved-record">Edit</button>
                        <button type="button" class="mini-action btn-danger" data-action="delete-saved-record">Delete</button>
                    </div>
                </td>
            </tr>
        `;
    }

    function renderSavedTable() {
        const records = buildSavedViewRecords();
        if (!records.length) {
            el.savedTableBody.innerHTML = '<tr><td class="empty-state" colspan="13">No saved movement records for this view.</td></tr>';
            el.savedTableSummary.textContent = state.savedViewMode === "date"
                ? `0 saved records for ${utils.formatDate(state.savedViewDate)}.`
                : "0 saved records.";
            return;
        }

        if (state.savedViewMode === "date") {
            el.savedTableBody.innerHTML = records.map(savedRecordHtml).join("");
        } else {
            let activeDate = "";
            let html = "";
            records.forEach((record) => {
                if (record.dateKey !== activeDate) {
                    activeDate = record.dateKey;
                    html += savedDateDividerHtml(record.dateKey);
                }
                html += savedRecordHtml(record);
            });
            el.savedTableBody.innerHTML = html;
        }

        const shiftSession = getShiftSession(state.defaultDate);
        if (state.currentShiftOnly && shiftSession && (state.savedViewMode === "all" || state.savedViewDate === state.defaultDate)) {
            el.savedTableSummary.textContent = state.savedViewMode === "date"
                ? `${records.length} current shift record${records.length === 1 ? "" : "s"} for ${utils.formatDate(state.defaultDate)} after S/N ${shiftSession.startSerial}.`
                : `${records.length} record${records.length === 1 ? "" : "s"} shown with current shift filtering for ${utils.formatDate(state.defaultDate)}.`;
            return;
        }

        el.savedTableSummary.textContent = state.savedViewMode === "date"
            ? `${records.length} saved record${records.length === 1 ? "" : "s"} for ${utils.formatDate(state.savedViewDate)}.`
            : `${records.length} saved record${records.length === 1 ? "" : "s"} in the register.`;
    }

    function renderStats() {
        const dayRecords = state.records.filter((record) => record.dateKey === state.defaultDate);
        const currentShiftRecords = getCurrentShiftRecords(state.records, state.defaultDate);
        const summary = utils.buildSummaryCounts(dayRecords, WATCH_OPTIONS, AIRCRAFT_OPTIONS);
        el.statsDateLabel.textContent = `Statistics for ${utils.formatDate(state.defaultDate)}.`;
        el.importedRecordCount.textContent = String(state.importedRecordCount || 0);
        el.totalRegisterCount.textContent = String(state.records.length);
        el.savedRecordCount.textContent = String(summary.totalRecords);
        el.currentShiftRecordCount.textContent = String(currentShiftRecords.length);
        el.aircraftStats.innerHTML = summary.aircraftCounts.map((item) => `
            <div class="count-chip">
                <span class="count-chip__label">${utils.escapeHtml(item.label)}</span>
                <span class="count-chip__value">${item.count}</span>
            </div>
        `).join("");
        el.watchStats.innerHTML = summary.watchCounts.map((item) => `
            <div class="count-chip">
                <span class="count-chip__label">${utils.escapeHtml(item.label)}</span>
                <span class="count-chip__value">${item.count}</span>
            </div>
        `).join("");
    }

    function renderShiftControls() {
        const session = getShiftSession(state.defaultDate);
        const currentShiftRecords = getCurrentShiftRecords(state.records, state.defaultDate);

        el.toggleShiftViewBtn.disabled = !session;
        el.toggleShiftViewBtn.textContent = state.currentShiftOnly ? "Show Full Day" : "Show Current Shift Only";

        if (!session) {
            el.shiftStatus.textContent = `No shift split has been started for ${utils.formatDate(state.defaultDate)}.`;
            el.shiftStatus.dataset.tone = "muted";
            return;
        }

        el.shiftStatus.textContent = state.currentShiftOnly
            ? `Current shift view is on for ${utils.formatDate(state.defaultDate)}. Showing ${currentShiftRecords.length} row(s) after S/N ${session.startSerial}.`
            : `Shift split saved for ${utils.formatDate(state.defaultDate)} after S/N ${session.startSerial}. Toggle current shift view when needed.`;
        el.shiftStatus.dataset.tone = state.currentShiftOnly ? "saved" : "draft";
    }

    function syncEntryRowDom(row, syncValues = false) {
        const tr = el.entrySheetBody.querySelector(`tr[data-row-id="${row.rowId}"]`);
        if (!tr) return;
        tr.dataset.statusTone = row.statusTone;
        const serialNode = tr.querySelector('[data-role="serial"]');
        const statusNode = tr.querySelector('[data-role="status"]');
        if (serialNode) serialNode.textContent = currentSerialLabel(row);
        if (statusNode) {
            statusNode.textContent = row.statusText;
            statusNode.dataset.tone = row.statusTone;
        }
        if (!syncValues) return;
        ENTRY_FIELD_NAMES.forEach((field) => {
            const input = tr.querySelector(`[data-field="${field}"]`);
            if (input) input.value = row[field] || "";
        });
    }

    function refreshAllEntrySerials() {
        state.entryRows.forEach((row) => syncEntryRowDom(row, false));
    }

    function rowMatchesRecord(row, record) {
        const validation = utils.validateEntryRow(row, state.defaultDate);
        if (validation.kind !== "valid") return false;
        const payload = validation.payload;
        return ENTRY_FIELD_NAMES.every((field) => payload[field] === record[field]);
    }

    function hydrateEntryRows(rawRows) {
        if (!Array.isArray(rawRows) || !rawRows.length) return createEntryRows(DEFAULT_ENTRY_ROWS);
        return rawRows.map((row) => {
            const hydrated = createDefaultEntryRow({ rowId: row.rowId || utils.createLocalId("entry") });
            hydrated.recordId = row.recordId || null;
            hydrated.lastSavedAt = row.lastSavedAt || "";
            ENTRY_FIELD_NAMES.forEach((field) => {
                hydrated[field] = row[field] || (field === "dateKey" ? state.defaultDate : "");
            });
            if (hydrated.recordId) {
                hydrated.statusText = row.lastSavedAt ? `Saved ${utils.formatClock(row.lastSavedAt)}` : "Saved";
                hydrated.statusTone = "saved";
            } else if (utils.rowHasUserData(hydrated, state.defaultDate)) {
                hydrated.statusText = "Waiting for required fields";
                hydrated.statusTone = "draft";
            }
            return hydrated;
        });
    }

    function trimTrailingBlankRows(minRows = DEFAULT_ENTRY_ROWS) {
        let nextLength = state.entryRows.length;
        while (nextLength > minRows) {
            const row = state.entryRows[nextLength - 1];
            if (row.recordId || utils.rowHasUserData(row, state.defaultDate)) break;
            nextLength -= 1;
        }
        if (nextLength !== state.entryRows.length) {
            state.entryRows = state.entryRows.slice(0, nextLength);
        }
    }

    function serializeEntryRows() {
        return state.entryRows.map((row) => {
            const payload = {
                rowId: row.rowId,
                recordId: row.recordId,
                lastSavedAt: row.lastSavedAt
            };
            ENTRY_FIELD_NAMES.forEach((field) => {
                payload[field] = row[field] || "";
            });
            return payload;
        });
    }

    function persistDraftRowsSoon() {
        global.clearTimeout(state.draftPersistTimer);
        state.draftPersistTimer = global.setTimeout(async () => {
            try {
                await store.setMeta("entryRows", serializeEntryRows());
            } catch (error) {
                console.error(error);
                setSaveStatus("Could not store row drafts", "error");
            }
        }, 200);
    }

    function setRowStatus(row, text, tone) {
        row.statusText = text;
        row.statusTone = tone;
        syncEntryRowDom(row, false);
    }

    async function refreshRecordsAndViews() {
        state.records = await store.getAllRecords();
        renderStats();
        renderShiftControls();
        renderSavedTable();
        refreshAllEntrySerials();
    }

    async function syncWorkbookNow() {
        if (!state.syncHandle || state.workbookSyncRunning) return;
        state.workbookSyncRunning = true;
        try {
            await workbook.writeWorkbookToHandle(state.records, state.syncHandle);
            setSyncStatus(`Excel auto sync updated ${state.syncHandleName || "the workbook"} at ${utils.formatClock(new Date().toISOString())}.`, "saved");
        } catch (error) {
            console.error(error);
            setSyncStatus(error.message || "Excel auto sync failed.", "error");
            pushToast(error.message || "Excel auto sync failed.", "error");
        } finally {
            state.workbookSyncRunning = false;
        }
    }

    function scheduleWorkbookSync() {
        if (!state.syncHandle) {
            setSyncStatus("Changes are saved locally. Connect auto sync if you want the Excel file rewritten too.", "draft");
            return;
        }
        global.clearTimeout(state.workbookSyncTimer);
        state.workbookSyncTimer = global.setTimeout(() => {
            void syncWorkbookNow();
        }, WORKBOOK_SYNC_DELAY_MS);
    }

    async function saveRow(rowId) {
        const row = state.entryRows.find((item) => item.rowId === rowId);
        if (!row) return;
        const validation = utils.validateEntryRow(row, state.defaultDate);

        if (validation.kind === "blank") {
            setRowStatus(row, row.recordId ? "Reset restores saved values" : "Ready", row.recordId ? "draft" : "idle");
            persistDraftRowsSoon();
            return;
        }

        if (validation.kind === "draft") {
            setRowStatus(row, validation.message, "draft");
            persistDraftRowsSoon();
            return;
        }

        if (validation.kind === "error") {
            setRowStatus(row, validation.message, "error");
            persistDraftRowsSoon();
            return;
        }

        setRowStatus(row, "Saving...", "saving");
        setSaveStatus("Saving row...", "draft");

        try {
            const saved = await store.upsertMovement(validation.payload, row.recordId);
            const savedRow = entryRowToRecordRow(saved, row.rowId);
            Object.assign(row, savedRow);
            syncEntryRowDom(row, true);
            await refreshRecordsAndViews();
            persistDraftRowsSoon();
            setSaveStatus(`Saved automatically at ${utils.formatClock(saved.updatedAt)}`, "saved");
            scheduleWorkbookSync();
        } catch (error) {
            console.error(error);
            setRowStatus(row, error.message || "Save failed", "error");
            setSaveStatus(error.message || "Could not save row", "error");
        }
    }

    function scheduleRowSave(rowId, immediate = false) {
        global.clearTimeout(state.rowSaveTimers.get(rowId));
        state.rowSaveTimers.delete(rowId);
        if (immediate) {
            void saveRow(rowId);
            return;
        }
        const timer = global.setTimeout(() => {
            state.rowSaveTimers.delete(rowId);
            void saveRow(rowId);
        }, AUTO_SAVE_DELAY_MS);
        state.rowSaveTimers.set(rowId, timer);
    }

    function updateRowFromInput(target) {
        const tr = target.closest("tr[data-row-id]");
        if (!tr) return null;
        const row = state.entryRows.find((item) => item.rowId === tr.dataset.rowId);
        if (!row) return null;
        const field = target.dataset.field;
        if (!field) return null;

        let nextValue = target.value;
        if (field === "arrivalTime" || field === "departureTime") {
            nextValue = utils.sanitizeTimeInput(nextValue);
            target.value = nextValue;
        }
        if (field === "soulsLanding" || field === "soulsTakeoff" || field === "crewOnBoard") {
            nextValue = String(nextValue || "").replace(/[^\d]/g, "");
            target.value = nextValue;
        }

        row[field] = nextValue;

        if (row.recordId) {
            setRowStatus(row, "Unsaved changes", "draft");
        } else if (utils.rowHasUserData(row, state.defaultDate)) {
            setRowStatus(row, "Waiting for required fields", "draft");
        } else {
            setRowStatus(row, "Ready", "idle");
        }

        persistDraftRowsSoon();
        return row;
    }

    function focusEntryRow(rowId) {
        const tr = el.entrySheetBody.querySelector(`tr[data-row-id="${rowId}"]`);
        if (!tr) return;
        const firstField = tr.querySelector("[data-field]");
        if (firstField) firstField.focus();
        tr.scrollIntoView({ behavior: "smooth", block: "center" });
    }

    function focusField(field) {
        if (!field) return;
        field.focus();
        if (typeof field.select === "function" && field.tagName !== "SELECT" && field.type !== "date") {
            field.select();
        }
    }

    function getRowFields(tr) {
        return Array.from(tr.querySelectorAll("[data-field]"));
    }

    function moveAcrossEntryFields(currentField, direction) {
        const tr = currentField.closest("tr[data-row-id]");
        if (!tr) return false;

        const fields = getRowFields(tr);
        const currentIndex = fields.indexOf(currentField);
        if (currentIndex === -1) return false;

        const sameRowTarget = fields[currentIndex + direction];
        if (sameRowTarget) {
            focusField(sameRowTarget);
            return true;
        }

        const rows = Array.from(el.entrySheetBody.querySelectorAll("tr[data-row-id]"));
        const rowIndex = rows.indexOf(tr);
        const adjacentRow = rows[rowIndex + direction];

        if (adjacentRow) {
            const adjacentFields = getRowFields(adjacentRow);
            focusField(direction > 0 ? adjacentFields[0] : adjacentFields[adjacentFields.length - 1]);
            return true;
        }

        if (direction > 0) {
            addEntryRows(1);
            const nextRows = Array.from(el.entrySheetBody.querySelectorAll("tr[data-row-id]"));
            const lastRow = nextRows[nextRows.length - 1];
            if (!lastRow) return false;
            focusField(getRowFields(lastRow)[0]);
            return true;
        }

        return false;
    }

    function findReusableEntryRow(recordId) {
        if (recordId) {
            const existing = state.entryRows.find((row) => row.recordId === recordId);
            if (existing) return existing;
        }
        return state.entryRows.find((row) => !row.recordId && !utils.rowHasUserData(row, state.defaultDate)) || null;
    }

    function addEntryRows(count) {
        state.entryRows.push(...createEntryRows(count));
        renderEntrySheet();
        persistDraftRowsSoon();
    }

    function applyDefaultDateToBlankRows(previousDate, nextDate) {
        state.entryRows.forEach((row) => {
            if (row.recordId) return;
            if (utils.rowHasUserData(row, previousDate)) return;
            row.dateKey = nextDate;
            row.statusText = "Ready";
            row.statusTone = "idle";
        });
    }

    async function startNewShift() {
        const currentDayRecords = state.records.filter((record) => record.dateKey === state.defaultDate);
        const lastSerial = currentDayRecords.reduce((max, record) => Math.max(max, record.serial), 0);

        state.shiftSessionsByDate[state.defaultDate] = {
            startSerial: lastSerial,
            startedAt: new Date().toISOString()
        };
        state.currentShiftOnly = true;
        state.savedViewMode = "date";
        state.savedViewDate = state.defaultDate;
        state.entryRows = createEntryRows(DEFAULT_ENTRY_ROWS);

        await persistSettings();
        await store.setMeta("entryRows", serializeEntryRows());
        el.savedViewMode.value = state.savedViewMode;
        el.savedViewDateInput.value = state.savedViewDate;
        el.savedViewDateInput.disabled = false;
        renderEntrySheet();
        renderStats();
        renderShiftControls();
        renderSavedTable();
        setSaveStatus(`New shift started for ${utils.formatDate(state.defaultDate)}`, "saved");
        pushToast(`New shift started after S/N ${lastSerial}.`, "success");
    }

    async function toggleShiftView() {
        const session = getShiftSession(state.defaultDate);
        if (!session) {
            pushToast("Start a new shift first for this working date.", "warning");
            return;
        }

        state.currentShiftOnly = !state.currentShiftOnly;
        await persistSettings();
        renderStats();
        renderShiftControls();
        renderSavedTable();
    }

    function resetEntryRow(rowId) {
        const rowIndex = state.entryRows.findIndex((row) => row.rowId === rowId);
        if (rowIndex === -1) return;
        const row = state.entryRows[rowIndex];
        const savedRecord = row.recordId ? state.records.find((record) => record.id === row.recordId) : null;
        state.entryRows[rowIndex] = savedRecord
            ? entryRowToRecordRow(savedRecord, row.rowId)
            : createDefaultEntryRow({ rowId: row.rowId });
        syncEntryRowDom(state.entryRows[rowIndex], true);
        persistDraftRowsSoon();
    }

    async function loadRecordIntoEntrySheet(recordId) {
        const record = state.records.find((item) => item.id === recordId);
        if (!record) {
            pushToast("That saved record no longer exists.", "warning");
            return;
        }

        let row = findReusableEntryRow(recordId);
        if (!row) {
            addEntryRows(ENTRY_ROW_INCREMENT);
            row = state.entryRows.find((item) => !item.recordId && !utils.rowHasUserData(item, state.defaultDate));
        }
        if (!row) return;

        const nextRow = entryRowToRecordRow(record, row.rowId);
        nextRow.statusText = "Loaded for editing";
        nextRow.statusTone = "draft";
        Object.assign(row, nextRow);
        syncEntryRowDom(row, true);
        persistDraftRowsSoon();
        focusEntryRow(row.rowId);
    }

    async function deleteSavedRecord(recordId) {
        const record = state.records.find((item) => item.id === recordId);
        if (!record) {
            pushToast("That saved record no longer exists.", "warning");
            return;
        }

        const ok = global.confirm(`Delete S/N ${record.serial} for ${utils.formatDate(record.dateKey)}?`);
        if (!ok) return;

        try {
            await store.deleteMovement(recordId);
            state.entryRows = state.entryRows.map((row) => (
                row.recordId === recordId ? createDefaultEntryRow({ rowId: row.rowId }) : row
            ));
            renderEntrySheet();
            await refreshRecordsAndViews();
            persistDraftRowsSoon();
            setSaveStatus(`Deleted S/N ${record.serial}`, "saved");
            scheduleWorkbookSync();
            pushToast("Saved record deleted.", "success");
        } catch (error) {
            console.error(error);
            setSaveStatus(error.message || "Delete failed", "error");
            pushToast(error.message || "Could not delete the record.", "error");
        }
    }

    async function connectAutoSync() {
        if (typeof global.showOpenFilePicker !== "function" && typeof global.showSaveFilePicker !== "function") {
            pushToast("This browser does not support direct Excel auto sync.", "warning");
            return;
        }

        try {
            let handle = null;
            if (typeof global.showOpenFilePicker === "function") {
                const [picked] = await global.showOpenFilePicker({
                    multiple: false,
                    types: [{ description: "Excel Workbook", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] } }]
                });
                handle = picked || null;
            } else {
                handle = await global.showSaveFilePicker({
                    suggestedName: DEFAULT_EXPORT_NAME,
                    types: [{ description: "Excel Workbook", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] } }]
                });
            }

            if (!handle) return;
            state.syncHandle = handle;
            state.syncHandleName = handle.name || DEFAULT_EXPORT_NAME;
            await store.setMeta("syncWorkbookHandle", handle);
            await store.setMeta("syncWorkbookName", state.syncHandleName);
            setSyncStatus(`Auto sync connected to ${state.syncHandleName}.`, "saved");
            await syncWorkbookNow();
            pushToast(`Auto sync connected to ${state.syncHandleName}.`, "success");
        } catch (error) {
            if (error && error.name === "AbortError") return;
            console.error(error);
            setSyncStatus(error.message || "Could not connect auto sync.", "error");
            pushToast(error.message || "Could not connect auto sync.", "error");
        }
    }

    async function downloadWorkbookFile() {
        try {
            setSaveStatus("Preparing Excel download...", "draft");
            await workbook.downloadWorkbook(state.records, state.workbookSourceName || DEFAULT_EXPORT_NAME);
            setSaveStatus(`Downloaded Excel at ${utils.formatClock(new Date().toISOString())}`, "saved");
            pushToast("Excel workbook downloaded.", "success");
        } catch (error) {
            console.error(error);
            setSaveStatus(error.message || "Excel download failed", "error");
            pushToast(error.message || "Excel download failed.", "error");
        }
    }

    async function importBundledWorkbook(forceReload = false) {
        if (forceReload && state.records.length) {
            const ok = global.confirm("Reloading the project workbook will replace the current saved register. Continue?");
            if (!ok) return false;
        }

        setSaveStatus("Loading workbook from project...", "draft");

        try {
            const bundled = await workbook.fetchBundledWorkbook();
            if (!bundled) {
                state.workbookSourceLabel = "Project workbook was not found in this folder.";
                el.workbookSourceStatus.textContent = state.workbookSourceLabel;
                setSaveStatus("Project workbook not found", "error");
                return false;
            }

            const imported = await workbook.parseWorkbookBuffer(bundled.buffer);
            await store.replaceAllRecords(imported);
            state.records = await store.getAllRecords();
            state.importedRecordCount = imported.length;
            state.shiftSessionsByDate = {};
            state.currentShiftOnly = false;
            state.entryRows = createEntryRows(DEFAULT_ENTRY_ROWS);
            await store.setMeta("entryRows", serializeEntryRows());
            await store.setMeta("workbookSourceName", bundled.name);
            await store.setMeta("importedRecordCount", state.importedRecordCount);
            await store.setMeta("shiftSessionsByDate", state.shiftSessionsByDate);
            await store.setMeta("currentShiftOnly", state.currentShiftOnly);
            state.workbookSourceName = bundled.name;
            state.workbookSourceLabel = `Loaded automatically from ${bundled.name}`;
            el.workbookSourceStatus.textContent = state.workbookSourceLabel;
            renderEntrySheet();
            await refreshRecordsAndViews();
            setSaveStatus(`Loaded ${imported.length} workbook rows`, "saved");
            pushToast(`Loaded ${imported.length} rows from ${bundled.name}.`, "success");
            return true;
        } catch (error) {
            console.error(error);
            state.workbookSourceLabel = error.message || "Project workbook could not be loaded.";
            el.workbookSourceStatus.textContent = state.workbookSourceLabel;
            setSaveStatus(error.message || "Workbook load failed", "error");
            pushToast(error.message || "Workbook load failed.", "error");
            return false;
        }
    }

    function populateAircraftSuggestions() {
        el.aircraftSuggestionList.innerHTML = AIRCRAFT_OPTIONS.map((option) => `<option value="${utils.escapeHtml(option)}"></option>`).join("");
    }

    async function persistSettings() {
        await store.setMeta("defaultDate", state.defaultDate);
        await store.setMeta("savedViewMode", state.savedViewMode);
        await store.setMeta("savedViewDate", state.savedViewDate);
        await store.setMeta("currentShiftOnly", state.currentShiftOnly);
        await store.setMeta("shiftSessionsByDate", state.shiftSessionsByDate);
    }

    function reconcileEntryRowsWithSavedRecords() {
        const map = recordMap();
        state.entryRows = state.entryRows.map((row) => {
            if (!row.recordId) {
                if (utils.rowHasUserData(row, state.defaultDate)) {
                    row.statusText = "Waiting for required fields";
                    row.statusTone = "draft";
                } else {
                    row.statusText = "Ready";
                    row.statusTone = "idle";
                }
                return row;
            }

            const saved = map.get(row.recordId);
            if (!saved) {
                row.recordId = null;
                row.lastSavedAt = "";
                row.statusText = utils.rowHasUserData(row, state.defaultDate) ? "Waiting for required fields" : "Ready";
                row.statusTone = utils.rowHasUserData(row, state.defaultDate) ? "draft" : "idle";
                return row;
            }

            row.lastSavedAt = saved.updatedAt;
            if (rowMatchesRecord(row, saved)) {
                row.statusText = `Saved ${utils.formatClock(saved.updatedAt)}`;
                row.statusTone = "saved";
            } else {
                row.statusText = "Unsaved changes";
                row.statusTone = "draft";
            }
            return row;
        });
    }

    function bindEvents() {
        el.addRowsBtn.addEventListener("click", () => {
            addEntryRows(ENTRY_ROW_INCREMENT);
            pushToast(`Added ${ENTRY_ROW_INCREMENT} more entry rows.`, "success");
        });

        el.startShiftBtn.addEventListener("click", () => {
            void startNewShift();
        });

        el.toggleShiftViewBtn.addEventListener("click", () => {
            void toggleShiftView();
        });

        el.reloadWorkbookBtn.addEventListener("click", () => {
            void importBundledWorkbook(true);
        });

        el.downloadWorkbookBtn.addEventListener("click", () => {
            void downloadWorkbookFile();
        });

        el.connectSyncBtn.addEventListener("click", () => {
            void connectAutoSync();
        });

        el.defaultDateInput.addEventListener("change", async () => {
            const previousDate = state.defaultDate;
            state.defaultDate = el.defaultDateInput.value || utils.todayYmd();
            if (state.savedViewMode === "date" && state.savedViewDate === previousDate) {
                state.savedViewDate = state.defaultDate;
                el.savedViewDateInput.value = state.savedViewDate;
            }
            if (!getShiftSession(state.defaultDate)) {
                state.currentShiftOnly = false;
            }
            applyDefaultDateToBlankRows(previousDate, state.defaultDate);
            renderEntrySheet();
            renderStats();
            renderShiftControls();
            renderSavedTable();
            refreshAllEntrySerials();
            await persistSettings();
            persistDraftRowsSoon();
        });

        el.savedViewMode.addEventListener("change", async () => {
            state.savedViewMode = el.savedViewMode.value;
            el.savedViewDateInput.disabled = state.savedViewMode !== "date";
            renderSavedTable();
            renderShiftControls();
            await persistSettings();
        });

        el.savedViewDateInput.addEventListener("change", async () => {
            state.savedViewDate = el.savedViewDateInput.value || utils.todayYmd();
            renderSavedTable();
            renderShiftControls();
            await persistSettings();
        });

        el.entrySheetBody.addEventListener("input", (event) => {
            const fieldTarget = event.target.closest("[data-field]");
            if (!fieldTarget) return;
            const row = updateRowFromInput(fieldTarget);
            if (!row) return;
            scheduleRowSave(row.rowId, false);
        });

        el.entrySheetBody.addEventListener("change", (event) => {
            const fieldTarget = event.target.closest("[data-field]");
            if (!fieldTarget) return;
            const row = updateRowFromInput(fieldTarget);
            if (!row) return;
            scheduleRowSave(row.rowId, true);
        });

        el.entrySheetBody.addEventListener("keydown", (event) => {
            const fieldTarget = event.target.closest("[data-field]");
            if (!fieldTarget) return;

            const moveBackward = event.key === "ArrowLeft" || (event.key === "Tab" && event.shiftKey);
            const moveForward = event.key === "ArrowRight" || (event.key === "Tab" && !event.shiftKey);
            if (!moveBackward && !moveForward) return;

            const moved = moveAcrossEntryFields(fieldTarget, moveBackward ? -1 : 1);
            if (moved) event.preventDefault();
        });

        el.entrySheetBody.addEventListener("click", (event) => {
            const button = event.target.closest("button[data-action]");
            if (!button) return;
            const tr = button.closest("tr[data-row-id]");
            if (!tr) return;
            if (button.dataset.action === "reset-entry-row") {
                resetEntryRow(tr.dataset.rowId);
            }
        });

        el.savedTableBody.addEventListener("click", (event) => {
            const button = event.target.closest("button[data-action]");
            if (!button) return;
            const tr = button.closest("tr[data-record-id]");
            if (!tr) return;
            if (button.dataset.action === "edit-saved-record") {
                void loadRecordIntoEntrySheet(tr.dataset.recordId);
                return;
            }
            if (button.dataset.action === "delete-saved-record") {
                void deleteSavedRecord(tr.dataset.recordId);
            }
        });
    }

    async function requestPersistentStorage() {
        if (navigator.storage && navigator.storage.persist) {
            try {
                await navigator.storage.persist();
            } catch (error) {
                console.warn("Persistent storage could not be requested.", error);
            }
        }
    }

    async function loadInitialState() {
        await store.openDb();
        await requestPersistentStorage();

        state.defaultDate = (await store.getMeta("defaultDate")) || utils.todayYmd();
        state.savedViewMode = (await store.getMeta("savedViewMode")) || "all";
        state.savedViewDate = (await store.getMeta("savedViewDate")) || state.defaultDate;
        state.importedRecordCount = Number((await store.getMeta("importedRecordCount")) || 0);
        state.currentShiftOnly = Boolean(await store.getMeta("currentShiftOnly"));
        state.shiftSessionsByDate = (await store.getMeta("shiftSessionsByDate")) || {};
        state.syncHandle = (await store.getMeta("syncWorkbookHandle")) || null;
        state.syncHandleName = (await store.getMeta("syncWorkbookName")) || "";
        state.workbookSourceName = (await store.getMeta("workbookSourceName")) || "";
        state.records = await store.getAllRecords();
        if (!state.importedRecordCount && state.workbookSourceName) {
            state.importedRecordCount = state.records.length;
        }
        state.entryRows = hydrateEntryRows(await store.getMeta("entryRows"));
        trimTrailingBlankRows(DEFAULT_ENTRY_ROWS);
        ensureMinimumEntryRows();

        if (!state.records.length) {
            await importBundledWorkbook(false);
            state.records = await store.getAllRecords();
        }

        if (!getShiftSession(state.defaultDate)) {
            state.currentShiftOnly = false;
        }

        reconcileEntryRowsWithSavedRecords();

        state.workbookSourceLabel = state.workbookSourceName
            ? `Loaded automatically from ${state.workbookSourceName}`
            : state.records.length
                ? "Using saved register data"
                : "No project workbook loaded";
        el.workbookSourceStatus.textContent = state.workbookSourceLabel;

        el.defaultDateInput.value = state.defaultDate;
        el.savedViewMode.value = state.savedViewMode;
        el.savedViewDateInput.value = state.savedViewDate;
        el.savedViewDateInput.disabled = state.savedViewMode !== "date";

        if (state.syncHandle) {
            setSyncStatus(`Auto sync ready for ${state.syncHandleName || "the selected workbook"}.`, "saved");
        } else {
            setSyncStatus("Excel auto sync is not connected.", "muted");
        }
    }

    async function init() {
        cacheDom();
        populateAircraftSuggestions();
        setSaveStatus("Starting...", "draft");
        await loadInitialState();
        renderEntrySheet();
        renderStats();
        renderShiftControls();
        renderSavedTable();
        bindEvents();
        setSaveStatus("Ready", "idle");
    }

    global.addEventListener("DOMContentLoaded", () => {
        init().catch((error) => {
            console.error(error);
            if (document.body) {
                setSaveStatus(error.message || "The page could not start", "error");
                pushToast(error.message || "The page could not start.", "error");
            } else {
                global.alert(error.message || "The page could not start.");
            }
        });
    });
})(window);
