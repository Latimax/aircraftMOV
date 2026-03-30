(function (global) {
    const app = global.AircraftMovementApp = global.AircraftMovementApp || {};
    const { WATCH_OPTIONS, AIRCRAFT_OPTIONS } = app.constants;

    function toYmdLocal(date) {
        return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
    }

    function todayYmd() {
        return toYmdLocal(new Date());
    }

    function formatDate(dateKey, separator = "-") {
        if (!dateKey) return "-";
        const [year, month, day] = String(dateKey).split("-");
        if (!year || !month || !day) return String(dateKey);
        return [day.padStart(2, "0"), month.padStart(2, "0"), year].join(separator);
    }

    function formatWorkbookDate(dateKey) {
        return formatDate(dateKey, "/");
    }

    function formatClock(value) {
        if (!value) return "--:--";
        const date = value instanceof Date ? value : new Date(value);
        if (Number.isNaN(date.getTime())) return "--:--";
        return date.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
    }

    function escapeHtml(value) {
        return String(value ?? "")
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#39;");
    }

    function createLocalId(prefix = "id") {
        if (global.crypto && typeof global.crypto.randomUUID === "function") {
            return `${prefix}-${global.crypto.randomUUID()}`;
        }
        return `${prefix}-${Date.now()}-${Math.random().toString(16).slice(2)}`;
    }

    function createBlankEntryRow(overrides = {}) {
        return {
            rowId: overrides.rowId || createLocalId("entry"),
            recordId: null,
            dateKey: "",
            watch: "",
            aircraftType: "",
            regNo: "",
            arrivalTime: "",
            departureTime: "",
            soulsLanding: "",
            soulsTakeoff: "",
            crewOnBoard: "",
            operatingCompany: "",
            destination: "",
            lastSavedAt: "",
            statusText: "Ready",
            statusTone: "idle",
            ...overrides
        };
    }

    function sanitizeTimeInput(value) {
        const text = String(value || "").trim();
        if (!text) return "";
        if (text === "-" || text === "--") return "-";
        return text.replace(/\D/g, "").slice(0, 4);
    }

    function normalizeTime(value) {
        const text = sanitizeTimeInput(value);
        if (!text || text === "-") return "-";
        if (text.length === 3) return `0${text}`;
        if (text.length === 4) return text;
        return text;
    }

    function isValidTime(value) {
        if (value === "-") return true;
        if (!/^\d{4}$/.test(value)) return false;
        const hour = Number.parseInt(value.slice(0, 2), 10);
        const minute = Number.parseInt(value.slice(2), 10);
        return hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59;
    }

    function parseOptionalCount(value) {
        const text = String(value ?? "").trim();
        if (!text) return 0;
        if (!/^\d+$/.test(text)) return null;
        return Number.parseInt(text, 10);
    }

    function padNumber(value, width) {
        return String(value).padStart(width, "0");
    }

    function rowHasUserData(row, defaultDate) {
        const meaningfulValues = [
            row.dateKey && row.dateKey !== defaultDate ? row.dateKey : "",
            row.watch,
            row.aircraftType,
            row.regNo,
            row.arrivalTime && row.arrivalTime !== "-" ? row.arrivalTime : "",
            row.departureTime && row.departureTime !== "-" ? row.departureTime : "",
            row.soulsLanding && row.soulsLanding !== "0" ? row.soulsLanding : "",
            row.soulsTakeoff && row.soulsTakeoff !== "0" ? row.soulsTakeoff : "",
            row.crewOnBoard && row.crewOnBoard !== "0" ? row.crewOnBoard : "",
            row.operatingCompany,
            row.destination
        ];
        return meaningfulValues.some((value) => String(value || "").trim() !== "");
    }

    function validateEntryRow(row, defaultDate) {
        if (!rowHasUserData(row, defaultDate)) {
            return { kind: "blank" };
        }

        const dateKey = String(row.dateKey || defaultDate || "").trim();
        const watch = String(row.watch || "").trim();
        const aircraftType = String(row.aircraftType || "").trim();
        const regNo = String(row.regNo || "").trim().toUpperCase();
        const arrivalTime = normalizeTime(row.arrivalTime);
        const departureTime = normalizeTime(row.departureTime);
        const operatingCompany = String(row.operatingCompany || "").trim();
        const destination = String(row.destination || "").trim();
        const soulsLanding = parseOptionalCount(row.soulsLanding);
        const soulsTakeoff = parseOptionalCount(row.soulsTakeoff);
        const crewOnBoard = parseOptionalCount(row.crewOnBoard);

        if (!isValidTime(arrivalTime)) return { kind: "error", message: "Arrival must be HHMM or -" };
        if (!isValidTime(departureTime)) return { kind: "error", message: "Departure must be HHMM or -" };
        if (soulsLanding == null) return { kind: "error", message: "Souls landing must be a whole number" };
        if (soulsTakeoff == null) return { kind: "error", message: "Souls take off must be a whole number" };
        if (crewOnBoard == null) return { kind: "error", message: "Crew must be a whole number" };

        if (!watch || !aircraftType || !regNo || !operatingCompany || !destination) {
            return { kind: "draft", message: "Waiting for required fields" };
        }

        return {
            kind: "valid",
            payload: {
                dateKey,
                watch,
                aircraftType,
                regNo,
                arrivalTime,
                departureTime,
                soulsLanding: padNumber(soulsLanding, 3),
                soulsTakeoff: padNumber(soulsTakeoff, 3),
                crewOnBoard: padNumber(crewOnBoard, 2),
                operatingCompany,
                destination
            }
        };
    }

    function sortRecordsForDisplay(records) {
        return [...records].sort((left, right) => {
            if (left.dateKey !== right.dateKey) return right.dateKey.localeCompare(left.dateKey);
            return left.serial - right.serial;
        });
    }

    function sortRecordsForWorkbook(records) {
        return [...records].sort((left, right) => {
            if (left.dateKey !== right.dateKey) return left.dateKey.localeCompare(right.dateKey);
            return left.serial - right.serial;
        });
    }

    function countByLabel(records, key, baseLabels) {
        const counts = new Map(baseLabels.map((label) => [label, 0]));
        records.forEach((record) => {
            const label = String(record[key] || "").trim();
            if (!label) return;
            counts.set(label, (counts.get(label) || 0) + 1);
        });
        return Array.from(counts.entries()).map(([label, count]) => ({ label, count }));
    }

    function buildSummaryCounts(records, watchOptions = WATCH_OPTIONS, aircraftOptions = AIRCRAFT_OPTIONS) {
        const extraAircraft = Array.from(new Set(records.map((record) => String(record.aircraftType || "").trim()).filter(Boolean)))
            .filter((label) => !aircraftOptions.includes(label))
            .sort((left, right) => left.localeCompare(right));

        return {
            totalRecords: records.length,
            watchCounts: countByLabel(records, "watch", watchOptions),
            aircraftCounts: countByLabel(records, "aircraftType", [...aircraftOptions, ...extraAircraft])
        };
    }

    function normalizeHeader(value) {
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

    function normalizeImportedDate(value) {
        if (!value && value !== 0) return "";
        if (value instanceof Date && !Number.isNaN(value.getTime())) return toYmdLocal(value);
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

    app.utils = {
        toYmdLocal,
        todayYmd,
        formatDate,
        formatWorkbookDate,
        formatClock,
        escapeHtml,
        createLocalId,
        createBlankEntryRow,
        sanitizeTimeInput,
        normalizeTime,
        isValidTime,
        rowHasUserData,
        validateEntryRow,
        sortRecordsForDisplay,
        sortRecordsForWorkbook,
        buildSummaryCounts,
        normalizeHeader,
        coerceCell,
        normalizeImportedDate
    };
})(window);
