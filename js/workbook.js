(function (global) {
    const app = global.AircraftMovementApp = global.AircraftMovementApp || {};
    const { WORKBOOK_CANDIDATES, DEFAULT_EXPORT_NAME, EXPORT_HEADERS } = app.constants;
    const utils = app.utils;

    async function fetchBundledWorkbook() {
        for (const name of WORKBOOK_CANDIDATES) {
            try {
                const response = await fetch(name, { cache: "no-store" });
                if (!response.ok) continue;
                const buffer = await response.arrayBuffer();
                if (!buffer.byteLength) continue;
                return { name: name.replace("./", ""), buffer };
            } catch (error) {
                console.warn(`Could not fetch ${name}.`, error);
            }
        }
        return null;
    }

    function createWorkbookShell() {
        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet("Aircraft Movement");
        sheet.columns = [
            { width: 8 }, { width: 16 }, { width: 18 }, { width: 18 }, { width: 16 }, { width: 18 },
            { width: 18 }, { width: 22 }, { width: 22 }, { width: 16 }, { width: 24 }, { width: 20 }
        ];
        sheet.views = [{ state: "frozen", ySplit: 1 }];
        const header = sheet.getRow(1);
        header.values = EXPORT_HEADERS;
        header.height = 24;
        header.eachCell((cell) => {
            cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1C3C23" } };
            cell.alignment = { vertical: "middle", horizontal: "center" };
            cell.border = {
                top: { style: "thin", color: { argb: "FF34533B" } },
                left: { style: "thin", color: { argb: "FF34533B" } },
                bottom: { style: "thin", color: { argb: "FF34533B" } },
                right: { style: "thin", color: { argb: "FF34533B" } }
            };
        });
        return { workbook, sheet };
    }

    function addDayBanner(sheet, rowNumber, dateKey) {
        sheet.mergeCells(`A${rowNumber}:L${rowNumber}`);
        const cell = sheet.getCell(`A${rowNumber}`);
        cell.value = `OPERATIONAL DAY - ${utils.formatDate(dateKey)}`;
        cell.font = { bold: true, color: { argb: "FF5B4700" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFBEAAE" } };
        cell.alignment = { vertical: "middle", horizontal: "left" };
        cell.border = {
            top: { style: "thin", color: { argb: "FFDAB85E" } },
            left: { style: "thin", color: { argb: "FFDAB85E" } },
            bottom: { style: "thin", color: { argb: "FFDAB85E" } },
            right: { style: "thin", color: { argb: "FFDAB85E" } }
        };
    }

    function addRecordRow(sheet, rowNumber, record) {
        const row = sheet.getRow(rowNumber);
        row.values = [
            record.serial,
            utils.formatWorkbookDate(record.dateKey),
            record.watch,
            record.aircraftType,
            record.regNo,
            record.arrivalTime,
            record.departureTime,
            record.soulsLanding,
            record.soulsTakeoff,
            record.crewOnBoard,
            record.operatingCompany,
            record.destination
        ];
        row.eachCell((cell) => {
            cell.alignment = { vertical: "middle", horizontal: "center" };
            cell.border = { bottom: { style: "thin", color: { argb: "FFE3E7E4" } } };
        });
        row.getCell(11).alignment = { vertical: "middle", horizontal: "left" };
        row.getCell(12).alignment = { vertical: "middle", horizontal: "left" };
    }

    async function buildWorkbook(records) {
        const { workbook, sheet } = createWorkbookShell();
        const ordered = utils.sortRecordsForWorkbook(records);
        if (!ordered.length) {
            sheet.mergeCells("A2:L2");
            sheet.getCell("A2").value = "No aircraft movements available.";
            return workbook;
        }

        let currentDate = "";
        let rowNumber = 2;
        ordered.forEach((record) => {
            if (record.dateKey !== currentDate) {
                currentDate = record.dateKey;
                addDayBanner(sheet, rowNumber, record.dateKey);
                rowNumber += 1;
            }
            addRecordRow(sheet, rowNumber, record);
            rowNumber += 1;
        });
        return workbook;
    }

    async function parseWorkbookBuffer(buffer) {
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(buffer);
        const sheet = workbook.worksheets[0];
        if (!sheet) throw new Error("The workbook does not contain any worksheet.");

        const headerLookup = new Map();
        sheet.getRow(1).eachCell((cell, index) => {
            headerLookup.set(utils.normalizeHeader(utils.coerceCell(cell.value)), index);
        });

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
            const hit = values.find((value) => headerLookup.has(utils.normalizeHeader(value)));
            if (hit) columns[key] = headerLookup.get(utils.normalizeHeader(hit));
        });

        if (!columns.dateKey || !columns.watch || !columns.aircraftType || !columns.regNo) {
            throw new Error("The workbook headers do not match the expected aircraft movement format.");
        }

        const imported = [];
        const serialMap = new Map();

        sheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return;
            const firstCell = String(utils.coerceCell(row.getCell(1).value) || "").trim().toUpperCase();
            if (firstCell.startsWith("OPERATIONAL DAY")) return;

            const dateKey = utils.normalizeImportedDate(utils.coerceCell(row.getCell(columns.dateKey).value));
            if (!dateKey) return;

            const readText = (columnNumber) => {
                if (!columnNumber) return "";
                const value = utils.coerceCell(row.getCell(columnNumber).value);
                return value instanceof Date ? utils.toYmdLocal(value) : String(value || "").trim();
            };

            const serial = (serialMap.get(dateKey) || 0) + 1;
            serialMap.set(dateKey, serial);

            imported.push({
                id: utils.createLocalId("record"),
                serial,
                orderKey: imported.length + 1,
                dateKey,
                watch: readText(columns.watch),
                aircraftType: readText(columns.aircraftType),
                regNo: readText(columns.regNo).toUpperCase(),
                arrivalTime: utils.normalizeTime(readText(columns.arrivalTime)),
                departureTime: utils.normalizeTime(readText(columns.departureTime)),
                soulsLanding: String(readText(columns.soulsLanding) || "0").padStart(3, "0"),
                soulsTakeoff: String(readText(columns.soulsTakeoff) || "0").padStart(3, "0"),
                crewOnBoard: String(readText(columns.crewOnBoard) || "0").padStart(2, "0"),
                operatingCompany: readText(columns.operatingCompany),
                destination: readText(columns.destination),
                createdAt: new Date().toISOString(),
                updatedAt: new Date().toISOString()
            });
        });

        return imported;
    }

    function downloadBlob(blob, name) {
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = name;
        link.click();
        global.setTimeout(() => URL.revokeObjectURL(link.href), 1200);
    }

    async function downloadWorkbook(records, name = DEFAULT_EXPORT_NAME) {
        const workbook = await buildWorkbook(records);
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
        downloadBlob(blob, name);
    }

    async function ensurePermission(handle, mode = "readwrite") {
        if (!handle || typeof handle.queryPermission !== "function") return true;
        if (await handle.queryPermission({ mode }) === "granted") return true;
        return (await handle.requestPermission({ mode })) === "granted";
    }

    async function writeWorkbookToHandle(records, handle) {
        const allowed = await ensurePermission(handle, "readwrite");
        if (!allowed) throw new Error("Permission to write the Excel workbook was denied.");
        const workbook = await buildWorkbook(records);
        const buffer = await workbook.xlsx.writeBuffer();
        const writable = await handle.createWritable();
        await writable.write(buffer);
        await writable.close();
    }

    app.workbook = {
        fetchBundledWorkbook,
        parseWorkbookBuffer,
        downloadWorkbook,
        writeWorkbookToHandle
    };
})(window);
