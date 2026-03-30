(function (global) {
    const app = global.AircraftMovementApp = global.AircraftMovementApp || {};
    const { DB_NAME, DB_VERSION, RECORD_STORE, META_STORE } = app.constants;
    const { createLocalId } = app.utils;

    let db = null;

    function requestToPromise(request) {
        return new Promise((resolve, reject) => {
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
        });
    }

    function transactionDone(tx) {
        return new Promise((resolve, reject) => {
            tx.oncomplete = () => resolve();
            tx.onerror = () => reject(tx.error);
            tx.onabort = () => reject(tx.error);
        });
    }

    async function openDb() {
        if (db) return db;
        db = await new Promise((resolve, reject) => {
            const request = indexedDB.open(DB_NAME, DB_VERSION);
            request.onupgradeneeded = () => {
                const database = request.result;
                let recordStore = null;
                if (!database.objectStoreNames.contains(RECORD_STORE)) {
                    recordStore = database.createObjectStore(RECORD_STORE, { keyPath: "id" });
                } else {
                    recordStore = request.transaction.objectStore(RECORD_STORE);
                }
                if (!recordStore.indexNames.contains("dateKey")) {
                    recordStore.createIndex("dateKey", "dateKey", { unique: false });
                }
                if (!recordStore.indexNames.contains("dateSerial")) {
                    recordStore.createIndex("dateSerial", ["dateKey", "serial"], { unique: false });
                }
                if (!database.objectStoreNames.contains(META_STORE)) {
                    database.createObjectStore(META_STORE, { keyPath: "key" });
                }
            };
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
        });
        return db;
    }

    async function getMeta(key) {
        const database = await openDb();
        const tx = database.transaction(META_STORE, "readonly");
        const result = await requestToPromise(tx.objectStore(META_STORE).get(key));
        await transactionDone(tx);
        return result ? result.value : null;
    }

    async function setMeta(key, value) {
        const database = await openDb();
        const tx = database.transaction(META_STORE, "readwrite");
        tx.objectStore(META_STORE).put({ key, value });
        await transactionDone(tx);
    }

    async function getRecord(id) {
        const database = await openDb();
        const tx = database.transaction(RECORD_STORE, "readonly");
        const result = await requestToPromise(tx.objectStore(RECORD_STORE).get(id));
        await transactionDone(tx);
        return result || null;
    }

    async function getAllRecords() {
        const database = await openDb();
        const tx = database.transaction(RECORD_STORE, "readonly");
        const rows = await requestToPromise(tx.objectStore(RECORD_STORE).getAll());
        await transactionDone(tx);
        return rows || [];
    }

    async function putRecord(record) {
        const database = await openDb();
        const tx = database.transaction(RECORD_STORE, "readwrite");
        tx.objectStore(RECORD_STORE).put(record);
        await transactionDone(tx);
    }

    async function deleteRecord(id) {
        const database = await openDb();
        const tx = database.transaction(RECORD_STORE, "readwrite");
        tx.objectStore(RECORD_STORE).delete(id);
        await transactionDone(tx);
    }

    async function getRecordsByDate(dateKey) {
        const database = await openDb();
        const tx = database.transaction(RECORD_STORE, "readonly");
        const rows = await requestToPromise(tx.objectStore(RECORD_STORE).index("dateKey").getAll(IDBKeyRange.only(dateKey)));
        await transactionDone(tx);
        return rows || [];
    }

    async function renumberDate(dateKey) {
        if (!dateKey) return;
        const rows = await getRecordsByDate(dateKey);
        rows.sort((left, right) => {
            const leftOrder = Number(left.orderKey || 0);
            const rightOrder = Number(right.orderKey || 0);
            if (leftOrder !== rightOrder) return leftOrder - rightOrder;
            return String(left.createdAt || "").localeCompare(String(right.createdAt || ""));
        });

        const database = await openDb();
        const tx = database.transaction(RECORD_STORE, "readwrite");
        const store = tx.objectStore(RECORD_STORE);
        rows.forEach((record, index) => {
            store.put({ ...record, serial: index + 1 });
        });
        await transactionDone(tx);
    }

    async function upsertMovement(payload, recordId = null) {
        const now = new Date().toISOString();
        const existing = recordId ? await getRecord(recordId) : null;

        if (existing) {
            const nextRecord = {
                ...existing,
                ...payload,
                updatedAt: now,
                orderKey: existing.dateKey === payload.dateKey ? existing.orderKey : (Date.now() + Math.random())
            };
            await putRecord(nextRecord);
            await renumberDate(existing.dateKey);
            if (existing.dateKey !== payload.dateKey) {
                await renumberDate(payload.dateKey);
            }
            return getRecord(nextRecord.id);
        }

        const nextRecord = {
            id: createLocalId("record"),
            ...payload,
            serial: 0,
            orderKey: Date.now() + Math.random(),
            createdAt: now,
            updatedAt: now
        };
        await putRecord(nextRecord);
        await renumberDate(payload.dateKey);
        return getRecord(nextRecord.id);
    }

    async function deleteMovement(recordId) {
        const existing = await getRecord(recordId);
        if (!existing) return;
        await deleteRecord(recordId);
        await renumberDate(existing.dateKey);
    }

    async function replaceAllRecords(records) {
        const database = await openDb();
        const tx = database.transaction(RECORD_STORE, "readwrite");
        const store = tx.objectStore(RECORD_STORE);
        store.clear();
        records.forEach((record, index) => {
            store.put({
                id: record.id || createLocalId("record"),
                serial: record.serial || 0,
                orderKey: Number(record.orderKey || 0) || index + 1,
                createdAt: record.createdAt || new Date().toISOString(),
                updatedAt: record.updatedAt || new Date().toISOString(),
                ...record
            });
        });
        await transactionDone(tx);
    }

    app.store = {
        openDb,
        getMeta,
        setMeta,
        getRecord,
        getAllRecords,
        upsertMovement,
        deleteMovement,
        replaceAllRecords
    };
})(window);
