(function (global) {
    const app = global.AircraftMovementApp = global.AircraftMovementApp || {};

    app.constants = Object.freeze({
        DB_NAME: "aircraft-movement-register",
        DB_VERSION: 2,
        RECORD_STORE: "records",
        META_STORE: "meta",
        DEFAULT_ENTRY_ROWS: 25,
        ENTRY_ROW_INCREMENT: 25,
        AUTO_SAVE_DELAY_MS: 500,
        WORKBOOK_SYNC_DELAY_MS: 900,
        DEFAULT_EXPORT_NAME: "aircraft_movement.xlsx",
        WORKBOOK_CANDIDATES: [
            "./aircraft_movement.xlsx",
            "./Aircraft_movement.xlsx",
            "./Aircraft_movement_updated.xlsx",
            "./aircraft_movement.xlsm",
            "./Aircraft_movement_updated.xlsm"
        ],
        WATCH_OPTIONS: ["Red", "Blue", "Green", "White"],
        AIRCRAFT_OPTIONS: [
            "B737", "A320", "B777", "A330", "B787", "A380",
            "A319", "A321", "E190", "E175", "CRJ900", "ATR72",
            "B757", "B767", "MD80", "C130", "G550", "DH8D"
        ],
        ENTRY_FIELD_NAMES: [
            "dateKey",
            "watch",
            "aircraftType",
            "regNo",
            "arrivalTime",
            "departureTime",
            "soulsLanding",
            "soulsTakeoff",
            "crewOnBoard",
            "operatingCompany",
            "destination"
        ],
        EXPORT_HEADERS: [
            "S/N",
            "DATE",
            "WATCH ON DUTY",
            "AIRCRAFT TYPE",
            "REG. NO",
            "TIME OF ARRIVAL",
            "TIME OF DEPARTURE",
            "SOULS ON BOARD LANDING",
            "SOULS ON BOARD TAKE OFF",
            "CREW ON BOARD",
            "OPERATING COMPANY",
            "DESTINATION"
        ]
    });
})(window);
