const APP_VERSION = "3.2";

const SHEET_NAMES = {
  config: "CONFIG",
  logs: "LOGS",
  rotationCare: "ROTATION_CARE",
  rotationSales: "ROTATION_SALES",
  rotationStart: "ROTATION_START",
  absences: "ABSENCES",
  history: "HISTORY"
};

const DEFAULT_CONFIG = {
  DEBUG: "false",
  LOG_LEVEL: "ERROR",
  TIMEZONE: "Europe/Paris",
  SLACK_POST_CHANNEL_ID: "",
  WEEKLY_ANNOUNCEMENT_CHANNEL_ID: "C08976543",
  LIEN_VISIO_WEEKLY: "",
  WEEKLY_MEET_URL: "",
  WEEKLY_SLIDES_FOLDER_ID: "",
  LIEN_REMOTE: "",
  JOUR_FERIES_FRANCAIS_API: "https://calendrier.api.gouv.fr/jours-feries/metropole/{year}.json",
  JOUR_OFF_INDY: "",
  MANAGERS_SLACK_GROUP_IDS_SALES: "",
  MANAGERS_SLACK_GROUP_IDS_CARE: "",
  MANAGERS_SLACK_USER_IDS: "",
  ROTATION_START_DATE: "2026-01-05",
  CARE_WEEK_OFFSET: "0",
  MONDAY_POST_HOUR: "09",
  MONDAY_POST_MINUTE: "05",
  ABSENCE_SOURCE: "GSHEET",
  ABSENCE_WEBAPP_URL: ""
};

const CONFIG = {
  get(key, defaultValue) {
    LOGGER.debug("START", { action: "CONFIG.get", key: key, defaultValue: defaultValue });
    const startMs = Date.now();
    try {
      if (!key || typeof key !== "string") {
        LOGGER.error("CONFIG.get", "Invalid key", "");
        return defaultValue;
      }

      const sheet = getOrCreateSheet_(SHEET_NAMES.config);
      const values = sheet.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][0]) === key) {
          const value = values[i][1];
          LOGGER.info("SUCCESS", {
            action: "CONFIG.get",
            key: key,
            found: true,
            durationMs: Date.now() - startMs
          });
          return value;
        }
      }

      LOGGER.info("SUCCESS", {
        action: "CONFIG.get",
        key: key,
        found: false,
        durationMs: Date.now() - startMs
      });
      return defaultValue;
    } catch (e) {
      LOGGER.error("CONFIG.get", e.message, e.stack);
      return defaultValue;
    }
  },

  set(key, value) {
    LOGGER.debug("START", { action: "CONFIG.set", key: key, value: value });
    const startMs = Date.now();
    try {
      if (!key || typeof key !== "string") {
        LOGGER.error("CONFIG.set", "Invalid key", "");
        return null;
      }

      const sheet = getOrCreateSheet_(SHEET_NAMES.config);
      const values = sheet.getDataRange().getValues();
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][0]) === key) {
          sheet.getRange(i + 1, 2).setValue(value);
          if (key === "DEBUG" || key === "LOG_LEVEL") {
            LOGGER_DEBUG_MODE_CACHE = null;
            LOGGER_LOG_LEVEL_CACHE = null;
          }
          LOGGER.info("SUCCESS", { action: "CONFIG.set", key: key, updated: true, durationMs: Date.now() - startMs });
          return value;
        }
      }

      sheet.appendRow([key, value, ""]);
      if (key === "DEBUG" || key === "LOG_LEVEL") {
        LOGGER_DEBUG_MODE_CACHE = null;
        LOGGER_LOG_LEVEL_CACHE = null;
      }
      LOGGER.info("SUCCESS", { action: "CONFIG.set", key: key, updated: false, durationMs: Date.now() - startMs });
      return value;
    } catch (e) {
      LOGGER.error("CONFIG.set", e.message, e.stack);
      return null;
    }
  },

  initDefaults() {
    LOGGER.debug("START", { action: "CONFIG.initDefaults" });
    const startMs = Date.now();
    try {
      const configSheet = getOrCreateSheet_(SHEET_NAMES.config);
      const headers = [["key", "value", "description"]];
      if (configSheet.getLastRow() === 0) {
        configSheet.getRange(1, 1, 1, 3).setValues(headers);
      }

      const existing = configSheet.getDataRange().getValues();
      const existingKeys = {};
      for (let i = 1; i < existing.length; i++) {
        existingKeys[String(existing[i][0])] = true;
      }

      const rowsToAppend = [];
      Object.keys(DEFAULT_CONFIG).forEach(function(key) {
        if (!existingKeys[key]) {
          rowsToAppend.push([key, DEFAULT_CONFIG[key], ""]);
        }
      });

      if (rowsToAppend.length > 0) {
        configSheet.getRange(configSheet.getLastRow() + 1, 1, rowsToAppend.length, 3).setValues(rowsToAppend);
      }

      removeDeprecatedConfigKeys_(configSheet, [
        "LUCCA_ENABLED",
        "LUCCA_BASE_URL",
        "LUCCA_API_TOKEN",
        "SLACK_WEBHOOK_URL",
        "SLACK_CHANNEL",
        "WEEKLY_ANNOUNCEMENT_CHANNEL",
        "MANAGERS_TAGS",
        "CARE_LEADS_TAGS",
        "SALES_LEADS_TAGS",
        "SALES_LEADS_SLACK_USER_IDS"
      ]);
      ensureBusinessSheets_();
      LOGGER.info("SUCCESS", { action: "CONFIG.initDefaults", inserted: rowsToAppend.length, durationMs: Date.now() - startMs });
      return true;
    } catch (e) {
      LOGGER.error("CONFIG.initDefaults", e.message, e.stack);
      return false;
    }
  }
};

function ensureBusinessSheets_() {
  LOGGER.debug("START", { action: "ensureBusinessSheets_" });
  const startMs = Date.now();
  try {
    const logsSheet = getOrCreateSheet_(SHEET_NAMES.logs);
    if (logsSheet.getLastRow() === 0) {
      logsSheet.getRange(1, 1, 1, 10).setValues([[
        "timestamp",
        "level",
        "action",
        "status",
        "message",
        "version",
        "durationMs",
        "user",
        "trigger",
        "meta"
      ]]);
    }

    const rotationCareSheet = getOrCreateSheet_(SHEET_NAMES.rotationCare);
    if (rotationCareSheet.getLastRow() === 0) {
      rotationCareSheet.getRange(1, 1, 1, 5).setValues([["order", "name", "slackUserId", "email", "active"]]);
      rotationCareSheet.getRange(2, 1, 3, 5).setValues([
        [1, "Lead Care 1", "UAAAAAAA1", "care.lead.1@company.com", "true"],
        [2, "Lead Care 2", "UAAAAAAA2", "care.lead.2@company.com", "true"],
        [3, "Lead Care 3", "UAAAAAAA3", "care.lead.3@company.com", "true"]
      ]);
    }

    const rotationSalesSheet = getOrCreateSheet_(SHEET_NAMES.rotationSales);
    if (rotationSalesSheet.getLastRow() === 0) {
      rotationSalesSheet.getRange(1, 1, 1, 5).setValues([["order", "name", "slackUserId", "email", "active"]]);
      rotationSalesSheet.getRange(2, 1, 3, 5).setValues([
        [1, "Lead Sales 1", "UBBBBBBB1", "sales.lead.1@company.com", "true"],
        [2, "Lead Sales 2", "UBBBBBBB2", "sales.lead.2@company.com", "true"],
        [3, "Lead Sales 3", "UBBBBBBB3", "sales.lead.3@company.com", "true"]
      ]);
    }

    const rotationStartSheet = getOrCreateSheet_(SHEET_NAMES.rotationStart);
    if (rotationStartSheet.getLastRow() === 0) {
      rotationStartSheet.getRange(1, 1, 1, 5).setValues([["order", "name", "slackUserId", "email", "active"]]);
      rotationStartSheet.getRange(2, 1, 2, 5).setValues([
        [1, "Starter 1", "UCCCCCCC1", "starter.1@company.com", "true"],
        [2, "Starter 2", "UCCCCCCC2", "starter.2@company.com", "true"]
      ]);
    }

    migrateLegacyRotationCare_();

    const absencesSheet = getOrCreateSheet_(SHEET_NAMES.absences);
    if (absencesSheet.getLastRow() === 0) {
      absencesSheet.getRange(1, 1, 1, 6).setValues([["email", "name", "startDate", "endDate", "source", "notes"]]);
    }

    const historySheet = getOrCreateSheet_(SHEET_NAMES.history);
    if (historySheet.getLastRow() === 0) {
      historySheet.getRange(1, 1, 1, 13).setValues([[
        "runAt",
        "mondayDate",
        "team",
        "primaryName",
        "primarySlackUserId",
        "backupName",
        "backupSlackUserId",
        "starterName",
        "starterSlackUserId",
        "status",
        "reason",
        "postedBy",
        "rawPayload"
      ]]);
    }

    LOGGER.info("SUCCESS", { action: "ensureBusinessSheets_", durationMs: Date.now() - startMs });
    return true;
  } catch (e) {
    LOGGER.error("ensureBusinessSheets_", e.message, e.stack);
    return false;
  }
}

function migrateLegacyRotationCare_() {
  LOGGER.debug("START", { action: "migrateLegacyRotationCare_" });
  const startMs = Date.now();
  try {
    const ss = getTargetSpreadsheet_(true);
    const legacySheet = ss ? ss.getSheetByName("ROTATION") : null;
    const careSheet = getOrCreateSheet_(SHEET_NAMES.rotationCare);
    if (!legacySheet || !careSheet) {
      LOGGER.info("SUCCESS", { action: "migrateLegacyRotationCare_", migrated: false, reason: "missing_sheet", durationMs: Date.now() - startMs });
      return false;
    }
    const careValues = careSheet.getDataRange().getValues();
    const careHasOnlySeed =
      careValues.length === 4 &&
      String(careValues[0][0]) === "order" &&
      String(careValues[1][1]).indexOf("Lead Care") === 0;
    if (careSheet.getLastRow() > 1 && !careHasOnlySeed) {
      LOGGER.info("SUCCESS", { action: "migrateLegacyRotationCare_", migrated: false, reason: "target_not_empty", durationMs: Date.now() - startMs });
      return false;
    }

    const values = legacySheet.getDataRange().getValues();
    if (values.length <= 1) {
      LOGGER.info("SUCCESS", { action: "migrateLegacyRotationCare_", migrated: false, reason: "legacy_empty", durationMs: Date.now() - startMs });
      return false;
    }

    careSheet.clearContents();
    careSheet.getRange(1, 1, values.length, values[0].length).setValues(values);
    LOGGER.info("SUCCESS", { action: "migrateLegacyRotationCare_", migrated: true, rows: values.length - 1, durationMs: Date.now() - startMs });
    return true;
  } catch (e) {
    LOGGER.error("migrateLegacyRotationCare_", e.message, e.stack);
    return false;
  }
}

function removeDeprecatedConfigKeys_(configSheet, keysToRemove) {
  LOGGER.debug("START", { action: "removeDeprecatedConfigKeys_", keysToRemove: keysToRemove });
  const startMs = Date.now();
  try {
    if (!configSheet || !Array.isArray(keysToRemove) || keysToRemove.length === 0) {
      LOGGER.info("SUCCESS", { action: "removeDeprecatedConfigKeys_", removed: 0, durationMs: Date.now() - startMs });
      return 0;
    }

    const values = configSheet.getDataRange().getValues();
    const rowsToDelete = [];
    for (let i = 1; i < values.length; i++) {
      const key = String(values[i][0] || "");
      if (keysToRemove.indexOf(key) !== -1) {
        rowsToDelete.push(i + 1);
      }
    }

    for (let j = rowsToDelete.length - 1; j >= 0; j--) {
      configSheet.deleteRow(rowsToDelete[j]);
    }

    LOGGER.info("SUCCESS", { action: "removeDeprecatedConfigKeys_", removed: rowsToDelete.length, durationMs: Date.now() - startMs });
    return rowsToDelete.length;
  } catch (e) {
    LOGGER.error("removeDeprecatedConfigKeys_", e.message, e.stack);
    return 0;
  }
}

function getOrCreateSheet_(sheetName, silent) {
  if (!silent) {
    LOGGER.debug("START", { action: "getOrCreateSheet_", sheetName: sheetName });
  }
  const startMs = Date.now();
  try {
    const ss = getTargetSpreadsheet_(silent);
    if (!ss) {
      if (!silent) {
        LOGGER.error("getOrCreateSheet_", "No target spreadsheet found. Set Script Property SPREADSHEET_ID.", "");
      }
      return null;
    }

    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    if (!silent) {
      LOGGER.info("SUCCESS", { action: "getOrCreateSheet_", sheetName: sheetName, durationMs: Date.now() - startMs });
    }
    return sheet;
  } catch (e) {
    if (!silent) {
      LOGGER.error("getOrCreateSheet_", e.message, e.stack);
    }
    return null;
  }
}

function getTargetSpreadsheet_(silent) {
  if (!silent) {
    LOGGER.debug("START", { action: "getTargetSpreadsheet_" });
  }
  const startMs = Date.now();
  try {
    const props = PropertiesService.getScriptProperties();
    const spreadsheetId = String(props.getProperty("SPREADSHEET_ID") || "").trim();
    if (spreadsheetId) {
      const ssById = SpreadsheetApp.openById(spreadsheetId);
      if (!silent) {
        LOGGER.info("SUCCESS", { action: "getTargetSpreadsheet_", mode: "by_id", durationMs: Date.now() - startMs });
      }
      return ssById;
    }

    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) {
      if (!silent) {
        LOGGER.info("SUCCESS", { action: "getTargetSpreadsheet_", mode: "active", durationMs: Date.now() - startMs });
      }
      return active;
    }

    if (!silent) {
      LOGGER.info("SUCCESS", { action: "getTargetSpreadsheet_", mode: "none", durationMs: Date.now() - startMs });
    }
    return null;
  } catch (e) {
    if (!silent) {
      LOGGER.error("getTargetSpreadsheet_", e.message, e.stack);
    }
    return null;
  }
}
