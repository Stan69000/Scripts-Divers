const AUDIT_TIMEOUT_MS = 300000;
let LOGGER_DEBUG_MODE_CACHE = null;
let LOGGER_LOG_LEVEL_CACHE = null;

const LOGGER = {
  debug(action, data) {
    try {
      if (!shouldLogLevel_("DEBUG")) {
        return;
      }
      this._write("DEBUG", action, "OK", this._safeStringify(data || {}), 0, data || {});
    } catch (e) {
      console.error("LOGGER.debug failure", e);
    }
  },

  info(action, data) {
    try {
      if (!shouldLogLevel_("INFO")) {
        return;
      }
      this._write("INFO", action, "OK", this._safeStringify(data || {}), this._extractDuration(data), data || {});
    } catch (e) {
      console.error("LOGGER.info failure", e);
    }
  },

  error(action, message, stack) {
    try {
      this._write("ERROR", action, "KO", String(message || ""), 0, { stack: stack || "" });
    } catch (e) {
      console.error("LOGGER.error failure", e);
    }
  },

  _write(level, action, status, message, durationMs, meta) {
    try {
      const now = new Date();
      const user = Session.getActiveUser().getEmail() || "unknown";
      const trigger = detectTriggerSource_();
      const row = [
        now,
        level,
        action,
        status,
        message,
        APP_VERSION,
        durationMs || 0,
        user,
        trigger,
        this._safeStringify(meta || {})
      ];

      console.log(JSON.stringify(row));

      const logsSheet = getOrCreateSheet_(SHEET_NAMES.logs, true);
      if (logsSheet) {
        if (logsSheet.getLastRow() === 0) {
          logsSheet.appendRow([
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
          ]);
        }
        logsSheet.appendRow(row);
      }
    } catch (e) {
      console.error("LOGGER._write failure", e);
    }
  },

  _extractDuration(data) {
    try {
      if (!data || typeof data !== "object") {
        return 0;
      }
      return Number(data.durationMs || 0);
    } catch (e) {
      return 0;
    }
  },

  _safeStringify(data) {
    try {
      return JSON.stringify(data);
    } catch (e) {
      return String(data);
    }
  }
};

function getDebugEnabledWithoutLogger_() {
  try {
    if (LOGGER_DEBUG_MODE_CACHE !== null) {
      return LOGGER_DEBUG_MODE_CACHE;
    }
    const ss = getTargetSpreadsheet_(true);
    if (!ss) {
      LOGGER_DEBUG_MODE_CACHE = true;
      return LOGGER_DEBUG_MODE_CACHE;
    }
    const configSheet = ss.getSheetByName(SHEET_NAMES.config);
    if (!configSheet || configSheet.getLastRow() < 2) {
      LOGGER_DEBUG_MODE_CACHE = true;
      return LOGGER_DEBUG_MODE_CACHE;
    }
    const values = configSheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === "DEBUG") {
        LOGGER_DEBUG_MODE_CACHE = String(values[i][1]) === "true";
        return LOGGER_DEBUG_MODE_CACHE;
      }
    }
    LOGGER_DEBUG_MODE_CACHE = true;
    return LOGGER_DEBUG_MODE_CACHE;
  } catch (e) {
    LOGGER_DEBUG_MODE_CACHE = true;
    return LOGGER_DEBUG_MODE_CACHE;
  }
}

function shouldLogLevel_(requestedLevel) {
  try {
    const rank = {
      ERROR: 1,
      INFO: 2,
      DEBUG: 3
    };
    const current = getLogLevelWithoutLogger_();
    const currentRank = rank[current] || 1;
    const requestedRank = rank[String(requestedLevel || "ERROR")] || 1;
    return requestedRank <= currentRank;
  } catch (e) {
    return String(requestedLevel || "ERROR") === "ERROR";
  }
}

function getLogLevelWithoutLogger_() {
  try {
    if (LOGGER_LOG_LEVEL_CACHE !== null) {
      return LOGGER_LOG_LEVEL_CACHE;
    }

    const ss = getTargetSpreadsheet_(true);
    if (!ss) {
      LOGGER_LOG_LEVEL_CACHE = "INFO";
      return LOGGER_LOG_LEVEL_CACHE;
    }
    const configSheet = ss.getSheetByName(SHEET_NAMES.config);
    if (!configSheet || configSheet.getLastRow() < 2) {
      LOGGER_LOG_LEVEL_CACHE = "INFO";
      return LOGGER_LOG_LEVEL_CACHE;
    }

    const values = configSheet.getDataRange().getValues();
    let debugValue = "";
    for (let i = 1; i < values.length; i++) {
      const key = String(values[i][0] || "");
      if (key === "LOG_LEVEL") {
        const level = String(values[i][1] || "").trim().toUpperCase();
        if (level === "ERROR" || level === "INFO" || level === "DEBUG") {
          LOGGER_LOG_LEVEL_CACHE = level;
          return LOGGER_LOG_LEVEL_CACHE;
        }
      }
      if (key === "DEBUG") {
        debugValue = String(values[i][1] || "").trim().toLowerCase();
      }
    }

    if (debugValue === "true") {
      LOGGER_LOG_LEVEL_CACHE = "DEBUG";
      return LOGGER_LOG_LEVEL_CACHE;
    }

    LOGGER_LOG_LEVEL_CACHE = "INFO";
    return LOGGER_LOG_LEVEL_CACHE;
  } catch (e) {
    LOGGER_LOG_LEVEL_CACHE = "INFO";
    return LOGGER_LOG_LEVEL_CACHE;
  }
}

const TRIGGERS = {
  install() {
    LOGGER.debug("START", { action: "TRIGGERS.install" });
    const startMs = Date.now();
    try {
      let accepted = true;
      try {
        const ui = SpreadsheetApp.getUi();
        const response = ui.alert(
          "Installation des triggers",
          "Créer/mettre à jour les triggers (hebdo + audit + cleanup) ?",
          ui.ButtonSet.YES_NO
        );
        accepted = response === ui.Button.YES;
      } catch (uiError) {
        accepted = true;
      }

      if (!accepted) {
        LOGGER.info("SUCCESS", { action: "TRIGGERS.install", accepted: false, durationMs: Date.now() - startMs });
        return false;
      }

      this.removeManaged();

      const hour = Number(CONFIG.get("MONDAY_POST_HOUR", "9"));
      const minute = Number(CONFIG.get("MONDAY_POST_MINUTE", "5"));
      const timezone = String(CONFIG.get("TIMEZONE", "Europe/Paris"));
      ScriptApp.newTrigger("weeklyKickoffRun")
        .timeBased()
        .onWeekDay(ScriptApp.WeekDay.MONDAY)
        .atHour(hour)
        .nearMinute(minute)
        .inTimezone(timezone)
        .create();

      ScriptApp.newTrigger("auditMain")
        .timeBased()
        .everyDays(1)
        .atHour(6)
        .inTimezone(timezone)
        .create();

      ScriptApp.newTrigger("cleanupRun")
        .timeBased()
        .everyHours(1)
        .create();

      LOGGER.info("SUCCESS", { action: "TRIGGERS.install", accepted: true, durationMs: Date.now() - startMs });
      return true;
    } catch (e) {
      LOGGER.error("TRIGGERS.install", e.message, e.stack);
      return false;
    }
  },

  removeManaged() {
    LOGGER.debug("START", { action: "TRIGGERS.removeManaged" });
    const startMs = Date.now();
    try {
      const managedHandlers = {
        weeklyKickoffRun: true,
        auditMain: true,
        cleanupRun: true
      };
      const allTriggers = ScriptApp.getProjectTriggers();
      for (let i = 0; i < allTriggers.length; i++) {
        const fnName = allTriggers[i].getHandlerFunction();
        if (managedHandlers[fnName]) {
          ScriptApp.deleteTrigger(allTriggers[i]);
        }
      }
      LOGGER.info("SUCCESS", { action: "TRIGGERS.removeManaged", deleted: true, durationMs: Date.now() - startMs });
      return true;
    } catch (e) {
      LOGGER.error("TRIGGERS.removeManaged", e.message, e.stack);
      return false;
    }
  }
};

function postSlackMessage_(payload) {
  LOGGER.debug("START", { action: "postSlackMessage_", payload: payload });
  const startMs = Date.now();
  try {
    if (!payload || typeof payload !== "object") {
      LOGGER.error("postSlackMessage_", "Invalid payload", "");
      return null;
    }

    const webhook = resolveSlackWebhookUrl_();
    if (!webhook) {
      LOGGER.error("postSlackMessage_", "Missing SLACK_WEBHOOK_URL in Script Properties", "");
      return null;
    }

    const safePayload = sanitizeSlackPayload_(payload);
    const safeText = String(safePayload && safePayload.text ? safePayload.text : "");
    LOGGER.info("SUCCESS", {
      action: "postSlackMessage_.payloadCheck",
      containsAbsenceLines: /Si tu es en conges un lundi prochainement, pense a le noter|URL absences:/i.test(safeText),
      durationMs: 0
    });
    const options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(safePayload),
      muteHttpExceptions: true
    };
    const res = UrlFetchApp.fetch(webhook, options);
    const code = res.getResponseCode();
    LOGGER.info("SUCCESS", { action: "postSlackMessage_", code: code, durationMs: Date.now() - startMs });
    return code;
  } catch (e) {
    LOGGER.error("postSlackMessage_", e.message, e.stack);
    return null;
  }
}

function sanitizeSlackPayload_(payload) {
  LOGGER.debug("START", { action: "sanitizeSlackPayload_" });
  const startMs = Date.now();
  try {
    const out = Object.assign({}, payload || {});
    if (typeof out.text === "string") {
      out.text = sanitizeSlackText_(out.text);
    }
    LOGGER.info("SUCCESS", { action: "sanitizeSlackPayload_", hasText: typeof out.text === "string", durationMs: Date.now() - startMs });
    return out;
  } catch (e) {
    LOGGER.error("sanitizeSlackPayload_", e.message, e.stack);
    return payload;
  }
}

function sanitizeSlackText_(text) {
  LOGGER.debug("START", { action: "sanitizeSlackText_" });
  const startMs = Date.now();
  try {
    const input = String(text || "");
    const lines = input.split("\n");
    const filtered = [];
    for (let i = 0; i < lines.length; i++) {
      const line = String(lines[i] || "").trim();
      if (!line) {
        filtered.push(lines[i]);
        continue;
      }
      if (/^Si tu es en conges un lundi prochainement, pense a le noter/i.test(line)) {
        continue;
      }
      if (/^URL absences:/i.test(line)) {
        continue;
      }
      filtered.push(lines[i]);
    }
    const out = filtered.join("\n").replace(/\n{3,}/g, "\n\n").trim();
    LOGGER.info("SUCCESS", { action: "sanitizeSlackText_", durationMs: Date.now() - startMs });
    return out;
  } catch (e) {
    LOGGER.error("sanitizeSlackText_", e.message, e.stack);
    return String(text || "");
  }
}

function resolveSlackWebhookUrl_() {
  LOGGER.debug("START", { action: "resolveSlackWebhookUrl_" });
  const startMs = Date.now();
  try {
    const props = PropertiesService.getScriptProperties();
    const fromProps = String(props.getProperty("SLACK_WEBHOOK_URL") || "").trim();
    if (fromProps) {
      LOGGER.info("SUCCESS", { action: "resolveSlackWebhookUrl_", source: "script_properties", durationMs: Date.now() - startMs });
      return fromProps;
    }

    const legacy = String(CONFIG.get("SLACK_WEBHOOK_URL", "") || "").trim();
    if (legacy) {
      LOGGER.info("SUCCESS", { action: "resolveSlackWebhookUrl_", source: "config_legacy", durationMs: Date.now() - startMs });
      return legacy;
    }

    LOGGER.info("SUCCESS", { action: "resolveSlackWebhookUrl_", source: "none", durationMs: Date.now() - startMs });
    return "";
  } catch (e) {
    LOGGER.error("resolveSlackWebhookUrl_", e.message, e.stack);
    return "";
  }
}

function detectTriggerSource_() {
  try {
    return Session.getTemporaryActiveUserKey() ? "TRIGGER_OR_USER" : "SYSTEM";
  } catch (e) {
    return "UNKNOWN";
  }
}
