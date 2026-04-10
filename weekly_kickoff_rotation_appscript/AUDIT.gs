function auditMain() {
  LOGGER.debug("START", { action: "auditMain" });
  const startMs = Date.now();
  try {
    const checks = {
      configReady: checkConfigReady_(),
      rotationReady: checkRotationReady_(),
      historyReady: checkHistoryReady_()
    };
    LOGGER.info("SUCCESS", { action: "auditMain", checks: checks, durationMs: Date.now() - startMs });
    return checks;
  } catch (e) {
    LOGGER.error("auditMain", e.message, e.stack);
    return {};
  }
}

function cleanupRun() {
  LOGGER.debug("START", { action: "cleanupRun" });
  const startMs = Date.now();
  try {
    const logsSheet = getOrCreateSheet_(SHEET_NAMES.logs);
    const values = logsSheet.getDataRange().getValues();
    if (values.length <= 1) {
      LOGGER.info("SUCCESS", { action: "cleanupRun", deleted: 0, durationMs: Date.now() - startMs });
      return 0;
    }

    const now = new Date();
    const keepRows = [values[0]];
    let deleted = 0;

    for (let i = 1; i < values.length; i++) {
      const ts = new Date(values[i][0]);
      const ageMs = now.getTime() - ts.getTime();
      const olderThan30Days = ageMs > 30 * 24 * 60 * 60 * 1000;
      if (isNaN(ts.getTime()) || olderThan30Days) {
        deleted++;
      } else {
        keepRows.push(values[i]);
      }
    }

    logsSheet.clearContents();
    logsSheet.getRange(1, 1, keepRows.length, keepRows[0].length).setValues(keepRows);
    LOGGER.info("SUCCESS", { action: "cleanupRun", deleted: deleted, durationMs: Date.now() - startMs });
    return deleted;
  } catch (e) {
    LOGGER.error("cleanupRun", e.message, e.stack);
    return 0;
  }
}

function checkConfigReady_() {
  LOGGER.debug("START", { action: "checkConfigReady_" });
  const startMs = Date.now();
  try {
    const requiredKeys = ["ROTATION_START_DATE"];
    const missing = requiredKeys.filter(function(key) {
      return String(CONFIG.get(key, "")) === "";
    });
    const webhookOk = resolveSlackWebhookUrl_() !== "";
    if (!webhookOk) {
      missing.push("SLACK_WEBHOOK_URL(script_properties)");
    }
    const ok = missing.length === 0;
    LOGGER.info("SUCCESS", { action: "checkConfigReady_", ok: ok, missing: missing, webhookOk: webhookOk, durationMs: Date.now() - startMs });
    return { ok: ok, missing: missing, webhookOk: webhookOk };
  } catch (e) {
    LOGGER.error("checkConfigReady_", e.message, e.stack);
    return { ok: false, missing: ["ERROR"] };
  }
}

function checkRotationReady_() {
  LOGGER.debug("START", { action: "checkRotationReady_" });
  const startMs = Date.now();
  try {
    const careRows = getCareRotation_();
    const salesRows = getSalesRotation_();
    const startRows = getStartRotation_();
    const ok = careRows.length >= 2 && salesRows.length >= 2 && startRows.length >= 1;
    LOGGER.info("SUCCESS", {
      action: "checkRotationReady_",
      ok: ok,
      careCount: careRows.length,
      salesCount: salesRows.length,
      startCount: startRows.length,
      durationMs: Date.now() - startMs
    });
    return { ok: ok, careCount: careRows.length, salesCount: salesRows.length, startCount: startRows.length };
  } catch (e) {
    LOGGER.error("checkRotationReady_", e.message, e.stack);
    return { ok: false, careCount: 0, salesCount: 0, startCount: 0 };
  }
}

function checkHistoryReady_() {
  LOGGER.debug("START", { action: "checkHistoryReady_" });
  const startMs = Date.now();
  try {
    const sheet = getOrCreateSheet_(SHEET_NAMES.history);
    const hasHeaders = sheet.getLastRow() >= 1;
    LOGGER.info("SUCCESS", { action: "checkHistoryReady_", ok: hasHeaders, durationMs: Date.now() - startMs });
    return { ok: hasHeaders };
  } catch (e) {
    LOGGER.error("checkHistoryReady_", e.message, e.stack);
    return { ok: false };
  }
}
