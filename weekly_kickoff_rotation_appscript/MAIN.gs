function install() {
  LOGGER.debug("START", { action: "install" });
  const startMs = Date.now();
  try {
    const ok = CONFIG.initDefaults();
    if (!ok) {
      LOGGER.error("install", "CONFIG.initDefaults failed", "");
      return false;
    }

    const triggerOk = TRIGGERS.install();
    LOGGER.info("INSTALL_OK", { action: "install", triggerOk: triggerOk, durationMs: Date.now() - startMs });
    return true;
  } catch (e) {
    LOGGER.error("install", e.message, e.stack);
    return false;
  }
}

function onOpen() {
  LOGGER.debug("START", { action: "onOpen" });
  const startMs = Date.now();
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("IT-INDY")
      .addItem("Install", "install")
      .addItem("Reinstall Triggers", "installTriggersOnly")
      .addItem("Dry Run Weekly", "dryRunWeeklyKickoff")
      .addItem("Run Weekly Now", "weeklyKickoffRun")
      .addItem("Sync Lucca IDs", "syncLuccaUserIdsInRotations")
      .addItem("Debug Lucca Juliette", "debugLuccaAbsenceJuliette")
      .addItem("Audit Now", "auditMain")
      .addItem("Cleanup Logs", "cleanupRun")
      .addToUi();
    LOGGER.info("SUCCESS", { action: "onOpen", durationMs: Date.now() - startMs });
  } catch (e) {
    LOGGER.error("onOpen", e.message, e.stack);
  }
}

function installTriggersOnly() {
  LOGGER.debug("START", { action: "installTriggersOnly" });
  const startMs = Date.now();
  try {
    const ok = TRIGGERS.install();
    LOGGER.info("SUCCESS", { action: "installTriggersOnly", ok: ok, durationMs: Date.now() - startMs });
    return ok;
  } catch (e) {
    LOGGER.error("installTriggersOnly", e.message, e.stack);
    return false;
  }
}

function setSpreadsheetId(spreadsheetId) {
  LOGGER.debug("START", { action: "setSpreadsheetId", spreadsheetId: spreadsheetId });
  const startMs = Date.now();
  try {
    if (!spreadsheetId || typeof spreadsheetId !== "string") {
      LOGGER.error("setSpreadsheetId", "Invalid spreadsheetId", "");
      return false;
    }
    PropertiesService.getScriptProperties().setProperty("SPREADSHEET_ID", spreadsheetId.trim());
    LOGGER.info("SUCCESS", { action: "setSpreadsheetId", durationMs: Date.now() - startMs });
    return true;
  } catch (e) {
    LOGGER.error("setSpreadsheetId", e.message, e.stack);
    return false;
  }
}

function setSlackWebhookUrl(webhookUrl) {
  LOGGER.debug("START", { action: "setSlackWebhookUrl", webhookUrl: webhookUrl ? "provided" : "" });
  const startMs = Date.now();
  try {
    const value = String(webhookUrl || "").trim();
    if (!value) {
      LOGGER.error("setSlackWebhookUrl", "Invalid webhookUrl", "");
      return false;
    }
    PropertiesService.getScriptProperties().setProperty("SLACK_WEBHOOK_URL", value);
    LOGGER.info("SUCCESS", { action: "setSlackWebhookUrl", durationMs: Date.now() - startMs });
    return true;
  } catch (e) {
    LOGGER.error("setSlackWebhookUrl", e.message, e.stack);
    return false;
  }
}

function clearSlackWebhookUrl() {
  LOGGER.debug("START", { action: "clearSlackWebhookUrl" });
  const startMs = Date.now();
  try {
    PropertiesService.getScriptProperties().deleteProperty("SLACK_WEBHOOK_URL");
    LOGGER.info("SUCCESS", { action: "clearSlackWebhookUrl", durationMs: Date.now() - startMs });
    return true;
  } catch (e) {
    LOGGER.error("clearSlackWebhookUrl", e.message, e.stack);
    return false;
  }
}
