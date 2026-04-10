function weeklyKickoffRun(forceRun) {
  LOGGER.debug("START", { action: "weeklyKickoffRun", forceRun: forceRun });
  const startMs = Date.now();
  try {
    const runDate = new Date();
    const runGate = validateWeeklyRunWindow_(runDate, forceRun);
    const mondayDate = runGate.mondayDate;
    const team = getTeamForWeek_(mondayDate);
    if (!runGate.allowed) {
      saveHistory_(mondayDate, team, null, null, "BLOCKED", runGate.runWindowReason, runGate.runMode, runGate.runWindowReason);
      LOGGER.info("SUCCESS", {
        action: "weeklyKickoffRun",
        status: "BLOCKED",
        runMode: runGate.runMode,
        reason: runGate.runWindowReason,
        durationMs: Date.now() - startMs
      });
      return { status: "BLOCKED", reason: runGate.runWindowReason };
    }

    const skipRule = getEffectiveSkipRuleForRun_(mondayDate, runGate);
    if (skipRule.skip) {
      saveHistory_(mondayDate, team, null, null, "SKIPPED", skipRule.reason, runGate.runMode, runGate.runWindowReason);
      LOGGER.info("SUCCESS", {
        action: "weeklyKickoffRun",
        status: "SKIPPED",
        runMode: runGate.runMode,
        reason: skipRule.reason,
        durationMs: Date.now() - startMs
      });
      return { status: "SKIPPED", reason: skipRule.reason };
    }

    const nomination = selectLeadForTeam_(team, mondayDate);
    const starter = selectStarterForWeek_(mondayDate);
    if (!nomination || !nomination.primary) {
      LOGGER.error("weeklyKickoffRun", "No nomination available", "");
      saveHistory_(mondayDate, team, nomination, starter, "ERROR", "No nomination", runGate.runMode, runGate.runWindowReason);
      return null;
    }

    const messageText = buildSlackMessage_(team, mondayDate, nomination, starter);
    const postChannelId = String(CONFIG.get("SLACK_POST_CHANNEL_ID", ""));
    const payload = {
      channel: postChannelId || undefined,
      text: messageText
    };
    const code = postSlackMessage_(payload);
    const status = code === 200 ? "POSTED" : "FAILED_POST";
    let reason = "Slack code=" + code;
    if (runGate.runMode === "holiday_reschedule_tuesday") {
      reason += " (reschedule mardi)";
    }
    if (runGate.runMode === "force") {
      reason += " (force)";
    }
    saveHistory_(mondayDate, team, nomination, starter, status, reason, runGate.runMode, runGate.runWindowReason);

    LOGGER.info("SUCCESS", {
      action: "weeklyKickoffRun",
      status: status,
      runMode: runGate.runMode,
      durationMs: Date.now() - startMs
    });
    return nomination;
  } catch (e) {
    LOGGER.error("weeklyKickoffRun", e.message, e.stack);
    return null;
  }
}

function dryRunWeeklyKickoff() {
  LOGGER.debug("START", { action: "dryRunWeeklyKickoff" });
  const startMs = Date.now();
  try {
    const runDate = new Date();
    const runGate = validateWeeklyRunWindow_(runDate, false);
    const mondayDate = runGate.mondayDate;
    const team = getTeamForWeek_(mondayDate);
    if (!runGate.allowed) {
      const blockedMessage = "Dry run bloque: " + runGate.runWindowReason + ".";
      console.log("DRY_RUN_MESSAGE:\n" + blockedMessage);
      LOGGER.info("SUCCESS", {
        action: "dryRunWeeklyKickoff",
        skipped: true,
        runMode: runGate.runMode,
        reason: runGate.runWindowReason,
        durationMs: Date.now() - startMs
      });
      return blockedMessage;
    }

    const skipRule = getEffectiveSkipRuleForRun_(mondayDate, runGate);
    if (skipRule.skip) {
      const skipMessage = "Dry run: aucun message Slack, weekly saute (" + skipRule.reason + ").";
      console.log("DRY_RUN_MESSAGE:\n" + skipMessage);
      LOGGER.info("SUCCESS", {
        action: "dryRunWeeklyKickoff",
        skipped: true,
        runMode: runGate.runMode,
        reason: skipRule.reason,
        durationMs: Date.now() - startMs
      });
      return skipMessage;
    }

    const nomination = selectLeadForTeam_(team, mondayDate);
    const starter = selectStarterForWeek_(mondayDate);
    const messageText = buildSlackMessage_(team, mondayDate, nomination, starter);
    try {
      SpreadsheetApp.getUi().alert("Dry run", messageText, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (uiError) {
      LOGGER.info("SUCCESS", { action: "dryRunWeeklyKickoff.uiFallback", reason: "no_ui_context" });
      console.log("DRY_RUN_MESSAGE:\n" + messageText);
    }
    LOGGER.info("SUCCESS", { action: "dryRunWeeklyKickoff", durationMs: Date.now() - startMs });
    return messageText;
  } catch (e) {
    LOGGER.error("dryRunWeeklyKickoff", e.message, e.stack);
    return null;
  }
}

function validateWeeklyRunWindow_(runDate, forceRun) {
  LOGGER.debug("START", { action: "validateWeeklyRunWindow_", runDate: runDate, forceRun: forceRun });
  const startMs = Date.now();
  try {
    const tz = String(CONFIG.get("TIMEZONE", "Europe/Paris"));
    const currentDate = runDate instanceof Date && !isNaN(runDate.getTime()) ? new Date(runDate.getTime()) : new Date();
    const currentTzDate = new Date(Utilities.formatDate(currentDate, tz, "yyyy-MM-dd'T'HH:mm:ss"));
    const mondayDate = getMondayForDate_(currentDate);
    const holiday = getFrenchHolidayNameForDate_(mondayDate);
    const targetHour = Number(CONFIG.get("MONDAY_POST_HOUR", "9"));
    const targetMinute = Number(CONFIG.get("MONDAY_POST_MINUTE", "5"));
    const targetDay = holiday ? 2 : 1;
    const targetDate = new Date(mondayDate.getTime());
    if (targetDay === 2) {
      targetDate.setDate(targetDate.getDate() + 1);
    }
    targetDate.setHours(targetHour, targetMinute, 0, 0);

    const windowStart = new Date(targetDate.getTime() - (45 * 60 * 1000));
    const windowEnd = new Date(targetDate.getTime() + (180 * 60 * 1000));
    const targetIso = Utilities.formatDate(targetDate, tz, "yyyy-MM-dd HH:mm");

    if (forceRun === true) {
      const forceResult = {
        allowed: true,
        mondayDate: mondayDate,
        runMode: "force",
        runWindowReason: "forceRun=true",
        holiday: holiday
      };
      LOGGER.info("SUCCESS", { action: "validateWeeklyRunWindow_", allowed: true, runMode: "force", durationMs: Date.now() - startMs });
      return forceResult;
    }

    if (currentTzDate.getDay() !== targetDay) {
      const wrongDayResult = {
        allowed: false,
        mondayDate: mondayDate,
        runMode: "blocked",
        runWindowReason: "outside_allowed_day target=" + targetIso + (holiday ? " holiday_reschedule" : ""),
        holiday: holiday
      };
      LOGGER.info("SUCCESS", { action: "validateWeeklyRunWindow_", allowed: false, runMode: "blocked", durationMs: Date.now() - startMs });
      return wrongDayResult;
    }

    if (currentTzDate.getTime() < windowStart.getTime() || currentTzDate.getTime() > windowEnd.getTime()) {
      const wrongWindowResult = {
        allowed: false,
        mondayDate: mondayDate,
        runMode: "blocked",
        runWindowReason: "outside_allowed_time_window target=" + targetIso,
        holiday: holiday
      };
      LOGGER.info("SUCCESS", { action: "validateWeeklyRunWindow_", allowed: false, runMode: "blocked", durationMs: Date.now() - startMs });
      return wrongWindowResult;
    }

    const runMode = holiday ? "holiday_reschedule_tuesday" : "scheduled_monday";
    const successResult = {
      allowed: true,
      mondayDate: mondayDate,
      runMode: runMode,
      runWindowReason: holiday ? "monday_holiday_rescheduled_to_tuesday" : "standard_monday_window",
      holiday: holiday
    };
    LOGGER.info("SUCCESS", { action: "validateWeeklyRunWindow_", allowed: true, runMode: runMode, durationMs: Date.now() - startMs });
    return successResult;
  } catch (e) {
    LOGGER.error("validateWeeklyRunWindow_", e.message, e.stack);
    return {
      allowed: false,
      mondayDate: getCurrentMonday_(),
      runMode: "blocked",
      runWindowReason: "validate_run_window_error",
      holiday: ""
    };
  }
}

function getEffectiveSkipRuleForRun_(mondayDate, runGate) {
  LOGGER.debug("START", { action: "getEffectiveSkipRuleForRun_", mondayDate: mondayDate, runGate: runGate });
  const startMs = Date.now();
  try {
    if (runGate && runGate.runMode === "force") {
      LOGGER.info("SUCCESS", { action: "getEffectiveSkipRuleForRun_", skip: false, durationMs: Date.now() - startMs });
      return { skip: false, reason: "" };
    }

    const tz = String(CONFIG.get("TIMEZONE", "Europe/Paris"));
    const isoDate = Utilities.formatDate(mondayDate, tz, "yyyy-MM-dd");
    const indyOffDates = parseDateCsv_(String(CONFIG.get("JOUR_OFF_INDY", "")));
    if (indyOffDates.indexOf(isoDate) !== -1) {
      const reasonOff = "JOUR_OFF_INDY " + isoDate;
      LOGGER.info("SUCCESS", { action: "getEffectiveSkipRuleForRun_", skip: true, reason: reasonOff, durationMs: Date.now() - startMs });
      return { skip: true, reason: reasonOff };
    }

    const holiday = getFrenchHolidayNameForDate_(mondayDate);
    if (holiday && (!runGate || runGate.runMode !== "holiday_reschedule_tuesday")) {
      const reasonHoliday = "JOUR_FERIE " + isoDate + " (" + holiday + ")";
      LOGGER.info("SUCCESS", { action: "getEffectiveSkipRuleForRun_", skip: true, reason: reasonHoliday, durationMs: Date.now() - startMs });
      return { skip: true, reason: reasonHoliday };
    }

    LOGGER.info("SUCCESS", { action: "getEffectiveSkipRuleForRun_", skip: false, durationMs: Date.now() - startMs });
    return { skip: false, reason: "" };
  } catch (e) {
    LOGGER.error("getEffectiveSkipRuleForRun_", e.message, e.stack);
    return { skip: false, reason: "" };
  }
}

function getTeamForWeek_(mondayDate) {
  LOGGER.debug("START", { action: "getTeamForWeek_", mondayDate: mondayDate });
  const startMs = Date.now();
  try {
    if (!(mondayDate instanceof Date)) {
      LOGGER.error("getTeamForWeek_", "Invalid mondayDate", "");
      return "care";
    }

    const anchorDateRaw = CONFIG.get("ROTATION_START_DATE", "2026-01-05");
    const careOffset = Number(CONFIG.get("CARE_WEEK_OFFSET", "0"));
    const anchorDate = parseConfigDate_(anchorDateRaw);
    if (!anchorDate) {
      LOGGER.error("getTeamForWeek_", "Invalid ROTATION_START_DATE", "");
      return "care";
    }
    const diffMs = mondayDate.getTime() - anchorDate.getTime();
    const weekIndex = Math.floor(diffMs / (7 * 24 * 60 * 60 * 1000));
    const isCare = (weekIndex + careOffset) % 2 === 0;
    const team = isCare ? "care" : "sales";

    LOGGER.info("SUCCESS", { action: "getTeamForWeek_", weekIndex: weekIndex, team: team, durationMs: Date.now() - startMs });
    return team;
  } catch (e) {
    LOGGER.error("getTeamForWeek_", e.message, e.stack);
    return "care";
  }
}

function parseConfigDate_(value) {
  LOGGER.debug("START", { action: "parseConfigDate_", value: value });
  const startMs = Date.now();
  try {
    if (value instanceof Date && !isNaN(value.getTime())) {
      const parsedDate = stripTime_(value);
      LOGGER.info("SUCCESS", { action: "parseConfigDate_", mode: "date_object", durationMs: Date.now() - startMs });
      return parsedDate;
    }

    const raw = String(value || "").trim();
    if (!raw) {
      LOGGER.info("SUCCESS", { action: "parseConfigDate_", mode: "empty", durationMs: Date.now() - startMs });
      return null;
    }

    const firstPart = raw.indexOf(" ") !== -1 ? raw.split(" ")[0] : raw;
    const normalized = firstPart.indexOf("T") !== -1 ? firstPart.split("T")[0] : firstPart;
    const candidate = new Date(normalized + "T00:00:00");
    if (isNaN(candidate.getTime())) {
      LOGGER.info("SUCCESS", { action: "parseConfigDate_", mode: "invalid", durationMs: Date.now() - startMs });
      return null;
    }
    const parsed = stripTime_(candidate);
    LOGGER.info("SUCCESS", { action: "parseConfigDate_", mode: "string", durationMs: Date.now() - startMs });
    return parsed;
  } catch (e) {
    LOGGER.error("parseConfigDate_", e.message, e.stack);
    return null;
  }
}

function selectLeadForTeam_(team, mondayDate) {
  LOGGER.debug("START", { action: "selectLeadForTeam_", team: team, mondayDate: mondayDate });
  const startMs = Date.now();
  try {
    const absenceData = getAbsenceDataForDate_(mondayDate);
    const mainSelection = selectFromTeamWithAbsences_(team, absenceData, mondayDate);
    if (mainSelection.primary) {
      const resultMain = {
        primary: mainSelection.primary,
        backup: mainSelection.backup,
        candidates: mainSelection.candidates,
        available: mainSelection.available,
        teamUsed: team,
        fallbackApplied: false
      };
      LOGGER.info("SUCCESS", {
        action: "selectLeadForTeam_",
        team: team,
        teamUsed: team,
        primary: resultMain.primary ? resultMain.primary.slackUserId : "",
        backup: resultMain.backup ? resultMain.backup.slackUserId : "",
        fallbackApplied: false,
        durationMs: Date.now() - startMs
      });
      return resultMain;
    }

    if (team === "sales") {
      const fallbackSelection = selectFromTeamWithAbsences_("care", absenceData, mondayDate);
      if (fallbackSelection.primary) {
        const resultFallback = {
          primary: fallbackSelection.primary,
          backup: fallbackSelection.backup,
          candidates: fallbackSelection.candidates,
          available: fallbackSelection.available,
          teamUsed: "care",
          fallbackApplied: true,
          fallbackReason: "Aucun Sales disponible, relais Care."
        };
        LOGGER.info("SUCCESS", {
          action: "selectLeadForTeam_",
          team: team,
          teamUsed: "care",
          primary: resultFallback.primary ? resultFallback.primary.slackUserId : "",
          backup: resultFallback.backup ? resultFallback.backup.slackUserId : "",
          fallbackApplied: true,
          durationMs: Date.now() - startMs
        });
        return resultFallback;
      }
    }

    const result = {
      primary: null,
      backup: null,
      candidates: mainSelection.candidates || [],
      available: mainSelection.available || [],
      teamUsed: team,
      fallbackApplied: false
    };

    LOGGER.info("SUCCESS", {
      action: "selectLeadForTeam_",
      team: team,
      teamUsed: team,
      primary: "",
      backup: "",
      fallbackApplied: false,
      durationMs: Date.now() - startMs
    });
    return result;
  } catch (e) {
    LOGGER.error("selectLeadForTeam_", e.message, e.stack);
    return { primary: null, backup: null, candidates: [], available: [], teamUsed: team, fallbackApplied: false };
  }
}

function selectFromTeamWithAbsences_(team, absentEmails, mondayDate) {
  LOGGER.debug("START", { action: "selectFromTeamWithAbsences_", team: team, absenceData: absentEmails, mondayDate: mondayDate });
  const startMs = Date.now();
  try {
    const absenceLookup = buildAbsenceLookup_(absentEmails);
    const candidates = getRotationForTeam_(team);
    if (!candidates || candidates.length === 0) {
      LOGGER.info("SUCCESS", { action: "selectFromTeamWithAbsences_", team: team, hasCandidates: false, durationMs: Date.now() - startMs });
      return { primary: null, backup: null, candidates: [], available: [] };
    }
    const enrichedAbsenceLookup = enrichAbsenceLookupWithLuccaCandidates_(absenceLookup, candidates, mondayDate);

    const lastAssignedSlackUserId = getLastAssignedSlackUserIdForTeam_(team);
    let startIndex = 0;
    if (lastAssignedSlackUserId) {
      for (let i = 0; i < candidates.length; i++) {
        if (candidates[i].slackUserId === lastAssignedSlackUserId) {
          startIndex = (i + 1) % candidates.length;
          break;
        }
      }
    }

    const ordered = rotateArrayFromIndex_(candidates, startIndex);
    const available = ordered.filter(function(item) {
      return item.active && !isPersonAbsent_(item, enrichedAbsenceLookup);
    });
    const primary = available.length > 0 ? available[0] : null;
    const backup = available.length > 1 ? available[1] : null;

    LOGGER.info("SUCCESS", {
      action: "selectFromTeamWithAbsences_",
      team: team,
      primary: primary ? primary.slackUserId : "",
      backup: backup ? backup.slackUserId : "",
      durationMs: Date.now() - startMs
    });
    return {
      primary: primary,
      backup: backup,
      candidates: ordered,
      available: available
    };
  } catch (e) {
    LOGGER.error("selectFromTeamWithAbsences_", e.message, e.stack);
    return { primary: null, backup: null, candidates: [], available: [] };
  }
}

function enrichAbsenceLookupWithLuccaCandidates_(absenceLookup, candidates, mondayDate) {
  LOGGER.debug("START", {
    action: "enrichAbsenceLookupWithLuccaCandidates_",
    candidatesCount: candidates ? candidates.length : 0,
    mondayDate: mondayDate
  });
  const startMs = Date.now();
  try {
    const source = String(CONFIG.get("ABSENCE_SOURCE", "GSHEET") || "").toUpperCase().trim();
    if (source !== "LUCCA") {
      LOGGER.info("SUCCESS", { action: "enrichAbsenceLookupWithLuccaCandidates_", skipped: "source_not_lucca", durationMs: Date.now() - startMs });
      return absenceLookup || { emails: {}, luccaUserIds: {}, names: {} };
    }
    if (!(mondayDate instanceof Date)) {
      LOGGER.info("SUCCESS", { action: "enrichAbsenceLookupWithLuccaCandidates_", skipped: "invalid_monday", durationMs: Date.now() - startMs });
      return absenceLookup || { emails: {}, luccaUserIds: {}, names: {} };
    }

    const baseLookup = absenceLookup || { emails: {}, luccaUserIds: {}, names: {} };
    if (!baseLookup.luccaUserIds) {
      baseLookup.luccaUserIds = {};
    }

    const idsToCheck = [];
    const seen = {};
    const list = Array.isArray(candidates) ? candidates : [];
    for (let i = 0; i < list.length; i++) {
      const id = normalizeLuccaUserId_(list[i] ? list[i].luccaUserId : "");
      if (!id || seen[id] || baseLookup.luccaUserIds[id]) {
        continue;
      }
      seen[id] = true;
      idsToCheck.push(id);
    }
    if (idsToCheck.length === 0) {
      LOGGER.info("SUCCESS", { action: "enrichAbsenceLookupWithLuccaCandidates_", checked: 0, added: 0, durationMs: Date.now() - startMs });
      return baseLookup;
    }

    const luccaCfg = getLuccaConfig_();
    if (!luccaCfg.baseUrl || !luccaCfg.token) {
      LOGGER.info("SUCCESS", { action: "enrichAbsenceLookupWithLuccaCandidates_", skipped: "missing_lucca_config", durationMs: Date.now() - startMs });
      return baseLookup;
    }
    const headers = {
      Authorization: "Lucca application=" + String(luccaCfg.token),
      Accept: "application/json"
    };
    const timezone = String(CONFIG.get("TIMEZONE", "Europe/Paris"));
    const dateString = Utilities.formatDate(mondayDate, timezone, "yyyy-MM-dd");

    let added = 0;
    for (let j = 0; j < idsToCheck.length; j++) {
      const luccaUserId = idsToCheck[j];
      const refs = getPagedItemsLucca_(
        luccaCfg.baseUrl,
        "/api/v3/leaves",
        {
          date: "between," + dateString + "," + dateString,
          "leavePeriod.ownerId": luccaUserId
        },
        headers,
        1
      );
      if (refs.length > 0) {
        baseLookup.luccaUserIds[luccaUserId] = true;
        added += 1;
      }
    }

    LOGGER.info("SUCCESS", {
      action: "enrichAbsenceLookupWithLuccaCandidates_",
      checked: idsToCheck.length,
      added: added,
      durationMs: Date.now() - startMs
    });
    return baseLookup;
  } catch (e) {
    LOGGER.error("enrichAbsenceLookupWithLuccaCandidates_", e.message, e.stack);
    return absenceLookup || { emails: {}, luccaUserIds: {}, names: {} };
  }
}

function getCareRotation_() {
  LOGGER.debug("START", { action: "getCareRotation_" });
  const startMs = Date.now();
  try {
    const rows = readRotationSheet_(SHEET_NAMES.rotationCare, "getCareRotation_");
    LOGGER.info("SUCCESS", { action: "getCareRotation_", count: rows.length, durationMs: Date.now() - startMs });
    return rows;
  } catch (e) {
    LOGGER.error("getCareRotation_", e.message, e.stack);
    return [];
  }
}

function getSalesRotation_() {
  LOGGER.debug("START", { action: "getSalesRotation_" });
  const startMs = Date.now();
  try {
    const rows = readRotationSheet_(SHEET_NAMES.rotationSales, "getSalesRotation_");
    LOGGER.info("SUCCESS", { action: "getSalesRotation_", count: rows.length, durationMs: Date.now() - startMs });
    return rows;
  } catch (e) {
    LOGGER.error("getSalesRotation_", e.message, e.stack);
    return [];
  }
}

function getStartRotation_() {
  LOGGER.debug("START", { action: "getStartRotation_" });
  const startMs = Date.now();
  try {
    const rows = readRotationSheet_(SHEET_NAMES.rotationStart, "getStartRotation_");
    LOGGER.info("SUCCESS", { action: "getStartRotation_", count: rows.length, durationMs: Date.now() - startMs });
    return rows;
  } catch (e) {
    LOGGER.error("getStartRotation_", e.message, e.stack);
    return [];
  }
}

function getRotationForTeam_(team) {
  LOGGER.debug("START", { action: "getRotationForTeam_", team: team });
  const startMs = Date.now();
  try {
    const rows = team === "care" ? getCareRotation_() : getSalesRotation_();
    LOGGER.info("SUCCESS", { action: "getRotationForTeam_", team: team, count: rows.length, durationMs: Date.now() - startMs });
    return rows;
  } catch (e) {
    LOGGER.error("getRotationForTeam_", e.message, e.stack);
    return [];
  }
}

function readRotationSheet_(sheetName, actionName) {
  LOGGER.debug("START", { action: "readRotationSheet_", sheetName: sheetName, actionName: actionName });
  const startMs = Date.now();
  try {
    const sheet = getOrCreateSheet_(sheetName);
    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) {
      LOGGER.info("SUCCESS", { action: "readRotationSheet_", sheetName: sheetName, count: 0, durationMs: Date.now() - startMs });
      return [];
    }

    const rows = [];
    for (let i = 1; i < values.length; i++) {
      const order = Number(values[i][0]);
      const name = String(values[i][1] || "");
      const slackUserId = String(values[i][2] || "").trim();
      const email = String(values[i][3] || "").toLowerCase().trim();
      const active = parseBoolean_(values[i][4], true);
      const luccaUserId = normalizeLuccaUserId_(values[i][5]);
      if (!name || !slackUserId || !email) {
        continue;
      }
      rows.push({
        order: isNaN(order) ? 9999 : order,
        name: name,
        slackUserId: slackUserId,
        email: email,
        active: active,
        luccaUserId: luccaUserId
      });
    }

    rows.sort(function(a, b) { return a.order - b.order; });
    LOGGER.info("SUCCESS", {
      action: "readRotationSheet_",
      sourceAction: actionName,
      sheetName: sheetName,
      count: rows.length,
      durationMs: Date.now() - startMs
    });
    return rows;
  } catch (e) {
    LOGGER.error("readRotationSheet_", e.message, e.stack);
    return [];
  }
}

function getAbsentEmailsForDate_(targetDate) {
  LOGGER.debug("START", { action: "getAbsentEmailsForDate_", targetDate: targetDate });
  const startMs = Date.now();
  try {
    if (!(targetDate instanceof Date)) {
      LOGGER.error("getAbsentEmailsForDate_", "Invalid targetDate", "");
      return [];
    }

    const absenceData = getAbsenceDataForDate_(targetDate);
    LOGGER.info("SUCCESS", { action: "getAbsentEmailsForDate_", source: "proxy", count: absenceData.emails.length, durationMs: Date.now() - startMs });
    return absenceData.emails;
  } catch (e) {
    LOGGER.error("getAbsentEmailsForDate_", e.message, e.stack);
    return [];
  }
}

function getAbsenceDataForDate_(targetDate) {
  LOGGER.debug("START", { action: "getAbsenceDataForDate_", targetDate: targetDate });
  const startMs = Date.now();
  try {
    if (!(targetDate instanceof Date)) {
      LOGGER.error("getAbsenceDataForDate_", "Invalid targetDate", "");
      return { emails: [], luccaUserIds: [], names: [] };
    }

    const source = String(CONFIG.get("ABSENCE_SOURCE", "GSHEET") || "").toUpperCase().trim();
    if (source === "LUCCA") {
      const luccaRows = fetchLuccaAbsences_(targetDate);
      LOGGER.info("SUCCESS", {
        action: "getAbsenceDataForDate_",
        source: "LUCCA",
        emails: luccaRows.emails.length,
        luccaUserIds: luccaRows.luccaUserIds.length,
        names: (luccaRows.names || []).length,
        durationMs: Date.now() - startMs
      });
      return luccaRows;
    }

    const sheet = getOrCreateSheet_(SHEET_NAMES.absences);
    const values = sheet.getDataRange().getValues();
    const absent = [];
    for (let i = 1; i < values.length; i++) {
      const email = String(values[i][0] || "").toLowerCase();
      const startDate = new Date(values[i][2]);
      const endDate = new Date(values[i][3]);
      if (!email || isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
        continue;
      }
      if (targetDate >= stripTime_(startDate) && targetDate <= stripTime_(endDate)) {
        absent.push(email);
      }
    }

    const result = { emails: absent, luccaUserIds: [], names: [] };
    LOGGER.info("SUCCESS", { action: "getAbsenceDataForDate_", source: "GSHEET", count: absent.length, durationMs: Date.now() - startMs });
    return result;
  } catch (e) {
    LOGGER.error("getAbsenceDataForDate_", e.message, e.stack);
    return { emails: [], luccaUserIds: [], names: [] };
  }
}

function fetchLuccaAbsences_(targetDate) {
  LOGGER.debug("START", { action: "fetchLuccaAbsences_", targetDate: targetDate });
  const startMs = Date.now();
  try {
    const luccaCfg = getLuccaConfig_();
    const baseUrl = luccaCfg.baseUrl;
    const token = luccaCfg.token;
    if (!baseUrl || !token) {
      LOGGER.info("SUCCESS", { action: "fetchLuccaAbsences_", reason: "missing_config", durationMs: Date.now() - startMs });
      return { emails: [], luccaUserIds: [], names: [] };
    }

    const dateString = Utilities.formatDate(targetDate, String(CONFIG.get("TIMEZONE", "Europe/Paris")), "yyyy-MM-dd");
    const headers = {
      Authorization: "Lucca application=" + token,
      Accept: "application/json"
    };

    const refs = getPagedItemsLucca_(
      baseUrl,
      "/api/v3/leaves",
      {
        date: "between," + dateString + "," + dateString,
        "leavePeriod.ownerId": "notequal,0"
      },
      headers,
      200
    );

    // First pass directly from /leaves refs (more resilient when detail calls are rate-limited).
    const luccaIdsFromRefs = [];
    const namesFromRefs = [];
    for (let r = 0; r < refs.length; r++) {
      const ref = refs[r];
      const refOwnerId = extractLeaveOwnerId_(ref);
      if (refOwnerId) {
        luccaIdsFromRefs.push(String(refOwnerId));
      }
      const refOwnerName = normalizePersonName_(extractLeaveOwnerName_(ref));
      if (refOwnerName) {
        namesFromRefs.push(refOwnerName);
      }
    }

    const details = fetchLuccaDetailsWithRetry_(refs, headers, {
      maxRetries: 5,
      baseDelayMs: 700,
      maxDelayMs: 12000,
      minGapMs: 100
    });

    const luccaIds = luccaIdsFromRefs.slice();
    const names = namesFromRefs.slice();
    for (let i = 0; i < details.length; i++) {
      const ownerId = extractLeaveOwnerId_(details[i]);
      if (ownerId) {
        luccaIds.push(String(ownerId));
      }
      const ownerName = extractLeaveOwnerName_(details[i]);
      const normalizedOwnerName = normalizePersonName_(ownerName);
      if (normalizedOwnerName) {
        names.push(normalizedOwnerName);
      }
    }
    const uniqueLuccaIds = uniqueCsvValues_(luccaIds);
    const userMap = fetchLuccaUsersByIdMap_(uniqueLuccaIds);
    const emails = [];
    for (let j = 0; j < uniqueLuccaIds.length; j++) {
      const email = String(userMap[uniqueLuccaIds[j]] || "").toLowerCase().trim();
      if (email) {
        emails.push(email);
      }
    }

    LOGGER.info("SUCCESS", {
      action: "fetchLuccaAbsences_",
      refs: refs.length,
      details: details.length,
      luccaUserIdsFromRefs: uniqueCsvValues_(luccaIdsFromRefs).length,
      luccaUserIds: uniqueLuccaIds.length,
      emails: emails.length,
      durationMs: Date.now() - startMs
    });
    return { emails: uniqueCsvValues_(emails), luccaUserIds: uniqueLuccaIds, names: uniqueCsvValues_(names) };
  } catch (e) {
    LOGGER.error("fetchLuccaAbsences_", e.message, e.stack);
    return { emails: [], luccaUserIds: [], names: [] };
  }
}

function extractLeaveOwnerName_(lv) {
  LOGGER.debug("START", { action: "extractLeaveOwnerName_", lv: lv });
  const startMs = Date.now();
  try {
    if (!lv) {
      LOGGER.info("SUCCESS", { action: "extractLeaveOwnerName_", found: false, durationMs: Date.now() - startMs });
      return "";
    }
    const ownerName =
      (lv.leavePeriod && lv.leavePeriod.owner && (lv.leavePeriod.owner.name || (Array.isArray(lv.leavePeriod.owner) ? lv.leavePeriod.owner[1] : ""))) ||
      (lv.owner && (lv.owner.name || (Array.isArray(lv.owner) ? lv.owner[1] : ""))) ||
      (lv.user && (lv.user.name || (Array.isArray(lv.user) ? lv.user[1] : ""))) ||
      "";
    const result = String(ownerName || "").trim();
    LOGGER.info("SUCCESS", { action: "extractLeaveOwnerName_", found: !!result, durationMs: Date.now() - startMs });
    return result;
  } catch (e) {
    LOGGER.error("extractLeaveOwnerName_", e.message, e.stack);
    return "";
  }
}

function isLuccaEnabled_() {
  LOGGER.debug("START", { action: "isLuccaEnabled_" });
  const startMs = Date.now();
  try {
    const cfg = getLuccaConfig_();
    const enabled = !!(cfg.baseUrl && cfg.token);
    LOGGER.info("SUCCESS", { action: "isLuccaEnabled_", enabled: enabled, durationMs: Date.now() - startMs });
    return enabled;
  } catch (e) {
    LOGGER.error("isLuccaEnabled_", e.message, e.stack);
    return false;
  }
}

function getLuccaConfig_() {
  LOGGER.debug("START", { action: "getLuccaConfig_" });
  const startMs = Date.now();
  try {
    const baseUrl = String(
      getScriptProperty_("URL_LUCCA", getScriptProperty_("LUCCA_BASE_URL", ""))
    ).replace(/\/+$/, "");
    const token = String(
      getScriptProperty_("LUCCA_TOKEN", getScriptProperty_("API_LUCCA_RH", getScriptProperty_("LUCCA_API_TOKEN", "")))
    );
    const result = { baseUrl: baseUrl, token: token };
    LOGGER.info("SUCCESS", { action: "getLuccaConfig_", hasBaseUrl: !!baseUrl, hasToken: !!token, durationMs: Date.now() - startMs });
    return result;
  } catch (e) {
    LOGGER.error("getLuccaConfig_", e.message, e.stack);
    return { baseUrl: "", token: "" };
  }
}

function syncLuccaUserIdsInRotations() {
  LOGGER.debug("START", { action: "syncLuccaUserIdsInRotations" });
  const startMs = Date.now();
  try {
    const targetEmails = getRotationEmailsForLuccaSync_();
    const luccaMaps = fetchLuccaUsersMaps_(targetEmails);
    const updatedCare = syncLuccaUserIdsOnSheet_(SHEET_NAMES.rotationCare, luccaMaps);
    const updatedSales = syncLuccaUserIdsOnSheet_(SHEET_NAMES.rotationSales, luccaMaps);
    const updatedStart = syncLuccaUserIdsOnSheet_(SHEET_NAMES.rotationStart, luccaMaps);
    const result = {
      updated: updatedCare + updatedSales + updatedStart,
      perSheet: {
        rotationCare: updatedCare,
        rotationSales: updatedSales,
        rotationStart: updatedStart
      }
    };
    LOGGER.info("SUCCESS", { action: "syncLuccaUserIdsInRotations", result: result, durationMs: Date.now() - startMs });
    return result;
  } catch (e) {
    LOGGER.error("syncLuccaUserIdsInRotations", e.message, e.stack);
    return { updated: 0, perSheet: {} };
  }
}

function getRotationEmailsForLuccaSync_() {
  LOGGER.debug("START", { action: "getRotationEmailsForLuccaSync_" });
  const startMs = Date.now();
  try {
    const emails = [];
    const all = getCareRotation_().concat(getSalesRotation_()).concat(getStartRotation_());
    for (let i = 0; i < all.length; i++) {
      const email = String(all[i].email || "").toLowerCase().trim();
      if (email) {
        emails.push(email);
      }
    }
    const uniqueEmails = uniqueCsvValues_(emails);
    LOGGER.info("SUCCESS", { action: "getRotationEmailsForLuccaSync_", count: uniqueEmails.length, durationMs: Date.now() - startMs });
    return uniqueEmails;
  } catch (e) {
    LOGGER.error("getRotationEmailsForLuccaSync_", e.message, e.stack);
    return [];
  }
}

function syncLuccaUserIdsOnSheet_(sheetName, luccaMaps) {
  LOGGER.debug("START", { action: "syncLuccaUserIdsOnSheet_", sheetName: sheetName });
  const startMs = Date.now();
  try {
    const sheet = getOrCreateSheet_(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) {
      LOGGER.info("SUCCESS", { action: "syncLuccaUserIdsOnSheet_", sheetName: sheetName, updated: 0, durationMs: Date.now() - startMs });
      return 0;
    }
    const header = sheet.getRange(1, 1, 1, Math.max(6, sheet.getLastColumn())).getValues()[0];
    if (String(header[5] || "").trim() !== "luccaUserId") {
      sheet.getRange(1, 6).setValue("luccaUserId");
    }
    const rowCount = sheet.getLastRow() - 1;
    const values = sheet.getRange(2, 1, rowCount, 6).getValues();
    const luccaCol = [];
    let updated = 0;
    const byEmail = luccaMaps && luccaMaps.byEmail ? luccaMaps.byEmail : {};
    const byName = luccaMaps && luccaMaps.byName ? luccaMaps.byName : {};
    for (let i = 0; i < values.length; i++) {
      const name = String(values[i][1] || "").trim();
      const email = String(values[i][3] || "").toLowerCase().trim();
      const currentId = normalizeLuccaUserId_(values[i][5]);
      const targetIdByEmail = normalizeLuccaUserId_(byEmail[email] || "");
      const targetIdByName = normalizeLuccaUserId_(byName[normalizePersonName_(name)] || "");
      const targetId = targetIdByEmail || targetIdByName;
      const nextId = targetId || currentId;
      if (nextId !== currentId && targetId) {
        updated += 1;
      }
      luccaCol.push([nextId]);
    }
    if (luccaCol.length > 0) {
      sheet.getRange(2, 6, luccaCol.length, 1).setValues(luccaCol);
    }
    LOGGER.info("SUCCESS", { action: "syncLuccaUserIdsOnSheet_", sheetName: sheetName, updated: updated, durationMs: Date.now() - startMs });
    return updated;
  } catch (e) {
    LOGGER.error("syncLuccaUserIdsOnSheet_", e.message, e.stack);
    return 0;
  }
}

function fetchLuccaUsersByEmailMap_(targetEmails) {
  LOGGER.debug("START", { action: "fetchLuccaUsersByEmailMap_", targetEmails: targetEmails });
  const startMs = Date.now();
  try {
    const maps = fetchLuccaUsersMaps_(targetEmails);
    const map = maps.byEmail;
    LOGGER.info("SUCCESS", { action: "fetchLuccaUsersByEmailMap_", users: maps.usersCount || 0, mapped: Object.keys(map).length, durationMs: Date.now() - startMs });
    return map;
  } catch (e) {
    LOGGER.error("fetchLuccaUsersByEmailMap_", e.message, e.stack);
    return {};
  }
}

function fetchLuccaUsersByIdMap_(luccaIds) {
  LOGGER.debug("START", { action: "fetchLuccaUsersByIdMap_", luccaIds: luccaIds });
  const startMs = Date.now();
  try {
    const ids = Array.isArray(luccaIds) ? luccaIds : [];
    if (ids.length === 0) {
      LOGGER.info("SUCCESS", { action: "fetchLuccaUsersByIdMap_", mapped: 0, durationMs: Date.now() - startMs });
      return {};
    }
    const maps = fetchLuccaUsersMaps_([]);
    const emailMap = maps.byEmail;
    const idMap = {};
    const usersByEmailKeys = Object.keys(emailMap);
    if (usersByEmailKeys.length === 0) {
      LOGGER.info("SUCCESS", { action: "fetchLuccaUsersByIdMap_", mapped: 0, durationMs: Date.now() - startMs });
      return {};
    }

    for (let i = 0; i < ids.length; i++) {
      const id = normalizeLuccaUserId_(ids[i]);
      if (id && maps.byId[id]) {
        idMap[id] = maps.byId[id];
      }
    }
    LOGGER.info("SUCCESS", { action: "fetchLuccaUsersByIdMap_", requested: ids.length, mapped: Object.keys(idMap).length, durationMs: Date.now() - startMs });
    return idMap;
  } catch (e) {
    LOGGER.error("fetchLuccaUsersByIdMap_", e.message, e.stack);
    return {};
  }
}

function fetchLuccaUsersMaps_(targetEmails) {
  LOGGER.debug("START", { action: "fetchLuccaUsersMaps_", targetEmails: targetEmails });
  const startMs = Date.now();
  try {
    const luccaCfg = getLuccaConfig_();
    if (!luccaCfg.baseUrl || !luccaCfg.token) {
      LOGGER.info("SUCCESS", { action: "fetchLuccaUsersMaps_", reason: "missing_config", durationMs: Date.now() - startMs });
      return { byEmail: {}, byId: {}, byName: {}, usersCount: 0 };
    }
    const headers = {
      Authorization: "Lucca application=" + luccaCfg.token,
      Accept: "application/json"
    };
    const targetList = Array.isArray(targetEmails)
      ? uniqueCsvValues_(targetEmails.map(function(v) { return String(v || "").toLowerCase().trim(); }))
      : [];
    const users = getPagedItemsLucca_(luccaCfg.baseUrl, "/api/v3/users", {}, headers, 200);
    const byEmailRaw = {};
    const byIdRaw = {};
    const byNameRaw = {};
    for (let i = 0; i < users.length; i++) {
      const user = users[i] || {};
      const email = extractLuccaUserEmail_(user);
      const id = extractLuccaUserId_(user);
      const name = extractLuccaUserName_(user);
      if (email && id) {
        byEmailRaw[email] = id;
        byIdRaw[id] = email;
      }
      if (name && id) {
        byNameRaw[name] = id;
      }
    }
    const byEmail = {};
    const byId = {};
    const byName = {};
    if (targetList.length > 0) {
      for (let j = 0; j < targetList.length; j++) {
        const tEmail = targetList[j];
        const tId = byEmailRaw[tEmail];
        if (tId) {
          byEmail[tEmail] = tId;
          byId[tId] = tEmail;
        }
      }
    } else {
      const eKeys = Object.keys(byEmailRaw);
      for (let k = 0; k < eKeys.length; k++) {
        byEmail[eKeys[k]] = byEmailRaw[eKeys[k]];
      }
      const iKeys = Object.keys(byIdRaw);
      for (let m = 0; m < iKeys.length; m++) {
        byId[iKeys[m]] = byIdRaw[iKeys[m]];
      }
      const nKeys = Object.keys(byNameRaw);
      for (let n = 0; n < nKeys.length; n++) {
        byName[nKeys[n]] = byNameRaw[nKeys[n]];
      }
    }
    if (targetList.length > 0) {
      const nKeysTarget = Object.keys(byNameRaw);
      for (let p = 0; p < nKeysTarget.length; p++) {
        byName[nKeysTarget[p]] = byNameRaw[nKeysTarget[p]];
      }
    }
    LOGGER.info("SUCCESS", {
      action: "fetchLuccaUsersMaps_",
      mode: targetList.length > 0 ? "targeted_from_full_scan" : "full_scan",
      targetCount: targetList.length,
      users: users.length,
      byEmail: Object.keys(byEmail).length,
      byId: Object.keys(byId).length,
      byName: Object.keys(byName).length,
      durationMs: Date.now() - startMs
    });
    return { byEmail: byEmail, byId: byId, byName: byName, usersCount: users.length };
  } catch (e) {
    LOGGER.error("fetchLuccaUsersMaps_", e.message, e.stack);
    return { byEmail: {}, byId: {}, byName: {}, usersCount: 0 };
  }
}

function fetchLuccaUserByEmail_(baseUrl, headers, email) {
  LOGGER.debug("START", { action: "fetchLuccaUserByEmail_", email: email });
  const startMs = Date.now();
  try {
    const cleanEmail = String(email || "").toLowerCase().trim();
    if (!cleanEmail) {
      LOGGER.info("SUCCESS", { action: "fetchLuccaUserByEmail_", found: false, durationMs: Date.now() - startMs });
      return null;
    }
    const queryVariants = [
      { mail: "equal," + cleanEmail, paging: "0,1" },
      { email: "equal," + cleanEmail, paging: "0,1" },
      { mail: cleanEmail, paging: "0,1" },
      { email: cleanEmail, paging: "0,1" }
    ];
    const maxRetries = 4;
    for (let q = 0; q < queryVariants.length; q++) {
      const url = baseUrl + "/api/v3/users?" + toQueryStringLucca_(queryVariants[q]);
      for (let attempt = 0; attempt <= maxRetries; attempt++) {
        const resp = UrlFetchApp.fetch(url, {
          method: "get",
          headers: headers,
          muteHttpExceptions: true
        });
        const code = resp.getResponseCode();
        const txt = resp.getContentText() || "";
        if (code >= 200 && code < 300) {
          const json = txt ? JSON.parse(txt) : {};
          const items = asLuccaItems_(json);
          if (items.length > 0) {
            const user = items[0] || {};
            const foundEmail = extractLuccaUserEmail_(user);
            const foundId = extractLuccaUserId_(user);
            if (foundEmail && foundId) {
              LOGGER.info("SUCCESS", { action: "fetchLuccaUserByEmail_", found: true, email: foundEmail, durationMs: Date.now() - startMs });
              return { email: foundEmail, id: foundId };
            }
          }
          break;
        }
        if (code === 429 || code === 502 || code === 503 || code === 504) {
          const delay = Math.min(10000, 700 * Math.pow(2, attempt));
          Utilities.sleep(delay + Math.floor(Math.random() * 200));
          continue;
        }
        break;
      }
    }
    LOGGER.info("SUCCESS", { action: "fetchLuccaUserByEmail_", found: false, email: cleanEmail, durationMs: Date.now() - startMs });
    return null;
  } catch (e) {
    LOGGER.error("fetchLuccaUserByEmail_", e.message, e.stack);
    return null;
  }
}

function extractLuccaUserEmail_(user) {
  LOGGER.debug("START", { action: "extractLuccaUserEmail_", user: user });
  const startMs = Date.now();
  try {
    if (!user) {
      LOGGER.info("SUCCESS", { action: "extractLuccaUserEmail_", found: false, durationMs: Date.now() - startMs });
      return "";
    }
    const email = String(
      user.mail ||
      user.email ||
      user.workEmail ||
      (user.owner && user.owner.mail) ||
      ""
    ).toLowerCase().trim();
    LOGGER.info("SUCCESS", { action: "extractLuccaUserEmail_", found: !!email, durationMs: Date.now() - startMs });
    return email;
  } catch (e) {
    LOGGER.error("extractLuccaUserEmail_", e.message, e.stack);
    return "";
  }
}

function extractLuccaUserName_(user) {
  LOGGER.debug("START", { action: "extractLuccaUserName_", user: user });
  const startMs = Date.now();
  try {
    if (!user) {
      LOGGER.info("SUCCESS", { action: "extractLuccaUserName_", found: false, durationMs: Date.now() - startMs });
      return "";
    }
    const firstName = String(user.firstName || user.firstname || (user.owner && user.owner.firstName) || "").trim();
    const lastName = String(user.lastName || user.lastname || (user.owner && user.owner.lastName) || "").trim();
    const displayName = String(user.name || user.displayName || (user.owner && user.owner.name) || "").trim();
    const rawName = (firstName || lastName) ? (firstName + " " + lastName).trim() : displayName;
    const normalized = normalizePersonName_(rawName);
    LOGGER.info("SUCCESS", { action: "extractLuccaUserName_", found: !!normalized, durationMs: Date.now() - startMs });
    return normalized;
  } catch (e) {
    LOGGER.error("extractLuccaUserName_", e.message, e.stack);
    return "";
  }
}

function extractLuccaUserId_(user) {
  LOGGER.debug("START", { action: "extractLuccaUserId_", user: user });
  const startMs = Date.now();
  try {
    if (!user) {
      LOGGER.info("SUCCESS", { action: "extractLuccaUserId_", found: false, durationMs: Date.now() - startMs });
      return "";
    }
    const id = normalizeLuccaUserId_(
      user.id ||
      user.userId ||
      (user.owner && user.owner.id) ||
      (user.person && user.person.id) ||
      (user.url ? String(user.url).match(/(\d+)(?:\D*)$/) : "")
    );
    LOGGER.info("SUCCESS", { action: "extractLuccaUserId_", found: !!id, durationMs: Date.now() - startMs });
    return id;
  } catch (e) {
    LOGGER.error("extractLuccaUserId_", e.message, e.stack);
    return "";
  }
}

function getPagedItemsLucca_(baseUrl, path, query, headers, pageSize) {
  LOGGER.debug("START", { action: "getPagedItemsLucca_", path: path, query: query, pageSize: pageSize });
  const startMs = Date.now();
  try {
    const out = [];
    let offset = 0;
    const limit = pageSize || 200;
    const maxPages = 100;
    let pageCount = 0;
    let previousSignature = "";
    const maxRetries = 5;
    const baseDelayMs = 700;
    const maxDelayMs = 12000;
    while (true) {
      pageCount += 1;
      if (pageCount > maxPages) {
        LOGGER.error("getPagedItemsLucca_", "Max pages reached", "path=" + path);
        break;
      }
      const q = Object.assign({}, query || {}, { paging: offset + "," + limit });
      const url = baseUrl + path + "?" + toQueryStringLucca_(q);
      let pageJson = null;
      let pageOk = false;
      for (let attempt = 0; attempt <= maxRetries; attempt++) {
        const resp = UrlFetchApp.fetch(url, {
          method: "get",
          headers: headers,
          muteHttpExceptions: true
        });
        const code = resp.getResponseCode();
        const txt = resp.getContentText() || "";
        if (code >= 200 && code < 300) {
          pageJson = txt ? JSON.parse(txt) : {};
          pageOk = true;
          break;
        }
        if (code === 429 || code === 502 || code === 503 || code === 504) {
          const delay = Math.min(maxDelayMs, baseDelayMs * Math.pow(2, attempt));
          const jitter = Math.floor(Math.random() * 250);
          Utilities.sleep(delay + jitter);
          continue;
        }
        LOGGER.error("getPagedItemsLucca_", "HTTP " + code + " on " + path, txt);
        pageOk = false;
        break;
      }
      if (!pageOk) {
        LOGGER.error("getPagedItemsLucca_", "Page request failed after retries", "url=" + url);
        break;
      }
      const items = asLuccaItems_(pageJson);
      if (!items.length) {
        break;
      }
      const firstId = items[0] && items[0].id ? String(items[0].id) : "";
      const lastId = items[items.length - 1] && items[items.length - 1].id ? String(items[items.length - 1].id) : "";
      const signature = firstId + "|" + lastId + "|" + String(items.length);
      if (signature && signature === previousSignature) {
        LOGGER.error("getPagedItemsLucca_", "Repeated page signature detected", signature);
        break;
      }
      previousSignature = signature;
      Array.prototype.push.apply(out, items);
      if (items.length < limit) {
        break;
      }
      offset += limit;
    }
    LOGGER.info("SUCCESS", { action: "getPagedItemsLucca_", count: out.length, durationMs: Date.now() - startMs });
    return out;
  } catch (e) {
    LOGGER.error("getPagedItemsLucca_", e.message, e.stack);
    return [];
  }
}

function fetchLuccaDetailsWithRetry_(refs, headers, cfg) {
  LOGGER.debug("START", { action: "fetchLuccaDetailsWithRetry_", refsCount: refs ? refs.length : 0, cfg: cfg });
  const startMs = Date.now();
  try {
    const maxRetries = cfg.maxRetries || 5;
    const baseDelayMs = cfg.baseDelayMs || 600;
    const maxDelayMs = cfg.maxDelayMs || 10000;
    const minGapMs = cfg.minGapMs || 100;
    const out = [];
    const urls = (refs || []).map(function(r) {
      return r && r.url ? String(r.url) : "";
    }).filter(function(v) {
      return !!v;
    });
    let lastCallAt = 0;
    for (let i = 0; i < urls.length; i++) {
      const url = urls[i];
      let ok = false;
      for (let a = 0; a <= maxRetries; a++) {
        const now = Date.now();
        const wait = minGapMs - (now - lastCallAt);
        if (wait > 0) {
          Utilities.sleep(wait);
        }
        lastCallAt = Date.now();
        const resp = UrlFetchApp.fetch(url, {
          method: "get",
          headers: headers,
          muteHttpExceptions: true
        });
        const code = resp.getResponseCode();
        const txt = resp.getContentText() || "";
        if (code >= 200 && code < 300) {
          let json = null;
          try {
            json = txt ? JSON.parse(txt) : null;
          } catch (ignored) {
            json = null;
          }
          if (json && json.data && typeof json.data === "object") {
            out.push(json.data);
            ok = true;
          }
          break;
        }
        if (code === 429 || code === 502 || code === 503 || code === 504) {
          const backoff = Math.min(maxDelayMs, baseDelayMs * Math.pow(2, a));
          const jitter = Math.floor(Math.random() * 250);
          Utilities.sleep(backoff + jitter);
          continue;
        }
        break;
      }
      if (!ok) {
        continue;
      }
    }
    LOGGER.info("SUCCESS", { action: "fetchLuccaDetailsWithRetry_", refs: urls.length, details: out.length, durationMs: Date.now() - startMs });
    return out;
  } catch (e) {
    LOGGER.error("fetchLuccaDetailsWithRetry_", e.message, e.stack);
    return [];
  }
}

function extractLeaveOwnerId_(lv) {
  LOGGER.debug("START", { action: "extractLeaveOwnerId_", lv: lv });
  const startMs = Date.now();
  try {
    if (!lv) {
      LOGGER.info("SUCCESS", { action: "extractLeaveOwnerId_", found: false, durationMs: Date.now() - startMs });
      return "";
    }
    const value =
      (lv.leavePeriod && (
        lv.leavePeriod.ownerId ||
        (lv.leavePeriod.owner && lv.leavePeriod.owner.id) ||
        lv.leavePeriod.owner
      )) ||
      lv.ownerId ||
      (lv.owner && (lv.owner.id || lv.owner)) ||
      (lv.user && (lv.user.id || lv.user)) ||
      null;
    const normalized = normalizeLuccaUserId_(value);
    if (normalized) {
      LOGGER.info("SUCCESS", { action: "extractLeaveOwnerId_", found: true, durationMs: Date.now() - startMs });
      return normalized;
    }
    const rawId = String(lv.id || "");
    const match = rawId.match(/^(\d+)-/);
    const fallback = match ? String(match[1]) : "";
    LOGGER.info("SUCCESS", { action: "extractLeaveOwnerId_", found: !!fallback, durationMs: Date.now() - startMs });
    return fallback;
  } catch (e) {
    LOGGER.error("extractLeaveOwnerId_", e.message, e.stack);
    return "";
  }
}

function asLuccaItems_(json) {
  LOGGER.debug("START", { action: "asLuccaItems_", hasJson: !!json });
  const startMs = Date.now();
  try {
    let items = [];
    if (!json) {
      items = [];
    } else if (Array.isArray(json.data)) {
      items = json.data;
    } else if (json.data && Array.isArray(json.data.items)) {
      items = json.data.items;
    } else if (Array.isArray(json.items)) {
      items = json.items;
    }
    LOGGER.info("SUCCESS", { action: "asLuccaItems_", count: items.length, durationMs: Date.now() - startMs });
    return items;
  } catch (e) {
    LOGGER.error("asLuccaItems_", e.message, e.stack);
    return [];
  }
}

function toQueryStringLucca_(obj) {
  LOGGER.debug("START", { action: "toQueryStringLucca_", obj: obj });
  const startMs = Date.now();
  try {
    const q = Object.keys(obj || {})
      .filter(function(k) {
        return obj[k] !== undefined && obj[k] !== null && obj[k] !== "";
      })
      .map(function(k) {
        return encodeURIComponent(k) + "=" + encodeURIComponent(String(obj[k]));
      })
      .join("&");
    LOGGER.info("SUCCESS", { action: "toQueryStringLucca_", length: q.length, durationMs: Date.now() - startMs });
    return q;
  } catch (e) {
    LOGGER.error("toQueryStringLucca_", e.message, e.stack);
    return "";
  }
}

function buildAbsenceLookup_(absenceData) {
  LOGGER.debug("START", { action: "buildAbsenceLookup_", absenceData: absenceData });
  const startMs = Date.now();
  try {
    const data = absenceData || {};
    const emails = Array.isArray(data.emails) ? data.emails : [];
    const luccaUserIds = Array.isArray(data.luccaUserIds) ? data.luccaUserIds : [];
    const names = Array.isArray(data.names) ? data.names : [];
    const emailMap = {};
    const luccaMap = {};
    const nameMap = {};
    for (let i = 0; i < emails.length; i++) {
      const email = String(emails[i] || "").toLowerCase().trim();
      if (email) {
        emailMap[email] = true;
      }
    }
    for (let j = 0; j < luccaUserIds.length; j++) {
      const id = normalizeLuccaUserId_(luccaUserIds[j]);
      if (id) {
        luccaMap[id] = true;
      }
    }
    for (let k = 0; k < names.length; k++) {
      const n = normalizePersonName_(names[k]);
      if (n) {
        nameMap[n] = true;
      }
    }
    const result = { emails: emailMap, luccaUserIds: luccaMap, names: nameMap };
    LOGGER.info("SUCCESS", {
      action: "buildAbsenceLookup_",
      emails: Object.keys(emailMap).length,
      luccaUserIds: Object.keys(luccaMap).length,
      names: Object.keys(nameMap).length,
      durationMs: Date.now() - startMs
    });
    return result;
  } catch (e) {
    LOGGER.error("buildAbsenceLookup_", e.message, e.stack);
    return { emails: {}, luccaUserIds: {}, names: {} };
  }
}

function isPersonAbsent_(person, absenceLookup) {
  LOGGER.debug("START", { action: "isPersonAbsent_", person: person, absenceLookup: absenceLookup });
  const startMs = Date.now();
  try {
    const email = String(person && person.email ? person.email : "").toLowerCase().trim();
    const luccaUserId = normalizeLuccaUserId_(person ? person.luccaUserId : "");
    const normalizedName = normalizePersonName_(person && person.name ? person.name : "");
    const byEmail = !!(email && absenceLookup && absenceLookup.emails && absenceLookup.emails[email]);
    const byLucca = !!(luccaUserId && absenceLookup && absenceLookup.luccaUserIds && absenceLookup.luccaUserIds[luccaUserId]);
    const byName = !!(normalizedName && absenceLookup && absenceLookup.names && absenceLookup.names[normalizedName]);
    const absent = byEmail || byLucca || byName;
    LOGGER.info("SUCCESS", { action: "isPersonAbsent_", byEmail: byEmail, byLucca: byLucca, byName: byName, absent: absent, durationMs: Date.now() - startMs });
    return absent;
  } catch (e) {
    LOGGER.error("isPersonAbsent_", e.message, e.stack);
    return false;
  }
}

function normalizeLuccaUserId_(value) {
  LOGGER.debug("START", { action: "normalizeLuccaUserId_", value: value });
  const startMs = Date.now();
  try {
    let raw = value;
    if (Array.isArray(raw) && raw.length > 0) {
      // Lucca can return relation tuples like [id, label, url].
      // Always prioritize the first numeric candidate.
      let picked = "";
      for (let i = 0; i < raw.length; i++) {
        const item = raw[i];
        if (item === null || item === undefined) {
          continue;
        }
        if (typeof item === "number" && !isNaN(item)) {
          picked = String(item);
          break;
        }
        const strItem = String(item).trim();
        if (/^\d+$/.test(strItem)) {
          picked = strItem;
          break;
        }
        const matchItem = strItem.match(/(?:^|\/)(\d+)(?:\D*)$/);
        if (matchItem) {
          picked = String(matchItem[1]);
          break;
        }
      }
      raw = picked || raw[0] || "";
    }
    if (raw && typeof raw === "object") {
      raw = raw.id || raw.userId || raw.value || raw.url || "";
    }
    raw = String(raw || "").trim();
    if (!raw) {
      LOGGER.info("SUCCESS", { action: "normalizeLuccaUserId_", normalized: "", durationMs: Date.now() - startMs });
      return "";
    }
    let normalized = "";
    if (/^\d+$/.test(raw)) {
      normalized = raw;
    } else {
      const match = raw.match(/(\d+)(?:\D*)$/);
      normalized = match ? String(match[1]) : "";
    }
    LOGGER.info("SUCCESS", { action: "normalizeLuccaUserId_", normalized: normalized, durationMs: Date.now() - startMs });
    return normalized;
  } catch (e) {
    LOGGER.error("normalizeLuccaUserId_", e.message, e.stack);
    return "";
  }
}

function normalizePersonName_(value) {
  LOGGER.debug("START", { action: "normalizePersonName_", value: value });
  const startMs = Date.now();
  try {
    const raw = String(value || "").toLowerCase().trim();
    if (!raw) {
      LOGGER.info("SUCCESS", { action: "normalizePersonName_", normalized: "", durationMs: Date.now() - startMs });
      return "";
    }
    const normalized = raw
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-z0-9 ]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
    LOGGER.info("SUCCESS", { action: "normalizePersonName_", normalized: normalized, durationMs: Date.now() - startMs });
    return normalized;
  } catch (e) {
    LOGGER.error("normalizePersonName_", e.message, e.stack);
    return "";
  }
}

function uniqueCsvValues_(arr) {
  LOGGER.debug("START", { action: "uniqueCsvValues_", arr: arr });
  const startMs = Date.now();
  try {
    const list = Array.isArray(arr) ? arr : [];
    const seen = {};
    const out = [];
    for (let i = 0; i < list.length; i++) {
      const key = String(list[i] || "").trim();
      if (!key || seen[key]) {
        continue;
      }
      seen[key] = true;
      out.push(key);
    }
    LOGGER.info("SUCCESS", { action: "uniqueCsvValues_", count: out.length, durationMs: Date.now() - startMs });
    return out;
  } catch (e) {
    LOGGER.error("uniqueCsvValues_", e.message, e.stack);
    return [];
  }
}

function getScriptProperty_(key, defaultValue) {
  LOGGER.debug("START", { action: "getScriptProperty_", key: key, defaultValue: defaultValue });
  const startMs = Date.now();
  try {
    if (!key || typeof key !== "string") {
      LOGGER.error("getScriptProperty_", "Invalid key", "");
      return defaultValue;
    }
    const props = PropertiesService.getScriptProperties();
    const value = props.getProperty(key);
    const result = value === null ? defaultValue : value;
    LOGGER.info("SUCCESS", { action: "getScriptProperty_", key: key, found: value !== null, durationMs: Date.now() - startMs });
    return result;
  } catch (e) {
    LOGGER.error("getScriptProperty_", e.message, e.stack);
    return defaultValue;
  }
}

function getLastAssignedSlackUserIdForTeam_(team) {
  LOGGER.debug("START", { action: "getLastAssignedSlackUserIdForTeam_", team: team });
  const startMs = Date.now();
  try {
    const sheet = getOrCreateSheet_(SHEET_NAMES.history);
    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) {
      LOGGER.info("SUCCESS", { action: "getLastAssignedSlackUserIdForTeam_", slackUserId: "", durationMs: Date.now() - startMs });
      return "";
    }

    const headers = values[0].map(function(item) { return String(item || ""); });
    const teamIdx = headers.indexOf("team");
    const statusIdx = headers.indexOf("status");
    const primaryIdx = headers.indexOf("primarySlackUserId");
    const resolvedTeamIdx = teamIdx === -1 ? 2 : teamIdx;
    const resolvedStatusIdx = statusIdx === -1 ? 7 : statusIdx;
    const resolvedPrimaryIdx = primaryIdx === -1 ? 4 : primaryIdx;

    for (let i = values.length - 1; i >= 1; i--) {
      if (String(values[i][resolvedTeamIdx]) === team && String(values[i][resolvedStatusIdx]) === "POSTED") {
        const slackUserId = String(values[i][resolvedPrimaryIdx] || "");
        LOGGER.info("SUCCESS", { action: "getLastAssignedSlackUserIdForTeam_", slackUserId: slackUserId, durationMs: Date.now() - startMs });
        return slackUserId;
      }
    }
    LOGGER.info("SUCCESS", { action: "getLastAssignedSlackUserIdForTeam_", slackUserId: "", durationMs: Date.now() - startMs });
    return "";
  } catch (e) {
    LOGGER.error("getLastAssignedSlackUserIdForTeam_", e.message, e.stack);
    return "";
  }
}

function getLastAssignedStarterSlackUserId_() {
  LOGGER.debug("START", { action: "getLastAssignedStarterSlackUserId_" });
  const startMs = Date.now();
  try {
    const sheet = getOrCreateSheet_(SHEET_NAMES.history);
    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) {
      LOGGER.info("SUCCESS", { action: "getLastAssignedStarterSlackUserId_", slackUserId: "", durationMs: Date.now() - startMs });
      return "";
    }

    const headers = values[0].map(function(item) { return String(item || ""); });
    const statusIdx = headers.indexOf("status");
    const starterIdx = headers.indexOf("starterSlackUserId");
    if (statusIdx === -1 || starterIdx === -1) {
      LOGGER.info("SUCCESS", { action: "getLastAssignedStarterSlackUserId_", slackUserId: "", reason: "missing_columns", durationMs: Date.now() - startMs });
      return "";
    }

    for (let i = values.length - 1; i >= 1; i--) {
      if (String(values[i][statusIdx]) !== "POSTED") {
        continue;
      }
      const starterSlackUserId = String(values[i][starterIdx] || "");
      if (starterSlackUserId) {
        LOGGER.info("SUCCESS", { action: "getLastAssignedStarterSlackUserId_", slackUserId: starterSlackUserId, durationMs: Date.now() - startMs });
        return starterSlackUserId;
      }
    }

    LOGGER.info("SUCCESS", { action: "getLastAssignedStarterSlackUserId_", slackUserId: "", durationMs: Date.now() - startMs });
    return "";
  } catch (e) {
    LOGGER.error("getLastAssignedStarterSlackUserId_", e.message, e.stack);
    return "";
  }
}

function selectStarterForWeek_(mondayDate) {
  LOGGER.debug("START", { action: "selectStarterForWeek_", mondayDate: mondayDate });
  const startMs = Date.now();
  try {
    const candidates = getStartRotation_();
    if (!candidates || candidates.length === 0) {
      LOGGER.info("SUCCESS", { action: "selectStarterForWeek_", found: false, reason: "no_candidates", durationMs: Date.now() - startMs });
      return null;
    }

    const absenceData = getAbsenceDataForDate_(mondayDate);
    const absenceLookup = buildAbsenceLookup_(absenceData);
    const lastAssigned = getLastAssignedStarterSlackUserId_();
    let startIndex = 0;
    if (lastAssigned) {
      for (let i = 0; i < candidates.length; i++) {
        if (candidates[i].slackUserId === lastAssigned) {
          startIndex = (i + 1) % candidates.length;
          break;
        }
      }
    }

    const ordered = rotateArrayFromIndex_(candidates, startIndex);
    const available = ordered.filter(function(item) {
      return item.active && !isPersonAbsent_(item, absenceLookup);
    });
    const starter = available.length > 0 ? available[0] : null;
    LOGGER.info("SUCCESS", {
      action: "selectStarterForWeek_",
      found: !!starter,
      starterSlackUserId: starter ? starter.slackUserId : "",
      durationMs: Date.now() - startMs
    });
    return starter;
  } catch (e) {
    LOGGER.error("selectStarterForWeek_", e.message, e.stack);
    return null;
  }
}

function buildSlackMessage_(team, mondayDate, nomination, starter) {
  LOGGER.debug("START", { action: "buildSlackMessage_", team: team, mondayDate: mondayDate, nomination: nomination, starter: starter });
  const startMs = Date.now();
  try {
    if (!nomination || !nomination.primary) {
      LOGGER.error("buildSlackMessage_", "Invalid nomination", "");
      return "Aucun lead disponible pour cette semaine.";
    }

    const visioLink = String(CONFIG.get("LIEN_VISIO_WEEKLY", CONFIG.get("WEEKLY_MEET_URL", "")) || "").trim();
    const latestSlideLink = getWeeklySlideLinkByMondayDate_(mondayDate);
    const lienRemote = String(CONFIG.get("LIEN_REMOTE", "") || "").trim();
    const lines = [
      "*Weekly kick-off - Qui fait le CR ?*",
      "Hello",
      "",
      "Ceci est un message automatique pour savoir qui s'occupera du CR du Weekly Kick Off.",
      "",
      "Le code de la remote sera publié quelques minutes avant le Weekly dans un autre message sur ce même channel.",
      "",
      "Lead désigné: " + formatSlackUserMention_(nomination.primary.slackUserId),
      "Personne qui lance le Weekly: " + (starter ? formatSlackUserMention_(starter.slackUserId) : "Aucune personne disponible"),
      "Slide hebdo: " + (latestSlideLink || "Non trouve"),
      "Visio weekly: " + (visioLink || "Non trouve"),
      "Lien remote: " + (lienRemote || "Non trouve"),
      "",
      "Bonne journée"
    ];

    const text = lines.join("\n");

    LOGGER.info("SUCCESS", { action: "buildSlackMessage_", durationMs: Date.now() - startMs });
    return text;
  } catch (e) {
    LOGGER.error("buildSlackMessage_", e.message, e.stack);
    return "Erreur de génération du message weekly.";
  }
}

function getAbsenceUrlForSlack_() {
  LOGGER.debug("START", { action: "getAbsenceUrlForSlack_" });
  const startMs = Date.now();
  try {
    let url = "";

    if (!url) {
      try {
        const props = PropertiesService.getScriptProperties();
        url = String(props.getProperty("ABSENCE_WEBAPP_URL") || "").trim();
      } catch (ignored1) {
        url = "";
      }
    }

    if (!url) {
      url = String(CONFIG.get("ABSENCE_WEBAPP_URL", "") || "").trim();
    }

    if (!url) {
      try {
        if (typeof getAbsenceWebAppUrl === "function") {
          url = String(getAbsenceWebAppUrl() || "").trim();
        }
      } catch (ignored2) {
        url = "";
      }
    }

    if (!url) {
      try {
        const serviceUrl = String(ScriptApp.getService().getUrl() || "").trim();
        if (serviceUrl) {
          url = serviceUrl;
        }
      } catch (ignored3) {
        url = "";
      }
    }

    if (!url) {
      LOGGER.info("SUCCESS", { action: "getAbsenceUrlForSlack_", hasUrl: false, durationMs: Date.now() - startMs });
      return "";
    }

    if (/\/dev(?:[/?]|$)/.test(url)) {
      url = url.replace("/dev", "/exec");
    }
    if (!/[?&]page=absence(?:&|$)/.test(url)) {
      url += url.indexOf("?") === -1 ? "?page=absence" : "&page=absence";
    }
    LOGGER.info("SUCCESS", { action: "getAbsenceUrlForSlack_", hasUrl: true, durationMs: Date.now() - startMs });
    return url;
  } catch (e) {
    LOGGER.error("getAbsenceUrlForSlack_", e.message, e.stack);
    return "";
  }
}

function getWeeklySlideLinkByMondayDate_(mondayDate) {
  LOGGER.debug("START", { action: "getWeeklySlideLinkByMondayDate_", mondayDate: mondayDate });
  const startMs = Date.now();
  try {
    const folderId = String(CONFIG.get("WEEKLY_SLIDES_FOLDER_ID", "") || "").trim();
    if (!folderId) {
      LOGGER.info("SUCCESS", { action: "getWeeklySlideLinkByMondayDate_", found: false, reason: "missing_folder_id", durationMs: Date.now() - startMs });
      return "";
    }

    const tz = String(CONFIG.get("TIMEZONE", "Europe/Paris"));
    const targetToken = Utilities.formatDate(mondayDate, tz, "yyyyMMdd");
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByType(MimeType.GOOGLE_SLIDES);
    let matchedFile = null;
    let matchedUpdated = 0;

    while (files.hasNext()) {
      const file = files.next();
      const fileName = String(file.getName() || "");
      const normalizedName = fileName.replace(/[^0-9]/g, "");
      const matchesDateToken = fileName.indexOf(targetToken) !== -1 || normalizedName.indexOf(targetToken) !== -1;
      if (!matchesDateToken) {
        continue;
      }
      const updated = file.getLastUpdated().getTime();
      if (updated > matchedUpdated) {
        matchedUpdated = updated;
        matchedFile = file;
      }
    }

    if (!matchedFile) {
      LOGGER.info("SUCCESS", {
        action: "getWeeklySlideLinkByMondayDate_",
        found: false,
        reason: "no_matching_slide_for_date",
        targetToken: targetToken,
        durationMs: Date.now() - startMs
      });
      return "";
    }

    const link = "<" + matchedFile.getUrl() + "|" + matchedFile.getName() + ">";
    LOGGER.info("SUCCESS", {
      action: "getWeeklySlideLinkByMondayDate_",
      found: true,
      fileId: matchedFile.getId(),
      targetToken: targetToken,
      durationMs: Date.now() - startMs
    });
    return link;
  } catch (e) {
    LOGGER.error("getWeeklySlideLinkByMondayDate_", e.message, e.stack);
    return "";
  }
}

function shouldSkipWeeklyByCalendar_(mondayDate) {
  LOGGER.debug("START", { action: "shouldSkipWeeklyByCalendar_", mondayDate: mondayDate });
  const startMs = Date.now();
  try {
    const tz = String(CONFIG.get("TIMEZONE", "Europe/Paris"));
    const isoDate = Utilities.formatDate(mondayDate, tz, "yyyy-MM-dd");

    const indyOffDates = parseDateCsv_(String(CONFIG.get("JOUR_OFF_INDY", "")));
    if (indyOffDates.indexOf(isoDate) !== -1) {
      const reasonOff = "JOUR_OFF_INDY " + isoDate;
      LOGGER.info("SUCCESS", { action: "shouldSkipWeeklyByCalendar_", skip: true, reason: reasonOff, durationMs: Date.now() - startMs });
      return { skip: true, reason: reasonOff };
    }

    const holiday = getFrenchHolidayNameForDate_(mondayDate);
    if (holiday) {
      const reasonHoliday = "JOUR_FERIE " + isoDate + " (" + holiday + ")";
      LOGGER.info("SUCCESS", { action: "shouldSkipWeeklyByCalendar_", skip: true, reason: reasonHoliday, durationMs: Date.now() - startMs });
      return { skip: true, reason: reasonHoliday };
    }

    LOGGER.info("SUCCESS", { action: "shouldSkipWeeklyByCalendar_", skip: false, durationMs: Date.now() - startMs });
    return { skip: false, reason: "" };
  } catch (e) {
    LOGGER.error("shouldSkipWeeklyByCalendar_", e.message, e.stack);
    return { skip: false, reason: "" };
  }
}

function getFrenchHolidayNameForDate_(targetDate) {
  LOGGER.debug("START", { action: "getFrenchHolidayNameForDate_", targetDate: targetDate });
  const startMs = Date.now();
  try {
    const apiConfig = String(CONFIG.get("JOUR_FERIES_FRANCAIS_API", "") || "").trim();
    const disabled = !apiConfig || apiConfig.toLowerCase() === "false" || apiConfig === "0";
    if (disabled) {
      LOGGER.info("SUCCESS", { action: "getFrenchHolidayNameForDate_", enabled: false, durationMs: Date.now() - startMs });
      return "";
    }

    const tz = String(CONFIG.get("TIMEZONE", "Europe/Paris"));
    const year = Utilities.formatDate(targetDate, tz, "yyyy");
    const isoDate = Utilities.formatDate(targetDate, tz, "yyyy-MM-dd");
    const defaultTemplate = "https://calendrier.api.gouv.fr/jours-feries/metropole/{year}.json";
    const template = (apiConfig.toLowerCase() === "true" || apiConfig.toLowerCase() === "auto") ? defaultTemplate : apiConfig;
    const url = template.indexOf("{year}") !== -1 ? template.replace("{year}", year) : template;
    const cacheKey = "fr_holidays_" + year;
    const cache = CacheService.getScriptCache();
    const cached = cache.get(cacheKey);
    let body = null;
    if (cached) {
      try {
        body = JSON.parse(cached);
      } catch (ignored) {
        body = null;
      }
    }
    if (!body) {
      const response = UrlFetchApp.fetch(url, { method: "get", muteHttpExceptions: true });
      if (response.getResponseCode() !== 200) {
        LOGGER.error("getFrenchHolidayNameForDate_", "HTTP " + response.getResponseCode() + " on " + url, "");
        return "";
      }
      body = JSON.parse(response.getContentText() || "{}");
      try {
        cache.put(cacheKey, JSON.stringify(body), 21600);
      } catch (ignored2) {
        // best-effort cache only
      }
    }
    const holidayName = String(body[isoDate] || "");
    LOGGER.info("SUCCESS", { action: "getFrenchHolidayNameForDate_", date: isoDate, holiday: holidayName, durationMs: Date.now() - startMs });
    return holidayName;
  } catch (e) {
    LOGGER.error("getFrenchHolidayNameForDate_", e.message, e.stack);
    return "";
  }
}

function getManagersMentionsForTeam_(team) {
  LOGGER.debug("START", { action: "getManagersMentionsForTeam_", team: team });
  const startMs = Date.now();
  try {
    const groupKey = team === "care" ? "MANAGERS_SLACK_GROUP_IDS_CARE" : "MANAGERS_SLACK_GROUP_IDS_SALES";
    const groupIds = parseTagCsv_(String(CONFIG.get(groupKey, "")));
    if (groupIds.length > 0) {
      const groupMentions = groupIds.map(function(groupId) {
        return formatSlackGroupMention_(groupId);
      }).filter(function(item) {
        return item !== "";
      }).join(" ");
      LOGGER.info("SUCCESS", { action: "getManagersMentionsForTeam_", mode: "group", count: groupIds.length, durationMs: Date.now() - startMs });
      return groupMentions;
    }

    const userIds = parseTagCsv_(String(CONFIG.get("MANAGERS_SLACK_USER_IDS", "")));
    const userMentions = userIds.map(function(userId) {
      return formatSlackUserMention_(userId);
    }).filter(function(item) {
      return item !== "";
    }).join(" ");
    LOGGER.info("SUCCESS", { action: "getManagersMentionsForTeam_", mode: "user_fallback", count: userIds.length, durationMs: Date.now() - startMs });
    return userMentions;
  } catch (e) {
    LOGGER.error("getManagersMentionsForTeam_", e.message, e.stack);
    return "";
  }
}

function saveHistory_(mondayDate, team, nomination, starter, status, reason, runMode, runWindowReason) {
  LOGGER.debug("START", {
    action: "saveHistory_",
    mondayDate: mondayDate,
    team: team,
    status: status,
    reason: reason,
    starter: starter,
    runMode: runMode,
    runWindowReason: runWindowReason
  });
  const startMs = Date.now();
  try {
    const sheet = getOrCreateSheet_(SHEET_NAMES.history);
    ensureHistoryExtendedColumns_(sheet);
    const primaryName = nomination && nomination.primary ? nomination.primary.name : "";
    const primarySlackUserId = nomination && nomination.primary ? nomination.primary.slackUserId : "";
    const backupName = nomination && nomination.backup ? nomination.backup.name : "";
    const backupSlackUserId = nomination && nomination.backup ? nomination.backup.slackUserId : "";
    const starterName = starter ? starter.name : "";
    const starterSlackUserId = starter ? starter.slackUserId : "";
    const user = Session.getActiveUser().getEmail() || "unknown";
    const payload = JSON.stringify({
      nomination: nomination || {},
      starter: starter || null
    });

    sheet.appendRow([
      new Date(),
      mondayDate,
      team,
      primaryName,
      primarySlackUserId,
      backupName,
      backupSlackUserId,
      starterName,
      starterSlackUserId,
      status,
      reason,
      user,
      payload,
      String(runMode || ""),
      String(runWindowReason || "")
    ]);

    LOGGER.info("SUCCESS", { action: "saveHistory_", durationMs: Date.now() - startMs });
    return true;
  } catch (e) {
    LOGGER.error("saveHistory_", e.message, e.stack);
    return false;
  }
}

function getCurrentMonday_() {
  LOGGER.debug("START", { action: "getCurrentMonday_" });
  const startMs = Date.now();
  try {
    const monday = getMondayForDate_(new Date());
    LOGGER.info("SUCCESS", { action: "getCurrentMonday_", monday: monday, durationMs: Date.now() - startMs });
    return monday;
  } catch (e) {
    LOGGER.error("getCurrentMonday_", e.message, e.stack);
    return stripTime_(new Date());
  }
}

function getMondayForDate_(dateValue) {
  LOGGER.debug("START", { action: "getMondayForDate_", dateValue: dateValue });
  const startMs = Date.now();
  try {
    const tz = String(CONFIG.get("TIMEZONE", "Europe/Paris"));
    const baseDate = dateValue instanceof Date && !isNaN(dateValue.getTime()) ? dateValue : new Date();
    const todayInTz = new Date(Utilities.formatDate(baseDate, tz, "yyyy-MM-dd'T'00:00:00"));
    const day = todayInTz.getDay();
    const offset = day === 0 ? -6 : 1 - day;
    const monday = new Date(todayInTz);
    monday.setDate(todayInTz.getDate() + offset);
    const result = stripTime_(monday);
    LOGGER.info("SUCCESS", { action: "getMondayForDate_", monday: result, durationMs: Date.now() - startMs });
    return result;
  } catch (e) {
    LOGGER.error("getMondayForDate_", e.message, e.stack);
    return stripTime_(new Date());
  }
}

function debugLuccaAbsenceJuliette() {
  return debugLuccaAbsenceForEmail("juliette@indy.fr");
}

function debugLuccaAbsenceForEmail(email, mondayDateStr) {
  LOGGER.debug("START", { action: "debugLuccaAbsenceForEmail", email: email, mondayDateStr: mondayDateStr });
  const startMs = Date.now();
  try {
    const cleanEmail = String(email || "").toLowerCase().trim();
    if (!cleanEmail) {
      LOGGER.error("debugLuccaAbsenceForEmail", "Missing email", "");
      return null;
    }

    const timezone = String(CONFIG.get("TIMEZONE", "Europe/Paris"));
    let mondayDate = mondayDateStr ? parseConfigDate_(mondayDateStr) : getCurrentMonday_();
    if (!(mondayDate instanceof Date)) {
      mondayDate = getCurrentMonday_();
    }
    mondayDate = stripTime_(mondayDate);
    const mondayYmd = Utilities.formatDate(mondayDate, timezone, "yyyy-MM-dd");

    const luccaCfg = getLuccaConfig_();
    const headers = {
      Authorization: "Lucca application=" + String(luccaCfg.token || ""),
      Accept: "application/json"
    };
    const luccaUser = luccaCfg.baseUrl && luccaCfg.token
      ? fetchLuccaUserByEmail_(luccaCfg.baseUrl, headers, cleanEmail)
      : null;
    const luccaUserId = normalizeLuccaUserId_(luccaUser ? luccaUser.id : "");

    const absData = getAbsenceDataForDate_(mondayDate);
    const lookup = buildAbsenceLookup_(absData);

    let leavesCountForDay = 0;
    let leavesCountForUser = 0;
    if (luccaCfg.baseUrl && luccaCfg.token) {
      const refs = getPagedItemsLucca_(
        luccaCfg.baseUrl,
        "/api/v3/leaves",
        {
          date: "between," + mondayYmd + "," + mondayYmd,
          "leavePeriod.ownerId": "notequal,0"
        },
        headers,
        200
      );
      leavesCountForDay = refs.length;
      if (luccaUserId) {
        for (let i = 0; i < refs.length; i++) {
          if (normalizeLuccaUserId_(extractLeaveOwnerId_(refs[i])) === luccaUserId) {
            leavesCountForUser += 1;
          }
        }
      }
    }

    const rotations = []
      .concat(getCareRotation_().map(function(p) { return { sheet: SHEET_NAMES.rotationCare, person: p }; }))
      .concat(getSalesRotation_().map(function(p) { return { sheet: SHEET_NAMES.rotationSales, person: p }; }))
      .concat(getStartRotation_().map(function(p) { return { sheet: SHEET_NAMES.rotationStart, person: p }; }));

    const matches = rotations.filter(function(row) {
      return String(row.person.email || "").toLowerCase().trim() === cleanEmail;
    }).map(function(row) {
      const person = row.person;
      return {
        sheet: row.sheet,
        name: person.name,
        email: person.email,
        slackUserId: person.slackUserId,
        luccaUserId: normalizeLuccaUserId_(person.luccaUserId),
        active: !!person.active,
        absentByEngine: isPersonAbsent_(person, lookup)
      };
    });

    const result = {
      email: cleanEmail,
      monday: mondayYmd,
      luccaConfigured: !!(luccaCfg.baseUrl && luccaCfg.token),
      luccaUserFound: !!luccaUser,
      luccaUserId: luccaUserId,
      leavesCountForDay: leavesCountForDay,
      leavesCountForUser: leavesCountForUser,
      absenceDataCounts: {
        emails: (absData.emails || []).length,
        luccaUserIds: (absData.luccaUserIds || []).length,
        names: (absData.names || []).length
      },
      emailInAbsenceData: !!lookup.emails[cleanEmail],
      luccaIdInAbsenceData: !!(luccaUserId && lookup.luccaUserIds[luccaUserId]),
      rotationMatches: matches
    };

    LOGGER.info("SUCCESS", { action: "debugLuccaAbsenceForEmail", result: result, durationMs: Date.now() - startMs });
    console.log("DEBUG_LUCCA_ABSENCE_RESULT:\n" + JSON.stringify(result, null, 2));
    return result;
  } catch (e) {
    LOGGER.error("debugLuccaAbsenceForEmail", e.message, e.stack);
    return null;
  }
}

function rotateArrayFromIndex_(array, startIndex) {
  LOGGER.debug("START", { action: "rotateArrayFromIndex_", size: array ? array.length : 0, startIndex: startIndex });
  const startMs = Date.now();
  try {
    if (!Array.isArray(array) || array.length === 0) {
      LOGGER.info("SUCCESS", { action: "rotateArrayFromIndex_", size: 0, durationMs: Date.now() - startMs });
      return [];
    }
    const idx = ((startIndex % array.length) + array.length) % array.length;
    const rotated = array.slice(idx).concat(array.slice(0, idx));
    LOGGER.info("SUCCESS", { action: "rotateArrayFromIndex_", size: rotated.length, durationMs: Date.now() - startMs });
    return rotated;
  } catch (e) {
    LOGGER.error("rotateArrayFromIndex_", e.message, e.stack);
    return [];
  }
}

function splitCsvSafe_(value) {
  LOGGER.debug("START", { action: "splitCsvSafe_", value: value });
  const startMs = Date.now();
  try {
    if (!value) {
      LOGGER.info("SUCCESS", { action: "splitCsvSafe_", count: 0, durationMs: Date.now() - startMs });
      return [];
    }
    const rows = String(value).split(",").map(function(item) {
      return item.trim();
    }).filter(function(item) {
      return item !== "";
    });
    LOGGER.info("SUCCESS", { action: "splitCsvSafe_", count: rows.length, durationMs: Date.now() - startMs });
    return rows;
  } catch (e) {
    LOGGER.error("splitCsvSafe_", e.message, e.stack);
    return [];
  }
}

function parseTagCsv_(value) {
  LOGGER.debug("START", { action: "parseTagCsv_", value: value });
  const startMs = Date.now();
  try {
    const tags = splitCsvSafe_(value).map(function(item) {
      return String(item).trim();
    }).filter(function(item) {
      return item !== "";
    });
    LOGGER.info("SUCCESS", { action: "parseTagCsv_", count: tags.length, durationMs: Date.now() - startMs });
    return tags;
  } catch (e) {
    LOGGER.error("parseTagCsv_", e.message, e.stack);
    return [];
  }
}

function parseDateCsv_(value) {
  LOGGER.debug("START", { action: "parseDateCsv_", value: value });
  const startMs = Date.now();
  try {
    const dates = splitCsvSafe_(value).map(function(item) {
      return String(item).trim();
    }).filter(function(item) {
      return /^\d{4}-\d{2}-\d{2}$/.test(item);
    });
    LOGGER.info("SUCCESS", { action: "parseDateCsv_", count: dates.length, durationMs: Date.now() - startMs });
    return dates;
  } catch (e) {
    LOGGER.error("parseDateCsv_", e.message, e.stack);
    return [];
  }
}

function parseBoolean_(value, defaultValue) {
  LOGGER.debug("START", { action: "parseBoolean_", value: value, defaultValue: defaultValue });
  const startMs = Date.now();
  try {
    if (value === true || value === false) {
      LOGGER.info("SUCCESS", { action: "parseBoolean_", parsed: value, durationMs: Date.now() - startMs });
      return value;
    }
    const normalized = String(value || "").trim().toLowerCase();
    if (normalized === "true" || normalized === "1" || normalized === "yes") {
      LOGGER.info("SUCCESS", { action: "parseBoolean_", parsed: true, durationMs: Date.now() - startMs });
      return true;
    }
    if (normalized === "false" || normalized === "0" || normalized === "no") {
      LOGGER.info("SUCCESS", { action: "parseBoolean_", parsed: false, durationMs: Date.now() - startMs });
      return false;
    }
    LOGGER.info("SUCCESS", { action: "parseBoolean_", parsed: defaultValue, durationMs: Date.now() - startMs });
    return defaultValue;
  } catch (e) {
    LOGGER.error("parseBoolean_", e.message, e.stack);
    return defaultValue;
  }
}

function formatSlackUserMention_(slackUserId) {
  LOGGER.debug("START", { action: "formatSlackUserMention_", slackUserId: slackUserId });
  const startMs = Date.now();
  try {
    const cleanId = String(slackUserId || "").trim();
    if (!cleanId) {
      LOGGER.info("SUCCESS", { action: "formatSlackUserMention_", mention: "", durationMs: Date.now() - startMs });
      return "";
    }
    const mention = cleanId.indexOf("<@") === 0 ? cleanId : "<@" + cleanId + ">";
    LOGGER.info("SUCCESS", { action: "formatSlackUserMention_", mention: mention, durationMs: Date.now() - startMs });
    return mention;
  } catch (e) {
    LOGGER.error("formatSlackUserMention_", e.message, e.stack);
    return "";
  }
}

function formatSlackChannelMention_(slackChannelId) {
  LOGGER.debug("START", { action: "formatSlackChannelMention_", slackChannelId: slackChannelId });
  const startMs = Date.now();
  try {
    const cleanId = String(slackChannelId || "").trim();
    if (!cleanId) {
      LOGGER.info("SUCCESS", { action: "formatSlackChannelMention_", mention: "", durationMs: Date.now() - startMs });
      return "";
    }
    const mention = cleanId.indexOf("<#") === 0 ? cleanId : "<#" + cleanId + ">";
    LOGGER.info("SUCCESS", { action: "formatSlackChannelMention_", mention: mention, durationMs: Date.now() - startMs });
    return mention;
  } catch (e) {
    LOGGER.error("formatSlackChannelMention_", e.message, e.stack);
    return "";
  }
}

function formatSlackGroupMention_(groupId) {
  LOGGER.debug("START", { action: "formatSlackGroupMention_", groupId: groupId });
  const startMs = Date.now();
  try {
    const cleanId = String(groupId || "").trim();
    if (!cleanId) {
      LOGGER.info("SUCCESS", { action: "formatSlackGroupMention_", mention: "", durationMs: Date.now() - startMs });
      return "";
    }
    if (cleanId.indexOf("<!subteam^") === 0) {
      LOGGER.info("SUCCESS", { action: "formatSlackGroupMention_", mention: cleanId, durationMs: Date.now() - startMs });
      return cleanId;
    }
    const mention = "<!subteam^" + cleanId + ">";
    LOGGER.info("SUCCESS", { action: "formatSlackGroupMention_", mention: mention, durationMs: Date.now() - startMs });
    return mention;
  } catch (e) {
    LOGGER.error("formatSlackGroupMention_", e.message, e.stack);
    return "";
  }
}

function formatSlackLink_(url, label) {
  LOGGER.debug("START", { action: "formatSlackLink_", url: url, label: label });
  const startMs = Date.now();
  try {
    const cleanLabel = String(label || "lien").trim() || "lien";
    let cleanUrl = String(url || "").trim();
    if (!cleanUrl) {
      LOGGER.info("SUCCESS", { action: "formatSlackLink_", hasUrl: false, durationMs: Date.now() - startMs });
      return cleanLabel;
    }
    if (cleanUrl.indexOf("<") === 0 && cleanUrl.lastIndexOf(">") === cleanUrl.length - 1) {
      cleanUrl = cleanUrl.slice(1, -1);
    }
    if (cleanUrl.indexOf("|") !== -1) {
      cleanUrl = cleanUrl.split("|")[0];
    }
    const formatted = "<" + cleanUrl + "|" + cleanLabel + ">";
    LOGGER.info("SUCCESS", { action: "formatSlackLink_", hasUrl: true, durationMs: Date.now() - startMs });
    return formatted;
  } catch (e) {
    LOGGER.error("formatSlackLink_", e.message, e.stack);
    return String(label || "lien");
  }
}

function stripTime_(dateValue) {
  LOGGER.debug("START", { action: "stripTime_", dateValue: dateValue });
  const startMs = Date.now();
  try {
    const d = new Date(dateValue);
    d.setHours(0, 0, 0, 0);
    LOGGER.info("SUCCESS", { action: "stripTime_", durationMs: Date.now() - startMs });
    return d;
  } catch (e) {
    LOGGER.error("stripTime_", e.message, e.stack);
    return new Date();
  }
}
