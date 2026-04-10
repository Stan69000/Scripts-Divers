function weeklyKickoffRun() {
  LOGGER.debug("START", { action: "weeklyKickoffRun" });
  const startMs = Date.now();
  try {
    const mondayDate = getCurrentMonday_();
    const team = getTeamForWeek_(mondayDate);
    const skipRule = shouldSkipWeeklyByCalendar_(mondayDate);
    if (skipRule.skip) {
      saveHistory_(mondayDate, team, null, null, "SKIPPED", skipRule.reason);
      LOGGER.info("SUCCESS", { action: "weeklyKickoffRun", status: "SKIPPED", reason: skipRule.reason, durationMs: Date.now() - startMs });
      return { status: "SKIPPED", reason: skipRule.reason };
    }

    const nomination = selectLeadForTeam_(team, mondayDate);
    const starter = selectStarterForWeek_(mondayDate);
    if (!nomination || !nomination.primary) {
      LOGGER.error("weeklyKickoffRun", "No nomination available", "");
      saveHistory_(mondayDate, team, nomination, starter, "ERROR", "No nomination");
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
    saveHistory_(mondayDate, team, nomination, starter, status, "Slack code=" + code);

    LOGGER.info("SUCCESS", { action: "weeklyKickoffRun", status: status, durationMs: Date.now() - startMs });
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
    const mondayDate = getCurrentMonday_();
    const team = getTeamForWeek_(mondayDate);
    const skipRule = shouldSkipWeeklyByCalendar_(mondayDate);
    if (skipRule.skip) {
      const skipMessage = "Dry run: aucun message Slack, weekly saute (" + skipRule.reason + ").";
      console.log("DRY_RUN_MESSAGE:\n" + skipMessage);
      LOGGER.info("SUCCESS", { action: "dryRunWeeklyKickoff", skipped: true, reason: skipRule.reason, durationMs: Date.now() - startMs });
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
    const absentEmails = getAbsentEmailsForDate_(mondayDate);
    const mainSelection = selectFromTeamWithAbsences_(team, absentEmails);
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
      const fallbackSelection = selectFromTeamWithAbsences_("care", absentEmails);
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

function selectFromTeamWithAbsences_(team, absentEmails) {
  LOGGER.debug("START", { action: "selectFromTeamWithAbsences_", team: team, absentEmailsCount: absentEmails ? absentEmails.length : 0 });
  const startMs = Date.now();
  try {
    const candidates = getRotationForTeam_(team);
    if (!candidates || candidates.length === 0) {
      LOGGER.info("SUCCESS", { action: "selectFromTeamWithAbsences_", team: team, hasCandidates: false, durationMs: Date.now() - startMs });
      return { primary: null, backup: null, candidates: [], available: [] };
    }

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
      return item.active && absentEmails.indexOf(item.email.toLowerCase()) === -1;
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
      if (!name || !slackUserId || !email) {
        continue;
      }
      rows.push({
        order: isNaN(order) ? 9999 : order,
        name: name,
        slackUserId: slackUserId,
        email: email,
        active: active
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

    const source = String(CONFIG.get("ABSENCE_SOURCE", "GSHEET"));
    if (source === "LUCCA" && isLuccaEnabled_()) {
      const luccaRows = fetchLuccaAbsences_(targetDate);
      LOGGER.info("SUCCESS", { action: "getAbsentEmailsForDate_", source: source, count: luccaRows.length, durationMs: Date.now() - startMs });
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

    LOGGER.info("SUCCESS", { action: "getAbsentEmailsForDate_", source: "GSHEET", count: absent.length, durationMs: Date.now() - startMs });
    return absent;
  } catch (e) {
    LOGGER.error("getAbsentEmailsForDate_", e.message, e.stack);
    return [];
  }
}

function fetchLuccaAbsences_(targetDate) {
  LOGGER.debug("START", { action: "fetchLuccaAbsences_", targetDate: targetDate });
  const startMs = Date.now();
  try {
    const baseUrl = getScriptProperty_("LUCCA_BASE_URL", "");
    const token = getScriptProperty_("LUCCA_API_TOKEN", "");
    if (!baseUrl || !token) {
      LOGGER.info("SUCCESS", { action: "fetchLuccaAbsences_", reason: "missing_config", durationMs: Date.now() - startMs });
      return [];
    }

    const dateString = Utilities.formatDate(targetDate, String(CONFIG.get("TIMEZONE", "Europe/Paris")), "yyyy-MM-dd");
    const url = baseUrl + "/api/v3/leaves?date=" + encodeURIComponent(dateString);
    const options = {
      method: "get",
      headers: {
        Authorization: "lucca application=" + token
      },
      muteHttpExceptions: true
    };
    const res = UrlFetchApp.fetch(url, options);
    if (res.getResponseCode() !== 200) {
      LOGGER.error("fetchLuccaAbsences_", "Lucca HTTP " + res.getResponseCode(), "");
      return [];
    }

    const body = JSON.parse(res.getContentText());
    const emails = (body.data || []).map(function(row) {
      return String(row.mail || "").toLowerCase();
    }).filter(function(email) {
      return email !== "";
    });

    LOGGER.info("SUCCESS", { action: "fetchLuccaAbsences_", count: emails.length, durationMs: Date.now() - startMs });
    return emails;
  } catch (e) {
    LOGGER.error("fetchLuccaAbsences_", e.message, e.stack);
    return [];
  }
}

function isLuccaEnabled_() {
  LOGGER.debug("START", { action: "isLuccaEnabled_" });
  const startMs = Date.now();
  try {
    const enabled = getScriptProperty_("LUCCA_ENABLED", "false") === "true";
    LOGGER.info("SUCCESS", { action: "isLuccaEnabled_", enabled: enabled, durationMs: Date.now() - startMs });
    return enabled;
  } catch (e) {
    LOGGER.error("isLuccaEnabled_", e.message, e.stack);
    return false;
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

    const absentEmails = getAbsentEmailsForDate_(mondayDate);
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
      return item.active && absentEmails.indexOf(item.email.toLowerCase()) === -1;
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

    const teamForMessage = nomination && nomination.teamUsed ? nomination.teamUsed : team;
    const managers = getManagersMentionsForTeam_(teamForMessage);
    const visioLink = String(CONFIG.get("LIEN_VISIO_WEEKLY", CONFIG.get("WEEKLY_MEET_URL", "")) || "").trim();
    const latestSlideLink = getWeeklySlideLinkByMondayDate_(mondayDate);
    const lienRemote = String(CONFIG.get("LIEN_REMOTE", "") || "").trim();
    const absenceUrl = getAbsenceUrlForSlack_();
    const absenceCta = absenceUrl ? formatSlackLink_(absenceUrl, "ici") : "ici";
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
      "Managers pour info : " + (managers || "Non configurés"),
      "Slide hebdo: " + (latestSlideLink || "Non trouve"),
      "Visio weekly: " + (visioLink || "Non trouve"),
      "Lien remote: " + (lienRemote || "Non trouve"),
      "Si tu es en conges un lundi prochainement, pense a le noter " + absenceCta,
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

    const response = UrlFetchApp.fetch(url, { method: "get", muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) {
      LOGGER.error("getFrenchHolidayNameForDate_", "HTTP " + response.getResponseCode() + " on " + url, "");
      return "";
    }

    const body = JSON.parse(response.getContentText());
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

function saveHistory_(mondayDate, team, nomination, starter, status, reason) {
  LOGGER.debug("START", { action: "saveHistory_", mondayDate: mondayDate, team: team, status: status, reason: reason, starter: starter });
  const startMs = Date.now();
  try {
    const sheet = getOrCreateSheet_(SHEET_NAMES.history);
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
      payload
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
    const tz = String(CONFIG.get("TIMEZONE", "Europe/Paris"));
    const now = new Date();
    const todayInTz = new Date(Utilities.formatDate(now, tz, "yyyy-MM-dd'T'00:00:00"));
    const day = todayInTz.getDay();
    const offset = day === 0 ? -6 : 1 - day;
    const monday = new Date(todayInTz);
    monday.setDate(todayInTz.getDate() + offset);

    LOGGER.info("SUCCESS", { action: "getCurrentMonday_", monday: monday, durationMs: Date.now() - startMs });
    return stripTime_(monday);
  } catch (e) {
    LOGGER.error("getCurrentMonday_", e.message, e.stack);
    return stripTime_(new Date());
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
