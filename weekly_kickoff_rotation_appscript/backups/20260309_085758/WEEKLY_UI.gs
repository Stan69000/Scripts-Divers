function doGet(e) {
  LOGGER.debug("START", { action: "doGet", e: e });
  const startMs = Date.now();
  try {
    const page = e && e.parameter ? String(e.parameter.page || "start") : "start";
    if (page === "absence") {
      const absenceHtml = buildAbsencePageHtml_();
      const absenceOutput = HtmlService.createHtmlOutput(absenceHtml).setTitle("Absence Weekly");
      LOGGER.info("SUCCESS", { action: "doGet", page: page, known: true, durationMs: Date.now() - startMs });
      return absenceOutput;
    }

    if (page !== "start") {
      const output = HtmlService.createHtmlOutput("<h3>Page inconnue</h3>");
      LOGGER.info("SUCCESS", { action: "doGet", page: page, known: false, durationMs: Date.now() - startMs });
      return output;
    }

    const html = buildStartWeeklyPageHtml_();
    const outputOk = HtmlService.createHtmlOutput(html).setTitle("Start Weekly");
    LOGGER.info("SUCCESS", { action: "doGet", page: page, known: true, durationMs: Date.now() - startMs });
    return outputOk;
  } catch (e2) {
    LOGGER.error("doGet", e2.message, e2.stack);
    return HtmlService.createHtmlOutput("<h3>Erreur de chargement Start Weekly</h3>");
  }
}

function getAbsenceFormContext() {
  LOGGER.debug("START", { action: "getAbsenceFormContext" });
  const startMs = Date.now();
  try {
    const now = new Date();
    const timezone = String(CONFIG.get("TIMEZONE", "Europe/Paris"));
    const todayIso = Utilities.formatDate(now, timezone, "yyyy-MM-dd");
    const userEmail = Session.getActiveUser().getEmail() || "";
    const context = {
      email: userEmail,
      today: todayIso
    };
    LOGGER.info("SUCCESS", { action: "getAbsenceFormContext", hasEmail: !!userEmail, durationMs: Date.now() - startMs });
    return context;
  } catch (e) {
    LOGGER.error("getAbsenceFormContext", e.message, e.stack);
    return { email: "", today: "" };
  }
}

function submitAbsenceRequest(formData) {
  LOGGER.debug("START", { action: "submitAbsenceRequest", formData: formData });
  const startMs = Date.now();
  try {
    if (!formData || typeof formData !== "object") {
      LOGGER.error("submitAbsenceRequest", "Invalid formData", "");
      return { ok: false, message: "Formulaire invalide." };
    }

    const email = String(formData.email || "").trim().toLowerCase();
    const name = String(formData.name || "").trim();
    const rawAbsences = Array.isArray(formData.absences) ? formData.absences : [];
    if (!email) {
      LOGGER.error("submitAbsenceRequest", "Missing email", "");
      return { ok: false, message: "Email obligatoire." };
    }
    if (rawAbsences.length > 5) {
      LOGGER.error("submitAbsenceRequest", "Too many absences", "");
      return { ok: false, message: "Maximum 5 absences par envoi." };
    }

    const sheet = getOrCreateSheet_(SHEET_NAMES.absences);
    if (!sheet) {
      LOGGER.error("submitAbsenceRequest", "ABSENCES sheet unavailable", "");
      return { ok: false, message: "Impossible d'acceder a la feuille ABSENCES." };
    }

    const rowsToInsert = [];
    for (let i = 0; i < rawAbsences.length; i += 1) {
      const item = rawAbsences[i] || {};
      const startDate = String(item.startDate || "").trim();
      const endDate = String(item.endDate || "").trim();
      if (!startDate && !endDate) {
        continue;
      }
      if (!isValidDateIso_(startDate) || !isValidDateIso_(endDate)) {
        LOGGER.error("submitAbsenceRequest", "Invalid dates in row", "");
        return { ok: false, message: "Dates invalides (YYYY-MM-DD) sur la ligne " + (i + 1) + "." };
      }
      if (startDate > endDate) {
        LOGGER.error("submitAbsenceRequest", "startDate after endDate", "");
        return { ok: false, message: "Date debut > date fin sur la ligne " + (i + 1) + "." };
      }
      rowsToInsert.push([email, name, startDate, endDate, "WEBAPP", ""]);
    }

    if (rowsToInsert.length === 0) {
      LOGGER.error("submitAbsenceRequest", "No absence row provided", "");
      return { ok: false, message: "Ajoute au moins une absence." };
    }

    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToInsert.length, 6).setValues(rowsToInsert);
    LOGGER.info("SUCCESS", {
      action: "submitAbsenceRequest",
      email: email,
      count: rowsToInsert.length,
      durationMs: Date.now() - startMs
    });
    return { ok: true, message: rowsToInsert.length + " absence(s) enregistree(s)." };
  } catch (e) {
    LOGGER.error("submitAbsenceRequest", e.message, e.stack);
    return { ok: false, message: "Erreur lors de l'enregistrement." };
  }
}

function getAbsenceWebAppUrl() {
  LOGGER.debug("START", { action: "getAbsenceWebAppUrl" });
  const startMs = Date.now();
  try {
    const url = ScriptApp.getService().getUrl();
    if (!url) {
      LOGGER.info("SUCCESS", { action: "getAbsenceWebAppUrl", hasUrl: false, durationMs: Date.now() - startMs });
      return "";
    }
    const finalUrl = url + "?page=absence";
    LOGGER.info("SUCCESS", { action: "getAbsenceWebAppUrl", hasUrl: true, durationMs: Date.now() - startMs });
    return finalUrl;
  } catch (e) {
    LOGGER.error("getAbsenceWebAppUrl", e.message, e.stack);
    return "";
  }
}

function resolveAbsenceWebAppUrl() {
  LOGGER.debug("START", { action: "resolveAbsenceWebAppUrl" });
  const startMs = Date.now();
  try {
    const runtimeUrl = String(getAbsenceWebAppUrl() || "").trim();
    const props = PropertiesService.getScriptProperties();
    const propsUrl = String(props.getProperty("ABSENCE_WEBAPP_URL") || "").trim();
    const configUrl = String(CONFIG.get("ABSENCE_WEBAPP_URL", "") || "").trim();
    const chosen = runtimeUrl || propsUrl || configUrl;

    if (chosen) {
      const normalized = /[?&]page=absence(?:&|$)/.test(chosen)
        ? chosen
        : chosen + (chosen.indexOf("?") === -1 ? "?page=absence" : "&page=absence");
      LOGGER.info("SUCCESS", {
        action: "resolveAbsenceWebAppUrl",
        found: true,
        source: runtimeUrl ? "runtime" : (propsUrl ? "script_properties" : "config"),
        durationMs: Date.now() - startMs
      });
      console.log("ABSENCE_WEBAPP_URL: " + normalized);
      return normalized;
    }

    const missingMessage = "Aucune URL absences detectee. Verifie que la Web App est deployee (Deploy > Manage deployments).";
    LOGGER.info("SUCCESS", {
      action: "resolveAbsenceWebAppUrl",
      found: false,
      durationMs: Date.now() - startMs
    });
    console.log(missingMessage);
    return "";
  } catch (e) {
    LOGGER.error("resolveAbsenceWebAppUrl", e.message, e.stack);
    return "";
  }
}

function initAbsenceWebAppUrl() {
  LOGGER.debug("START", { action: "initAbsenceWebAppUrl" });
  const startMs = Date.now();
  try {
    const url = String(resolveAbsenceWebAppUrl() || "").trim();
    if (!url) {
      LOGGER.info("SUCCESS", {
        action: "initAbsenceWebAppUrl",
        saved: false,
        reason: "missing_url",
        durationMs: Date.now() - startMs
      });
      return { ok: false, message: "URL absences introuvable. Deploie/maj la Web App puis relance." };
    }

    const props = PropertiesService.getScriptProperties();
    props.setProperty("ABSENCE_WEBAPP_URL", url);
    CONFIG.set("ABSENCE_WEBAPP_URL", url);
    LOGGER.info("SUCCESS", { action: "initAbsenceWebAppUrl", saved: true, durationMs: Date.now() - startMs });
    console.log("ABSENCE_WEBAPP_URL enregistree: " + url);
    return { ok: true, url: url };
  } catch (e) {
    LOGGER.error("initAbsenceWebAppUrl", e.message, e.stack);
    return { ok: false, message: "Erreur initAbsenceWebAppUrl." };
  }
}

function getStartWeeklyContext() {
  LOGGER.debug("START", { action: "getStartWeeklyContext" });
  const startMs = Date.now();
  try {
    const mondayDate = getCurrentMonday_();
    const team = getTeamForWeek_(mondayDate);
    const skipRule = shouldSkipWeeklyByCalendar_(mondayDate);
    const meetUrl = getWeeklyVisioUrl_();
    const remoteBaseUrl = String(CONFIG.get("LIEN_REMOTE", "") || "").trim();
    let absenceUrl = "";
    try {
      if (typeof getAbsenceUrlForSlack_ === "function") {
        absenceUrl = String(getAbsenceUrlForSlack_() || "").trim();
      } else {
        absenceUrl = String(getAbsenceWebAppUrl() || "").trim();
      }
    } catch (ignored) {
      absenceUrl = String(getAbsenceWebAppUrl() || "").trim();
    }
    const timezone = String(CONFIG.get("TIMEZONE", "Europe/Paris"));
    const mondayIso = Utilities.formatDate(mondayDate, timezone, "yyyy-MM-dd");

    const result = {
      mondayDate: mondayIso,
      team: team,
      skip: skipRule.skip,
      skipReason: skipRule.reason,
      previewRemoteMessage: buildRemoteCodeMessage_(mondayIso, "", remoteBaseUrl, meetUrl),
      meetUrl: meetUrl,
      slideUrl: "",
      remoteBaseUrl: remoteBaseUrl,
      absenceUrl: absenceUrl
    };
    LOGGER.info("SUCCESS", { action: "getStartWeeklyContext", skip: skipRule.skip, durationMs: Date.now() - startMs });
    return result;
  } catch (e) {
    LOGGER.error("getStartWeeklyContext", e.message, e.stack);
    return {
      mondayDate: "",
      team: "",
      skip: false,
      skipReason: "",
      previewRemoteMessage: "",
      meetUrl: "",
      slideUrl: "",
      remoteBaseUrl: "",
      absenceUrl: ""
    };
  }
}

function getWeeklySlideUrlForUi() {
  LOGGER.debug("START", { action: "getWeeklySlideUrlForUi" });
  const startMs = Date.now();
  try {
    const mondayDate = getCurrentMonday_();
    const slideSlackLink = getWeeklySlideLinkByMondayDate_(mondayDate);
    const slideUrl = extractUrlFromSlackLink_(slideSlackLink);
    LOGGER.info("SUCCESS", { action: "getWeeklySlideUrlForUi", hasUrl: !!slideUrl, durationMs: Date.now() - startMs });
    return slideUrl;
  } catch (e) {
    LOGGER.error("getWeeklySlideUrlForUi", e.message, e.stack);
    return "";
  }
}

function publishRemoteCode(remoteCode) {
  LOGGER.debug("START", { action: "publishRemoteCode", remoteCode: remoteCode });
  const startMs = Date.now();
  try {
    const safeCode = String(remoteCode || "").trim();
    if (!safeCode) {
      LOGGER.error("publishRemoteCode", "Missing remote code", "");
      return { ok: false, reason: "missing_remote_code", meetText: "", slackCode: null };
    }

    const context = getStartWeeklyContext();
    const remoteBaseUrl = String(context.remoteBaseUrl || "");
    const visioUrl = String(context.meetUrl || "").trim();
    const message = buildRemoteCodeMessage_(context.mondayDate, safeCode, remoteBaseUrl, visioUrl);

    const code = postSlackMessage_({ text: message });
    const meetText = "Weekly remote code: " + safeCode + (remoteBaseUrl ? " | " + remoteBaseUrl : "") + (visioUrl ? " | Visio: " + visioUrl : "");
    const ok = code === 200;
    LOGGER.info("SUCCESS", { action: "publishRemoteCode", ok: ok, slackCode: code, durationMs: Date.now() - startMs });
    return { ok: ok, reason: ok ? "" : "slack_post_failed", meetText: meetText, slackCode: code, previewMessage: message };
  } catch (e) {
    LOGGER.error("publishRemoteCode", e.message, e.stack);
    return { ok: false, reason: "exception", meetText: "", slackCode: null, previewMessage: "" };
  }
}

function getStartWeeklyWebAppUrl() {
  LOGGER.debug("START", { action: "getStartWeeklyWebAppUrl" });
  const startMs = Date.now();
  try {
    const url = ScriptApp.getService().getUrl();
    if (!url) {
      LOGGER.info("SUCCESS", { action: "getStartWeeklyWebAppUrl", hasUrl: false, durationMs: Date.now() - startMs });
      return "";
    }
    const finalUrl = url + "?page=start";
    LOGGER.info("SUCCESS", { action: "getStartWeeklyWebAppUrl", hasUrl: true, durationMs: Date.now() - startMs });
    return finalUrl;
  } catch (e) {
    LOGGER.error("getStartWeeklyWebAppUrl", e.message, e.stack);
    return "";
  }
}

function buildStartWeeklyPageHtml_() {
  LOGGER.debug("START", { action: "buildStartWeeklyPageHtml_" });
  const startMs = Date.now();
  try {
    const html = [
      "<!doctype html><html><head><meta charset='utf-8'><title>Start Weekly</title>",
      "<style>",
      ":root{--bg:#f6f0e6;--ink:#191919;--muted:#5f6166;--navy:#202d3c;--pink:#ef5b86;--card:#fffdf9;--line:#e4dccf;}",
      "*{box-sizing:border-box;}",
      "body{margin:0;font-family:Arial,sans-serif;background:linear-gradient(180deg,#f7f2ea 0%,#f4eee4 100%);color:var(--ink);}",
      ".top{background:var(--navy);color:#fff;padding:12px 20px;font-size:14px;}",
      ".top strong{color:#ffd7e3;}",
      ".brand{display:flex;align-items:center;gap:10px;}",
      ".logo{width:154px;height:62px;border-radius:14px;background:var(--pink);color:#fff;display:flex;align-items:center;justify-content:center;font-weight:900;font-style:italic;font-size:56px;line-height:1;letter-spacing:-2px;box-shadow:0 10px 20px rgba(239,91,134,0.35);padding-bottom:6px;}",
      ".wrap{max-width:1100px;margin:0 auto;padding:18px 18px 28px;}",
      ".hero{display:grid;grid-template-columns:1.2fr 0.8fr;gap:18px;align-items:stretch;}",
      ".card{background:var(--card);border:1px solid var(--line);border-radius:16px;box-shadow:0 12px 32px rgba(32,45,60,0.08);}",
      ".main{padding:22px;}",
      "h1{font-size:44px;line-height:1.03;margin:0 0 10px;font-weight:800;letter-spacing:-0.8px;}",
      ".sub{font-size:18px;color:var(--muted);margin:0 0 18px;}",
      ".actions{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:12px;}",
      ".btn{display:inline-flex;align-items:center;justify-content:center;padding:12px 18px;border-radius:999px;border:1px solid var(--line);text-decoration:none;color:var(--ink);background:#fff;cursor:pointer;font-weight:700;}",
      ".btn.primary{background:var(--pink);border-color:var(--pink);color:#fff;}",
      ".btn.secondary{background:#fff;}",
      ".btn.disabled{opacity:.55;pointer-events:none;cursor:default;}",
      ".meta{display:flex;flex-wrap:wrap;gap:10px;margin-top:8px;font-size:13px;color:var(--muted);}",
      ".kpi{background:#fff;padding:10px 12px;border:1px solid var(--line);border-radius:10px;}",
      ".side{padding:18px;display:flex;flex-direction:column;gap:12px;}",
      ".side h2{margin:0 0 2px;font-size:22px;}",
      ".step{padding:12px;background:#fff;border:1px solid var(--line);border-radius:12px;font-size:15px;line-height:1.35;display:flex;gap:10px;align-items:flex-start;}",
      ".step .num{font-weight:800;color:#1f2f41;min-width:18px;}",
      ".step .txt{flex:1;}",
      ".remote-pill{display:inline-block;padding:7px 12px;border:1px solid #8e939b;border-radius:999px;background:#f7f8fa;color:#252a31;font-weight:700;font-size:14px;}",
      ".grid{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-top:16px;}",
      ".panel{padding:16px;}",
      "h2{margin:0 0 10px;font-size:20px;}",
      "label{display:block;font-size:13px;color:var(--muted);margin-bottom:6px;}",
      "input{width:100%;padding:10px 12px;border:1px solid var(--line);border-radius:10px;font-size:15px;}",
      "pre{white-space:pre-wrap;background:#fbfaf7;padding:12px;border-radius:10px;border:1px solid var(--line);min-height:100px;}",
      ".ok{color:#1f7a3e;font-weight:700;}",
      ".warn{color:#a25b00;font-weight:700;}",
      ".small{font-size:12px;color:var(--muted);}",
      "@media (max-width:950px){.hero{grid-template-columns:1fr;}.grid{grid-template-columns:1fr;}h1{font-size:34px;}}",
      "</style>",
      "</head><body>",
      "<div class='top'>Weekly automation active. <strong>Mode Start Weekly</strong></div>",
      "<div class='wrap'>",
      "<div class='brand' style='margin-bottom:12px;'>",
      "<div class='logo'>indy</div>",
      "</div>",
      "<section class='hero'>",
      "<article class='card main'>",
      "<h1>Lancer le Weekly Kick Off</h1>",
      "<p class='sub'>Preparez la visio, les slides et le code remote, puis publiez le message Slack final.</p>",
      "<div class='actions' id='quickActions'></div>",
      "<div class='meta'>",
      "<div class='kpi'>Semaine: <span id='kpiWeek'>-</span></div>",
      "<div class='kpi'>Statut: <span id='kpiStatus'>Chargement...</span></div>",
      "</div>",
      "</article>",
      "<aside class='card side'>",
      "<h2>Mode d'emploi</h2>",
      "<div class='step'><span class='num'>1.</span><div class='txt'>Ouvre les slides.</div></div>",
      "<div class='step'><span class='num'>2.</span><div class='txt'>Ouvre la remote.</div></div>",
      "<div class='step'><span class='num'>3.</span><div class='txt'>Clique sur <span class='remote-pill'>Present w/ Remote</span> en haut de Google Slides.</div></div>",
      "<div class='step'><span class='num'>4.</span><div class='txt'>Copie le code de la remote.</div></div>",
      "<div class='step'><span class='num'>5.</span><div class='txt'>Colle le code dans cette Web App.</div></div>",
      "<div class='step'><span class='num'>6.</span><div class='txt'>Clique <b>Publier les infos sur all-announcements</b>.</div></div>",
      "<div class='step'><span class='num'>7.</span><div class='txt'>Partage ton ecran (onglet des slides).</div></div>",
      "<div class='step'><span class='num'>8.</span><div class='txt'>Prends la parole vers 250 connectes et laisse jusqu'a 3 minutes de marge.</div></div>",
      "</aside>",
      "</section>",
      "<section class='grid'>",
      "<article class='card panel'>",
      "<h2>Preview message a envoyer</h2>",
      "<label for='remoteCode'>Code remote</label>",
      "<input id='remoteCode' placeholder='Ex: 589485' oninput='refreshPreview()'>",
      "<div class='actions' style='margin-top:10px;'>",
      "<button class='btn primary' onclick='publishCode()'>Publier les infos sur all-announcements</button>",
      "</div>",
      "<pre id='weeklyMessage'>Chargement...</pre>",
      "<p class='small'>Le message ci-dessus correspond au message <b>Weekly remote code</b> envoye sur Slack.</p>",
      "</article>",
      "<article class='card panel'>",
      "<h2>Resultat envoi</h2>",
      "<pre id='remoteResult'>Aucun envoi pour l'instant.</pre>",
      "<p class='small'>Le bloc <b>Meet chat</b> est pret a copier/coller dans le chat Google Meet.</p>",
      "</article>",
      "</section>",
      "</div>",
      "<script>",
      "function esc(s){return (s||'').toString();}",
      "var ctxCache=null;",
      "var pendingSlideUrl='';",
      "function renderActions(){",
      " if(!ctxCache){return;}",
      " var links=[];",
      " if(ctxCache.meetUrl){links.push('<a class=\"btn primary\" target=\"_blank\" href=\"'+esc(ctxCache.meetUrl)+'\">Lancer le Weekly (Meet)</a>');}",
      " if(ctxCache.slideUrl){links.push('<a class=\"btn secondary\" target=\"_blank\" href=\"'+esc(ctxCache.slideUrl)+'\">Open Slide</a>');}else{links.push('<span class=\"btn secondary disabled\">Open Slide</span>');}",
      " if(ctxCache.remoteBaseUrl){links.push('<a class=\"btn secondary\" target=\"_blank\" href=\"'+esc(ctxCache.remoteBaseUrl)+'\">Open Remote</a>');}",
      " if(ctxCache.absenceUrl){links.push('<a class=\"btn secondary\" target=\"_blank\" href=\"'+esc(ctxCache.absenceUrl)+'\">Open Absences</a>');}",
      " document.getElementById('quickActions').innerHTML=links.join(' ');",
      "}",
      "function buildRemotePreview(code){",
      " if(!ctxCache){return 'Chargement...';}",
      " var c=(code||'').trim()||'[CODE_REMOTE]';",
      " var lines=['*Weekly remote code*','Code: `'+c+'`'];",
      " if(ctxCache.remoteBaseUrl){lines.push('Remote: '+ctxCache.remoteBaseUrl);}",
      " if(ctxCache.meetUrl){lines.push('Visio: '+ctxCache.meetUrl);}",
      " lines.push('Weekly: '+(ctxCache.mondayDate||''));",
      " return lines.join('\\n');",
      "}",
      "function render(ctx){",
      " ctxCache=ctx;",
      " if(pendingSlideUrl){ctxCache.slideUrl=pendingSlideUrl;}",
      " renderActions();",
      " document.getElementById('kpiWeek').textContent=ctx.mondayDate||'-';",
      " document.getElementById('kpiStatus').textContent=ctx.skip?('SKIP: '+ctx.skipReason):'OK';",
      " document.getElementById('kpiStatus').className=ctx.skip?'warn':'ok';",
      " document.getElementById('weeklyMessage').textContent=buildRemotePreview(document.getElementById('remoteCode').value);",
      "}",
      "function refreshPreview(){document.getElementById('weeklyMessage').textContent=buildRemotePreview(document.getElementById('remoteCode').value);}",
      "function load(){",
      " google.script.run.withSuccessHandler(render).getStartWeeklyContext();",
      " google.script.run.withSuccessHandler(function(slideUrl){pendingSlideUrl=slideUrl||'';if(ctxCache){ctxCache.slideUrl=pendingSlideUrl;renderActions();}}).getWeeklySlideUrlForUi();",
      "}",
      "function publishCode(){",
      " var code=document.getElementById('remoteCode').value;",
      " if(!code||!code.trim()){document.getElementById('remoteResult').textContent='Code remote obligatoire.';return;}",
      " google.script.run.withSuccessHandler(function(res){",
      "   document.getElementById('remoteResult').textContent = res.ok ? ('Slack OK\\n\\nMessage envoye:\\n'+(res.previewMessage||'')+'\\n\\nMeet chat:\\n'+res.meetText) : ('Echec: '+res.reason);",
      " }).publishRemoteCode(code);",
      "}",
      "load();",
      "</script></body></html>"
    ].join("");
    LOGGER.info("SUCCESS", { action: "buildStartWeeklyPageHtml_", durationMs: Date.now() - startMs });
    return html;
  } catch (e) {
    LOGGER.error("buildStartWeeklyPageHtml_", e.message, e.stack);
    return "<h3>Erreur rendering Start Weekly</h3>";
  }
}

function buildAbsencePageHtml_() {
  LOGGER.debug("START", { action: "buildAbsencePageHtml_" });
  const startMs = Date.now();
  try {
    const html = [
      "<!doctype html><html><head><meta charset='utf-8'><title>Absent à un weekly ? Déclare ton absence ici</title>",
      "<style>",
      "*{box-sizing:border-box;}",
      "body{margin:0;font-family:Arial,sans-serif;background:#f7f2ea;color:#1f2430;}",
      ".wrap{max-width:680px;margin:36px auto;padding:0 16px;}",
      ".card{background:#fff;border:1px solid #e7ddcf;border-radius:14px;padding:18px;box-shadow:0 10px 26px rgba(0,0,0,.06);}",
      "h1{margin:0 0 8px;font-size:30px;line-height:1.15;}",
      ".sub{margin:0 0 16px;color:#5c6470;}",
      "label{display:block;margin:10px 0 6px;font-weight:700;font-size:13px;}",
      "input{width:100%;padding:10px 12px;border:1px solid #d9d0c2;border-radius:10px;font-size:14px;min-width:0;}",
      "input[type='date']{padding-right:34px;}",
      ".row{display:grid;grid-template-columns:1fr 1fr;gap:10px;}",
      ".multi{margin-top:8px;display:grid;gap:8px;}",
      ".lineTag{display:inline-flex;align-items:center;justify-content:center;min-width:24px;height:24px;border-radius:999px;background:#f3ede1;border:1px solid #e0d5c3;font-size:12px;font-weight:700;color:#5b6470;}",
      ".r5{display:grid;grid-template-columns:24px 1fr 1fr;gap:8px;align-items:end;}",
      ".r5 > div{min-width:0;}",
      ".actions{margin-top:14px;display:flex;gap:10px;align-items:center;}",
      "button{padding:10px 14px;border:0;border-radius:999px;background:#ef5b86;color:#fff;font-weight:800;cursor:pointer;}",
      "#status{font-size:13px;color:#2b3340;}",
      ".small{margin-top:10px;color:#6b7582;font-size:12px;}",
      "@media (max-width:700px){.row,.r5{grid-template-columns:1fr;}.lineTag{display:none;}}",
      "</style></head><body>",
      "<div class='wrap'><div class='card'>",
      "<h1>Absent à un weekly ? Déclare ton absence ici</h1>",
      "<label for='email'>Email</label>",
      "<input id='email' placeholder='prenom.nom@indy.fr'>",
      "<label for='name'>Nom</label>",
      "<input id='name' placeholder='Prénom Nom'>",
      "<label>Absences (jusqu'à 5)</label>",
      "<div class='multi'>",
      "<div class='r5'><span class='lineTag'>1</span><div><label for='startDate1'>Date début</label><input id='startDate1' type='date'></div><div><label for='endDate1'>Date fin</label><input id='endDate1' type='date'></div></div>",
      "<div class='r5'><span class='lineTag'>2</span><div><label for='startDate2'>Date début</label><input id='startDate2' type='date'></div><div><label for='endDate2'>Date fin</label><input id='endDate2' type='date'></div></div>",
      "<div class='r5'><span class='lineTag'>3</span><div><label for='startDate3'>Date début</label><input id='startDate3' type='date'></div><div><label for='endDate3'>Date fin</label><input id='endDate3' type='date'></div></div>",
      "<div class='r5'><span class='lineTag'>4</span><div><label for='startDate4'>Date début</label><input id='startDate4' type='date'></div><div><label for='endDate4'>Date fin</label><input id='endDate4' type='date'></div></div>",
      "<div class='r5'><span class='lineTag'>5</span><div><label for='startDate5'>Date début</label><input id='startDate5' type='date'></div><div><label for='endDate5'>Date fin</label><input id='endDate5' type='date'></div></div>",
      "</div>",
      "<div class='actions'>",
      "<button onclick='submitForm()'>Enregistrer l'absence</button>",
      "<span id='status'></span>",
      "</div>",
      "<div class='small'>Tu peux déclarer jusqu'à 5 absences. Cela ne concerne que le Weekly et ne sera partagé nulle part ailleurs.</div>",
      "</div></div>",
      "<script>",
      "function load(){",
      " google.script.run.withSuccessHandler(function(ctx){",
      "   if(ctx.email){document.getElementById('email').value=ctx.email;}",
      "   if(ctx.today){document.getElementById('startDate1').value=ctx.today;document.getElementById('endDate1').value=ctx.today;}",
      " }).getAbsenceFormContext();",
      "}",
      "function submitForm(){",
      " var abs=[];",
      " for(var i=1;i<=5;i++){",
      "   abs.push({",
      "     startDate:document.getElementById('startDate'+i).value,",
      "     endDate:document.getElementById('endDate'+i).value",
      "   });",
      " }",
      " var payload={",
      "   email:document.getElementById('email').value,",
      "   name:document.getElementById('name').value,",
      "   absences:abs",
      " };",
      " document.getElementById('status').textContent='Enregistrement...';",
      " google.script.run.withSuccessHandler(function(res){",
      "   document.getElementById('status').textContent=(res&&res.message)?res.message:'Terminé.';",
      " }).submitAbsenceRequest(payload);",
      "}",
      "load();",
      "</script></body></html>"
    ].join("");
    LOGGER.info("SUCCESS", { action: "buildAbsencePageHtml_", durationMs: Date.now() - startMs });
    return html;
  } catch (e) {
    LOGGER.error("buildAbsencePageHtml_", e.message, e.stack);
    return "<h3>Erreur chargement formulaire absence</h3>";
  }
}

function isValidDateIso_(value) {
  LOGGER.debug("START", { action: "isValidDateIso_", value: value });
  const startMs = Date.now();
  try {
    const iso = String(value || "").trim();
    const ok = /^\d{4}-\d{2}-\d{2}$/.test(iso);
    LOGGER.info("SUCCESS", { action: "isValidDateIso_", ok: ok, durationMs: Date.now() - startMs });
    return ok;
  } catch (e) {
    LOGGER.error("isValidDateIso_", e.message, e.stack);
    return false;
  }
}

function getWeeklyVisioUrl_() {
  LOGGER.debug("START", { action: "getWeeklyVisioUrl_" });
  const startMs = Date.now();
  try {
    const rawUrl = String(CONFIG.get("LIEN_VISIO_WEEKLY", CONFIG.get("WEEKLY_MEET_URL", "")) || "").trim();
    const url = normalizeMeetUrl_(rawUrl);
    LOGGER.info("SUCCESS", { action: "getWeeklyVisioUrl_", hasUrl: !!url, durationMs: Date.now() - startMs });
    return url;
  } catch (e) {
    LOGGER.error("getWeeklyVisioUrl_", e.message, e.stack);
    return "";
  }
}

function normalizeMeetUrl_(value) {
  LOGGER.debug("START", { action: "normalizeMeetUrl_", value: value });
  const startMs = Date.now();
  try {
    const raw = String(value || "").trim();
    if (!raw) {
      LOGGER.info("SUCCESS", { action: "normalizeMeetUrl_", normalized: "", durationMs: Date.now() - startMs });
      return "";
    }
    if (/^https?:\/\//i.test(raw)) {
      LOGGER.info("SUCCESS", { action: "normalizeMeetUrl_", normalized: raw, durationMs: Date.now() - startMs });
      return raw;
    }
    if (/^meet\.google\.com\//i.test(raw)) {
      const normalizedMeet = "https://" + raw;
      LOGGER.info("SUCCESS", { action: "normalizeMeetUrl_", normalized: normalizedMeet, durationMs: Date.now() - startMs });
      return normalizedMeet;
    }
    if (/^\/\//.test(raw)) {
      const normalizedProto = "https:" + raw;
      LOGGER.info("SUCCESS", { action: "normalizeMeetUrl_", normalized: normalizedProto, durationMs: Date.now() - startMs });
      return normalizedProto;
    }
    LOGGER.info("SUCCESS", { action: "normalizeMeetUrl_", normalized: raw, durationMs: Date.now() - startMs });
    return raw;
  } catch (e) {
    LOGGER.error("normalizeMeetUrl_", e.message, e.stack);
    return "";
  }
}

function buildRemoteCodeMessage_(mondayDateIso, remoteCode, remoteBaseUrl, visioUrl) {
  LOGGER.debug("START", {
    action: "buildRemoteCodeMessage_",
    mondayDateIso: mondayDateIso,
    hasRemoteCode: !!remoteCode,
    hasRemoteBaseUrl: !!remoteBaseUrl,
    hasVisioUrl: !!visioUrl
  });
  const startMs = Date.now();
  try {
    const code = String(remoteCode || "").trim();
    const absenceUrl = String(getAbsenceWebAppUrl() || "").trim();
    const lines = ["*Weekly remote code*"];
    if (code) {
      lines.push("Code: `" + code + "`");
    } else {
      lines.push("Code: `[CODE_REMOTE]`");
    }
    if (remoteBaseUrl) {
      lines.push("Remote: " + remoteBaseUrl);
    }
    if (visioUrl) {
      lines.push("Visio: " + visioUrl);
    }
    if (mondayDateIso) {
      lines.push("Weekly: " + mondayDateIso);
    }
    if (absenceUrl) {
      lines.push("Si tu es en conges un lundi prochainement, pense a le noter " + formatSlackLink_(absenceUrl, "ici"));
      lines.push("URL absences: " + absenceUrl);
    }
    const message = lines.join("\n");
    LOGGER.info("SUCCESS", { action: "buildRemoteCodeMessage_", durationMs: Date.now() - startMs });
    return message;
  } catch (e) {
    LOGGER.error("buildRemoteCodeMessage_", e.message, e.stack);
    return "*Weekly remote code*";
  }
}

function extractUrlFromSlackLink_(slackLink) {
  LOGGER.debug("START", { action: "extractUrlFromSlackLink_", slackLink: slackLink });
  const startMs = Date.now();
  try {
    const value = String(slackLink || "").trim();
    if (!value) {
      LOGGER.info("SUCCESS", { action: "extractUrlFromSlackLink_", found: false, durationMs: Date.now() - startMs });
      return "";
    }
    const match = value.match(/^<([^|>]+)\|?.*>$/);
    const url = match ? String(match[1]) : value;
    LOGGER.info("SUCCESS", { action: "extractUrlFromSlackLink_", found: !!url, durationMs: Date.now() - startMs });
    return url;
  } catch (e) {
    LOGGER.error("extractUrlFromSlackLink_", e.message, e.stack);
    return "";
  }
}
