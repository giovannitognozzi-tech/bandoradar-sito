/**
 * BandoRadar - Iscrizioni e invio digest (Google Apps Script) — versione 3
 *
 * 1. Iscrizione dal sito: salva l'impresa nel foglio "Iscrizioni" (stato "attiva"),
 *    manda la conferma e, subito dopo, una seconda email con le gare aperte adatte a lei.
 * 2. Bandi aperti: ogni mattina la pipeline carica qui l'elenco delle gare partecipabili
 *    e la mappa ATECO -> CPV (POST azione "bandi"). Così la mail all'iscrizione parte
 *    all'istante, anche a PC spento.
 * 3. Elenco iscritti per la pipeline (GET ?azione=iscritti&token=...), con i bandi già
 *    mandati all'iscrizione, perché il digest del giorno dopo non li ripeta.
 * 4. Invio dei digest per conto della pipeline (POST azione "digest"), solo verso iscritti attivi.
 * 5. Disiscrizione con un clic dal link in fondo a ogni email.
 *
 * Dopo aver incollato questo codice: Esegui "installa", poi
 * Distribuisci > Gestisci distribuzioni > matita > Versione: Nuova versione > Distribuisci.
 */

const FOGLIO = "Iscrizioni";
const FOGLIO_BANDI = "Bandi aperti";
const FOGLIO_MAPPA = "Mappa ATECO";
const FOGLIO_INVII = "Invii all'iscrizione";
const INTESTAZIONE = ["Data iscrizione", "Email", "Partita IVA", "Ragione sociale",
                      "Codice ATECO", "Classifica SOA", "Area", "Consenso privacy", "Stato"];
const COL = { email: 1, piva: 2, ragione: 3, ateco: 4, soa: 5, area: 6, stato: 8 };   // indici 0-based
const INTESTAZIONE_BANDI = ["id", "fonte", "url", "cig", "ente", "regione", "oggetto", "cpv", "importo", "scadenza"];
const CLASSIFICHE_EUR = { "I": 258000, "II": 516000, "III": 1033000, "III-bis": 1500000, "IV": 2582000,
                          "IV-bis": 3500000, "V": 5165000, "VI": 10329000, "VII": 15494000, "VIII": Infinity };
const GIORNI_MINIMI = 7;      // stessa regola della pipeline: almeno 7 giorni per preparare l'offerta
const MAX_BANDI_MAIL = 5;
const NOME_MITTENTE = "BandoRadar";

/** Da eseguire dall'editor dopo ogni aggiornamento: fogli, permessi e codice segreto. */
function installa() {
  foglio_();
  MailApp.getRemainingDailyQuota();
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty("TOKEN")) props.setProperty("TOKEN", Utilities.getUuid().replace(/-/g, ""));
  Logger.log("Pronto. Email ancora inviabili oggi: " + MailApp.getRemainingDailyQuota());
  Logger.log("CODICE SEGRETO per la pipeline (BANDORADAR_TOKEN): " + props.getProperty("TOKEN"));
}

// ================================================================ GET

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.azione === "cancella") return paginaCancellazione_(p.p, p.k);
  if (p.azione === "iscritti") {
    if (!tokenValido_(p.token)) return json_({ ok: false, errore: "non autorizzato" });
    return json_({ ok: true, iscritti: iscrittiAttivi_() });
  }
  return ContentService.createTextOutput("BandoRadar: iscrizioni attive.");
}

// ================================================================ POST

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents || "{}");
    if (d.azione === "digest") return inviaDigest_(d);
    if (d.azione === "bandi") return caricaBandi_(d);
    return iscrivi_(d);
  } catch (err) {
    console.error(err);
    return json_({ ok: false, errore: "Errore temporaneo: riprova tra qualche minuto." });
  }
}

function iscrivi_(d) {
  if (d.botcheck) return json_({ ok: true, n_bandi: 0 });          // trappola anti-bot: fingiamo successo

  const email = String(d.email || "").trim().toLowerCase();
  const piva = String(d.partita_iva || "").replace(/\s+/g, "").replace(/^IT/i, "");
  const ateco = normalizzaAteco_(d.ateco);
  const soa = CLASSIFICHE_EUR.hasOwnProperty(String(d.classifica_soa || "")) ? String(d.classifica_soa) : "";
  const ragione = testo_(d.ragione_sociale, 150);
  const area = testo_(d.area, 40) || "Tutta Italia";

  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) return json_({ ok: false, errore: "Indirizzo email non valido." });
  if (!pivaValida_(piva)) return json_({ ok: false, errore: "Partita IVA non valida." });
  if (!ateco) return json_({ ok: false, errore: "Codice ATECO non valido (esempio: 43.21.01)." });
  if (d.consenso_privacy !== true) return json_({ ok: false, errore: "Serve il consenso privacy." });

  let giaIscritta = false;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = foglio_();
    const riga = [new Date(), email, "'" + piva, ragione, "'" + ateco, soa, area, "sì", "attiva"];
    const n = rigaDi_(sh, piva);
    if (n > 0) { giaIscritta = true; sh.getRange(n, 1, 1, riga.length).setValues([riga]); }
    else sh.appendRow(riga);
  } finally {
    lock.releaseLock();
  }

  let nBandi = 0;
  const cache = CacheService.getScriptCache();          // anti-abuso: una tornata di email ogni 6 ore per indirizzo
  if (!cache.get("conf:" + email)) {
    const bandi = bandiPer_(ateco, soa, area, piva);
    nBandi = bandi.length;
    inviaConferma_(email, piva, ragione, ateco, soa, area, giaIscritta, nBandi);
    if (nBandi) {
      inviaBandiIscrizione_(email, ragione || ("P.IVA " + piva), piva, bandi);
      registraInvii_(piva, bandi);
    }
    cache.put("conf:" + email, "1", 21600);
    avvisaTitolare_(email, piva, ragione, ateco, soa, area, giaIscritta, nBandi);
  }
  return json_({ ok: true, giaIscritta: giaIscritta, n_bandi: nBandi });
}

/** La pipeline carica ogni mattina le gare aperte e la mappa ATECO (sostituisce le precedenti). */
function caricaBandi_(d) {
  if (!tokenValido_(d.token)) return json_({ ok: false, errore: "non autorizzato" });
  const bandi = Array.isArray(d.bandi) ? d.bandi : [];
  const mappa = Array.isArray(d.mappa) ? d.mappa : [];
  const righeB = bandi.map(function (b) {
    return [b.id, b.fonte, b.url, b.cig || "", b.ente || "", b.regione || "", testo_(b.oggetto, 300),
            (b.cpv || []).join(" "), b.importo == null ? "" : String(b.importo), b.scadenza || ""];
  });
  scriviTabella_(FOGLIO_BANDI, INTESTAZIONE_BANDI, righeB);
  if (mappa.length) {
    scriviTabella_(FOGLIO_MAPPA, ["prefisso", "descrizione", "cpv", "parole_chiave"],
      mappa.map(function (m) { return [m.prefisso, m.descrizione, (m.cpv || []).join(" "), (m.parole || []).join("|")]; }));
  }
  return json_({ ok: true, bandi: righeB.length, mappa: mappa.length });
}

/** Invio di un digest per conto della pipeline: solo con codice segreto e solo a iscritti attivi. */
function inviaDigest_(d) {
  if (!tokenValido_(d.token)) return json_({ ok: false, errore: "non autorizzato" });
  const to = String(d.to || "").trim().toLowerCase();
  const iscritto = iscrittiAttivi_().filter(function (x) { return x.email === to; })[0];
  if (!iscritto) return json_({ ok: false, errore: "destinatario non iscritto o non attivo" });
  if (MailApp.getRemainingDailyQuota() < 1) return json_({ ok: false, errore: "quota email giornaliera esaurita" });
  const html = String(d.html || "") + piedeCancellazione_(iscritto.partita_iva);
  MailApp.sendEmail({ to: to, subject: testo_(d.subject, 200), htmlBody: html,
                      body: "Il tuo digest BandoRadar è in formato HTML.", name: NOME_MITTENTE });
  return json_({ ok: true, quota_rimanente: MailApp.getRemainingDailyQuota() });
}

// ================================================================ MATCHING ALL'ISCRIZIONE

/** Stesse regole del filtro della pipeline: scadenza, importo SOA, area, CPV o parole chiave. */
function bandiPer_(ateco, soa, area, piva) {
  const voce = voceAteco_(ateco);
  if (!voce) return [];
  const tetto = soa ? CLASSIFICHE_EUR[soa] * 1.2 : Infinity;
  const limite = new Date(); limite.setHours(0, 0, 0, 0); limite.setDate(limite.getDate() + GIORNI_MINIMI);
  const giaMandati = inviiPer_(piva);
  const tutte = leggiTabella_(FOGLIO_BANDI);
  const out = [];
  tutte.forEach(function (b) {
    if (giaMandati[b.id]) return;
    if (b.scadenza && new Date(b.scadenza + "T00:00:00") < limite) return;
    const importo = b.importo === "" ? null : Number(b.importo);
    if (importo && importo > tetto) return;
    if (area !== "Tutta Italia" && b.regione && b.regione !== area) return;
    const cpv = String(b.cpv || "").split(" ").filter(String);
    const cpvOk = cpv.some(function (c) { return voce.cpv.some(function (p) { return c.indexOf(p) === 0; }); });
    const testo = (b.oggetto + " " + b.ente).toLowerCase();
    const kw = voce.parole.filter(function (k) { return k && testo.indexOf(k.toLowerCase()) >= 0; });
    if (!cpvOk && !kw.length) return;
    b.punteggio = (cpvOk ? 60 : 0) + Math.min(kw.length, 4) * 10;
    b.motivo = [cpvOk ? "settore (CPV) compatibile" : "", kw.length ? "parole chiave: " + kw.slice(0, 3).join(", ") : ""]
               .filter(String).join("; ");
    out.push(b);
  });
  out.sort(function (a, b) {
    return b.punteggio - a.punteggio || String(a.scadenza || "9999").localeCompare(String(b.scadenza || "9999"));
  });
  return out.slice(0, MAX_BANDI_MAIL);
}

function voceAteco_(ateco) {
  const cifre = String(ateco || "").replace(/[^\d]/g, "");
  const mappa = {};
  leggiTabella_(FOGLIO_MAPPA).forEach(function (m) { mappa[String(m.prefisso)] = m; });
  for (let n = cifre.length; n >= 2; n--) {
    const m = mappa[cifre.slice(0, n)];
    if (m) return { cpv: String(m.cpv).split(" ").filter(String), parole: String(m.parole_chiave).split("|").filter(String) };
  }
  return null;
}

function inviiPer_(piva) {
  const fatti = {};
  leggiTabella_(FOGLIO_INVII).forEach(function (r) { if (String(r.partita_iva) === piva) fatti[r.bando_id] = true; });
  return fatti;
}

function registraInvii_(piva, bandi) {
  const sh = foglioConIntestazione_(FOGLIO_INVII, ["partita_iva", "bando_id", "data"]);
  const ora = new Date();
  bandi.forEach(function (b) { sh.appendRow(["'" + piva, b.id, ora]); });
}

// ================================================================ EMAIL

function inviaConferma_(email, piva, ragione, ateco, soa, area, giaIscritta, nBandi) {
  const nome = ragione || "la tua impresa";
  const oggetto = giaIscritta ? "BandoRadar: dati aggiornati" : "BandoRadar: iscrizione completata";
  const frase = nBandi
    ? "In un'email a parte trovi già " + (nBandi === 1 ? "la gara aperta adatta" : "le " + nBandi + " gare aperte più adatte") + " alla tua impresa. "
      + "Da domani ti scriviamo ogni mattina in cui ne escono di nuove."
    : "Al momento non ci sono gare aperte adatte alla tua impresa: ogni mattina controlliamo i nuovi bandi e ti scriviamo appena ne esce una.";
  const html =
    '<div style="font-family:Arial,sans-serif;max-width:560px;color:#16243A;line-height:1.55">' +
    '<h2 style="color:#0B2545;margin:0 0 12px">' + (giaIscritta ? "Dati aggiornati" : "Benvenuto in BandoRadar") + '</h2>' +
    '<p>L\'iscrizione di <b>' + esc_(nome) + '</b> è completa.</p>' +
    '<table style="border-collapse:collapse;margin:14px 0;font-size:14px">' +
      rigaTab_("Partita IVA", piva) + rigaTab_("Codice ATECO", ateco) +
      (soa ? rigaTab_("Classifica SOA", soa) : "") + rigaTab_("Area", area) +
    '</table>' +
    '<p>' + frase + ' Nei giorni senza gare utili non ti disturbiamo.</p>' +
    '<p style="margin-top:20px">— BandoRadar</p></div>' + piedeCancellazione_(piva);
  const testo = "L'iscrizione di " + nome + " (P.IVA " + piva + ") a BandoRadar è completa. " + frase +
    " Per annullare l'iscrizione: " + linkCancellazione_(piva);
  MailApp.sendEmail({ to: email, subject: oggetto, body: testo, htmlBody: html, name: NOME_MITTENTE });
}

function inviaBandiIscrizione_(email, nome, piva, bandi) {
  const oggi = Utilities.formatDate(new Date(), "Europe/Rome", "dd/MM");
  const oggetto = "📡 " + bandi.length + (bandi.length === 1 ? " bando aperto" : " bandi aperti") + " per te (" + oggi + ") — " + nome;
  const schede = bandi.map(schedaBando_).join("");
  const html =
    '<div style="font-family:Arial,sans-serif;max-width:600px;color:#16243A;line-height:1.5">' +
    '<h2 style="color:#0B2545;margin:0 0 4px">📡 BandoRadar</h2>' +
    '<p style="color:#586A82;margin-top:0">Ecco le gare aperte adatte a <b>' + esc_(nome) + '</b> in questo momento.</p>' +
    schede + '</div>' + piedeCancellazione_(piva);
  MailApp.sendEmail({ to: email, subject: oggetto, htmlBody: html,
                      body: "Le gare aperte per te sono nella versione HTML di questa email.", name: NOME_MITTENTE });
}

function schedaBando_(b) {
  const importo = b.importo === "" ? "non indicato" : "€ " + Number(b.importo).toLocaleString("it-IT", { maximumFractionDigits: 0 });
  const scad = b.scadenza ? b.scadenza.split("-").reverse().join("/") : "non indicata";
  const azione = b.fonte === "TED"
    ? '<a href="' + esc_(b.url) + '" style="display:inline-block;background:#1a56db;color:#fff;padding:9px 16px;border-radius:7px;text-decoration:none;font-size:14px">Apri il bando →</a>'
    : '<div style="padding:10px 12px;background:#f0f4ff;border-radius:7px"><div style="font-size:12px;color:#555">Codice CIG — cercalo sul portale dell\'ente o della tua regione:</div>' +
      '<div style="font-size:19px;font-weight:700;letter-spacing:2px;color:#1a56db;font-family:Consolas,monospace">' + esc_(b.cig || "n/d") + '</div></div>';
  return '<div style="border:1px solid #e0e0e0;border-radius:10px;padding:16px;margin:14px 0">' +
    '<div style="font-size:15px;font-weight:700;color:#0B2545;margin-bottom:6px">' + esc_(b.oggetto) + '</div>' +
    '<div style="font-size:14px;color:#333">🏛️ ' + esc_(b.ente || "n/d") + (b.regione ? " · " + esc_(b.regione) : "") + '<br>' +
    '💶 ' + importo + ' · 📅 scadenza offerte ' + scad + '<br>' +
    '<span style="color:#586A82">🎯 ' + esc_(b.motivo) + '</span></div>' +
    '<div style="margin-top:12px">' + azione + '</div></div>';
}

function avvisaTitolare_(email, piva, ragione, ateco, soa, area, giaIscritta, nBandi) {
  const io = Session.getEffectiveUser().getEmail();
  if (!io) return;
  MailApp.sendEmail(io, (giaIscritta ? "Iscrizione aggiornata: " : "Nuova iscrizione: ") + (ragione || piva),
    [ "Email: " + email, "P.IVA: " + piva, "Ragione sociale: " + (ragione || "-"), "ATECO: " + ateco,
      "Classifica SOA: " + (soa || "-"), "Area: " + area, "Bandi inviati all'iscrizione: " + nBandi, "",
      "Elenco completo nel foglio Iscrizioni." ].join("\n"),
    { name: NOME_MITTENTE });
}

function piedeCancellazione_(piva) {
  return '<p style="font-family:Arial,sans-serif;color:#8A98AB;font-size:12px;margin-top:28px;border-top:1px solid #E3E9F1;padding-top:12px">' +
    'Ricevi questa email perché la tua impresa è iscritta a BandoRadar. ' +
    '<a href="' + linkCancellazione_(piva) + '" style="color:#8A98AB">Annulla l\'iscrizione</a></p>';
}

function linkCancellazione_(piva) {
  return ScriptApp.getService().getUrl() + "?azione=cancella&p=" + encodeURIComponent(piva) + "&k=" + firma_(piva);
}

function paginaCancellazione_(piva, firma) {
  let titolo = "Link non valido", testo = "Il link di cancellazione non è valido. Rispondi a una nostra email per annullare l'iscrizione.";
  if (piva && firma && firma === firma_(String(piva))) {
    const sh = foglio_();
    const n = rigaDi_(sh, String(piva));
    if (n > 0) sh.getRange(n, COL.stato + 1).setValue("cancellata");
    titolo = "Iscrizione annullata";
    testo = "Non riceverai più email da BandoRadar. Se cambi idea, puoi iscriverti di nuovo dal sito.";
  }
  return HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;max-width:480px;margin:60px auto;color:#16243A">' +
    '<h2 style="color:#0B2545">' + titolo + '</h2><p>' + testo + '</p></div>').setTitle("BandoRadar");
}

// ================================================================ FOGLI

function foglio_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(FOGLIO);
  if (sh) {
    const intest = sh.getRange(1, 1, 1, INTESTAZIONE.length).getValues()[0].join("|");
    if (intest !== INTESTAZIONE.join("|")) {
      if (sh.getLastRow() <= 1) { sh.clear(); sh.appendRow(INTESTAZIONE); }
      else { sh.setName(FOGLIO + " (vecchio formato)"); sh = null; }    // conserva i dati vecchi
    }
  }
  if (!sh) {
    sh = ss.insertSheet(FOGLIO);
    sh.appendRow(INTESTAZIONE);
  }
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, INTESTAZIONE.length).setFontWeight("bold");
  return sh;
}

function foglioConIntestazione_(nome, intestazione) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(nome);
  if (!sh) {
    sh = ss.insertSheet(nome);
    sh.appendRow(intestazione);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Sostituisce il contenuto di una scheda; tutto come testo (CPV e date non vanno convertiti). */
function scriviTabella_(nome, intestazione, righe) {
  const sh = foglioConIntestazione_(nome, intestazione);
  sh.clearContents();
  const tutte = [intestazione].concat(righe);
  const rng = sh.getRange(1, 1, tutte.length, intestazione.length);
  rng.setNumberFormat("@");
  rng.setValues(tutte);
  sh.setFrozenRows(1);
}

function leggiTabella_(nome) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nome);
  if (!sh) return [];
  const righe = sh.getDataRange().getValues();
  if (righe.length < 2) return [];
  const chiavi = righe[0].map(String);
  return righe.slice(1).map(function (r) {
    const o = {};
    chiavi.forEach(function (k, i) { o[k] = r[i] instanceof Date ? Utilities.formatDate(r[i], "Europe/Rome", "yyyy-MM-dd") : String(r[i]); });
    return o;
  });
}

function rigaDi_(sh, piva) {
  const righe = sh.getDataRange().getValues();
  for (let i = 1; i < righe.length; i++) if (String(righe[i][COL.piva]) === piva) return i + 1;
  return -1;
}

function iscrittiAttivi_() {
  const righe = foglio_().getDataRange().getValues();
  const invii = {};
  leggiTabella_(FOGLIO_INVII).forEach(function (r) {
    (invii[String(r.partita_iva)] = invii[String(r.partita_iva)] || []).push(r.bando_id);
  });
  const out = [];
  for (let i = 1; i < righe.length; i++) {
    const r = righe[i];
    if (String(r[COL.stato]).trim().toLowerCase() !== "attiva") continue;
    const piva = String(r[COL.piva]);
    out.push({ email: String(r[COL.email]).trim().toLowerCase(), partita_iva: piva,
               ragione_sociale: String(r[COL.ragione] || ""), ateco: String(r[COL.ateco] || ""),
               classifica_soa: String(r[COL.soa] || ""), area: String(r[COL.area] || "Tutta Italia"),
               gia_inviati: invii[piva] || [] });
  }
  return out;
}

// ================================================================ UTILITÀ

function tokenValido_(t) {
  const giusto = PropertiesService.getScriptProperties().getProperty("TOKEN");
  return !!giusto && t === giusto;
}

function firma_(piva) {
  const token = PropertiesService.getScriptProperties().getProperty("TOKEN") || "";
  const sig = Utilities.computeHmacSha256Signature(String(piva), token);
  return Utilities.base64EncodeWebSafe(sig).replace(/=+$/, "").slice(0, 22);
}

function normalizzaAteco_(v) {
  const cifre = String(v || "").replace(/[^\d]/g, "");
  if (cifre.length < 2 || cifre.length > 6) return "";
  return cifre.match(/.{1,2}/g).join(".");
}

/** Partita IVA italiana: 11 cifre con cifra di controllo (stesso algoritmo della pagina). */
function pivaValida_(s) {
  if (!/^\d{11}$/.test(s)) return false;
  let tot = 0;
  for (let i = 0; i < 10; i++) {
    let c = +s[i];
    if (i % 2 === 1) { c *= 2; if (c > 9) c -= 9; }
    tot += c;
  }
  return (10 - (tot % 10)) % 10 === +s[10];
}

function testo_(v, max) { return String(v || "").replace(/[\r\n\t]+/g, " ").trim().slice(0, max); }
function esc_(s) { return String(s).replace(/[&<>"']/g, function (c) { return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]; }); }
function rigaTab_(k, v) {
  return '<tr><td style="padding:3px 14px 3px 0;color:#586A82">' + k + '</td><td style="padding:3px 0"><b>' + esc_(v) + '</b></td></tr>';
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
