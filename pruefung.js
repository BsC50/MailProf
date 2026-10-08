/*
 * Mail-Prüfer – gemeinsame Prüflogik
 * Wird sowohl beim Senden (Smart Alert) als auch in der Seitenleiste benutzt.
 * Bewusst ohne Module/Build-Schritt geschrieben, damit es überall in Outlook läuft.
 */
/* global Office */

var MailPruefer = (function () {
  "use strict";

  // ---------- Einstellungen ----------
  var STANDARD = {
    name: "",                 // leer = Vorname aus dem Outlook-Profil
    rechtschreibung: true,
    sprache: "de-CH",          // Schweizer Rechtschreibung (ss statt ß)
    ignorierteWoerter: []      // eigene Fachbegriffe, die nie als Fehler gelten
  };

  function ladeEinstellungen() {
    var e = {};
    for (var k in STANDARD) e[k] = STANDARD[k];
    try {
      var gespeichert = Office.context.roamingSettings.get("mailpruefer");
      if (gespeichert) for (var k2 in gespeichert) e[k2] = gespeichert[k2];
    } catch (x) { /* ausserhalb von Outlook (Tests) */ }
    if (!e.name) {
      try {
        var voll = Office.context.mailbox.userProfile.displayName || "";
        e.name = voll.split(/[\s,]+/)[0] || "";
      } catch (x) { /* ignorieren */ }
    }
    return e;
  }

  function speichereEinstellungen(e, fertig) {
    Office.context.roamingSettings.set("mailpruefer", e);
    Office.context.roamingSettings.saveAsync(function () { if (fertig) fertig(); });
  }

  // ---------- Text vorbereiten ----------
  // Nur den neu geschriebenen Teil prüfen, nicht den zitierten Verlauf.
  var ZITAT_MARKER = [
    /^-{2,}\s*(Ursprüngliche Nachricht|Original Message|Weitergeleitete Nachricht|Forwarded message)/im,
    /^_{5,}\s*$/m,
    /^\s*(Von|From)\s*:.*\n\s*(Gesendet|Sent|Datum|Date)\s*:/im,
    /^\s*Am .{5,80}(schrieb|hat .{1,60} geschrieben)\s*.*:\s*$/im,
    /^\s*On .{5,80}wrote:\s*$/im
  ];

  function eigenerText(text) {
    text = (text || "").replace(/\r\n?/g, "\n").replace(/ /g, " ");
    var ende = text.length;
    ZITAT_MARKER.forEach(function (re) {
      var m = re.exec(text);
      if (m && m.index < ende) ende = m.index;
    });
    return text.slice(0, ende);
  }

  function zeilen(text) {
    return text.split("\n").map(function (z) { return z.trim(); });
  }

  function ersteNichtLeereZeile(text) {
    var z = zeilen(text);
    for (var i = 0; i < z.length; i++) if (z[i]) return z[i];
    return "";
  }

  function kurz(s, n) {
    n = n || 60;
    s = s.replace(/\s+/g, " ").trim();
    return s.length > n ? s.slice(0, n - 1) + "…" : s;
  }

  // ---------- 1. Anhang erwähnt, aber keiner dabei ----------
  var ANHANG_WOERTER = [
    /\bim\s+anhang\b/i, /\bals\s+anhang\b/i, /\banbei\b/i, /\bangehängt\w*/i,
    /\bbeigefügt\w*/i, /\bbeiliegend\w*/i, /\bin\s+der\s+beilage\b/i, /\bbeilage\b/i,
    /\banhängend\b/i, /\bhänge\b.{0,40}\ban\b/i, /\bsende\b.{0,40}\bmit\b/i,
    /\bschicke\b.{0,40}\bmit\b/i, /\bfinden\s+sie\b.{0,30}\b(anhang|beilage|dokument|pdf|datei)/i,
    /\battached\b/i, /\battachment\b/i, /\benclosed\b/i
  ];

  function pruefeAnhang(text, anzahlAnhaenge) {
    if (anzahlAnhaenge > 0) return [];
    var z = zeilen(text);
    for (var i = 0; i < z.length; i++) {
      for (var j = 0; j < ANHANG_WOERTER.length; j++) {
        var m = ANHANG_WOERTER[j].exec(z[i]);
        if (m) {
          return [{
            id: "anhang",
            art: "fehler",
            titel: "Anhang fehlt",
            detail: "Du erwähnst einen Anhang, aber es ist keine Datei angehängt.",
            fundstelle: kurz(z[i], 80),
            markiert: m[0],
            hilfe: "Hänge die Datei an (Büroklammer-Symbol) oder entferne den Hinweis."
          }];
        }
      }
    }
    return [];
  }

  // ---------- 1b. Passt der Dateiname zur Mail? ----------
  // Allgemeine Wörter in Dateinamen, die nichts über den Inhalt sagen
  var DATEI_ALLGEMEIN = (
    "scan scans scanned img image images bild bilder foto fotos photo photos pic dsc dscn dcim screenshot bildschirmfoto " +
    "whatsapp signal telegram export download downloads file files datei dateien dokument dokumente document documents doc docs " +
    "final finale endversion entwurf draft kopie copy version vers rev neu new old alt unbenannt untitled ohne titel " +
    "signed unterschrieben unterzeichnet sign attachment anhang anhaenge beilage beilagen anlage pdf docx xlsx pptx jpg jpeg png heic " +
    "und and the for von vom fuer for der die das des mit with ein eine to im in am an auf aus zu zum zur teil part seite page kw"
  ).split(" ");

  // Englisch / Deutsch / Schweizerdeutsch – Wörter, die zusammengehören
  var WORTGRUPPEN = [
    "school schule schul schulisch", "invoice rechnung bill faktura", "receipt quittung beleg", "contract vertrag vereinbarung agreement",
    "offer angebot offerte quote quotation", "insurance versicherung police", "health gesundheit krankenkasse kk krankenversicherung",
    "application bewerbung", "cv lebenslauf resume", "certificate zeugnis zertifikat bescheinigung", "agenda traktanden traktandenliste tagesordnung",
    "minutes protokoll", "presentation praesentation folien slides", "worksheet arbeitsblatt arbeitsblaetter", "exercise uebung aufgabe",
    "form formular", "letter brief schreiben", "confirmation bestaetigung", "appointment termin", "schedule plan zeitplan stundenplan",
    "tax steuer steuern steuererklaerung", "salary lohn gehalt", "payslip lohnabrechnung lohnausweis", "reminder mahnung erinnerung",
    "terms agb bedingungen", "report bericht", "notes notizen", "summary zusammenfassung", "feedback rueckmeldung", "questionnaire fragebogen umfrage",
    "coaching coach sitzung session", "training schulung weiterbildung kurs workshop", "doctor arzt aerztin arztzeugnis", "hospital spital klinik",
    "child kind kinder", "parents eltern", "rent miete mietvertrag", "travel reise", "ticket billett", "map karte plan", "photo foto"
  ].map(function (g) { return g.split(" "); });

  function normalisiere(s) {
    return s.toLowerCase().replace(/ä/g, "ae").replace(/ö/g, "oe").replace(/ü/g, "ue").replace(/ß/g, "ss")
      .replace(/[éèê]/g, "e").replace(/[àâ]/g, "a");
  }

  function dateiWoerter(name) {
    var ohneEndung = name.replace(/\.[a-z0-9]{1,5}$/i, "");
    var teile = ohneEndung
      .replace(/([a-zäöü])([A-ZÄÖÜ])/g, "$1 $2")            // KrankenkasseRechnung → Krankenkasse Rechnung
      .replace(/\d+/g, " ")                                // Praemie2026 → Praemie
      .split(/[\s_\-.,()[\]{}+#&]+/);
    return teile.map(normalisiere).filter(function (t) {
      return t.length >= 3 && !/\d/.test(t) && DATEI_ALLGEMEIN.indexOf(t) === -1;
    });
  }

  function passtWort(t, mailWoerter) {
    var kandidaten = [t];
    WORTGRUPPEN.forEach(function (g) {
      if (g.some(function (w) { return w === t || (w.length >= 5 && t.indexOf(w) === 0); })) kandidaten = kandidaten.concat(g);
    });
    return kandidaten.some(function (k) {
      return mailWoerter.some(function (m) {
        if (m === k) return true;
        var n = Math.min(m.length, k.length);
        if (n >= 5 && m.slice(0, 5) === k.slice(0, 5) && Math.abs(m.length - k.length) <= 3) return true; // Rechnung / Rechnungen
        if (k.length >= 4 && m.length > k.length && m.indexOf(k) !== -1) return true;   // Kasse ⊂ Krankenkasse
        if (m.length >= 5 && k.length > m.length && k.indexOf(m) !== -1) return true;   // Krankenkasse ⊂ Krankenkassenpraemie
        return false;
      });
    });
  }

  function pruefeAnhangPasst(text, betreff, dateinamen) {
    if (!dateinamen || !dateinamen.length) return [];
    var mailWoerter = normalisiere((betreff || "") + " " + text).split(/[^a-z]+/).filter(function (w) { return w.length >= 2; });
    if (mailWoerter.length < 5) return [];          // zu wenig Text zum Vergleichen
    var probleme = [];
    dateinamen.forEach(function (name) {
      var woerter = dateiWoerter(name);
      if (!woerter.length) return;                    // z. B. «Scan_2026.pdf», «IMG_1234.jpg»
      if (woerter.some(function (w) { return passtWort(w, mailWoerter); })) return;
      probleme.push({
        id: "anhang-passt-" + probleme.length,
        art: "hinweis",
        titel: "Passt der Anhang «" + kurz(name, 40) + "»?",
        detail: "Der Dateiname kommt weder im Betreff noch im Text vor. Ist das die richtige Datei?",
        hilfe: "Wenn die Datei stimmt, kannst du die Mail trotzdem senden."
      });
    });
    return probleme;
  }

  // ---------- 2. Anrede ----------
  var ANREDE_START = /^(sehr\s+geehrte[rs]?|geehrte[rs]?|liebe[rs]?|hallo|guten\s+(tag|morgen|abend)|grüezi|grüss\s+gott|grüß\s+gott|hoi|salü|sali|servus|moin|hi|hey|dear|hello|bonjour|chère?s?|buongiorno|ciao)\b/i;

  function pruefeAnrede(text, empfaengerNamen) {
    var probleme = [];
    var erste = ersteNichtLeereZeile(text);
    if (!erste) return probleme;

    if (!ANREDE_START.test(erste)) {
      probleme.push({
        id: "anrede-fehlt",
        art: "hinweis",
        titel: "Keine Anrede gefunden",
        detail: "Die Mail beginnt nicht mit einer Anrede wie «Liebe…», «Sehr geehrte…» oder «Hallo…».",
        fundstelle: kurz(erste, 80),
        hilfe: "Füge oben eine Anrede ein, z. B. «Liebe Frau Muster,»."
      });
      return probleme;
    }

    // Falsche Endungen: «Sehr geehrte Herr», «Liebe Herr», «Lieber Frau»
    var FALSCH = [
      { re: /\bsehr\s+geehrte\s+herr\b/i, alt: "Sehr geehrte Herr", neu: "Sehr geehrter Herr" },
      { re: /\bsehr\s+geehrter\s+frau\b/i, alt: "Sehr geehrter Frau", neu: "Sehr geehrte Frau" },
      { re: /\bliebe\s+herr\b/i, alt: "Liebe Herr", neu: "Lieber Herr" },
      { re: /\blieber\s+frau\b/i, alt: "Lieber Frau", neu: "Liebe Frau" },
      { re: /\bsehr\s+geehrten(?=\s+(damen|herren)\b)/i, alt: "Sehr geehrten", neu: "Sehr geehrte" },
      { re: /\bsehr\s+geerte?r?\b/i, alt: null, neu: null }
    ];
    FALSCH.forEach(function (f) {
      var m = f.re.exec(erste);
      if (!m) return;
      var neu = f.neu ? erste.replace(f.re, function (treffer) { return passeGrossschreibung(treffer, f.neu); }) : null;
      probleme.push({
        id: "anrede-endung",
        art: "fehler",
        titel: "Anrede falsch geschrieben",
        detail: f.neu ? "«" + m[0] + "» passt grammatikalisch nicht zusammen." : "«" + m[0] + "» ist falsch geschrieben (richtig: «geehrte»).",
        fundstelle: erste,
        markiert: m[0],
        ersetzen: neu ? { alt: erste, neu: neu } : null
      });
    });

    // Name in der Anrede passt zu keinem Empfänger?
    if (empfaengerNamen && empfaengerNamen.length) {
      var rest = erste.replace(ANREDE_START, "")
        .replace(/\b(herr|frau|dr\.?|prof\.?|liebe[rs]?|geehrte[rs]?|und|&|zusammen|alle|team|damen|herren|miteinander|mitenand)\b/gi, " ")
        .replace(/[,!.:;]/g, " ");
      var woerter = rest.split(/\s+/).filter(function (w) { return /^[A-ZÄÖÜ][\wäöüéèàçß-]{1,}$/.test(w); });
      if (woerter.length) {
        var namenText = empfaengerNamen.join(" ").toLowerCase();
        var unbekannt = woerter.filter(function (w) { return namenText.indexOf(w.toLowerCase()) === -1; });
        if (unbekannt.length === woerter.length && empfaengerNamen.join("").replace(/[^a-z]/gi, "").length > 0) {
          probleme.push({
            id: "anrede-name",
            art: "hinweis",
            titel: "Name in der Anrede prüfen",
            detail: "«" + unbekannt.join(" ") + "» kommt bei keinem Empfänger vor (" + kurz(empfaengerNamen.join(", "), 70) + "). Tippfehler oder falsche Person?",
            fundstelle: erste,
            markiert: unbekannt[0]
          });
        }
      }
    }

    // Komma nach der Anrede
    if (!/[,!:]\s*$/.test(erste) && erste.split(/\s+/).length <= 6) {
      probleme.push({
        id: "anrede-komma",
        art: "hinweis",
        titel: "Komma nach der Anrede fehlt",
        detail: "Nach einer Anrede steht normalerweise ein Komma.",
        fundstelle: erste,
        ersetzen: { alt: erste, neu: erste.replace(/[.;]?\s*$/, ",") }
      });
    }
    return probleme;
  }

  function passeGrossschreibung(original, neu) {
    if (original.charAt(0) === original.charAt(0).toLowerCase()) return neu.charAt(0).toLowerCase() + neu.slice(1);
    return neu;
  }

  // ---------- 3. Grussformel und Name ----------
  var GRUSS = /^(mit\s+)?(freundliche[nr]?|herzliche[nr]?|liebe[nr]?|beste[nr]?|viele[nr]?|sonnige[nr]?|schöne[nr]?|hochachtungsvoll|lg|mfg|vg|bg|hg|gruss|gruß|grüsse|grüße|grüessli|en\s+liebe\s+gruess|kind\s+regards|best\s+regards|regards|best|cheers|bis\s+bald|bis\s+dann|bis\s+morgen|bis\s+später|alles\s+liebe|alles\s+gute|herzlich)\b.{0,40}$/i;
  var GRUSS_FALSCH = [
    { re: /\bmit\s+freundliche\s+(grüsse|grüße|gruss|gruß)\b/i, neu: "Mit freundlichen Grüssen" },
    { re: /\bfreundlichen\s+(grüsse|grüße)\b(?!n)/i, neu: "Freundliche Grüsse", nurOhneMit: true },
    { re: /\bmit\s+herzliche\s+(grüsse|grüße)\b/i, neu: "Mit herzlichen Grüssen" },
    { re: /\bherzlichen\s+(grüsse|grüße)\b(?!n)/i, neu: "Herzliche Grüsse", nurOhneMit: true },
    { re: /\bmit\s+freundlichem\s+gruss?\b(?!e)/i, neu: "Mit freundlichem Gruss" }
  ];

  function pruefeGruss(text, einstellungen) {
    var probleme = [];
    var z = zeilen(text);
    var idx = -1;
    for (var i = z.length - 1; i >= 0; i--) {
      if (z[i] && GRUSS.test(z[i]) && z[i].length <= 45) { idx = i; break; }
    }

    if (idx === -1) {
      var inhalt = z.filter(Boolean);
      if (inhalt.length >= 3) {
        probleme.push({
          id: "gruss-fehlt",
          art: "hinweis",
          titel: "Keine Grussformel gefunden",
          detail: "Am Schluss fehlt eine Grussformel wie «Freundliche Grüsse» oder «Herzliche Grüsse».",
          fundstelle: kurz(inhalt[inhalt.length - 1], 80)
        });
      }
      return probleme;
    }

    var grussZeile = z[idx];
    GRUSS_FALSCH.forEach(function (f) {
      var m = f.re.exec(grussZeile);
      if (!m) return;
      if (f.nurOhneMit && /\bmit\b/i.test(grussZeile)) return;
      probleme.push({
        id: "gruss-falsch",
        art: "fehler",
        titel: "Grussformel falsch geschrieben",
        detail: "«" + m[0] + "» ist grammatikalisch falsch.",
        fundstelle: grussZeile,
        markiert: m[0],
        ersetzen: { alt: m[0], neu: f.neu }
      });
    });

    // Name nach der Grussformel?
    var name = (einstellungen.name || "").trim();
    var danach = z.slice(idx + 1).filter(Boolean).slice(0, 4);
    var namenGefunden = false;
    if (name) {
      var teile = name.toLowerCase().split(/\s+/);
      var block = (grussZeile + " " + danach.join(" ")).toLowerCase();
      namenGefunden = teile.some(function (t) { return t.length > 1 && block.indexOf(t) !== -1; });
    } else {
      namenGefunden = danach.length > 0;
    }
    if (!namenGefunden) {
      probleme.push({
        id: "name-fehlt",
        art: "fehler",
        titel: "Dein Name fehlt",
        detail: "Nach «" + kurz(grussZeile, 40) + "» steht " + (name ? "nicht «" + name + "»" : "kein Name") + ".",
        fundstelle: grussZeile,
        markiert: grussZeile,
        einfuegenNach: name ? { nach: grussZeile, text: name } : null
      });
    }
    return probleme;
  }

  // ---------- 4. Betreff ----------
  function pruefeBetreff(betreff) {
    if (betreff && betreff.trim()) return [];
    return [{
      id: "betreff",
      art: "fehler",
      titel: "Betreff ist leer",
      detail: "Die Mail hat keinen Betreff.",
      hilfe: "Schreibe oben in die Betreffzeile, worum es geht."
    }];
  }

  // ---------- 5. Rechtschreibung (lokal, nichts verlässt Outlook) ----------
  // Wörterbuch: igerman98 / Hunspell. Läuft vollständig im Add-in, ohne Internetdienst.
  var EIGENE_WOERTER = (
    "Coaching Coachings Coach Coaches Coachee Coachees Coachin Coachinnen coachen gecoacht Feedback Feedbacks " +
    "Meeting Meetings Workshop Workshops Team Teams Online online offline Zoom Session Sessions Mindset Burnout " +
    "Supervision Intervision Retreat Newsletter Podcast Webinar Webinare Link Links Website Webseite Homepage Login " +
    "Download Update Tool Tools Check-in Check-out Follow-up Call Calls Slot Slots Termin Termine Mail Mails E-Mail E-Mails " +
    "Grüessli Merci merci Hoi hoi Salü salü Sali sali Grüezi grüezi Ciao ciao Tschüss Tschau Znüni Zmittag Zvieri " +
    "Velo Natel Trottoir Billett Glace Poulet parkieren grillieren allfällig allfällige allfälligen allfälliger allfälliges " +
    "vorgängig vorgängige Pendenz Pendenzen Traktandum Traktanden traktandiert Couvert Spital Lernende Lernenden " +
    "OK ok okay Okay bzw usw usf etc ca inkl exkl evtl ev ggf vs Nr Tel Mob Hr Fr Dr Prof dipl lic MAS CAS DAS"
  ).split(" ");

  function erzeugeLader() {
    return function () {
      if (typeof HunspellLader === "undefined") return Promise.reject(new Error("kein Wörterbuch"));
      var basis = location.href.replace(/[?#].*$/, "").replace(/[^/]*$/, "");
      return HunspellLader.laden(basis).then(function (h) {
        if (!h.__eigeneGeladen) { EIGENE_WOERTER.forEach(function (w) { h.addWord(w); }); h.__eigeneGeladen = true; }
        return h;
      });
    };
  }
  var woerterbuch = { laden: erzeugeLader() };

  var BUCHSTABE = "A-Za-zÄÖÜäöüßÀ-ÖØ-öø-ÿ";
  var WORT_RE = new RegExp("[" + BUCHSTABE + "]+(?:[-'’][" + BUCHSTABE + "]+)*", "g");

  function grossAnfang(w) { return w.charAt(0).toUpperCase() + w.slice(1); }

  function istRichtig(h, w, schweiz) {
    if (w.length < 2) return true;
    if (h.spell(w)) return true;
    // Satzanfang: «Mit», «Heute» … auch klein geschrieben bekannt?
    var klein = w.charAt(0).toLowerCase() + w.slice(1);
    if (klein !== w && h.spell(klein)) return true;
    // Schweiz: «ss» statt «ß» (Grüsse → Grüße)
    if (schweiz && w.indexOf("ss") !== -1) {
      var teile = w.split("ss");
      var n = teile.length - 1;
      for (var maske = 1; maske < (1 << Math.min(n, 4)); maske++) {
        var v = teile[0];
        for (var i = 1; i <= n; i++) v += ((maske >> (i - 1)) & 1 ? "ß" : "ss") + teile[i];
        if (h.spell(v) || h.spell(v.charAt(0).toLowerCase() + v.slice(1))) return true;
      }
    }
    // Zusammengesetzte Wörter mit Fremdwörtern (z. B. Coachingsitzung, Feedbackrunde)
    if (w.length >= 8) {
      for (var j = 4; j <= w.length - 3; j++) {
        var vorne = w.slice(0, j);
        var hinten = grossAnfang(w.slice(j));
        if (h.spell(vorne) || EIGENE_WOERTER.indexOf(vorne) !== -1 || EIGENE_WOERTER.indexOf(vorne.replace(/s$/, "")) !== -1) {
          if (istRichtig(h, hinten, schweiz) || istRichtig(h, w.slice(j), schweiz)) return true;
        }
      }
    }
    return false;
  }

  function vorschlaegeFuer(h, w, schweiz) {
    var liste = h.suggest(w) || [];
    var aus = [];
    liste.forEach(function (v) {
      if (schweiz) v = v.replace(/ß/g, "ss");
      if (v !== w && aus.indexOf(v) === -1) aus.push(v);
    });
    return aus.slice(0, 3);
  }

  function fundstelleUm(text, start, laenge) {
    var anf = Math.max(0, text.lastIndexOf("\n", start) + 1);
    var end = text.indexOf("\n", start + laenge);
    if (end === -1) end = text.length;
    var satz = text.slice(anf, end);
    if (satz.length > 110) {
      var rel = start - anf;
      var s = Math.max(0, rel - 45);
      satz = (s > 0 ? "…" : "") + satz.slice(s, rel + laenge + 45) + (rel + laenge + 45 < satz.length ? "…" : "");
    }
    return satz.trim();
  }

  function pruefeRechtschreibung(text, einstellungen, extraIgnorieren, mitVorschlaegen) {
    if (!einstellungen.rechtschreibung || !text.trim()) return Promise.resolve([]);
    var schweiz = (einstellungen.sprache || "de-CH") === "de-CH";
    var ignor = {};
    (einstellungen.ignorierteWoerter || []).concat(extraIgnorieren || []).forEach(function (w) {
      String(w).split(/[\s,;<>@.()]+/).forEach(function (t) { if (t) ignor[t.toLowerCase()] = true; });
    });
    if (einstellungen.name) einstellungen.name.split(/\s+/).forEach(function (t) { ignor[t.toLowerCase()] = true; });

    // Signatur (alles nach der Grussformel: Name, Adresse, Telefon) nicht prüfen
    text = ohneSignatur(text);

    // Links, Mailadressen, Zahlen und Abkürzungen ausblenden
    var sauber = text
      .replace(/\b(https?:\/\/|www\.)\S+/gi, function (m) { return " ".repeat(m.length); })
      .replace(/\S+@\S+/g, function (m) { return " ".repeat(m.length); });

    return woerterbuch.laden().then(function (h) {
      var probleme = [];
      var gesehen = {};
      var m;
      WORT_RE.lastIndex = 0;
      while ((m = WORT_RE.exec(sauber)) && probleme.length < 25) {
        var wort = m[0];
        if (ignor[wort.toLowerCase()] || gesehen[wort]) continue;
        if (wort.length > 1 && wort === wort.toUpperCase()) continue;   // ABK, CEO
        var nachher = sauber.charAt(m.index + wort.length);
        if (nachher === "." && wort.length <= 4) continue;              // Abkürzung «Tel.», «bzw.»
        var start = m.index;
        // Schweiz: «ß» ist falsch
        if (schweiz && wort.indexOf("ß") !== -1) {
          gesehen[wort] = true;
          probleme.push({
            id: "rs-" + probleme.length, art: "fehler",
            titel: "«ß» in der Schweiz: «" + wort + "»",
            detail: "In der Schweiz schreibt man «ss» statt «ß».",
            fundstelle: fundstelleUm(text, start, wort.length), markiert: wort,
            vorschlaege: [wort.replace(/ß/g, "ss")], ersetzenWort: wort
          });
          continue;
        }
        var teile = wort.split(/[-'’]/);
        var falsch = !teile.every(function (t) { return ignor[t.toLowerCase()] || istRichtig(h, t, schweiz); });
        if (falsch && istRichtig(h, wort, schweiz)) falsch = false;
        if (!falsch) continue;
        gesehen[wort] = true;
        probleme.push({
          id: "rs-" + probleme.length, art: "fehler",
          titel: "Rechtschreibung: «" + kurz(wort, 30) + "»",
          detail: "Dieses Wort steht nicht im Wörterbuch. Tippfehler? (Namen und Fachwörter kannst du mit «Wort merken» speichern.)",
          fundstelle: fundstelleUm(text, start, wort.length), markiert: wort,
          vorschlaege: mitVorschlaegen ? vorschlaegeFuer(h, wort, schweiz) : [], ersetzenWort: wort
        });
      }
      return probleme;
    }).catch(function () {
      return [{
        id: "rs-aus", art: "info",
        titel: "Rechtschreibprüfung nicht verfügbar",
        detail: "Das Wörterbuch konnte nicht geladen werden. Anhang, Anrede und Grussformel wurden trotzdem geprüft."
      }];
    });
  }

  // ---------- 6. Grammatik & Schreibweise (lokale Regeln, wie Word) ----------
  var ABKUERZUNGEN = ("z b d h u a s o v g ca bzw usw usf inkl exkl zzgl ggf evtl bspw etc vs nr tel mob dr fr hr prof st " +
    "mio mrd std min max jh geb ehem abs art kap bd ff vgl sog resp spez allg betr frdl lic dipl").split(" ");

  function ohneSignatur(text) {
    var zz = text.split("\n");
    for (var g = zz.length - 1; g >= 0; g--) {
      var zt = zz[g].trim();
      if (zt && zt.length <= 45 && GRUSS.test(zt)) return zz.slice(0, g + 1).join("\n");
    }
    return text;
  }

  // Feste Wendungen und häufige Verwechslungen
  var WENDUNGEN = [
    { re: /\bim voraus\b/i, gross: "voraus", warum: "«Voraus» schreibt man in dieser Wendung gross." },
    { re: /\bdes weiteren\b/i, gross: "weiteren", warum: "«Weiteren» schreibt man hier gross." },
    { re: /\bbis auf weiteres\b/i, gross: "weiteres", warum: "«Weiteres» schreibt man hier gross." },
    { re: /\bohne weiteres\b/i, gross: "weiteres", warum: "«Weiteres» schreibt man hier gross." },
    { re: /\bim allgemeinen\b/i, gross: "allgemeinen", warum: "«Allgemeinen» schreibt man hier gross." },
    { re: /\bim nachhinein\b/i, gross: "nachhinein", warum: "«Nachhinein» schreibt man gross." },
    { re: /\bim folgenden\b/i, gross: "folgenden", warum: "«Folgenden» schreibt man hier gross." },
    { re: /\b(vielen|herzlichen|besten|tausend) dank\b/i, neu: null, gross: "dank", warum: "«Dank» ist ein Nomen und wird gross geschrieben." },
    { re: /\bseid (gestern|vorgestern|heute|wann|langem|kurzem|einigen|einer|einem|dem|der|letzte[mnr]?|jahren|wochen|monaten|tagen|anfang)\b/i, wort: "seid", neu: "seit", warum: "Zeitangabe: «seit» (mit t). «seid» kommt von «sein» (ihr seid)." },
    { re: /\bihr seit\b/i, wort: "seit", neu: "seid", warum: "«ihr seid» (von «sein») schreibt man mit d." },
    { re: /\b(ich|er|man) weis\b/i, wort: "weis", neu: "weiss", warum: "Von «wissen»: «ich weiss» (in Deutschland «weiß»)." },
    { re: /\bweis (ich|nicht)\b/i, wort: "weis", neu: "weiss", warum: "Von «wissen»: «weiss» (in Deutschland «weiß»)." },
    { re: /\bnicht war\?/i, wort: "war", neu: "wahr", warum: "«nicht wahr?» schreibt man mit h." },
    { re: /\b(ist|sein|wäre) (?:doch |ja |eben |halt )?war\b(?=[\s.,!?]|$)/i, wort: "war", neu: "wahr", warum: "Im Sinn von «richtig»: «wahr» mit h." },
    { re: /\bwider (sehen|melden|zurück|da|einmal|mal|gut|gesund|treffen|hören)\b/i, wort: "wider", neu: "wieder", warum: "Im Sinn von «noch einmal»: «wieder» mit ie." },
    { re: /\bauf widersehen\b/i, wort: "widersehen", neu: "Wiedersehen", warum: "«Auf Wiedersehen» schreibt man mit ie." },
    { re: /\bso das (ich|du|er|sie|es|wir|ihr|man)\b/i, wort: "so das", neu: "sodass", warum: "Als Folge («so dass»): «sodass» oder «so dass» – nie «das» mit einem s." },
    { re: /\b(damit|ohne|als|bevor|nachdem|weil) das (ich|du|er|wir|ihr|man)\b/i, wort: "das", neu: "dass", warum: "Hier ist die Konjunktion gemeint: «dass» mit ss." },
    { re: /\b(hoffe|hoffen|glaube|glauben|denke|denken|weiss|weiß|wissen|finde|finden|meine|befürchte|vermute|erwarte|sicher|froh|wichtig|gesagt|klar)\s*, das (ich|du|er|sie|es|wir|ihr|man) ([^.,!?\n]+ )?(kann|könnte|muss|müsste|soll|sollte|werde|wird|würde|habe|hat|hätte|bin|ist|wäre|sind|darf|will|wollte|können|müssen|werden|haben)(?=[\s.,!?]|$)/i, wort: ", das", neu: ", dass", warum: "Leitet «das» hier einen Nebensatz ein (ersetzbar durch «dass»)? Dann mit ss.", vorsicht: true },
  ];

  function grammatikProblem(text, start, laenge, titel, warum, alt, neu, art) {
    var stelle = fundstelleUm(text, start, laenge);
    var markiert = text.substr(start, laenge);
    var p = {
      id: "gr-" + start, art: art || "fehler", titel: titel, detail: warum,
      fundstelle: stelle, markiert: markiert.trim()
    };
    if (neu != null) p.ersetzen = { alt: alt, neu: neu, ganzesWort: /^[A-Za-zÄÖÜäöüß]/.test(alt) && /[A-Za-zÄÖÜäöüß]$/.test(alt) };
    return p;
  }

  function pruefeGrammatik(text, einstellungen) {
    if (!einstellungen.rechtschreibung || !text.trim()) return [];
    text = ohneSignatur(text);
    var sauber = text
      .replace(/\b(https?:\/\/|www\.)\S+/gi, function (m) { return " ".repeat(m.length); })
      .replace(/\S+@\S+/g, function (m) { return " ".repeat(m.length); });
    var probleme = [];
    var m, re;

    // 1. Doppelte Wörter «die die», «und und»
    re = new RegExp("(^|[^" + BUCHSTABE + "])([" + BUCHSTABE + "]{2,})([ \\t]+)\\2(?![" + BUCHSTABE + "])", "gi");
    while ((m = re.exec(sauber))) {
      var s1 = m.index + m[1].length;
      var doppelt = m[2] + m[3] + sauber.substr(s1 + m[2].length + m[3].length, m[2].length);
      probleme.push(grammatikProblem(text, s1, doppelt.length, "Doppeltes Wort: «" + m[2] + "»",
        "Das Wort steht zweimal hintereinander.", doppelt, m[2],
        /^(das|die|der|den|dem|des|sie)$/i.test(m[2]) ? "hinweis" : "fehler"));
      re.lastIndex = s1 + m[2].length;
    }

    // 2. Satzanfang klein nach . ! ?
    re = new RegExp("([" + BUCHSTABE + "0-9]*)([.!?])(\\s+)([a-zäöü][" + BUCHSTABE + "]*)", "g");
    while ((m = re.exec(sauber))) {
      var vorher = m[1], zeichen = m[2];
      if (zeichen === ".") {
        if (/^\d+$/.test(vorher)) continue;                        // «15. oktober» → Datum/Ordnungszahl
        if (ABKUERZUNGEN.indexOf(vorher.toLowerCase()) !== -1) continue;
        if (sauber.charAt(m.index + vorher.length - 1) === ".") continue; // «...»
        if (!vorher) continue;
      }
      var wortStart = m.index + vorher.length + 1 + m[3].length;
      var w = m[4];
      var gross = w.charAt(0).toUpperCase() + w.slice(1);
      probleme.push(grammatikProblem(text, wortStart, w.length, "Satzanfang: «" + w + "»",
        "Nach einem Satzende schreibt man gross weiter.", zeichen + m[3] + w, zeichen + m[3] + gross));
      probleme[probleme.length - 1].ersetzen.anzeige = gross;
    }

    // 3. Höflichkeitsform: «ihre», «ihnen» klein in einer Sie-Mail
    var erste = ersteNichtLeereZeile(text);
    var foermlich = /\b(sehr geehrte|geehrte|liebe frau|lieber herr|guten tag (frau|herr)|grüezi (frau|herr))\b/i.test(erste) ||
      /[a-zäöü,] (Sie|Ihnen)\b/.test(text);
    if (foermlich) {
      re = /(^|[^A-Za-zÄÖÜäöüß])(ihnen|ihre[nmrs]?)(?![A-Za-zÄÖÜäöüß])/g;
      while ((m = re.exec(sauber))) {
        var hs = m.index + m[1].length, hw = m[2];
        probleme.push(grammatikProblem(text, hs, hw.length, "Höflichkeitsform: «" + hw + "»",
          "Du siezt in dieser Mail. Ist die angeschriebene Person gemeint, schreibt man «" + hw.charAt(0).toUpperCase() + hw.slice(1) + "» gross.",
          hw, hw.charAt(0).toUpperCase() + hw.slice(1), "hinweis"));
      }
    }

    // 4. Feste Wendungen und Verwechslungen
    WENDUNGEN.forEach(function (r) {
      var g = new RegExp(r.re.source, r.re.flags.indexOf("g") === -1 ? r.re.flags + "g" : r.re.flags);
      var t;
      while ((t = g.exec(sauber))) {
        var ganz = t[0];
        var alt, neu, start, laenge;
        if (r.gross) {                       // nur ein Wort gross schreiben
          var pos = ganz.toLowerCase().lastIndexOf(r.gross);
          if (ganz.charAt(pos) !== r.gross.charAt(0)) continue;   // schon gross
          alt = ganz; neu = ganz.slice(0, pos) + r.gross.charAt(0).toUpperCase() + ganz.slice(pos + 1);
          start = t.index + pos; laenge = r.gross.length;
        } else if (r.wort) {                 // ein Wort in der Wendung ersetzen
          var p2 = ganz.toLowerCase().indexOf(r.wort.toLowerCase());
          var original = ganz.substr(p2, r.wort.length);
          var ersatz = original.charAt(0) === original.charAt(0).toUpperCase() && r.neu.charAt(0) === r.neu.charAt(0).toLowerCase()
            ? r.neu.charAt(0).toUpperCase() + r.neu.slice(1) : r.neu;
          alt = ganz.slice(0, p2 + r.wort.length).replace(/^, /, ", ");
          neu = ganz.slice(0, p2) + ersatz;
          start = t.index + p2; laenge = r.wort.length;
        } else {
          alt = ganz; neu = r.neu; start = t.index; laenge = ganz.length;
          if (alt === neu) continue;
        }
        probleme.push(grammatikProblem(text, start, laenge, "Grammatik: «" + text.substr(start, laenge) + "»",
          r.warum, alt, neu, r.vorsicht ? "hinweis" : "fehler"));
      }
    });

    // 5. Leerzeichen bei Satzzeichen
    re = /([A-Za-zÄÖÜäöüß]+)[ \t]+([,.!?;])(?=\s|$)/g;
    while ((m = re.exec(sauber))) {
      probleme.push(grammatikProblem(text, m.index, m[0].length, "Leerzeichen vor «" + m[2] + "»",
        "Vor einem Satzzeichen steht kein Leerzeichen.", m[0], m[1] + m[2], "hinweis"));
    }
    re = /([a-zäöüß]{2,})([,.])([A-ZÄÖÜa-zäöü][a-zäöüß]{2,})/g;
    while ((m = re.exec(sauber))) {
      if (m[2] === "." && /^[a-z]/.test(m[3])) continue;     // Dateinamen, Domains
      probleme.push(grammatikProblem(text, m.index, m[0].length, "Leerzeichen nach «" + m[2] + "» fehlt",
        "Nach einem Komma oder Punkt folgt ein Leerzeichen.", m[0], m[1] + m[2] + " " + m[3], "hinweis"));
    }

    // Gleiche Stelle nicht mehrfach melden
    var gesehen = {};
    return probleme.filter(function (p) {
      var k = p.fundstelle + "|" + p.markiert;
      if (gesehen[k]) return false;
      gesehen[k] = true;
      return true;
    }).slice(0, 20);
  }

  // ---------- Alles zusammen ----------
  function pruefeAlles(daten) {
    // daten: { text, betreff, anzahlAnhaenge, dateinamen[], empfaengerNamen[], einstellungen }
    var e = daten.einstellungen || ladeEinstellungen();
    var t = eigenerText(daten.text);
    var lokal = []
      .concat(pruefeBetreff(daten.betreff))
      .concat(pruefeAnhang(t, daten.anzahlAnhaenge || 0))
      .concat(pruefeAnhangPasst(t, daten.betreff, daten.dateinamen || []))
      .concat(pruefeAnrede(t, daten.empfaengerNamen || []))
      .concat(pruefeGruss(t, e))
      .concat(pruefeGrammatik(t, e));
    return pruefeRechtschreibung(t, e, daten.empfaengerNamen, daten.mitVorschlaegen).then(function (lt) {
      // Doppelte Meldungen (z. B. Grussformel) nicht zweimal zeigen
      var schon = lokal.map(function (p) { return (p.markiert || "").toLowerCase(); });
      lt = lt.filter(function (p) {
        return !p.markiert || !schon.some(function (s) { return s && s.indexOf(p.markiert.toLowerCase()) !== -1; });
      });
      return lokal.concat(lt);
    });
  }

  function zusammenfassung(probleme) {
    var relevant = probleme.filter(function (p) { return p.art !== "info"; });
    if (!relevant.length) return "";
    var teile = relevant.slice(0, 5).map(function (p) { return "• " + p.titel; });
    if (relevant.length > 5) teile.push("• … und " + (relevant.length - 5) + " weitere");
    var txt = "Bitte prüfe deine Mail:\n" + teile.join("\n") + "\n\nKlicke auf «Fehler ansehen» für Details und Korrekturvorschläge.";
    return txt.length > 495 ? txt.slice(0, 494) + "…" : txt;
  }

  // ---------- Daten aus Outlook holen ----------
  function async(fn) {
    return new Promise(function (ok, fehler) {
      fn(function (r) {
        if (r.status === Office.AsyncResultStatus.Succeeded) ok(r.value); else fehler(r.error);
      });
    });
  }

  function holeMailDaten() {
    var item = Office.context.mailbox.item;
    var text = async(function (cb) { item.body.getAsync(Office.CoercionType.Text, cb); });
    var betreff = async(function (cb) { item.subject.getAsync(cb); });
    var anh = async(function (cb) { item.getAttachmentsAsync(cb); }).catch(function () { return []; });
    var an = async(function (cb) { item.to.getAsync(cb); }).catch(function () { return []; });
    var cc = async(function (cb) { item.cc.getAsync(cb); }).catch(function () { return []; });
    return Promise.all([text, betreff, anh, an, cc]).then(function (w) {
      var anhaenge = (w[2] || []).filter(function (a) { return !a.isInline; });
      var namen = (w[3] || []).concat(w[4] || []).map(function (r) {
        return (r.displayName || "") + " " + (r.emailAddress || "").split("@")[0].replace(/[._-]/g, " ");
      });
      return { text: w[0], betreff: w[1], anzahlAnhaenge: anhaenge.length, dateinamen: anhaenge.map(function (a) { return a.name || ""; }), empfaengerNamen: namen };
    });
  }

  return {
    ladeEinstellungen: ladeEinstellungen,
    speichereEinstellungen: speichereEinstellungen,
    eigenerText: eigenerText,
    pruefeAnhangPasst: pruefeAnhangPasst,
    pruefeGrammatik: pruefeGrammatik,
    pruefeAlles: pruefeAlles,
    zusammenfassung: zusammenfassung,
    holeMailDaten: holeMailDaten,
    async: async,
    woerterbuch: woerterbuch
  };
})();

// ---------- Prüfung beim Klick auf «Senden» ----------
function beimSenden(event) {
  MailPruefer.holeMailDaten().then(function (daten) {
    daten.einstellungen = MailPruefer.ladeEinstellungen();
    return MailPruefer.pruefeAlles(daten);
  }).then(function (probleme) {
    var text = MailPruefer.zusammenfassung(probleme);
    if (!text) {
      event.completed({ allowEvent: true });
      return;
    }
    event.completed({
      allowEvent: false,
      errorMessage: text,
      cancelLabel: "Fehler ansehen",
      commandId: "mailprueferOeffnen",
      sendModeOverride: Office.MailboxEnums && Office.MailboxEnums.SendModeOverride
        ? Office.MailboxEnums.SendModeOverride.PromptUser : undefined
    });
  }).catch(function () {
    // Wenn die Prüfung selbst scheitert, nie das Senden blockieren
    event.completed({ allowEvent: true });
  });
}

if (typeof Office !== "undefined" && Office.actions && Office.actions.associate) {
  Office.actions.associate("beimSenden", beimSenden);
}
if (typeof module !== "undefined") module.exports = MailPruefer;
