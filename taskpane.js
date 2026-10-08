/* Mail-Prüfer – Seitenleiste */
/* global Office, MailPruefer */
(function () {
  "use strict";
  var $ = function (id) { return document.getElementById(id); };
  var einstellungen;

  Office.onReady(function () {
    einstellungen = MailPruefer.ladeEinstellungen();
    $("btn-pruefen").onclick = pruefen;
    $("btn-einstellungen").onclick = zeigeEinstellungen;
    $("btn-zurueck").onclick = zeigePruefung;
    $("btn-speichern").onclick = speichern;
    pruefen();
  });

  // ---------- Prüfen und anzeigen ----------
  function pruefen() {
    setzeStatus("Prüfe deine Mail …");
    $("liste").innerHTML = "";
    MailPruefer.holeMailDaten().then(function (daten) {
      daten.einstellungen = einstellungen;
      daten.mitVorschlaegen = true;
      return MailPruefer.pruefeAlles(daten);
    }).then(zeige).catch(function (e) {
      setzeStatus("Die Mail konnte nicht gelesen werden: " + (e && e.message ? e.message : e));
    });
  }

  function zeige(probleme) {
    var liste = $("liste");
    liste.innerHTML = "";
    var echte = probleme.filter(function (p) { return p.art !== "info"; });
    if (!echte.length) setzeStatus("✓ Alles in Ordnung – die Mail kann raus.", true);
    else setzeStatus(echte.length === 1 ? "1 Punkt zum Prüfen:" : echte.length + " Punkte zum Prüfen:");
    probleme.forEach(function (p) { liste.appendChild(karte(p)); });
  }

  function setzeStatus(text, ok) {
    var s = $("status");
    s.textContent = text;
    s.className = "status" + (ok ? " ok" : "");
  }

  function el(tag, klasse, text) {
    var e = document.createElement(tag);
    if (klasse) e.className = klasse;
    if (text != null) e.textContent = text;
    return e;
  }

  function karte(p) {
    var li = el("li", "karte " + p.art);
    li.appendChild(el("span", "etikett", p.art === "fehler" ? "Fehler" : p.art === "info" ? "Info" : "Hinweis"));
    li.appendChild(el("h3", null, p.titel));
    if (p.detail) li.appendChild(el("p", null, p.detail));

    if (p.fundstelle) {
      var f = el("div", "fundstelle");
      var pos = p.markiert ? p.fundstelle.indexOf(p.markiert) : -1;
      if (pos >= 0) {
        f.appendChild(document.createTextNode(p.fundstelle.slice(0, pos)));
        f.appendChild(el("mark", null, p.markiert));
        f.appendChild(document.createTextNode(p.fundstelle.slice(pos + p.markiert.length)));
      } else {
        f.textContent = p.fundstelle;
      }
      li.appendChild(f);
    }
    if (p.hilfe) li.appendChild(el("p", "hilfe", p.hilfe));

    var akt = el("div", "aktionen");
    if (p.vorschlaege && p.vorschlaege.length) {
      p.vorschlaege.forEach(function (v) {
        var b = el("button", "klein vorschlag", "→ " + v);
        b.title = "«" + p.ersetzenWort + "» durch «" + v + "» ersetzen";
        b.onclick = function () { korrigiere(li, function (doc) { return ersetzeText(doc, p.ersetzenWort, v, true); }); };
        akt.appendChild(b);
      });
    }
    if (p.ersetzen) {
      var b2 = el("button", "klein vorschlag", "→ " + kurz(p.ersetzen.anzeige || p.ersetzen.neu, 40));
      b2.onclick = function () { korrigiere(li, function (doc) { return ersetzeText(doc, p.ersetzen.alt, p.ersetzen.neu, !!p.ersetzen.ganzesWort); }); };
      akt.appendChild(b2);
    }
    if (p.einfuegenNach) {
      var b3 = el("button", "klein vorschlag", "«" + p.einfuegenNach.text + "» einfügen");
      b3.onclick = function () { korrigiere(li, function (doc) { return fuegeNamenEin(doc, p.einfuegenNach.nach, p.einfuegenNach.text); }); };
      akt.appendChild(b3);
    }
    if (p.ersetzenWort) {
      var b4 = el("button", "klein", "Wort merken");
      b4.title = "Dieses Wort künftig nie mehr als Fehler melden";
      b4.onclick = function () {
        einstellungen.ignorierteWoerter = (einstellungen.ignorierteWoerter || []).concat([p.ersetzenWort]);
        MailPruefer.speichereEinstellungen(einstellungen);
        li.classList.add("erledigt");
        akt.innerHTML = "";
        akt.appendChild(el("span", "hilfe", "Gemerkt – wird nicht mehr gemeldet."));
      };
      akt.appendChild(b4);
    }
    if (akt.childNodes.length) li.appendChild(akt);
    return li;
  }

  function kurz(s, n) { return s.length > n ? s.slice(0, n - 1) + "…" : s; }

  // ---------- Korrektur direkt im Mailtext ----------
  function korrigiere(li, aenderung) {
    var item = Office.context.mailbox.item;
    MailPruefer.async(function (cb) { item.body.getAsync(Office.CoercionType.Html, cb); }).then(function (html) {
      var doc = new DOMParser().parseFromString(html, "text/html");
      if (!aenderung(doc)) {
        var hinweis = li.querySelector(".aktionen");
        hinweis.innerHTML = "";
        hinweis.appendChild(el("span", "hilfe", "Diese Stelle konnte ich nicht automatisch ändern – bitte direkt im Text korrigieren."));
        return null;
      }
      var neu = "<!DOCTYPE html>" + doc.documentElement.outerHTML;
      return MailPruefer.async(function (cb) { item.body.setAsync(neu, { coercionType: Office.CoercionType.Html }, cb); });
    }).then(function (r) {
      if (r !== null) pruefen();
    }).catch(function (e) {
      setzeStatus("Änderung fehlgeschlagen: " + (e && e.message ? e.message : e));
    });
  }

  function textKnoten(doc) {
    var walker = doc.createTreeWalker(doc.body, NodeFilter.SHOW_TEXT, null);
    var knoten = [], n;
    while ((n = walker.nextNode())) {
      var p = n.parentNode && n.parentNode.nodeName;
      if (p !== "STYLE" && p !== "SCRIPT") knoten.push(n);
    }
    return knoten;
  }

  function maskiere(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"); }
  function normal(s) { return s.replace(/ /g, " "); }

  function ersetzeText(doc, alt, neu, ganzesWort) {
    var re = ganzesWort
      ? new RegExp("(^|[^\\wäöüÄÖÜß])(" + maskiere(alt).replace(/ /g, "[ \\u00a0]+") + ")(?![\\wäöüÄÖÜß])")
      : new RegExp("()(" + maskiere(alt).replace(/ /g, "[ \\u00a0]+") + ")");
    var knoten = textKnoten(doc);
    for (var i = 0; i < knoten.length; i++) {
      var t = knoten[i].nodeValue;
      var m = re.exec(t);
      if (m) {
        var start = m.index + m[1].length;
        knoten[i].nodeValue = t.slice(0, start) + neu + t.slice(start + m[2].length);
        return true;
      }
    }
    return false;
  }

  function fuegeNamenEin(doc, grussZeile, name) {
    var knoten = textKnoten(doc);
    var ziel = normal(grussZeile).trim();
    for (var i = knoten.length - 1; i >= 0; i--) {
      if (normal(knoten[i].nodeValue).trim() === ziel || normal(knoten[i].nodeValue).indexOf(ziel) !== -1) {
        var n = knoten[i];
        var block = n.parentNode;
        // Steht die Grussformel allein in einem Absatz, einen neuen Absatz mit dem Namen anlegen
        if (block && /^(P|DIV)$/.test(block.nodeName) && normal(block.textContent).trim() === ziel) {
          var kopie = block.cloneNode(false);
          kopie.textContent = name;
          block.parentNode.insertBefore(kopie, block.nextSibling);
        } else {
          var br = doc.createElement("br");
          n.parentNode.insertBefore(br, n.nextSibling);
          br.parentNode.insertBefore(doc.createTextNode(name), br.nextSibling);
        }
        return true;
      }
    }
    return false;
  }

  // ---------- Einstellungen ----------
  function zeigeEinstellungen() {
    $("e-name").value = einstellungen.name || "";
    $("e-rs").checked = !!einstellungen.rechtschreibung;
    $("e-sprache").value = einstellungen.sprache || "de-CH";
    $("e-woerter").value = (einstellungen.ignorierteWoerter || []).join("\n");
    $("ansicht-pruefung").hidden = true;
    $("ansicht-einstellungen").hidden = false;
  }
  function zeigePruefung() {
    $("ansicht-einstellungen").hidden = true;
    $("ansicht-pruefung").hidden = false;
  }
  function speichern() {
    einstellungen.name = $("e-name").value.trim();
    einstellungen.rechtschreibung = $("e-rs").checked;
    einstellungen.sprache = $("e-sprache").value;
    einstellungen.ignorierteWoerter = $("e-woerter").value.split(/\n+/).map(function (w) { return w.trim(); }).filter(Boolean);
    MailPruefer.speichereEinstellungen(einstellungen, function () {
      zeigePruefung();
      pruefen();
    });
  }
})();
