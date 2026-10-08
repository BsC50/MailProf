# Mail-Prüfer – Outlook-Add-in

Prüft jede Mail beim Klick auf **Senden**. Findet es etwas, hält Outlook die Mail an und zeigt
«Fehler ansehen» (öffnet die Seitenleiste mit Korrekturvorschlägen) oder «Trotzdem senden».

**Was geprüft wird**
- Anhang erwähnt («im Anhang», «anbei», «beigefügt» …), aber keine Datei angehängt
- Angehängte Datei passt nicht zur Mail: kommt der Dateiname (z. B. «school.pdf») weder im Betreff noch im Text vor,
  fragt der Prüfer nach. Englische und deutsche Wörter werden verknüpft (school = Schule, invoice = Rechnung).
  Allgemeine Namen wie «Scan_2026.pdf» oder «IMG_1234.jpg» werden übergangen. Den Inhalt der Datei liest der Prüfer nicht.
- Anrede fehlt, ist falsch («Sehr geehrte Herr», «Lieber Frau»), Name passt zu keinem Empfänger, Komma fehlt
- Grussformel falsch («Mit freundliche Grüsse») oder dein Name fehlt darunter
- Leerer Betreff
- Rechtschreibung mit eingebautem Wörterbuch (Schweizer Schreibweise: «ß» wird als Fehler gemeldet)
- Grammatik und Schreibweise mit festen Regeln, ähnlich wie in Word: Satzanfang klein, doppelte Wörter,
  Höflichkeitsform («ihnen» → «Ihnen» in Sie-Mails), «seid/seit», «das/dass» in typischen Fällen,
  «im voraus» → «im Voraus», «ich weis» → «ich weiss», Leerzeichen bei Satzzeichen u. a.

Zitierter Verlauf (alte Mails unterhalb von «Von: … Gesendet: …») wird nicht geprüft.

## Einrichten (einmalig, ca. 15 Minuten)

Outlook lädt Add-ins von einer Internetadresse. Die Dateien kommen deshalb kostenlos auf GitHub Pages.

1. **GitHub-Konto anlegen** auf github.com (kostenlos). Merke dir deinen Benutzernamen.
2. **Neues Repository**: oben rechts «+» → «New repository» → Name `mailpruefer` → «Public» → «Create repository».
3. **Dateien hochladen**: Auf der Repository-Seite «uploading an existing file» klicken, den *Inhalt* dieses
   Ordners (alle Dateien und den Ordner `assets`) hineinziehen → «Commit changes».
4. **Seite einschalten**: «Settings» → «Pages» → bei «Branch» `main` wählen → «Save».
   Nach 1–2 Minuten ist die Seite unter `https://DEIN-NAME.github.io/mailpruefer/` erreichbar.
5. **Manifest anpassen**: Claude deinen GitHub-Benutzernamen schicken – du bekommst eine passende `manifest.xml`.
6. **In Outlook installieren**: Im Browser `https://aka.ms/olksideload` öffnen → «Meine Add-Ins» →
   «Benutzerdefiniertes Add-In hinzufügen» → «Aus Datei hinzufügen» → `manifest.xml` wählen → «Installieren».
   Das gilt dann für Outlook im Web und das neue Outlook auf dem Computer.

## Benutzen
- Einfach Mails schreiben und senden. Ist alles in Ordnung, geht die Mail ohne Unterbrechung raus.
- Beim Verfassen gibt es in der Symbolleiste den Knopf **Mail prüfen** – damit kannst du jederzeit vorher prüfen.
- In der Seitenleiste: grüne Knöpfe übernehmen die Korrektur direkt im Text. «Wort merken» für Fachbegriffe.
- Zahnrad oben rechts: Name, Sprache, eigene Wörter, Rechtschreibprüfung ein/aus.

## Datenschutz
Alle Prüfungen laufen direkt in deinem Outlook – auch die Rechtschreibung (eingebautes Wörterbuch).
Der Mailtext wird an keinen fremden Dienst geschickt und vom Mail-Prüfer nirgends gespeichert.
- Auf GitHub liegen nur die Programmdateien, nie Mails.
- Gespeichert werden nur deine Einstellungen (Name, Sprache, gemerkte Wörter) – in deinem eigenen Outlook-Postfach.
  Tipp: Namen von Klient:innen besser nicht mit «Wort merken» speichern, sondern die Meldung einfach übergehen.

## Hinweise
- Bei einem Geschäftskonto (Microsoft 365 der Firma) kann die IT eigene Add-ins gesperrt haben.
- Die Grammatikprüfung arbeitet mit festen Regeln für die häufigsten Fehler. Sie versteht Sätze nicht so umfassend
  wie Word (z. B. «ich habe gegangen» oder jedes «das/dass» erkennt sie nicht), schickt dafür aber nichts ins Internet.
- Die Signatur unter der Grussformel (Name, Adresse, Telefon) wird bei der Rechtschreibung übersprungen.
- Das Wörterbuch (igerman98, GPL) und die Prüf-Software (hunspell-asm, MIT) liegen mit ihren Lizenzen im Ordner `woerterbuch`.
