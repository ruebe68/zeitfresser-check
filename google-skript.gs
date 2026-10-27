// Zeitfresser-Check · Ablage der Antworten in dieser Google-Tabelle
// Einfügen unter Erweiterungen → Apps Script, dann als Web-App bereitstellen
// (Ausführen als: Ich · Zugriff: Jeder). Die Web-App-Adresse kommt in index.html (ENDPOINT).

const BLATT = 'Antworten';

// Spalte → Feld im Bogen. Neue Fragen hier ergänzen, die Spalte wird automatisch angehängt.
const SPALTEN = [
  ['ID', 'id'], ['Eingang', '_eingang'], ['Aktualisiert', 'aktualisiert'], ['Status', 'status'],
  ['Firma', 'firma'], ['Name', 'person'], ['Rolle', 'rolle'], ['Branche', 'branche'], ['Anrede', 'anrede'],
  ['Stunden pro Jahr (Wiederholungen)', 'stundenProJahr'], ['Schritt', 'schritt'],
  ['Arbeitstag', 'a.alltag'], ['Bürozeit pro Woche', 'a.buerozeit'],
  ['Was nervt', 'a.nervt'], ['Letzte nervige Situation', 'a.nervt_konkret'],
  ['Wiederkehrende Aufgaben', '_aufgaben'],
  ['Größter Zeitfresser', 'a.zeitfresser'], ['Wird aufgeschoben', 'a.aufschieben'],
  ['Wo geht Geld verloren', 'a.geld'], ['Wo mehr rausholen', 'a.mehr_raus'],
  ['Kanäle', 'a.kanaele'], ['Anfragen pro Woche', 'a.anfragen_woche'], ['Immer gleiche Fragen', 'a.fragen_immer'],
  ['Was bleibt liegen bei Ausfall', 'a.ausfall'], ['Teamgröße', 'a.team'], ['Aufwand ums Team', 'a.team_aufwand'],
  ['Programme', 'a.werkzeuge'], ['Daten in den Abläufen', 'a.daten'],
  ['Nie allein', 'a.nie_allein'], ['Gern allein', 'a.gern_allein'], ['Offenheit (1–5)', 'a.offenheit'],
  ['Perfekter Arbeitstag', 'a.perfekter_tag'], ['Mit 10 Stunden mehr', 'a.zehn_stunden'],
  ['Nie wieder selbst machen', 'a.wunsch']
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = JSON.parse(e.postData.contents);
    if (!p || typeof p.id !== 'string' || !/^[a-z0-9-]{8,80}$/.test(p.id)) return antwort({ ok: false });
    const blatt = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BLATT)
      || SpreadsheetApp.getActiveSpreadsheet().insertSheet(BLATT);
    const kopf = kopfzeile(blatt);
    const zeile = kopf.map(function (titel) { return wert(p, titel); });
    const n = blatt.getLastRow();
    const ids = n > 1 ? blatt.getRange(2, 1, n - 1, 1).getValues().map(function (r) { return r[0]; }) : [];
    const i = ids.indexOf(p.id);
    const spalteEingang = kopf.indexOf('Eingang');
    if (i >= 0) {
      zeile[spalteEingang] = blatt.getRange(i + 2, spalteEingang + 1).getValue();
      blatt.getRange(i + 2, 1, 1, zeile.length).setValues([zeile]);
    } else {
      zeile[spalteEingang] = new Date();
      blatt.appendRow(zeile);
    }
    return antwort({ ok: true });
  } catch (err) {
    return antwort({ ok: false });
  } finally {
    lock.releaseLock();
  }
}

function kopfzeile(blatt) {
  let kopf = blatt.getLastRow() ? blatt.getRange(1, 1, 1, blatt.getLastColumn()).getValues()[0] : [];
  const fehlend = SPALTEN.map(function (s) { return s[0]; }).filter(function (t) { return kopf.indexOf(t) < 0; });
  if (fehlend.length) {
    blatt.getRange(1, kopf.length + 1, 1, fehlend.length).setValues([fehlend]).setFontWeight('bold');
    blatt.setFrozenRows(1);
    kopf = kopf.concat(fehlend);
  }
  return kopf;
}

function wert(p, titel) {
  const s = SPALTEN.filter(function (x) { return x[0] === titel; })[0];
  if (!s) return '';
  const feld = s[1];
  let v;
  if (feld === '_eingang') return '';
  if (feld === '_aufgaben') {
    v = (p.aufgaben || []).filter(function (t) { return t && t.name; }).map(function (t) {
      return '• ' + t.name + ' · ' + t.freq + ' · ' + (t.min || '?') + ' Min.';
    }).join('\n');
  } else if (feld.indexOf('a.') === 0) {
    v = (p.antworten || {})[feld.slice(2)];
    if (v && typeof v === 'object') {
      v = (v.auswahl || []).join(', ') + (v.sonst ? ((v.auswahl || []).length ? '; ' : '') + v.sonst : '');
    }
  } else {
    v = p[feld];
  }
  if (v === undefined || v === null) return '';
  if (typeof v === 'number') return v;
  v = String(v).slice(0, 5000);
  // Schutz vor Formeln in Freitext
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function antwort(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
