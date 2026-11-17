/**
 * Loulou Studio · Reformer Teacher Training · inschrijvingsflow
 *
 * 1. Website-formulier -> doPost: rij in de Sheet + mail naar Loulou + bevestiging naar inschrijver.
 * 2. In de Sheet: menu "Loulou" -> betalingsverzoek voorschot (50%) als PDF naar de deelnemer.
 * 3. Later: menu "Loulou" -> betalingsverzoek saldo (50%) als PDF naar de deelnemer.
 * Status en datums worden telkens in de Sheet bijgehouden.
 */

// ---------- INSTELLINGEN (invullen) ----------
var CFG = {
  loulouMail:   'info@louloustudio.be',
  afzender:     'Loulou Studio',
  bedrijf:      'Loulou Studio',                 // juridische naam (eenmanszaak)
  adres:        'Abdijmolenstraat 30, 9031 Drongen',
  btw:          'BE 0000.000.000',               // <-- invullen
  iban:         'BE00 0000 0000 0000',           // <-- invullen
  bic:          'XXXXBEBB',                      // <-- invullen
  editie:       'april 2027',
  data:         '3 & 4, 17 & 18 en 24 & 25 april 2027',
  prijsEarly:   1000,      // excl. btw
  prijsNormaal: 1250,      // excl. btw
  earlyTot:     new Date(2027, 0, 15, 23, 59),   // 15 januari 2027
  earlyPlaatsen: 5,
  btwPct:       0.21,
  voorschotNietTerug: 500,
  saldoTegen:   new Date(2027, 2, 6),            // 4 weken voor 3 april 2027
  betaaltermijnDagen: 7,
  refPrefix:    'TT27-'
};
var SHEET = 'Inschrijvingen';
var KOL = ['Datum','Editie','Voornaam','Achternaam','E-mail','Telefoon','Leeftijd','Woonplaats','Ervaring','Lesgeven',
           'Statuut','Factuur','Motivatie','Tarief','Referentie','Status','Voorschot gevraagd','Saldo gevraagd'];
var C = {};  KOL.forEach(function (k, i) { C[k] = i + 1; });

var BORDEAUX = '#59010E', LINNEN = '#F6F2EB', INK = '#141110', MUTED = '#4A423D', TAUPE = '#8E705B', LIJN = '#E4DBCE';

// ---------- 1. FORMULIER ----------
function doPost(e) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.website) return json({ ok: true });
    var req = ['voornaam','achternaam','email','telefoon','leeftijd','ervaring','lesgeven','statuut','motivatie'];
    for (var i = 0; i < req.length; i++) if (!d[req[i]]) return json({ ok: false, error: 'missing ' + req[i] });
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(d.email)) return json({ ok: false, error: 'email' });

    var sh = sheet();
    var nu = new Date();
    var early = nu <= CFG.earlyTot && telEarly(sh) < CFG.earlyPlaatsen;
    var nr = sh.getLastRow();                         // volgnummer = aantal inschrijvingen
    var ref = CFG.refPrefix + ('00' + nr).slice(-3);
    var rij = [nu, CFG.editie, d.voornaam, d.achternaam, d.email, d.telefoon, d.leeftijd, d.woonplaats || '', d.ervaring,
               d.lesgeven, d.statuut, d.factuur || '', d.motivatie, early ? 'Early bird' : 'Normaal', ref, 'Nieuw', '', ''];
    sh.appendRow(rij);
    d.tarief = early ? 'Early bird' : 'Normaal'; d.ref = ref;
    mailLoulou(d); mailBevestiging(d);
    return json({ ok: true });
  } catch (err) { return json({ ok: false, error: String(err) }); }
  finally { lock.releaseLock(); }
}

function sheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow(KOL); sh.getRange(1, 1, 1, KOL.length).setFontWeight('bold').setBackground(LIJN); sh.setFrozenRows(1);
    sh.getRange(2, C['Status'], 500, 1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['Nieuw','Voorschot gevraagd','Voorschot betaald','Saldo gevraagd','Volledig betaald','Geannuleerd','Wachtlijst'], true).build());
    sh.getRange(2, C['Tarief'], 500, 1).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(['Early bird','Normaal'], true).build());
  }
  return sh;
}
function telEarly(sh) {
  if (sh.getLastRow() < 2) return 0;
  var v = sh.getRange(2, C['Tarief'], sh.getLastRow() - 1, 2).getValues();
  return v.filter(function (r) { return r[0] === 'Early bird'; }).length;
}

// ---------- 2 & 3. BETALINGSVERZOEKEN (menu in de Sheet) ----------
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Loulou')
    .addItem('Stuur betalingsverzoek voorschot (geselecteerde rij)', 'stuurVoorschot')
    .addItem('Stuur betalingsverzoek saldo (geselecteerde rij)', 'stuurSaldo')
    .addSeparator()
    .addItem('Voorbeeld-PDF voorschot maken (niet versturen)', 'voorbeeldPdf')
    .addToUi();
}
function stuurVoorschot() { verzoek('voorschot', true); }
function stuurSaldo()     { verzoek('saldo', true); }
function voorbeeldPdf()   { verzoek('voorschot', false); }

function verzoek(soort, versturen) {
  var ui = SpreadsheetApp.getUi(), sh = sheet(), r = sh.getActiveRange().getRow();
  if (r < 2) { ui.alert('Selecteer eerst een rij met een deelnemer.'); return; }
  var v = sh.getRange(r, 1, 1, KOL.length).getValues()[0];
  var p = { voornaam: v[C['Voornaam']-1], achternaam: v[C['Achternaam']-1], email: v[C['E-mail']-1],
            woonplaats: v[C['Woonplaats']-1], factuur: v[C['Factuur']-1], tarief: v[C['Tarief']-1], ref: v[C['Referentie']-1] };
  var b = bedragen(p.tarief);
  var bestand = 'Betalingsverzoek ' + p.ref + ' ' + soort + ' - ' + p.voornaam + ' ' + p.achternaam + '.pdf';
  var pdf = Utilities.newBlob(pdfHtml(p, soort, b), 'text/html', 'x.html').getAs('application/pdf').setName(bestand);

  if (!versturen) {
    var f = DriveApp.createFile(pdf);
    ui.alert('Voorbeeld opgeslagen in je Google Drive:\n' + f.getName()); return;
  }
  var bedrag = soort === 'voorschot' ? b.voorschot : b.saldo;
  if (ui.alert('Betalingsverzoek ' + soort + ' (' + eur(bedrag) + ') sturen naar ' + p.email + '?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  MailApp.sendEmail({ to: p.email, replyTo: CFG.loulouMail, bcc: CFG.loulouMail, name: CFG.afzender,
    subject: (soort === 'voorschot' ? 'Betalingsverzoek voorschot' : 'Betalingsverzoek saldo') + ' · Reformer Teacher Training · ' + p.ref,
    htmlBody: mailVerzoek(p, soort, bedrag), attachments: [pdf] });
  sh.getRange(r, C['Status']).setValue(soort === 'voorschot' ? 'Voorschot gevraagd' : 'Saldo gevraagd');
  sh.getRange(r, C[soort === 'voorschot' ? 'Voorschot gevraagd' : 'Saldo gevraagd']).setValue(new Date());
  ui.alert('Verstuurd naar ' + p.email + '.');
}

function bedragen(tarief) {
  var excl = tarief === 'Early bird' ? CFG.prijsEarly : CFG.prijsNormaal;
  var btw = Math.round(excl * CFG.btwPct * 100) / 100, incl = excl + btw;
  var voorschot = Math.round(incl * 50) / 100;
  return { excl: excl, btw: btw, incl: incl, voorschot: voorschot, saldo: Math.round((incl - voorschot) * 100) / 100 };
}

function pdfHtml(p, soort, b) {
  var vandaag = new Date();
  var verval = soort === 'voorschot' ? new Date(vandaag.getTime() + CFG.betaaltermijnDagen * 864e5) : CFG.saldoTegen;
  var teBetalen = soort === 'voorschot' ? b.voorschot : b.saldo;
  var mededeling = p.ref + ' ' + p.voornaam + ' ' + p.achternaam;
  var td = 'padding:9px 0;border-bottom:1px solid ' + LIJN + ';font-size:12px;';
  return '<!doctype html><html><body style="font-family:Helvetica,Arial,sans-serif;color:' + INK + ';margin:0">' +
  '<div style="background:' + BORDEAUX + ';color:#fff;padding:28px 40px">' +
    '<table style="width:100%;color:#fff"><tr><td style="font-family:Georgia,serif;font-size:34px;color:#fff">loulou<div style="font-family:Helvetica,Arial;font-size:9px;letter-spacing:3px;opacity:.8">PILATES STUDIO</div></td>' +
    '<td style="text-align:right;font-size:11px;letter-spacing:2px;color:#fff">BETALINGSVERZOEK<div style="font-family:Georgia,serif;font-size:20px;letter-spacing:0;margin-top:4px">' + esc(p.ref) + (soort === 'saldo' ? '-S' : '-V') + '</div></td></tr></table></div>' +
  '<div style="padding:30px 40px">' +
    '<table style="width:100%;font-size:11px;color:' + MUTED + '"><tr>' +
      '<td style="vertical-align:top;width:50%"><div style="letter-spacing:2px;color:' + TAUPE + '">VAN</div><div style="margin-top:6px;color:' + INK + ';font-size:12px;line-height:1.6"><b>' + esc(CFG.bedrijf) + '</b><br>' + esc(CFG.adres) + '<br>Btw ' + esc(CFG.btw) + '<br>' + esc(CFG.loulouMail) + '</div></td>' +
      '<td style="vertical-align:top"><div style="letter-spacing:2px;color:' + TAUPE + '">AAN</div><div style="margin-top:6px;color:' + INK + ';font-size:12px;line-height:1.6"><b>' + esc(p.voornaam + ' ' + p.achternaam) + '</b><br>' + esc(p.woonplaats || '') + (p.factuur ? '<br>' + esc(p.factuur) : '') + '<br>' + esc(p.email) + '</div></td>' +
    '</tr></table>' +
    '<table style="width:100%;font-size:11px;color:' + MUTED + ';margin-top:22px"><tr>' +
      '<td>Datum<br><b style="color:' + INK + ';font-size:12px">' + dat(vandaag) + '</b></td>' +
      '<td>Te betalen voor<br><b style="color:' + INK + ';font-size:12px">' + dat(verval) + '</b></td>' +
      '<td>Referentie<br><b style="color:' + INK + ';font-size:12px">' + esc(p.ref) + '</b></td></tr></table>' +
    '<table style="width:100%;border-collapse:collapse;margin-top:28px">' +
      '<tr><td style="' + td + 'color:' + TAUPE + ';letter-spacing:2px;font-size:10px">OMSCHRIJVING</td><td style="' + td + 'text-align:right;color:' + TAUPE + ';letter-spacing:2px;font-size:10px">BEDRAG</td></tr>' +
      '<tr><td style="' + td + '"><b>Reformer Teacher Training · ' + esc(CFG.editie) + '</b><br><span style="color:' + MUTED + '">' + esc(CFG.data) + ' · tarief ' + esc(p.tarief.toLowerCase()) + '</span></td><td style="' + td + 'text-align:right">' + eur(b.excl) + '</td></tr>' +
      '<tr><td style="' + td + 'color:' + MUTED + '">Btw 21%</td><td style="' + td + 'text-align:right;color:' + MUTED + '">' + eur(b.btw) + '</td></tr>' +
      '<tr><td style="' + td + '"><b>Totaal opleiding incl. btw</b></td><td style="' + td + 'text-align:right"><b>' + eur(b.incl) + '</b></td></tr>' +
      (soort === 'saldo' ? '<tr><td style="' + td + 'color:' + MUTED + '">Reeds betaald (voorschot 50%)</td><td style="' + td + 'text-align:right;color:' + MUTED + '">– ' + eur(b.voorschot) + '</td></tr>' : '') +
    '</table>' +
    '<div style="background:' + LINNEN + ';margin-top:24px;padding:20px 24px"><table style="width:100%"><tr>' +
      '<td style="font-size:11px;color:' + MUTED + ';letter-spacing:2px">' + (soort === 'voorschot' ? 'NU TE BETALEN · 50% VOORSCHOT' : 'NU TE BETALEN · SALDO 50%') + '</td>' +
      '<td style="text-align:right;font-family:Georgia,serif;font-size:30px;color:' + BORDEAUX + '">' + eur(teBetalen) + '</td></tr></table></div>' +
    '<table style="width:100%;margin-top:22px;font-size:12px;line-height:1.7"><tr>' +
      '<td style="vertical-align:top;width:50%"><div style="font-size:10px;letter-spacing:2px;color:' + TAUPE + '">BETAALGEGEVENS</div>IBAN <b>' + esc(CFG.iban) + '</b><br>BIC ' + esc(CFG.bic) + '<br>Begunstigde ' + esc(CFG.bedrijf) + '<br>Mededeling <b>' + esc(mededeling) + '</b></td>' +
      '<td style="vertical-align:top;color:' + MUTED + ';font-size:11px">' +
        (soort === 'voorschot'
          ? 'Van dit voorschot is ' + eur(CFG.voorschotNietTerug) + ' niet terugbetaalbaar. Je plaats is definitief zodra we je betaling ontvangen. Het saldo van ' + eur(b.saldo) + ' betaal je uiterlijk ' + dat(CFG.saldoTegen) + '.'
          : 'Met deze betaling is je opleiding volledig betaald. We kijken ernaar uit je te verwelkomen op ' + esc(CFG.data.split(',')[0]) + '.') +
        '<br><br>Annuleren kan tot 2 maanden voor de start, met terugbetaling min het voorschot van ' + eur(CFG.voorschotNietTerug) + '. Binnen 2 maanden voor de start is annuleren niet meer mogelijk.</td>' +
    '</tr></table>' +
    '<div style="margin-top:30px;font-size:9px;color:' + TAUPE + '">Dit is een betalingsverzoek, geen factuur. Je factuur ontvang je na betaling.</div>' +
  '</div></body></html>';
}

// ---------- MAILS ----------
function kader(inhoud) {
  return '<div style="background:' + LINNEN + ';padding:32px 16px;font-family:Helvetica,Arial,sans-serif">' +
    '<div style="max-width:560px;margin:0 auto;background:#fff;border:1px solid ' + LIJN + ';border-radius:6px;overflow:hidden">' +
    '<div style="background:' + BORDEAUX + ';padding:26px 32px;color:#fff;font-family:Georgia,serif;font-size:30px">loulou' +
    '<div style="font-family:Helvetica,Arial,sans-serif;font-size:10px;letter-spacing:.3em;margin-top:4px;opacity:.8">PILATES STUDIO</div></div>' +
    '<div style="padding:32px">' + inhoud + '</div>' +
    '<div style="padding:18px 32px;border-top:1px solid ' + LIJN + ';color:' + TAUPE + ';font-size:12px">' + esc(CFG.bedrijf) + ' · ' + esc(CFG.adres) + ' · louloustudio.be</div></div></div>';
}
function rij(l, w) {
  return '<tr><td style="padding:10px 0;border-bottom:1px solid ' + LIJN + ';color:' + MUTED + ';font-size:12px;letter-spacing:.12em;text-transform:uppercase;width:42%;vertical-align:top">' + l +
         '</td><td style="padding:10px 0;border-bottom:1px solid ' + LIJN + ';color:' + INK + ';font-size:15px">' + esc(w || '–') + '</td></tr>';
}
function gegevens(d) {
  return '<table style="width:100%;border-collapse:collapse;margin-top:18px">' + rij('Naam', d.voornaam + ' ' + d.achternaam) + rij('E-mail', d.email) +
    rij('Telefoon', d.telefoon) + rij('Leeftijd', d.leeftijd) + rij('Woonplaats', d.woonplaats) + rij('Pilates-ervaring', d.ervaring) +
    rij('Daarna lesgeven', d.lesgeven) + rij('Statuut', d.statuut) + rij('Factuur', d.factuur) + rij('Motivatie', d.motivatie) + '</table>';
}
function mailLoulou(d) {
  var html = kader('<div style="color:' + TAUPE + ';font-size:12px;letter-spacing:.2em;text-transform:uppercase">Nieuwe inschrijving · ' + esc(d.ref) + ' · ' + esc(d.tarief) + '</div>' +
    '<h1 style="font-family:Georgia,serif;font-weight:normal;color:' + BORDEAUX + ';font-size:28px;margin:10px 0 0">' + esc(d.voornaam + ' ' + d.achternaam) + '</h1>' + gegevens(d) +
    '<p style="color:' + MUTED + ';font-size:14px;margin-top:22px">Volgende stap: bel of mail ' + esc(d.voornaam) + ', en stuur daarna in de Google Sheet via menu <b>Loulou → Stuur betalingsverzoek voorschot</b>.</p>');
  MailApp.sendEmail({ to: CFG.loulouMail, replyTo: d.email, name: CFG.afzender, subject: 'Nieuwe inschrijving Teacher Training: ' + d.voornaam + ' ' + d.achternaam, htmlBody: html });
}
function mailBevestiging(d) {
  var html = kader('<h1 style="font-family:Georgia,serif;font-weight:normal;color:' + BORDEAUX + ';font-size:28px;margin:0">Bedankt voor je inschrijving, ' + esc(d.voornaam) + '</h1>' +
    '<p style="color:' + MUTED + ';font-size:15px;line-height:1.6;margin-top:16px">We hebben je inschrijving voor de <b>Reformer Teacher Training (' + esc(CFG.editie) + ')</b> goed ontvangen. We nemen binnen enkele werkdagen persoonlijk contact met je op. Daarna ontvang je een betalingsverzoek voor het voorschot.</p>' +
    '<p style="color:' + MUTED + ';font-size:15px;line-height:1.6">Je plaats is definitief na betaling van 50% (waarvan ' + eur(CFG.voorschotNietTerug) + ' niet-terugbetaalbaar voorschot). De overige 50% betaal je uiterlijk 4 weken voor de start.</p>' +
    '<div style="color:' + TAUPE + ';font-size:12px;letter-spacing:.2em;text-transform:uppercase;margin-top:26px">Jouw gegevens</div>' + gegevens(d) +
    '<p style="color:' + MUTED + ';font-size:15px;margin-top:24px">Vragen? Antwoord gewoon op deze mail.</p><p style="color:' + INK + ';font-size:15px;margin-top:6px">Tot snel,<br>Loulou Studio</p>');
  MailApp.sendEmail({ to: d.email, replyTo: CFG.loulouMail, name: CFG.afzender, subject: 'Je inschrijving voor de Reformer Teacher Training', htmlBody: html });
}
function mailVerzoek(p, soort, bedrag) {
  return kader('<h1 style="font-family:Georgia,serif;font-weight:normal;color:' + BORDEAUX + ';font-size:26px;margin:0">Hoi ' + esc(p.voornaam) + ',</h1>' +
    '<p style="color:' + MUTED + ';font-size:15px;line-height:1.6;margin-top:16px">' +
    (soort === 'voorschot'
      ? 'Leuk dat je erbij bent! In de bijlage vind je het betalingsverzoek voor het voorschot van je Reformer Teacher Training. Zodra we <b>' + eur(bedrag) + '</b> ontvangen, is je plaats definitief.'
      : 'De start van je Reformer Teacher Training komt dichterbij! In de bijlage vind je het betalingsverzoek voor het saldo van <b>' + eur(bedrag) + '</b>, te betalen uiterlijk ' + dat(CFG.saldoTegen) + '.') +
    '</p><p style="color:' + MUTED + ';font-size:15px;line-height:1.6">Vermeld bij je overschrijving de mededeling <b>' + esc(p.ref + ' ' + p.voornaam + ' ' + p.achternaam) + '</b>.</p>' +
    '<p style="color:' + INK + ';font-size:15px;margin-top:20px">Tot snel,<br>Loulou Studio</p>');
}

// ---------- HULP ----------
function eur(n) { return '€' + Number(n).toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.').replace(',00', ''); }
function dat(d) { var m = ['januari','februari','maart','april','mei','juni','juli','augustus','september','oktober','november','december']; return d.getDate() + ' ' + m[d.getMonth()] + ' ' + d.getFullYear(); }
function esc(s) { return String(s == null ? '' : s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/\n/g,'<br>'); }
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
