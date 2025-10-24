/**
 * Utility functions for handling late registration assignments.
 *
 * These helpers are designed to be shared across multiple Sheets via the
 * shared library. The main entry point is `late_onEditHandler`, which can
 * be called from an onEdit trigger. It watches a specific column (the
 * "Équipe" column by default) and, when a value is assigned, sends an
 * email notification to a list of recipients. The email includes the
 * player's name, passport number and assigned team, and uses the email
 * address of the user who made the edit as both the sender and reply‑to.
 *
 * To use this in a bound spreadsheet script:
 *   1. Define a configuration object with at least `sheetName` and
 *      `recipients` set. Other keys default to sensible French column
 *      headers.
 *   2. Create an onEdit trigger that calls `SI_LIB.late_onEditHandler(e, cfg)`
 *      where `cfg` is your configuration object. `SI_LIB` is the name
 *      you assign to the library when adding it to your project.
 *
 * Example local wrapper in a bound script:
 *
 * function onEdit(e) {
 *   const cfg = {
 *     sheetName: 'Inscriptions',
 *     recipients: ['responsable1@exemple.com','responsable2@exemple.com'],
 *     colTeam: 'Équipe',
 *     colPassport: 'Passeport #',
 *     colPrenom: 'Prénom',
 *     colNom: 'Nom',
 *     colMailSent: 'Courriel envoyé' // optional
 *   };
 *   SI_LIB.late_onEditHandler(e, cfg);
 * }
 */

/**
 * Handler for onEdit triggers to detect team assignments and send an email
 * notification.  It accepts the onEdit event object and a configuration
 * object. If the edited cell is in the configured team column and the
 * value is non‑empty, it will gather the row data and delegate to
 * `late_sendAssignmentEmail`.
 *
 * @param {Object} e   The onEdit event object (Apps Script passes this).
 * @param {Object} cfg Configuration object with at least the following
 *                      properties:
 *                      - sheetName   {string} Name of the sheet to monitor.
 *                      - recipients  {Array<string>} List of email
 *                                     addresses to notify.
 *                      - colTeam     {string} Header for the team column.
 *                                     Defaults to 'Équipe'.
 *                      - colPassport {string} Header for passport column.
 *                                     Defaults to 'Passeport #'.
 *                      - colPrenom   {string} Header for first name column.
 *                                     Defaults to 'Prénom'.
 *                      - colNom      {string} Header for last name column.
 *                                     Defaults to 'Nom'.
 *                      - colMailSent {string} (optional) Header for column
 *                                     where a timestamp should be written
 *                                     after sending the email.
 */
function late_onEditHandler(e, cfg) {
  if (!e || !cfg) return;
  var range = e.range;
  // Only react to single‑cell edits
  if (!range || range.getNumRows() !== 1 || range.getNumColumns() !== 1) return;
  var sheet = range.getSheet();
  var sheetName = String(cfg.sheetName || 'Inscriptions');
  if (!sheet || sheet.getName() !== sheetName) return;
  var row = range.getRow();
  // Ignore header row
  if (row === 1) return;

  // Retrieve headers
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  var colTeamName = String(cfg.colTeam || 'Équipe');
  var colIndexTeam = headers.indexOf(colTeamName);
  if (colIndexTeam < 0) return;
  // Only react if the edited cell is in the team column
  if (range.getColumn() !== colIndexTeam + 1) return;

  var newVal = String(range.getValue() || '').trim();
  // Do nothing if the cell was cleared
  if (!newVal) return;

  // Collect row data into an object keyed by header names
  var rowValues = sheet.getRange(row, 1, 1, headers.length).getValues()[0];
  var data = {};
  for (var i = 0; i < headers.length; i++) {
    data[headers[i]] = rowValues[i];
  }
  data.__row = row;
  late_sendAssignmentEmail(data, cfg);
}

/**
 * Sends an email notification when a team assignment occurs.  Uses the
 * configuration object to determine column names, recipients and
 * optionally updates a 'Courriel envoyé' column with a timestamp.
 *
 * @param {Object} rowData  An object containing the row values keyed by
 *                          header names, as built in late_onEditHandler.
 * @param {Object} cfg      The configuration object (see above).
 */

function late_sendAssignmentEmail(rowData, cfg) {
  if (!rowData || !cfg) return;

  var team      = String(rowData[cfg.colTeam || 'Équipe'] || '').trim();
  var prenom    = String(rowData[cfg.colPrenom || 'Prénom'] || '').trim();
  var nom       = String(rowData[cfg.colNom || 'Nom'] || '').trim();
  var passport  = String(rowData[cfg.colPassport || 'Passeport #'] || '').trim();
  if (!team || !prenom || !nom || !passport) return;

  var recipients = Array.isArray(cfg.recipients) ? cfg.recipients : [];
  if (!recipients.length) return;
  var to = recipients.filter(function(addr){ return addr && String(addr).trim(); }).join(',');
  if (!to) return;

  // Classeur/feuille actifs
  var ss, sheet, ssId, gid, rowNum, a1, rowLink, sheetLink;
  try {
    ss     = SpreadsheetApp.getActive();
    sheet  = ss.getSheetByName(String(cfg.sheetName || 'Inscriptions'));
    ssId   = ss.getId();
    gid    = sheet ? sheet.getSheetId() : null;
    rowNum = Number(rowData.__row || 0);

    // Viser la cellule de la colonne Équipe sur la bonne ligne (sinon fallback colonne A)
    var headers = sheet ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String) : [];
    var colTeamIdx = headers.indexOf(String(cfg.colTeam || 'Équipe')) + 1; // 1-based
    var colLetter  = colTeamIdx > 0 ? late_columnLetterFromIndex_(colTeamIdx) : 'A';

    a1 = (rowNum > 1) ? (colLetter + rowNum + ':' + colLetter + rowNum) : '';
    var base = 'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(ssId) + '/edit';
    sheetLink = ssId ? base : '';
    rowLink   = (ssId && gid != null && a1) ? (base + '#gid=' + gid + '&range=' + encodeURIComponent(a1)) : '';
  } catch (_) {
    // pas de lien si échec
    rowLink = '';
    sheetLink = '';
  }

  // Courriel
  var sender = '';
  try { sender = Session.getActiveUser().getEmail(); } catch (_) { sender = ''; }

  var subject = 'Assignation d’équipe – ' + prenom + ' ' + nom;
  var passportDisplay = passport.replace(/^'/, '');

  var bodyLines = [
    'Bonjour,',
    '',
    'Un joueur vient d’être assigné :',
    '',
    'Nom : ' + prenom + ' ' + nom,
    'Passeport : ' + passportDisplay,
    'Équipe assignée : ' + team,
    '',
    rowLink   ? ('Lien (ligne) : ' + rowLink)   : '',
    sheetLink ? ('Feuille : ' + sheetLink)      : '',
    '',
    '— ' + (sender || '')
  ].filter(Boolean);

  try {
    MailApp.sendEmail({
      to: to,
      subject: subject,
      body: bodyLines.join('\n'),
      name: sender || undefined,
      replyTo: sender || undefined
    });
  } catch (err) {
    console.error('late_sendAssignmentEmail error:', err);
    throw err;
  }

  // Horodater l’envoi (optionnel)
  if (cfg.colMailSent && sheet && rowNum > 1) {
    var allHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
    var colMailSentIdx = allHeaders.indexOf(String(cfg.colMailSent));
    if (colMailSentIdx >= 0) {
      var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'America/Toronto', 'yyyy-MM-dd HH:mm');
      sheet.getRange(rowNum, colMailSentIdx + 1).setValue(stamp);
    }
  }
}

function late_columnLetterFromIndex_(idx1based) {
  // idx1based: 1=A, 2=B, ...
  var s = '';
  var n = Number(idx1based || 0);
  while (n > 0) {
    var r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s || 'A';
}
