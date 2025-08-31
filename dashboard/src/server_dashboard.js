/** ---------- CONFIG ---------- */
const DASH_SHEETS = {
  INSCRIPTIONS : 'INSCRIPTIONS',
  ARTICLES     : 'ARTICLES',
  ERREURS      : 'ERREURS',
  OUTBOX       : 'MAIL_OUTBOX',
  IMPORT_LOG   : 'IMPORT_LOG'
};

function openSeasonSpreadsheet_(){
  // Utilise TA fonction existante fournie dans le projet
  const ssId = getSeasonId_();
  return SpreadsheetApp.openById(ssId);
}

/** Renvoie le nombre de lignes non vides (hors en-tête) d’une feuille si elle existe. */
function countRows_(ss, name){
  const sh = ss.getSheetByName(name);
  if(!sh) return 0;
  const last = sh.getLastRow();
  if(last <= 1) return 0;
  return last - 1; // on assume en-tête à la ligne 1
}

/** Lis un tableau (2D) avec header + rows. Renvoie {headers, rows} */
function readTable_(ss, name){
  const sh = ss.getSheetByName(name);
  if(!sh) return { headers: [], rows: [] };
  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if(lr < 1 || lc < 1) return { headers: [], rows: [] };
  const values = sh.getRange(1,1,lr,lc).getValues();
  if(values.length === 0) return { headers: [], rows: [] };
  const headers = (values.shift() || []).map(h => String(h||'').trim());
  return { headers, rows: values };
}

function findCol_(headers, candidates){
  const set = headers.map(h => h.toString().trim().toLowerCase());
  for (const name of candidates){
    const i = set.indexOf(name.toLowerCase());
    if(i >= 0) return i;
  }
  return -1;
}

/**
 * KPI pour l’accueil (onglet Saison).
 */
function getDashboardMetrics(){
  const ss = openSeasonSpreadsheet_();

  const inscriptionsTotal = countRows_(ss, DASH_SHEETS.INSCRIPTIONS);
  const articlesTotal     = countRows_(ss, DASH_SHEETS.ARTICLES);

  // ERREURS
  let erreursTotal = 0, erreursDernierImport = null, errSubtitle = 'Dernier import';
  const errT = readTable_(ss, DASH_SHEETS.ERREURS);
  erreursTotal = errT.rows.length;

  if (errT.rows.length > 0){
    const dateCol = findCol_(errT.headers, ['date','timestamp','time','datetime']);
    if (dateCol >= 0){
      const map = {};
      errT.rows.forEach(r=>{
        const cell = r[dateCol];
        const iso = (cell instanceof Date)
          ? cell.toISOString().slice(0,10)
          : String(cell||'').trim();
        map[iso] = (map[iso]||0) + 1;
      });
      const dates = Object.keys(map).sort();
      const last = dates[dates.length-1];
      erreursDernierImport = map[last];
      errSubtitle = last ? `Le ${last}` : 'Dernier import';
    }
  }

  // OUTBOX
  let outboxPending = 0, outboxSubtitle = 'Prêts à l’envoi';
  const outT = readTable_(ss, DASH_SHEETS.OUTBOX);
  if (outT.rows.length){
    const statusCol = findCol_(outT.headers, ['status','etat','state']);
    const sentAtCol = findCol_(outT.headers, ['sent_at','envoye_le','sent','date_envoi']);
    outboxPending = outT.rows.filter(r=>{
      const st = statusCol >= 0 ? String(r[statusCol]||'').toUpperCase() : '';
      const sent = sentAtCol >= 0 ? r[sentAtCol] : '';
      const isSent = (st === 'SENT' || st === 'ENVOYE' || (sent && String(sent).trim() !== ''));
      return !isSent;
    }).length;
  }

  return {
    inscriptionsTotal,
    inscriptionsSubtitle: 'Total cumul.',
    articlesTotal,
    articlesSubtitle: 'Actifs',
    erreursTotal,
    erreursDernierImport,
    erreursSubtitle: errSubtitle,
    outboxPending,
    outboxSubtitle
  };
}

/**
 * Activité récente (10 lignes) : essaie IMPORT_LOG puis fallback ERREURS.
 */
function getRecentActivity(){
  const ss = openSeasonSpreadsheet_();

  // IMPORT_LOG prioritaire
  const logT = readTable_(ss, DASH_SHEETS.IMPORT_LOG);
  if (logT.rows.length){
    const cDate = findCol_(logT.headers, ['date','timestamp','time','datetime']);
    const cType = findCol_(logT.headers, ['type','level','categorie']);
    const cMsg  = findCol_(logT.headers, ['message','details','detail','info','texte']);
    return logT.rows.slice(-10).reverse().map(r=>({
      date: formatDate_(cDate>=0 ? r[cDate] : ''),
      type: cType>=0 ? String(r[cType]||'') : '',
      details: cMsg>=0 ? String(r[cMsg]||'') : ''
    }));
  }

  // Fallback ERREURS
  const errT = readTable_(ss, DASH_SHEETS.ERREURS);
  if (errT.rows.length){
    const cDate = findCol_(errT.headers, ['date','timestamp','time','datetime']);
    const cType = findCol_(errT.headers, ['type','code','error','erreur']);
    const cMsg  = findCol_(errT.headers, ['details','message','info','texte']);
    return errT.rows.slice(-10).reverse().map(r=>({
      date: formatDate_(cDate>=0 ? r[cDate] : ''),
      type: cType>=0 ? String(r[cType]||'') : 'ERREUR',
      details: cMsg>=0 ? String(r[cMsg]||'') : ''
    }));
  }

  return [];
}

function formatDate_(v){
  if (v instanceof Date) {
    const pad = n=>('0'+n).slice(-2);
    return `${v.getFullYear()}-${pad(v.getMonth()+1)}-${pad(v.getDate())} ${pad(v.getHours())}:${pad(v.getMinutes())}`;
  }
  return String(v||'');
}
