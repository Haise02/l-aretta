/**
 * L'Aretta — Lead capture form "Prenotazione 2 giugno 2026"
 *
 * QUESTO FILE NON GIRA IN LOCALE. È IL CODICE DA INCOLLARE
 * dentro Google Apps Script (script.google.com), legato al Google Sheet
 * dove vuoi che arrivino le prenotazioni.
 *
 * === SETUP (5 minuti) ===
 *
 * 1. Crea un nuovo Google Sheet (sheet.google.com → vuoto). Nominalo:
 *    "L'Aretta — Prenotazioni 2 giugno 2026"
 *
 * 2. Dentro al Sheet vai a: Estensioni → Apps Script
 *    Si apre l'editor Apps Script. Cancella tutto il contenuto del file
 *    Code.gs e incolla TUTTO il contenuto di questo file qui sotto.
 *
 * 3. (opzionale) cambia NOTIFICA_EMAIL qui sotto con la tua mail.
 *    Se la lasci vuota, niente mail di alert (solo riga su Sheet).
 *
 * 4. Salva (Ctrl+S). Dai un nome al progetto, es: "Aretta Lead".
 *
 * 5. Clicca Deploy → New deployment.
 *      Type:     "Web app"  (clicca l'ingranaggio se non lo vedi)
 *      Description: "v1"
 *      Execute as: "Me"
 *      Who has access: "Anyone"   ← MOLTO IMPORTANTE
 *    Premi "Deploy". Google ti chiede l'autorizzazione la prima volta:
 *    accetta tutto (è tuo, gestisci solo il tuo Sheet).
 *
 * 6. Copia l'URL "Web app URL" che ti dà Google (lungo, finisce con /exec).
 *
 * 7. Apri evento-2-giugno.html, cerca il testo:
 *      REPLACE_WITH_APPS_SCRIPT_URL
 *    e sostituiscilo con l'URL copiato al punto 6.
 *
 * 8. Push su GitHub, redeploy Vercel. Fatto: le prenotazioni arrivano
 *    sul Sheet in tempo reale.
 *
 * === MODIFICHE ===
 * Se aggiorni questo script in futuro, devi fare:
 * Deploy → Manage deployments → matita modifica → New version → Deploy.
 * (Solo creando una NUOVA versione del deployment, l'URL serve i nuovi cambiamenti.)
 */

// CONFIGURAZIONE
const NOTIFICA_EMAIL = ''; // Es: 'tu@gmail.com'. Lascia '' per disattivare le mail.
const SHEET_NAME = 'Prenotazioni'; // Nome del foglio dentro al Spreadsheet

// Intestazioni colonne (in ordine)
const HEADERS = [
  'Timestamp',
  'Nome',
  'Telefono',
  'Email',
  'Persone',
  'Note',
  'Evento',
  'User Agent'
];

/**
 * doPost — riceve la POST del form
 */
function doPost(e) {
  try {
    const params = e.parameter || {};

    const row = [
      new Date(),
      String(params.nome || '').trim(),
      String(params.telefono || '').trim(),
      String(params.email || '').trim(),
      String(params.persone || '').trim(),
      String(params.note || '').trim(),
      String(params.evento || 'Pranzo 2 giugno 2026').trim(),
      String((e.headers && e.headers['User-Agent']) || '').trim()
    ];

    const sheet = getOrCreateSheet_();
    sheet.appendRow(row);

    // Notifica email (opzionale)
    if (NOTIFICA_EMAIL) {
      sendNotifica_(row);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * doGet — utile per test diretto da browser
 */
function doGet() {
  return ContentService.createTextOutput(
    'L\'Aretta lead capture endpoint — funziona. Usa POST per inviare dati.'
  );
}

/**
 * Recupera o crea il foglio "Prenotazioni" con le header
 */
function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // Se la riga 1 è vuota, scrivi le header
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    // Formatta riga header
    const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange
      .setFontWeight('bold')
      .setBackground('#1B1611')
      .setFontColor('#F3EBD8');
    sheet.setFrozenRows(1);
    // Auto-resize colonne
    for (let i = 1; i <= HEADERS.length; i++) sheet.autoResizeColumn(i);
  }

  return sheet;
}

/**
 * Manda mail di notifica all'owner
 */
function sendNotifica_(row) {
  const [ts, nome, tel, email, persone, note, evento] = row;
  const subject = `🍝 Nuova prenotazione · ${nome} · ${persone} persone`;
  const body = [
    `Nuova richiesta di prenotazione per "${evento}":`,
    ``,
    `• Nome:      ${nome}`,
    `• Telefono:  ${tel}`,
    `• Email:     ${email}`,
    `• Persone:   ${persone}`,
    `• Note:      ${note || '(nessuna)'}`,
    ``,
    `Ricevuta il: ${ts.toLocaleString('it-IT')}`,
    ``,
    `— L'Aretta lead capture`
  ].join('\n');

  try {
    MailApp.sendEmail(NOTIFICA_EMAIL, subject, body);
  } catch (err) {
    // se MailApp è esaurita o non autorizzata, non blocca il salvataggio
    console.error('Mail non inviata:', err);
  }
}

/**
 * Test rapido — esegui questa funzione dall'editor (in alto: scegli "test_"
 * dal dropdown e premi ▶) per simulare una prenotazione di prova.
 */
function test_() {
  const fakeEvent = {
    parameter: {
      nome: 'Mario Rossi (TEST)',
      telefono: '+39 333 1234567',
      email: 'mario.test@example.com',
      persone: '4',
      note: 'Una vegetariana, un seggiolone',
      evento: 'Pranzo 2 giugno 2026 — TEST'
    }
  };
  const res = doPost(fakeEvent);
  console.log(res.getContent());
}
