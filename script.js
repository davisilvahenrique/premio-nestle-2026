// ═══════════════════════════════════════════════════════════════════════════
//  PRÊMIO NESTLÉ 2026 — Google Apps Script
//  Cole este código em: Google Sheets → Extensões → Apps Script
// ═══════════════════════════════════════════════════════════════════════════

const SHEET_NAME = 'Votos';

// Colunas na planilha (na ordem)
const COLUMNS = [
  'Timestamp',
  'Nome',
  'Amigo Nestlé do Ano',
  'AmigA Nestlé do Ano',
  'Golden Boy',
  'Pior Amigo do Ano',
  'Resenha do Ano',
  'Prêmio CEO',
  'Story do Samuca',
  'Prêmio Neve',
  'Prêmio Piscineira',
  'Fioti do Ano',
  'Casal do Ano',
  'Prêmio SAKA',
  'Versão do AcacIA',
];

// Mapeamento campo JSON → coluna
const KEY_MAP = {
  timestamp:          'Timestamp',
  nome:               'Nome',
  amigo_ano:          'Amigo Nestlé do Ano',
  amiga_ano:          'AmigA Nestlé do Ano',
  golden_boy:         'Golden Boy',
  pior_amigo:         'Pior Amigo do Ano',
  resenha_ano:        'Resenha do Ano',
  premio_ceo:         'Prêmio CEO',
  story_samuca:       'Story do Samuca',
  premio_neve:        'Prêmio Neve',
  premio_piscineira:  'Prêmio Piscineira',
  fioti_ano:          'Fioti do Ano',
  casal_ano:          'Casal do Ano',
  premio_saka:        'Prêmio SAKA',
  versao_acacio:      'Versão do AcacIA',
};

// ── Setup inicial: cria a aba e o cabeçalho se não existir ──────────────────
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(COLUMNS);
    sheet.getRange(1, 1, 1, COLUMNS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
}

// ── Recebe o POST do formulário ─────────────────────────────────────────────
function doPost(e) {
  // Trava: evita que votos simultâneos se misturem ou se percam
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet();

    // Verifica voto duplicado pelo nome (case-insensitive)
    const nomeLower = (data.nome || '').toLowerCase().trim();
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]).toLowerCase().trim() === nomeLower) {
        return jsonResponse({ status: 'already_voted' });
      }
    }

    // Monta a linha na ordem das colunas
    const row = COLUMNS.map(col => {
      const key = Object.keys(KEY_MAP).find(k => KEY_MAP[k] === col);
      return key ? (data[key] || '') : '';
    });

    sheet.appendRow(row);
    SpreadsheetApp.flush();

    return jsonResponse({ status: 'ok' });
  } catch (err) {
    return jsonResponse({ status: 'error', message: err.message });
  } finally {
    lock.releaseLock();
  }
}

// ── Helpers ─────────────────────────────────────────────────────────────────
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(COLUMNS);
    sheet.getRange(1, 1, 1, COLUMNS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
