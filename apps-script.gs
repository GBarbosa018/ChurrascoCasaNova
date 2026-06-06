/**
 * Backend de reservas — Lista de Presentes (Google Sheets como "banco de dados").
 *
 * COMO USAR:
 * 1. Crie uma planilha nova no Google Sheets.
 * 2. Menu: Extensões → Apps Script.
 * 3. Apague o conteúdo padrão e cole TODO este arquivo.
 * 4. Clique em "Implantar" → "Nova implantação" → tipo "App da Web".
 *      - Executar como: Eu (sua conta)
 *      - Quem pode acessar: Qualquer pessoa
 * 5. Copie a URL terminada em /exec e cole na constante API_URL do index.html.
 *
 * A planilha terá as colunas: id | nome | reservadoPor | data
 * (criadas automaticamente na primeira reserva).
 */

const SHEET_NAME = 'Reservas';

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['id', 'nome', 'reservadoPor', 'data']);
  }
  return sheet;
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Lê todas as reservas. Chamado pelo site ao carregar. */
function doGet() {
  const sheet = getSheet();
  const rows = sheet.getDataRange().getValues(); // inclui cabeçalho
  const reservas = [];
  for (let i = 1; i < rows.length; i++) {
    const [id, nome, reservadoPor] = rows[i];
    if (id !== '' && reservadoPor) {
      reservas.push({ id: id, reservadoPor: reservadoPor });
    }
  }
  return jsonOutput({ ok: true, reservas: reservas });
}

/** Registra uma reserva. Chamado pelo site ao confirmar. */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000); // evita duas pessoas reservarem ao mesmo tempo
  try {
    const body = JSON.parse(e.postData.contents);
    const id = Number(body.id);
    const nome = String(body.nome || '').trim();

    if (!id || !nome) {
      return jsonOutput({ ok: false, reason: 'invalid' });
    }

    const sheet = getSheet();
    const rows = sheet.getDataRange().getValues();

    // Já reservado? Devolve quem reservou.
    for (let i = 1; i < rows.length; i++) {
      if (Number(rows[i][0]) === id && rows[i][2]) {
        return jsonOutput({ ok: false, reason: 'taken', reservadoPor: rows[i][2] });
      }
    }

    // Grava a reserva.
    sheet.appendRow([id, nome, nome, new Date()]);
    return jsonOutput({ ok: true });
  } catch (err) {
    return jsonOutput({ ok: false, reason: 'error', message: String(err) });
  } finally {
    lock.releaseLock();
  }
}
