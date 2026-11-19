/**
 * ЕДИНЫЙ ЦЕНТР УПРАВЛЕНИЯ ПРОДАЖАМИ И СКЛАДОМ (CTRM «Капитал-Плюс» v7.2)
 * Файл: Main.js — Главное меню (5 базовых листов), утилиты и маршрутизатор триггеров.
 */

function isWarehouseSheetName_(name) {
  const clean = String(name || '').toLowerCase().replace(/^\d+\.\s*/, '').trim();
  return clean === 'склад';
}

function isCrmSheetName_(name) {
  const compact = String(name || '').toLowerCase().replace(/\s+/g, '');
  const targets = [
    '1.обзвон(база)',
    '2.золотойсписок(киты)',
    '3.листожидания',
    '4.архивлидов'
  ];
  return targets.includes(compact);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Управление')
    .addItem('Открыть калькулятор отгрузки (ГОСТ / Весы / Торг)', 'openCalculatorSidebar')
    .addItem('Сформировать прайс-лист (HTML / PDF)', 'openPriceListModal')
    .addItem('Напоминалка и Карта партнёра (HTML / Мессенджер)', 'openFollowUpModal')
    .addSeparator()
    .addItem('Обратный Радар: кому продать этот Ø?', 'openReverseRadarForActivePipe')
    .addItem('Радар: сверить Склад ↔ Ожидание и Киты', 'matchWarehouseWithWaitingList')
    .addItem('Пробить контрагента по ИНН (Rusprofile / Checko)', 'openInnCheckDialog')
    .addItem('Сборка Холдинга (объединить юрлица 1 ЛПР)', 'mergeSelectedWhalesIntoHolding')
    .addItem('Проверить дубли ИНН на листе', 'highlightAllDuplicateInnsManual')
    .addSeparator()
    .addItem('Отсортировать склад по диаметрам', 'sortWarehouseData')
    .addItem('Добавить клиента в Контакты (iPhone)', 'addSelectedRowToContacts')
    .addToUi();

  try {
    const startMs = Date.now();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = ss.getActiveSheet();
    const activeRange = activeSheet ? activeSheet.getActiveRange() : null;
    const activeRow = activeRange ? activeRange.getRow() : 2;
    const activeCol = activeRange ? activeSheet.getActiveRange().getColumn() : 1;

    const crmSheets = ["1. Обзвон (База)", "3. Лист ожидания", "2. Золотой список (Киты)"];
    if (activeSheet) {
      const actName = activeSheet.getName().trim();
      const idx = crmSheets.indexOf(actName);
      if (idx > 0) {
        crmSheets.splice(idx, 1);
        crmSheets.unshift(actName);
      }
    }

    for (let i = 0; i < crmSheets.length; i++) {
      if (Date.now() - startMs > 20000) break;
      const sh = getSheetByNameSafe(ss, crmSheets[i]);
      if (sh) {
        if (typeof syncUnpackedNotesOnSheetSilent_ === 'function') {
          syncUnpackedNotesOnSheetSilent_(sh);
        }
        if (typeof sortAndHighlightDeadlinesOnSheet_ === 'function') {
          sortAndHighlightDeadlinesOnSheet_(sh);
        } else if (typeof highlightDuplicateInnsOnSheet_ === 'function') {
          highlightDuplicateInnsOnSheet_(sh);
        }
      }
    }

    if (activeSheet && activeRow >= 1 && activeCol >= 1) {
      const safeRow = Math.min(activeRow, Math.max(1, activeSheet.getMaxRows()));
      const safeCol = Math.min(activeCol, Math.max(1, activeSheet.getMaxColumns()));
      activeSheet.setActiveSelection(activeSheet.getRange(safeRow, safeCol));
    }
  } catch (e) {}
}

// =================================================================
// ЗАЩИТА И ОЧИСТКА ДАННЫХ (DEFENSIVE UTILITIES)
// =================================================================

function safeParseFloat(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  const cleaned = val.toString()
    .replace(/[\s\u00A0]/g, '')
    .replace(',', '.');
  const num = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
}

function cleanCompanyName(name) {
  if (!name) return '';
  let str = name.toString().trim();

  str = str.replace(/""+/g, '"');

  if ((str.startsWith('"') && str.endsWith('"')) || (str.startsWith("'") && str.endsWith("'"))) {
    const inner = str.substring(1, str.length - 1).trim();
    if (inner.includes('"') || inner.toUpperCase().startsWith('ООО') || inner.toUpperCase().startsWith('АО')) {
      str = inner;
    }
  }

  str = str.replace(/^(ООО|ЗАО|ОАО|ПАО|АО|ИП)"/i, '$1 "');

  const quotes = (str.match(/"/g) || []).length;
  if (quotes % 2 !== 0 && !str.endsWith('"')) {
    str = str + '"';
  }

  return str.trim();
}

function isValidINN(inn) {
  if (!inn) return false;
  const cleanInn = inn.toString().replace(/\D/g, '');
  return cleanInn.length === 10 || cleanInn.length === 12;
}

function getSheetByNameSafe(ss, targetName) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const direct = ss.getSheetByName(targetName);
  if (direct) return direct;
  const targetClean = targetName.toLowerCase().replace(/\s+/g, '');
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().toLowerCase().replace(/\s+/g, '') === targetClean) {
      return sheets[i];
    }
  }
  return null;
}

// =================================================================
// ЕДИНЫЙ ТРИГГЕР ИЗМЕНЕНИЙ (МАРШРУТИЗАТОР СОБЫТИЙ)
// =================================================================

function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName().trim();

  if (isCrmSheetName_(sheetName)) {
    if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) {
      if (typeof handleMultiCellCrmEdit_ === 'function') {
        handleMultiCellCrmEdit_(e, sheet);
      }
      return;
    }
    if (typeof handleCallingSheetEdit === 'function') {
      handleCallingSheetEdit(e, sheet);
    }
  } else if (isWarehouseSheetName_(sheetName)) {
    if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) {
      if (typeof handleMultiCellWarehouseEdit_ === 'function') {
        handleMultiCellWarehouseEdit_(e, sheet);
      }
      return;
    }
    if (typeof handleWarehouseSheetEdit === 'function') {
      handleWarehouseSheetEdit(e, sheet);
    }
  }
}

// =================================================================
// СКЛАДСКИЕ УТИЛИТЫ
// =================================================================

function sortWarehouseData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = typeof getWarehouseSheet_ === 'function' ? getWarehouseSheet_(ss) : getSheetByNameSafe(ss, 'Склад');
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;

  const range = sheet.getRange(3, 1, lastRow - 2, 15);
  range.sort([
    {column: 4, ascending: true},
    {column: 5, ascending: true}
  ]);

  ss.toast('Трубы выстроены строго по возрастанию диаметра!', 'Склад отсортирован', 4);
}