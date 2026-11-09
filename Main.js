/**
 * ЕДИНЫЙ ЦЕНТР УПРАВЛЕНИЯ ПРОДАЖАМИ И СКЛАДОМ (CTRM «Капитал-Плюс» v6.4)
 * Файл: Main.gs — Главное меню, глобальные утилиты и маршрутизатор триггеров
 * с автоматическим сбросом подсветок и авто-подъемом очереди при обновлении страницы.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📱 Управление')
    .addItem('⚡ Открыть калькулятор отгрузки', 'openCalculatorSidebar')
    .addItem('📄 Сформировать прайс-лист (HTML / PDF)', 'openPriceListModal')
    .addSeparator()
    .addItem('🎯 Радар: сверить Склад ↔ Ожидание и Киты', 'matchWarehouseWithWaitingList')
    .addItem('🔍 Проверить дубли ИНН на листе', 'highlightAllDuplicateInnsManual')
    .addItem('🧹 Отсортировать склад по диаметрам', 'sortWarehouseData')
    .addSeparator()
    .addItem('➕ Добавить клиента в Контакты (iPhone)', 'addSelectedRowToContacts')
    .addToUi();

  try {
    const startMs = Date.now();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheet = ss.getActiveSheet();
    const activeRange = activeSheet ? activeSheet.getActiveRange() : null;
    const activeRow = activeRange ? activeRange.getRow() : 2;
    const activeCol = activeRange ? activeRange.getColumn() : 1;

    const crmSheets = ["1. Обзвон (База)", "3. Лист ожидания", "2. Золотой список (Киты)"];
    // Сначала обрабатываем активный лист, чтобы пользователь сразу видел актуальную очередь
    if (activeSheet) {
      const actName = activeSheet.getName().trim();
      const idx = crmSheets.indexOf(actName);
      if (idx > 0) {
        crmSheets.splice(idx, 1);
        crmSheets.unshift(actName);
      }
    }

    for (let i = 0; i < crmSheets.length; i++) {
      // Защита от таймаута 30 сек в простом триггере onOpen
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

    // Возвращаем фокус на ячейку пользователя, если он стоял на активном листе
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

  const crmSheets = [
    "1. Обзвон (База)",
    "2. Золотой список (Киты)",
    "3. Лист ожидания",
    "4.Архив лидов"
  ];

  if (crmSheets.includes(sheetName)) {
    // Если выделено несколько ячеек (например, очистка через Delete/Backspace или вставка Ctrl+V)
    if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) {
      if (typeof handleMultiCellCrmEdit_ === 'function') {
        handleMultiCellCrmEdit_(e, sheet);
      }
      return;
    }
    if (typeof handleCallingSheetEdit === 'function') {
      handleCallingSheetEdit(e, sheet);
    }
  } else if (sheetName === "Склад") {
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
  const sheet = getSheetByNameSafe(ss, 'Склад');
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return;

  const range = sheet.getRange(3, 1, lastRow - 2, 15);

  range.sort([
    {column: 4, ascending: true},
    {column: 5, ascending: true}
  ]);

  ss.toast('Трубы выстроены строго по возрастанию диаметра!', '✅ Склад отсортирован', 4);
}