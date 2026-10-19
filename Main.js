/**
 * ЕДИНЫЙ ЦЕНТР УПРАВЛЕНИЯ ПРОДАЖАМИ И СКЛАДОМ (CTRM «Капитал-Плюс» v6.1)
 * Файл: Main.gs — Главное меню, глобальные утилиты и маршрутизатор триггеров
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📱 Управление')
    .addItem('⚡ Открыть калькулятор отгрузки', 'openCalculatorSidebar')
    .addItem('📄 Сформировать прайс-лист (HTML / PDF)', 'openPriceListModal')
    .addSeparator()
    .addItem('🎯 Радар: сверить Склад ↔ Ожидание и Киты', 'matchWarehouseWithWaitingList')
    .addItem('🗂️ Упаковать примечания (K) в карточки Досье (B)', 'packAllNotesIntoDossierOnActiveSheet')
    .addItem('🔍 Проверить дубли ИНН на листе', 'highlightAllDuplicateInnsManual')
    .addItem('🧹 Отсортировать склад по диаметрам', 'sortWarehouseData')
    .addSeparator()
    .addItem('➕ Добавить клиента в Контакты (iPhone)', 'addSelectedRowToContacts')
    .addToUi();

  // Автоматическая подсветка дублей ИНН на активном листе при открытии таблицы
  try {
    const activeSheet = SpreadsheetApp.getActiveSheet();
    if (activeSheet && typeof highlightDuplicateInnsOnSheet_ === 'function') {
      const crmSheets = ["1. Обзвон (База)", "2. Золотой список (Киты)", "3. Лист ожидания", "4.Архив лидов"];
      if (crmSheets.includes(activeSheet.getName().trim())) {
        highlightDuplicateInnsOnSheet_(activeSheet);
      }
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
    // Даже если вставили сразу пачку строк (Ctrl+V), автоматически подсвечиваем дубли ИНН
    if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) {
      if (typeof highlightDuplicateInnsOnSheet_ === 'function') {
        highlightDuplicateInnsOnSheet_(sheet);
      }
      return;
    }
    if (typeof handleCallingSheetEdit === 'function') {
      handleCallingSheetEdit(e, sheet);
    }
  } else if (sheetName === "Склад") {
    if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) return;
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