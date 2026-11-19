/**
 * ЕДИНЫЙ ЦЕНТР УПРАВЛЕНИЯ ПРОДАЖАМИ И СКЛАДОМ (CTRM «Капитал-Плюс» v7.1)
 * Файл: Supply.js — Лист «7. Закупки» (Стакан рынка, Сбор дефицита из CRM и Обратный радар)
 * Полная интеграция с листами «Склад», «2. Золотой список (Киты)» и «3. Лист ожидания».
 */

const SUPPLY_SHEET_NAME = '7. Закупки';

const SUPPLY_HEADERS = [
  'ID лота',                     // A (1)
  'ИНН Поставщика',              // B (2)
  'Поставщик / Контакт',         // C (3)
  'Размер (ØхСтенка)',           // D (4)
  'Тип трубы',                   // E (5)
  'Объем (тн)',                  // F (6)
  'Базис отгрузки (Город)',      // G (7)
  'Цена входа (₽/тн)',           // H (8)
  'Логистика (₽/тн)',            // I (9)
  'Себестоимость с дост. (₽/тн)',// J (10)
  'НДС / Оплата',                // K (11)
  'Бронь / Актуально до',        // L (12)
  'Статус лота',                 // M (13)
  'Кому продать (Мэтчинг CRM)',  // N (14)
  'Примечание (Стенка / ГОСТ)'   // O (15)
];

function setupSupplySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = typeof getSheetByNameSafe === 'function' ? getSheetByNameSafe(ss, SUPPLY_SHEET_NAME) : ss.getSheetByName(SUPPLY_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SUPPLY_SHEET_NAME);
  }

  const maxCols = SUPPLY_HEADERS.length;
  if (sheet.getMaxColumns() < maxCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), maxCols - sheet.getMaxColumns());
  }

  const headerRange = sheet.getRange(1, 1, 1, maxCols);
  headerRange
    .setValues([SUPPLY_HEADERS])
    .setBackground('#0f172a')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontFamily('Arial')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  sheet.setRowHeight(1, 38);
  sheet.setFrozenRows(1);

  const widths = [95, 115, 210, 130, 125, 95, 160, 125, 120, 145, 130, 120, 170, 330, 240];
  widths.forEach((w, idx) => sheet.setColumnWidth(idx + 1, w));

  const numRows = Math.max(sheet.getMaxRows() - 1, 100);
  const dataRange = sheet.getRange(2, 1, numRows, maxCols);
  dataRange.setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');

  sheet.getRange(2, 1, numRows, 13).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  sheet.getRange(2, 14, numRows, 2).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  sheet.getRange(2, 2, numRows, 1).setNumberFormat('@');         // B: ИНН
  sheet.getRange(2, 4, numRows, 1).setNumberFormat('@');         // D: Размер
  sheet.getRange(2, 6, numRows, 1).setNumberFormat('0.00');      // F: Объем (тн)
  sheet.getRange(2, 8, numRows, 3).setNumberFormat('#,##0 "₽"'); // H, I, J: Цены
  sheet.getRange(2, 12, numRows, 1).setNumberFormat('dd.MM.yyyy');// L: Бронь до

  const pipeTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Новая', 'Лежалая', 'Восстановленная', 'Б/у', 'НКТ'], true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 5, numRows, 1).setDataValidation(pipeTypeRule);

  const vatRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['С НДС 22%', 'Без НДС', 'Нал / Карта', 'Отсрочка / Бронь'], true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 11, numRows, 1).setDataValidation(vatRule);

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList([
      'Поиск (Дефицит CRM)',
      'В рынке (Предложение)',
      'Под бронью',
      'Выкуплено (В транзите)',
      'Архив / Ушло'
    ], true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 13, numRows, 1).setDataValidation(statusRule);

  sheet.clearConditionalFormatRules();
  const statusColRange = sheet.getRange(2, 13, numRows, 1);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Поиск (Дефицит CRM)')
      .setBackground('#fee2e2').setFontColor('#991b1b')
      .setRanges([statusColRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('В рынке (Предложение)')
      .setBackground('#dbeafe').setFontColor('#1e40af')
      .setRanges([statusColRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Под бронью')
      .setBackground('#fef3c7').setFontColor('#92400e')
      .setRanges([statusColRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Выкуплено (В транзите)')
      .setBackground('#dcfce7').setFontColor('#166534')
      .setRanges([statusColRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Архив / Ушло')
      .setBackground('#f1f5f9').setFontColor('#64748b')
      .setRanges([statusColRange]).build()
  ];
  sheet.setConditionalFormatRules(rules);

  ss.toast('Лист «7. Закупки» готов к работе!', '7. Закупки', 4);
}

/**
 * Авто-сбор дефицита из «2. Золотой список (Киты)» и «3. Лист ожидания»
 * Проверяет живые заявки в H против реальных остатков листа «Склад»
 */
function collectDeficitFromCRM() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let supplySheet = typeof getSheetByNameSafe === 'function' ? getSheetByNameSafe(ss, SUPPLY_SHEET_NAME) : ss.getSheetByName(SUPPLY_SHEET_NAME);
  if (!supplySheet) {
    setupSupplySheet();
    supplySheet = ss.getSheetByName(SUPPLY_SHEET_NAME);
  }

  // Используем штатную карту диаметров со листа «Склад» (строки с 3-й)
  const stockByDiam = typeof getAvailableStockDiamsMap_ === 'function' ? getAvailableStockDiamsMap_(ss) : {};

  const existingDeficits = new Map();
  if (supplySheet.getLastRow() > 1) {
    const supData = supplySheet.getRange(2, 1, supplySheet.getLastRow() - 1, 15).getValues();
    supData.forEach((r, idx) => {
      const size = String(r[3] || '').trim();
      const status = String(r[12] || '').trim();
      if (size && status === 'Поиск (Дефицит CRM)') {
        existingDeficits.set(size, idx + 2);
      }
    });
  }

  const crmSheets = ['2. Золотой список (Киты)', '3. Лист ожидания'];
  const deficitMap = new Map();

  crmSheets.forEach(sheetName => {
    const sh = typeof getSheetByNameSafe === 'function' ? getSheetByNameSafe(ss, sheetName) : ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return;
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();
    const tag = sheetName.indexOf('Киты') !== -1 ? '[КИТ]' : '[ОЖИД]';

    rows.forEach(row => {
      const inn = String(row[0] || '').trim();
      const company = String(row[1] || '').trim();
      const city = typeof extractShortCity_ === 'function' ? extractShortCity_(row[2]) : String(row[2] || '').trim();
      const need = String(row[7] || '').trim();
      const pipeType = String(row[8] || '').trim();

      if (!company || !need || need.toLowerCase() === 'нет' || need === '—' || need === '-') return;

      const items = need.split(/[,;]+/).map(s => s.trim()).filter(Boolean);
      items.forEach(item => {
        const diams = typeof extractCleanPipeDiams_ === 'function' ? extractCleanPipeDiams_(item) : [];
        if (diams.length === 0) return;
        const d = diams[0];

        // Если диаметра нет на листе «Склад» — добавляем в Дефицит Закупок
        if (!stockByDiam[d]) {
          const cleanSize = typeof formatPipeNeed_ === 'function' ? formatPipeNeed_(item) : item;
          if (!deficitMap.has(cleanSize)) {
            deficitMap.set(cleanSize, { pipeType: pipeType || 'Б/у', clients: [] });
          }
          const entry = deficitMap.get(cleanSize);
          const clientLine = `${tag} ${company}${city ? ' (' + city + ')' : (inn ? ' (ИНН ' + inn + ')' : '')}`;
          if (!entry.clients.includes(clientLine)) {
            entry.clients.push(clientLine);
          }
        }
      });
    });
  });

  let addedCount = 0;
  let updatedCount = 0;
  const newRows = [];
  const stamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone() || 'Europe/Moscow', 'ddMM');

  deficitMap.forEach((info, sizeKey) => {
    const clientsText = info.clients.join('\n');
    if (existingDeficits.has(sizeKey)) {
      const rowIndex = existingDeficits.get(sizeKey);
      supplySheet.getRange(rowIndex, 14).setValue(clientsText);
      updatedCount++;
    } else {
      addedCount++;
      const lotId = `DEF-${stamp}-${addedCount}`;
      newRows.push([
        lotId,
        '',
        'ДЕФИЦИТ ПОД ЗАЯВКУ CRM',
        sizeKey,
        info.pipeType,
        '',
        '',
        '',
        '',
        '',
        'С НДС 22%',
        '',
        'Поиск (Дефицит CRM)',
        clientsText,
        'Авто-сбор из колонки H'
      ]);
    }
  });

  if (newRows.length > 0) {
    const startRow = supplySheet.getLastRow() + 1;
    if (supplySheet.getMaxRows() < startRow + newRows.length) {
      supplySheet.insertRowsAfter(supplySheet.getMaxRows(), newRows.length + 10);
    }
    supplySheet.getRange(startRow, 1, newRows.length, SUPPLY_HEADERS.length).setValues(newRows);
  }

  ss.toast(`Добавлено новых позиций дефицита: ${addedCount} | Обновлено: ${updatedCount}`, 'Сбор дефицита CRM', 5);
}

/**
 * Обратный радар («Кому продать этот Ø?»):
 * Сверяет все позиции на листе «7. Закупки» с живыми заявками (H) и историей покупок 1С (S)
 */
function matchSupplyWithDemand() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const supplySheet = typeof getSheetByNameSafe === 'function' ? getSheetByNameSafe(ss, SUPPLY_SHEET_NAME) : ss.getSheetByName(SUPPLY_SHEET_NAME);
  if (!supplySheet || supplySheet.getLastRow() < 2) {
    ss.toast('На листе «7. Закупки» пока нет позиций.', '7. Закупки', 4);
    return;
  }

  const demandByDiam = {};
  const ensureDiam_ = (d) => {
    if (!demandByDiam[d]) demandByDiam[d] = { hot: [], history1C: [] };
  };

  const whalesSheet = typeof getSheetByNameSafe === 'function' ? getSheetByNameSafe(ss, '2. Золотой список (Киты)') : ss.getSheetByName('2. Золотой список (Киты)');
  if (whalesSheet && whalesSheet.getLastRow() > 1) {
    const wCols = Math.min(whalesSheet.getLastColumn(), 19);
    const wData = whalesSheet.getRange(2, 1, whalesSheet.getLastRow() - 1, wCols).getValues();
    wData.forEach(r => {
      const company = String(r[1] || '').trim();
      const needH   = String(r[7] || '').trim();
      const tonsQ   = wCols >= 17 ? safeParseFloat(r[16]) : 0;
      const hist1C  = wCols >= 19 ? String(r[18] || '').trim() : '';
      if (!company) return;

      if (needH && needH.toLowerCase() !== 'нет' && needH !== '—') {
        const diamsH = extractCleanPipeDiams_(needH);
        diamsH.forEach(d => {
          ensureDiam_(d);
          demandByDiam[d].hot.push(`[ЗАЯВКА КИТ] ${company} (${needH})`);
        });
      }

      if (hist1C) {
        const diamsS = extractCleanPipeDiams_(hist1C);
        diamsS.forEach(d => {
          ensureDiam_(d);
          if (demandByDiam[d].history1C.length < 6) {
            const tnLabel = tonsQ > 0 ? ` ${Math.round(tonsQ)}т` : '';
            demandByDiam[d].history1C.push(`${company}${tnLabel}`);
          }
        });
      }
    });
  }

  const waitSheet = typeof getSheetByNameSafe === 'function' ? getSheetByNameSafe(ss, '3. Лист ожидания') : ss.getSheetByName('3. Лист ожидания');
  if (waitSheet && waitSheet.getLastRow() > 1) {
    const waitData = waitSheet.getRange(2, 1, waitSheet.getLastRow() - 1, 14).getValues();
    waitData.forEach(r => {
      const company = String(r[1] || '').trim();
      const needH   = String(r[7] || '').trim();
      const histN   = String(r[13] || '').trim();
      if (!company) return;

      if (needH && needH.toLowerCase() !== 'нет' && needH !== '—') {
        const diamsH = extractCleanPipeDiams_(needH);
        diamsH.forEach(d => {
          ensureDiam_(d);
          demandByDiam[d].hot.push(`[ЗАЯВКА ОЖИД] ${company} (${needH})`);
        });
      } else if (histN) {
        const diamsN = extractCleanPipeDiams_(histN);
        diamsN.forEach(d => {
          ensureDiam_(d);
          if (demandByDiam[d].history1C.length < 6 && !demandByDiam[d].history1C.includes(company)) {
            demandByDiam[d].history1C.push(company);
          }
        });
      }
    });
  }

  const numRows = supplySheet.getLastRow() - 1;
  const supData = supplySheet.getRange(2, 1, numRows, 15).getValues();
  const outColN = [];

  supData.forEach(r => {
    const sizeD = String(r[3] || '').trim();
    const diams = typeof extractCleanPipeDiams_ === 'function' ? extractCleanPipeDiams_(sizeD) : [];
    if (diams.length === 0) {
      outColN.push([r[13] || '']);
      return;
    }
    const d = diams[0];
    const matched = demandByDiam[d];
    if (!matched || (matched.hot.length === 0 && matched.history1C.length === 0)) {
      outColN.push(['Спрос в базе не найден']);
      return;
    }
    const lines = [];
    if (matched.hot.length > 0) lines.push(matched.hot.join(' | '));
    if (matched.history1C.length > 0) lines.push('Профиль отгрузок / CRM: ' + matched.history1C.join(', '));
    outColN.push([lines.join('\n')]);
  });

  supplySheet.getRange(2, 14, numRows, 1).setValues(outColN);
  ss.toast('Мэтчинг «7. Закупки» ↔ Спрос CRM и отгрузок выполнен!', 'Обратный радар', 5);
}

function handleSupplySheetEdit(e, sheet) {
  if (!e || !e.range) return;
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < 2) return;

  if (col === 3 || col === 4) {
    const rowVals = sheet.getRange(row, 1, 1, 15).getValues()[0];
    const hasSupplier = String(rowVals[2] || '').trim() !== '';
    const hasSize = String(rowVals[3] || '').trim() !== '';
    if (!hasSupplier && !hasSize) {
      sheet.getRange(row, 1, 1, 15).clearContent().setBackground(null);
      return;
    }
  }

  // Авто-форматирование размера в колонке D (например: 5308 -> 530х8)
  if (col === 4) {
    const rawSize = String(e.value !== undefined ? e.value : sheet.getRange(row, 4).getValue() || '').trim();
    if (rawSize && typeof formatPipeNeed_ === 'function') {
      const formatted = formatPipeNeed_(rawSize);
      if (formatted && formatted !== rawSize) {
        sheet.getRange(row, 4).setNumberFormat('@').setValue(formatted);
      }
    }
    if (!sheet.getRange(row, 1).getValue()) {
      const stamp = Utilities.formatDate(new Date(), 'Europe/Moscow', 'ddMM');
      sheet.getRange(row, 1).setValue(`SUP-${stamp}-${row - 1}`);
    }
    if (!sheet.getRange(row, 13).getValue()) {
      sheet.getRange(row, 13).setValue('В рынке (Предложение)');
    }
  }

  // Авто-пересчет Себестоимости с доставкой: J (10) = H (8) + I (9)
  if (col === 8 || col === 9) {
    const priceIn = safeParseFloat(sheet.getRange(row, 8).getValue());
    const logistics = safeParseFloat(sheet.getRange(row, 9).getValue());
    const totalCell = sheet.getRange(row, 10);
    if (priceIn > 0 || logistics > 0) {
      totalCell.setValue(priceIn + logistics);
    } else {
      totalCell.clearContent();
    }
  }
}