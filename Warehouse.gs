/**
 * ЕДИНЫЙ ЦЕНТР УПРАВЛЕНИЯ ПРОДАЖАМИ И СКЛАДОМ
 * Файл: Warehouse.gs — Работа со складом (15 колонок A–O), расчеты, цех и калькулятор
 *
 * СТРУКТУРА ЛИСТА «СКЛАД» (15 колонок):
 * A (1):  Категория
 * B (2):  Позиция
 * C (3):  Состояние
 * D (4):  Ø мм
 * E (5):  Стенка мм
 * F (6):  ГОСТ / Сырье
 * G (7):  Закупка ₽/тн
 * H (8):  Продажа ₽/тн
 * I (9):  Продажа ₽/м
 * J (10): Наличие
 * K (11): Остаток тн
 * L (12): Остаток м
 * M (13): Сумма ₽
 * N (14): Примечания
 * O (15): Обновлено
 */

const MIN_MARGIN_PER_TON = 3000;

function openCalculatorSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Calculator')
    .setTitle('Калькулятор • Маржа • Цех')
    .setWidth(380);
  SpreadsheetApp.getUi().showSidebar(html);
}

function normalizePipeName(str) {
  if (!str) return '';
  return str.toString()
    .toLowerCase()
    .replace(/[xX]/g, 'х')
    .replace(/[\s\u00A0]+/g, ' ')
    .trim();
}

function formatCleanPipeTitle_(str) {
  if (!str) return '';
  return str.toString()
    .replace(/(\d)\s*[xX*]\s*(\d)/g, '\$1х\$2')
    .replace(/[\s\u00A0]+/g, ' ')
    .trim();
}

/**
 * Быстрое чтение склада за 1 обращение к листу (A3:O)
 */
function getPipesData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = typeof getSheetByNameSafe === 'function' ? getSheetByNameSafe(ss, 'Склад') : ss.getSheetByName('Склад');
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];

  const values = sheet.getRange(3, 1, lastRow - 2, 15).getValues();

  return values.map((row, idx) => {
    const rawName = (row[1] || '').toString().trim(); // B
    const diam    = safeParseFloat(row[3]);           // D
    const wall    = safeParseFloat(row[4]);           // E
    const weightPerMeterKg = (diam > 0 && wall > 0) ? (diam - wall) * wall * 0.0246615 : 0;

    return {
      rowIndex: idx + 3,
      category: (row[0] || '').toString().trim(),     // A
      name: rawName,                                  // B
      condition: (row[2] || '').toString().trim(),    // C
      diam: diam,                                     // D
      wall: wall,                                     // E
      gost: (row[5] || '').toString().trim(),         // F
      buyPriceTn: safeParseFloat(row[6]),             // G: Закупка ₽/тн
      priceTn: safeParseFloat(row[7]),                // H: Продажа ₽/тн
      priceM: safeParseFloat(row[8]),                 // I: Продажа ₽/м
      status: (row[9] || '').toString().trim(),       // J: Наличие
      stockTn: safeParseFloat(row[10]),               // K: Остаток тн
      stockM: safeParseFloat(row[11]),                // L: Остаток м
      notes: (row[13] || '').toString().trim(),       // N: Примечания
      weightPerMeterKg: weightPerMeterKg,
      minMarginPerTon: MIN_MARGIN_PER_TON
    };
  }).filter(p => p.name.length > 0 && !p.name.toUpperCase().startsWith('ИТОГО'));
}

/**
 * Пакетное списание остатков при отгрузке через калькулятор
 */
function executeBatchDeduction(cartItems) {
  if (!cartItems || cartItems.length === 0) return "Корзина пуста!";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = typeof getSheetByNameSafe === 'function' ? getSheetByNameSafe(ss, 'Склад') : ss.getSheetByName('Склад');
  if (!sheet) throw new Error("Лист «Склад» не найден!");

  let report = [];
  const now = new Date();
  let totalRevenue = 0;
  let totalMargin = 0;
  let hasBuyData = false;

  cartItems.forEach(item => {
    const row = item.rowIndex;

    const actualName = sheet.getRange(row, 2).getValue().toString().trim();
    if (normalizePipeName(actualName) !== normalizePipeName(item.name)) {
      throw new Error(`⚠ Позиция в строке ${row} изменилась (ожидалась "${item.name}", сейчас "${actualName}"). Обновите калькулятор.`);
    }

    const currentM = safeParseFloat(sheet.getRange(row, 12).getValue()); // L: Остаток м
    const deductM  = safeParseFloat(item.meters);
    const deductTn = safeParseFloat(item.tons);
    const newM     = Math.max(0, Math.round((currentM - deductM) * 10) / 10);

    const diam = safeParseFloat(sheet.getRange(row, 4).getValue()); // D
    const wall = safeParseFloat(sheet.getRange(row, 5).getValue()); // E
    const weightPerMeterTn = ((diam - wall) * wall * 0.0246615) / 1000;

    const buyPriceTn  = safeParseFloat(sheet.getRange(row, 7).getValue()); // G: Закупка
    const dealPriceTn = safeParseFloat(item.dealPriceTn) || safeParseFloat(sheet.getRange(row, 8).getValue()); // H: Продажа

    if (buyPriceTn > 0 && (dealPriceTn - buyPriceTn) < MIN_MARGIN_PER_TON) {
      throw new Error(
        `⛔ СТОП-ЛОСС по позиции «${item.name}»!\n` +
        `Закупка: ${buyPriceTn.toLocaleString('ru-RU')} ₽/тн, Цена сделки: ${dealPriceTn.toLocaleString('ru-RU')} ₽/тн.\n` +
        `Маржа ниже порога ${MIN_MARGIN_PER_TON.toLocaleString('ru-RU')} ₽/тн!`
      );
    }

    if (newM > 0) {
      const newTn = Math.round(newM * weightPerMeterTn * 100) / 100;
      sheet.getRange(row, 12).setValue(newM);          // L: Остаток м
      sheet.getRange(row, 11).setValue(newTn);         // K: Остаток тн
      sheet.getRange(row, 10).setValue('В наличии');   // J: Наличие

      const currentPriceTn = safeParseFloat(sheet.getRange(row, 8).getValue()); // H: Продажа
      if (currentPriceTn > 0) {
        sheet.getRange(row, 13).setValue(Math.round(currentPriceTn * newTn)); // M: Сумма
      } else {
        sheet.getRange(row, 13).clearContent();
      }
    } else {
      sheet.getRange(row, 12).clearContent();
      sheet.getRange(row, 11).clearContent();
      sheet.getRange(row, 13).clearContent();
      sheet.getRange(row, 10).setValue('Нет');
    }

    sheet.getRange(row, 15).setValue(now).setNumberFormat('dd.MM.yyyy HH:mm'); // O: Обновлено

    const itemRevenue = Math.round(dealPriceTn * deductTn);
    totalRevenue += itemRevenue;

    if (buyPriceTn > 0) {
      hasBuyData = true;
      const itemMargin = Math.round((dealPriceTn - buyPriceTn) * deductTn);
      totalMargin += itemMargin;
      report.push(`• ${item.name}: -${deductM} м (~${deductTn.toFixed(3)} тн) | Маржа: +${itemMargin.toLocaleString('ru-RU')} ₽`);
    } else {
      report.push(`• ${item.name}: -${deductM} м (~${deductTn.toFixed(3)} тн)`);
    }
  });

  let summaryText = `✅ Списано со склада на сумму: ${totalRevenue.toLocaleString('ru-RU')} ₽\n`;
  if (hasBuyData) {
    summaryText += `💰 Чистая маржа со сделки: +${totalMargin.toLocaleString('ru-RU')} ₽\n`;
  }
  return summaryText + `\n` + report.join("\n");
}

/**
 * Пакетное оприходование (Приёмка) на склад со средневзвешенной закупкой
 */
function executeBatchAddition(cartItems) {
  if (!cartItems || cartItems.length === 0) return "Корзина пуста!";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = typeof getSheetByNameSafe === 'function' ? getSheetByNameSafe(ss, 'Склад') : ss.getSheetByName('Склад');
  if (!sheet) throw new Error("Лист «Склад» не найден!");

  let report = [];
  const now = new Date();

  cartItems.forEach(item => {
    let row;

    if (item.isNew) {
      row = sheet.getLastRow() + 1;

      const diam = safeParseFloat(item.diam);
      const wall = safeParseFloat(item.wall);
      if (diam <= 0 || wall <= 0) throw new Error(`Ошибка геометрии у ${item.name}`);

      const weightPerMeterTn = ((diam - wall) * wall * 0.0246615) / 1000;
      const buyPriceTn  = safeParseFloat(item.buyPriceTn);
      const sellPriceTn = safeParseFloat(item.priceTn) || (buyPriceTn > 0 ? buyPriceTn + MIN_MARGIN_PER_TON : 0);
      const addTn       = safeParseFloat(item.tons);
      const addM        = Math.max(0, Math.round((addTn / weightPerMeterTn) * 10) / 10);
      const gostValue   = item.gost ? item.gost : "—";
      const cleanName   = formatCleanPipeTitle_(item.name);

      sheet.getRange(row, 1).setValue(item.category);  // A
      sheet.getRange(row, 2).setValue(cleanName);      // B
      sheet.getRange(row, 3).setValue(item.condition); // C
      sheet.getRange(row, 4).setValue(diam);           // D
      sheet.getRange(row, 5).setValue(wall);           // E
      sheet.getRange(row, 6).setValue(gostValue);      // F

      if (buyPriceTn > 0) {
        sheet.getRange(row, 7).setValue(buyPriceTn);   // G: Закупка
      }
      if (sellPriceTn > 0) {
        sheet.getRange(row, 8).setValue(sellPriceTn);  // H: Продажа тн
        if (weightPerMeterTn > 0) {
          sheet.getRange(row, 9).setValue(Math.round(sellPriceTn * weightPerMeterTn)); // I: Продажа м
        }
        sheet.getRange(row, 13).setValue(Math.round(sellPriceTn * addTn)); // M: Сумма
      }

      sheet.getRange(row, 10).setValue("В наличии");   // J: Наличие
      sheet.getRange(row, 11).setValue(addTn);         // K: Остаток тн
      sheet.getRange(row, 12).setValue(addM);          // L: Остаток м
      if (item.notes) {
        sheet.getRange(row, 14).setValue(item.notes);  // N: Примечания
      }

      formatNewWarehouseRow_(sheet, row, diam, wall, now);
      report.push(`➕ НОВАЯ: ${cleanName} (${item.condition}, ${gostValue}) — ${addTn.toFixed(3)} тн`);
    } else {
      row = item.rowIndex;
      const actualName = sheet.getRange(row, 2).getValue().toString().trim();

      if (normalizePipeName(actualName) !== normalizePipeName(item.name)) {
        throw new Error(`⚠ Позиция в строке ${row} изменилась (ожидалась "${item.name}"). Обновите окно.`);
      }

      const currentM  = safeParseFloat(sheet.getRange(row, 12).getValue()); // L
      const currentTn = safeParseFloat(sheet.getRange(row, 11).getValue()); // K
      const addTn     = safeParseFloat(item.tons);

      const diam = safeParseFloat(sheet.getRange(row, 4).getValue());
      const wall = safeParseFloat(sheet.getRange(row, 5).getValue());
      const weightPerMeterTn = ((diam - wall) * wall * 0.0246615) / 1000;

      const addM  = Math.max(0, Math.round((addTn / weightPerMeterTn) * 10) / 10);
      const newM  = Math.max(0, Math.round((currentM + addM) * 10) / 10);
      const newTn = Math.round(newM * weightPerMeterTn * 100) / 100;

      // Средневзвешенная закупочная цена при пополнении существующей позиции
      const incomingBuyTn = safeParseFloat(item.buyPriceTn);
      if (incomingBuyTn > 0) {
        const oldBuyTn = safeParseFloat(sheet.getRange(row, 7).getValue());
        const weightedBuy = (currentTn > 0 && oldBuyTn > 0)
          ? Math.round(((oldBuyTn * currentTn) + (incomingBuyTn * addTn)) / (currentTn + addTn))
          : Math.round(incomingBuyTn);
        sheet.getRange(row, 7).setValue(weightedBuy); // G: Закупка
      }

      sheet.getRange(row, 12).setValue(newM);        // L
      sheet.getRange(row, 11).setValue(newTn);       // K
      sheet.getRange(row, 10).setValue('В наличии'); // J

      const currentPriceTn = safeParseFloat(sheet.getRange(row, 8).getValue()); // H
      if (currentPriceTn > 0) {
        sheet.getRange(row, 13).setValue(Math.round(currentPriceTn * newTn));   // M
      }

      sheet.getRange(row, 15).setValue(now).setNumberFormat('dd.MM.yyyy HH:mm'); // O
      report.push(`📥 ${item.name}: +${addTn.toFixed(3)} тн (Стало: ${newTn.toFixed(3)} тн)`);
    }
  });

  return `✅ Успешно оприходовано!\n\n` + report.join("\n");
}

/**
 * Переработка (Цех): Б/у ➔ Восстановленная (со средневзвешенной себестоимостью)
 */
function executeBatchRestoration(cartItems) {
  if (!cartItems || cartItems.length === 0) return "Буфер цеха пуст!";

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = typeof getSheetByNameSafe === 'function' ? getSheetByNameSafe(ss, 'Склад') : ss.getSheetByName('Склад');
  if (!sheet) throw new Error("Лист «Склад» не найден!");

  const now = new Date();
  let report = [];

  cartItems.forEach(item => {
    const srcRow = item.rowIndex;
    const actualName = sheet.getRange(srcRow, 2).getValue().toString().trim();
    if (normalizePipeName(actualName) !== normalizePipeName(item.name)) {
      throw new Error(`⚠ Исходная позиция Б/у в строке ${srcRow} изменилась. Обновите калькулятор.`);
    }

    const diam = safeParseFloat(sheet.getRange(srcRow, 4).getValue());
    const wall = safeParseFloat(sheet.getRange(srcRow, 5).getValue());
    const weightPerMeterTn = ((diam - wall) * wall * 0.0246615) / 1000;
    if (weightPerMeterTn <= 0) throw new Error(`Ошибка геометрии у ${actualName}`);

    // 1. Списываем с исходной строки Б/у
    const srcCurrentM = safeParseFloat(sheet.getRange(srcRow, 12).getValue());
    const moveM       = safeParseFloat(item.meters);
    const moveTn      = safeParseFloat(item.tons);
    const srcNewM     = Math.max(0, Math.round((srcCurrentM - moveM) * 10) / 10);

    if (srcNewM > 0) {
      const srcNewTn = Math.round(srcNewM * weightPerMeterTn * 100) / 100;
      sheet.getRange(srcRow, 12).setValue(srcNewM);
      sheet.getRange(srcRow, 11).setValue(srcNewTn);
      sheet.getRange(srcRow, 10).setValue('В наличии');
      const srcSellTn = safeParseFloat(sheet.getRange(srcRow, 8).getValue());
      if (srcSellTn > 0) {
        sheet.getRange(srcRow, 13).setValue(Math.round(srcSellTn * srcNewTn));
      }
    } else {
      sheet.getRange(srcRow, 12).clearContent();
      sheet.getRange(srcRow, 11).clearContent();
      sheet.getRange(srcRow, 13).clearContent();
      sheet.getRange(srcRow, 10).setValue('Нет');
    }
    sheet.getRange(srcRow, 15).setValue(now).setNumberFormat('dd.MM.yyyy HH:mm');

    // 2. Считаем новую себестоимость Восстановленной трубы (Закупка Б/у + Работа)
    const srcBuyTn      = safeParseFloat(sheet.getRange(srcRow, 7).getValue());
    const baseCostTn    = safeParseFloat(item.baseBuyPriceTn) || srcBuyTn;
    const serviceCostTn = safeParseFloat(item.serviceCostTn) || 0;
    const restoredBuyTn = baseCostTn > 0 ? Math.round(baseCostTn + serviceCostTn) : 0;

    const targetCat     = (item.targetCategory || sheet.getRange(srcRow, 1).getValue() || 'ЭСВ').toString().trim();
    const targetGost    = (item.targetGost || sheet.getRange(srcRow, 6).getValue() || '—').toString().trim();
    const targetSellTn  = safeParseFloat(item.targetSellPriceTn) || (restoredBuyTn > 0 ? restoredBuyTn + MIN_MARGIN_PER_TON : 0);
    const targetNotes   = (item.targetNotes || 'Пескоструй, фаска').toString().trim();
    const cleanPipeName = formatCleanPipeTitle_(`Труба ${diam}х${wall}`);

    // 3. Ищем существующую строку «Восстановленная» с тем же Ø и Стенкой
    const lastRow = sheet.getLastRow();
    const allRows = sheet.getRange(3, 1, lastRow - 2, 15).getValues();
    let targetRow = -1;

    for (let r = 0; r < allRows.length; r++) {
      const rDiam = safeParseFloat(allRows[r][3]);
      const rWall = safeParseFloat(allRows[r][4]);
      const rCond = String(allRows[r][2] || '').trim().toLowerCase();
      if (rDiam === diam && rWall === wall && rCond.includes('восст')) {
        targetRow = r + 3;
        break;
      }
    }

    if (targetRow !== -1) {
      const oldM     = safeParseFloat(sheet.getRange(targetRow, 12).getValue());
      const oldTn    = safeParseFloat(sheet.getRange(targetRow, 11).getValue());
      const oldBuyTn = safeParseFloat(sheet.getRange(targetRow, 7).getValue());

      const updatedM  = Math.round((oldM + moveM) * 10) / 10;
      const updatedTn = Math.round(updatedM * weightPerMeterTn * 100) / 100;

      sheet.getRange(targetRow, 1).setValue(targetCat);
      sheet.getRange(targetRow, 6).setValue(targetGost);
      sheet.getRange(targetRow, 10).setValue('В наличии');
      sheet.getRange(targetRow, 12).setValue(updatedM);
      sheet.getRange(targetRow, 11).setValue(updatedTn);

      if (restoredBuyTn > 0) {
        const weightedBuy = (oldTn > 0 && oldBuyTn > 0)
          ? Math.round(((oldBuyTn * oldTn) + (restoredBuyTn * moveTn)) / (oldTn + moveTn))
          : restoredBuyTn;
        sheet.getRange(targetRow, 7).setValue(weightedBuy);
      }

      const finalSellTn = targetSellTn || safeParseFloat(sheet.getRange(targetRow, 8).getValue());
      if (finalSellTn > 0) {
        sheet.getRange(targetRow, 8).setValue(finalSellTn);
        sheet.getRange(targetRow, 9).setValue(Math.round(finalSellTn * weightPerMeterTn));
        sheet.getRange(targetRow, 13).setValue(Math.round(finalSellTn * updatedTn));
      }

      if (targetNotes) sheet.getRange(targetRow, 14).setValue(targetNotes);
      sheet.getRange(targetRow, 15).setValue(now).setNumberFormat('dd.MM.yyyy HH:mm');

      report.push(`⚙️ ${cleanPipeName}: ${moveTn.toFixed(3)} тн переведено из Б/у в «Восстановленная» (строка ${targetRow}).`);
    } else {
      const newRow = sheet.getLastRow() + 1;
      sheet.getRange(newRow, 1).setValue(targetCat);
      sheet.getRange(newRow, 2).setValue(cleanPipeName);
      sheet.getRange(newRow, 3).setValue("Восстановленная");
      sheet.getRange(newRow, 4).setValue(diam);
      sheet.getRange(newRow, 5).setValue(wall);
      sheet.getRange(newRow, 6).setValue(targetGost);

      if (restoredBuyTn > 0) sheet.getRange(newRow, 7).setValue(restoredBuyTn);
      if (targetSellTn > 0) {
        sheet.getRange(newRow, 8).setValue(targetSellTn);
        sheet.getRange(newRow, 9).setValue(Math.round(targetSellTn * weightPerMeterTn));
        sheet.getRange(newRow, 13).setValue(Math.round(targetSellTn * moveTn));
      }

      sheet.getRange(newRow, 10).setValue("В наличии");
      sheet.getRange(newRow, 11).setValue(moveTn);
      sheet.getRange(newRow, 12).setValue(moveM);
      sheet.getRange(newRow, 14).setValue(targetNotes);

      formatNewWarehouseRow_(sheet, newRow, diam, wall, now);
      report.push(`⚙️ НОВАЯ ВОССТ.: ${cleanPipeName} — ${moveTn.toFixed(3)} тн.`);
    }
  });

  if (typeof sortWarehouseData === 'function') {
    sortWarehouseData();
  }

  return `✅ Переработка (Б/у ➔ Восстановленная) выполнена!\n\n` + report.join("\n\n");
}

/**
 * Копирование форматов и чип-списков из строки 3 (без навязывания списков в колонку G)
 */
function formatNewWarehouseRow_(sheet, row, diam, wall, now) {
  const templateRow = 3;

  sheet.getRange(templateRow, 1, 1, 15).copyTo(
    sheet.getRange(row, 1, 1, 15),
    SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
    false
  );

  const colsToCopyValidation = [1, 3, 6, 10];
  colsToCopyValidation.forEach(colIdx => {
    sheet.getRange(templateRow, colIdx).copyTo(
      sheet.getRange(row, colIdx),
      SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION,
      false
    );
  });

  sheet.getRange(row, 7).clearDataValidation();

  sheet.getRange(row, 4, 1, 3).setHorizontalAlignment("center");
  sheet.getRange(row, 10, 1, 3).setHorizontalAlignment("center");
  sheet.getRange(row, 4).setNumberFormat((diam % 1 === 0) ? "0" : "0.0");
  sheet.getRange(row, 5).setNumberFormat((wall % 1 === 0) ? "0" : "0.0");
  sheet.setRowHeight(row, 32);
  sheet.getRange(row, 15).setValue(now).setNumberFormat('dd.MM.yyyy HH:mm');
}

/**
 * ЕДИНЫЙ ТРИГГЕР ДЛЯ ЛИСТА «СКЛАД» (15 колонок)
 */
function handleWarehouseSheetEdit(e, sheet) {
  if (!e || !e.range || e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) return;

  const row = e.range.rowStart;
  const col = e.range.columnStart;
  if (row < 3) return;

  const now = new Date();
  const ss = sheet.getParent();

  // 1. АВТО-СОЗДАНИЕ СТРОКИ: Ввод названия в колонку B (2 - Позиция)
  if (col === 2) {
    const rawInput = (e.value !== undefined ? e.value : sheet.getRange(row, 2).getValue()).toString().trim();

    if (rawInput && !rawInput.toUpperCase().startsWith('ИТОГО')) {
      const match = rawInput.match(/(\d{2,4})\s*(?:[хxX*\/x\-\s])\s*(\d+(?:[.,]\d+)?)(.*)/);
      if (match) {
        const diam = parseFloat(match[1]);
        const wall = parseFloat(match[2].replace(',', '.'));
        const extraText = match[3] ? match[3].trim() : "";

        if (diam > 0 && wall > 0) {
          let standardName = `Труба ${diam}х${wall}`;
          if (extraText) standardName += ` ${extraText}`;

          sheet.getRange(row, 2).setValue(standardName);
          sheet.getRange(row, 4, 1, 2).setValues([[diam, wall]]);

          const weightPerMeterTn = ((diam - wall) * wall * 0.0246615) / 1000;
          const currentPriceTn = safeParseFloat(sheet.getRange(row, 8).getValue()); // H

          if (currentPriceTn > 0 && weightPerMeterTn > 0) {
            sheet.getRange(row, 9).setValue(Math.round(currentPriceTn * weightPerMeterTn)); // I
          }

          const currentM = safeParseFloat(sheet.getRange(row, 12).getValue()); // L
          if (currentM > 0 && weightPerMeterTn > 0) {
            const newTn = Math.round(currentM * weightPerMeterTn * 100) / 100;
            sheet.getRange(row, 11).setValue(newTn); // K
            if (currentPriceTn > 0) {
              sheet.getRange(row, 13).setValue(Math.round(currentPriceTn * newTn)); // M
            }
          }
        }
      }

      if (!sheet.getRange(row, 10).getValue()) {
        sheet.getRange(row, 10).setValue('В наличии');
      }

      const currentDiam = safeParseFloat(sheet.getRange(row, 4).getValue());
      const currentWall = safeParseFloat(sheet.getRange(row, 5).getValue());
      formatNewWarehouseRow_(sheet, row, currentDiam, currentWall, now);

      const currentFilter = sheet.getFilter();
      if (currentFilter && currentFilter.getRange().getLastRow() < row) {
        currentFilter.remove();
        sheet.getRange(2, 1, row - 1, 15).createFilter();
      }
    }
    return;
  }

  // 2. РУЧНАЯ ПРАВКА ГЕОМЕТРИИ (Ø в D / 4 или Стенка в E / 5)
  if (col === 4 || col === 5) {
    const geom = sheet.getRange(row, 4, 1, 2).getValues()[0];
    const diam = safeParseFloat(geom[0]);
    const wall = safeParseFloat(geom[1]);

    if (diam > 0 && wall > 0) {
      const currentName = sheet.getRange(row, 2).getValue().toString().trim();
      let extraText = "";
      const match = currentName.match(/(?:\d{2,4})\s*(?:[хxX*\/x\-\s])\s*(?:\d+(?:[.,]\d+)?)(.*)/);
      if (match && match[1]) extraText = match[1].trim();

      let updatedName = `Труба ${diam}х${wall}`;
      if (extraText) updatedName += ` ${extraText}`;
      sheet.getRange(row, 2).setValue(updatedName);

      const weightPerMeterTn = ((diam - wall) * wall * 0.0246615) / 1000;
      const currentPriceTn = safeParseFloat(sheet.getRange(row, 8).getValue()); // H

      if (currentPriceTn > 0) {
        sheet.getRange(row, 9).setValue(Math.round(currentPriceTn * weightPerMeterTn)); // I
      }

      const currentM = safeParseFloat(sheet.getRange(row, 12).getValue()); // L
      if (currentM > 0) {
        const newTn = Math.round(currentM * weightPerMeterTn * 100) / 100;
        sheet.getRange(row, 11).setValue(newTn); // K
        if (currentPriceTn > 0) {
          sheet.getRange(row, 13).setValue(Math.round(currentPriceTn * newTn)); // M
        }
      }

      sheet.getRange(row, 4).setNumberFormat((diam % 1 === 0) ? "0" : "0.0");
      sheet.getRange(row, 5).setNumberFormat((wall % 1 === 0) ? "0" : "0.0");
      sheet.getRange(row, 15).setValue(now).setNumberFormat('dd.MM.yyyy HH:mm');
    }
    return;
  }

  // 3. ДВУХСТОРОННИЙ ПЕРЕСЧЕТ: G(7), H(8), I(9), K(11), L(12)
  if (col === 7 || col === 8 || col === 9 || col === 11 || col === 12) {
    const rowData = sheet.getRange(row, 4, 1, 10).getValues()[0];
    const diam = safeParseFloat(rowData[0]); // D (4)
    const wall = safeParseFloat(rowData[1]); // E (5)

    if (!diam || !wall) return;
    const weightPerMeterTn = ((diam - wall) * wall * 0.0246615) / 1000;
    if (weightPerMeterTn <= 0) return;

    const rawVal = e.value !== undefined ? e.value : rowData[col - 4];
    const val = safeParseFloat(rawVal);
    const buyPriceTn = safeParseFloat(col === 7 ? val : rowData[3]); // G (7)

    // Изменение Закупки (G / 7)
    if (col === 7) {
      const curSellTn = safeParseFloat(rowData[4]); // H (8)
      if (val > 0 && curSellTn > 0 && (curSellTn - val) < MIN_MARGIN_PER_TON) {
        ss.toast(`⚠️ Маржа (${(curSellTn - val).toLocaleString('ru-RU')} ₽/тн) ниже порога Стоп-Лосса (${MIN_MARGIN_PER_TON} ₽/тн)!`, "⛔ Контроль маржи", 5);
      }
      sheet.getRange(row, 15).setValue(now).setNumberFormat('dd.MM.yyyy HH:mm');
      return;
    }

    // Изменение Продажи ₽/тн (H / 8)
    if (col === 8) {
      if (val > 0) {
        if (buyPriceTn > 0 && (val - buyPriceTn) < MIN_MARGIN_PER_TON) {
          ss.toast(`⚠️ СТОП-ЛОСС! Маржа ${(val - buyPriceTn).toLocaleString('ru-RU')} ₽/тн (минимум ${MIN_MARGIN_PER_TON} ₽/тн)!`, "⛔ Низкая маржа", 5);
        }
        sheet.getRange(row, 9).setValue(Math.round(val * weightPerMeterTn)); // I
        const currentTn = safeParseFloat(rowData[7]); // K (11)
        if (currentTn > 0) {
          sheet.getRange(row, 13).setValue(Math.round(val * currentTn)); // M
        }
      } else {
        sheet.getRange(row, 9).clearContent();
        sheet.getRange(row, 13).clearContent();
      }
    }
    // Изменение Продажи ₽/м (I / 9)
    else if (col === 9) {
      if (val > 0) {
        const priceTn = Math.round(val / weightPerMeterTn);
        sheet.getRange(row, 8).setValue(priceTn); // H
        const currentTn = safeParseFloat(rowData[7]); // K (11)
        if (currentTn > 0) {
          sheet.getRange(row, 13).setValue(Math.round(priceTn * currentTn)); // M
        }
      } else {
        sheet.getRange(row, 8).clearContent();
        sheet.getRange(row, 13).clearContent();
      }
    }
    // Изменение Метров (L / 12) -> пересчет Тонн (K / 11)
    else if (col === 12) {
      if (val > 0) {
        const newTn = Math.round(val * weightPerMeterTn * 100) / 100;
        sheet.getRange(row, 11).setValue(newTn);
        sheet.getRange(row, 10).setValue('В наличии');
        const currentPriceTn = safeParseFloat(rowData[4]); // H (8)
        if (currentPriceTn > 0) {
          sheet.getRange(row, 13).setValue(Math.round(currentPriceTn * newTn));
        }
      } else {
        sheet.getRange(row, 11).clearContent();
        sheet.getRange(row, 10).setValue('Нет');
        sheet.getRange(row, 13).clearContent();
      }
    }
    // Изменение Тонн (K / 11) -> пересчет Метров (L / 12)
    else if (col === 11) {
      if (val > 0) {
        const newM = Math.round((val / weightPerMeterTn) * 10) / 10;
        sheet.getRange(row, 12).setValue(newM);
        sheet.getRange(row, 10).setValue('В наличии');
        const currentPriceTn = safeParseFloat(rowData[4]); // H (8)
        if (currentPriceTn > 0) {
          sheet.getRange(row, 13).setValue(Math.round(currentPriceTn * val));
        }
      } else {
        sheet.getRange(row, 12).clearContent();
        sheet.getRange(row, 10).setValue('Нет');
        sheet.getRange(row, 13).clearContent();
      }
    }

    sheet.getRange(row, 15).setValue(now).setNumberFormat('dd.MM.yyyy HH:mm');
  }
}