/**
 * ЕДИНЫЙ ЦЕНТР УПРАВЛЕНИЯ ПРОДАЖАМИ И СКЛАДОМ
 * Файл: Calling.gs — Маршрутизация лидов, анти-дубль ИНН, телефоны и конвейер обзвона (Стандарт A–N)
 */

function handleCallingSheetEdit(e, sheet) {
  const row = e.range.rowStart;
  const col = e.range.columnStart;
  const sheetName = sheet.getName().trim();

  // Игнорируем заголовки (строка 1)
  if (row < 2) return;

  // =================================================================
  // 0. КОНТРОЛЬ И ПОДСВЕТКА ДУБЛЕЙ ИНН (Колонка A / 1)
  // =================================================================
  if (col === 1) {
    const rawInn = e.value !== undefined ? String(e.value) : String(sheet.getRange(row, 1).getValue());
    const cleanInn = rawInn.replace(/\D/g, '');
    if (cleanInn && cleanInn !== rawInn.trim()) {
      sheet.getRange(row, 1).setNumberFormat('@').setValue(cleanInn);
    }
    highlightDuplicateInnsOnSheet_(sheet);
    if (cleanInn.length >= 10) {
      checkInnAcrossAllSheets_(sheet.getParent(), sheetName, row, cleanInn);
    }
    return;
  }

  // =================================================================
  // 1. АВТО-ОЧИСТКА НАЗВАНИЯ КОМПАНИИ (Колонка B / 2)
  // =================================================================
  if (col === 2) {
    const compVal = e.value !== undefined ? String(e.value) : String(sheet.getRange(row, col).getValue());
    if (compVal && typeof cleanCompanyName === 'function') {
      const cleaned = cleanCompanyName(compVal);
      if (cleaned && cleaned !== compVal) {
        sheet.getRange(row, col).setValue(cleaned);
      }
    }
    return;
  }

  // =================================================================
  // 2. ЖЕЛЕЗОБЕТОННОЕ ФОРМАТИРОВАНИЕ ТЕЛЕФОНОВ: МОБИЛЬНЫЕ + ГОРОДСКИЕ (Колонка E / 5)
  // =================================================================
  if (col === 5) {
    const cell = sheet.getRange(row, col);
    const phoneVal = e.value !== undefined ? String(e.value) : String(cell.getValue());
    if (phoneVal) {
      const formatted = formatRussianPhones_(phoneVal);
      if (formatted) {
        cell.setNumberFormat('@'); // Защита от превращения +7 в формулу
        if (formatted !== phoneVal.trim()) {
          cell.setValue(formatted);
        }
      }
    }
    return;
  }

  // =================================================================
  // 3. МАРШРУТИЗАЦИЯ ЛИДОВ И КОНВЕЙЕР ОБЗВОНА (Колонка M / 13)
  // =================================================================
  if (col !== 13) return;

  const decisionVal = e.value !== undefined ? String(e.value) : String(sheet.getRange(row, col).getValue());
  const decision = decisionVal.trim().toLowerCase();
  if (!decision) return;

  const ss = sheet.getParent();
  const timeZone = ss.getSpreadsheetTimeZone() || "Europe/Moscow";
  const now = new Date();
  const nowStr = Utilities.formatDate(now, timeZone, "dd.MM.yyyy HH:mm");

  // Считываем все 14 колонок (A..N)
  const rowData = sheet.getRange(row, 1, 1, 14).getValues()[0];
  const inn         = String(rowData[0] || '').trim();
  const company     = typeof cleanCompanyName === 'function' ? cleanCompanyName(rowData[1]) : String(rowData[1] || '').trim();
  const city        = String(rowData[2] || '').trim();
  const lpr         = String(rowData[3] || '').trim();
  const phone       = formatRussianPhones_(String(rowData[4] || '').trim());
  const email       = String(rowData[5] || '').trim();
  const sphere      = String(rowData[6] || '').trim(); // G(7): Сфера
  const need        = String(rowData[7] || '').trim(); // H(8): Потребность
  const pipeType    = String(rowData[8] || '').trim(); // I(9): Тип трубы
  const application = String(rowData[9] || '').trim(); // J(10): Под что берут
  let notes         = String(rowData[10] || '').trim(); // K(11): Примечание (ручное)
  let deadline      = rowData[11];                      // L(12): Дедлайн
  let history       = String(rowData[13] || '').trim(); // N(14): История (системные логи)

  let formattedDeadline = "";
  if (deadline instanceof Date) {
    formattedDeadline = Utilities.formatDate(deadline, timeZone, "yyyy-MM-dd");
  } else if (deadline) {
    formattedDeadline = String(deadline).trim();
  }

  // === ЗАЩИТА ПРАВИЛА №6: ЗАПРЕТ РУЧНОГО ПЕРЕВОДА В КИТЫ ===
  if (decision.includes("кит") || decision.includes("золот")) {
    sheet.getRange(row, 13).clearContent();
    ss.toast(
      "В «Киты» клиент попадает ТОЛЬКО автоматически после оплаты заказа! Выберите «Выставлен счет» или «В Ожидание».",
      "⛔ Правило №6 (Закрытый клуб)",
      6
    );
    return;
  }

  // === ВЕТКА 1: ВЫСТАВЛЕН СЧЕТ ===
  if (decision.includes("выставлен счет") || decision.includes("счет") || decision.includes("счёт") || decision.includes("заказ")) {
    const tag = `[💰 Выставлен счет (Контроль оплаты) - ${nowStr}]`;
    const updatedHistory = prependHistoryLog_(history, tag);
    const nextCheckDate = getNextWorkingDateStr_(now, 1, timeZone);

    if (sheetName === "1. Обзвон (База)") {
      const targetSheet = getSheetByNameSafe(ss, "3. Лист ожидания");
      if (!targetSheet) return ss.toast("Ошибка: лист Ожидания не найден", "Ошибка");

      insertAndFormatRow_(targetSheet, [
        inn, company, city, lpr, phone, email,
        sphere, need, pipeType, application,
        notes, nextCheckDate, "", updatedHistory
      ]);
      sheet.deleteRow(row);
      highlightDuplicateInnsOnSheet_(sheet);
      ss.toast(`«${company || inn}» -> Счет зафиксирован! Перенесен в Ожидание (контроль завтра).`, "💰 Выставлен счет", 5);
    } else {
      sheet.getRange(row, 14).setValue(updatedHistory);
      sheet.getRange(row, 12).setValue(nextCheckDate);
      sheet.getRange(row, 13).clearContent();
      moveRowToBottom_(sheet, row);
      ss.toast(`«${company || inn}» -> Счет зафиксирован в Истории! Опущен вниз списка.`, "💰 Выставлен счет", 4);
    }
    return;
  }

  // === ВЕТКА 2: ВОЗВРАТ ИЗ АРХИВА В РАБОТУ ===
  if (sheetName === "4.Архив лидов" && (decision.includes("обзвон") || decision.includes("ожидан") || decision.includes("вернуть"))) {
    const toWaiting = decision.includes("ожидан");
    const targetName = toWaiting ? "3. Лист ожидания" : "1. Обзвон (База)";
    const targetSheet = getSheetByNameSafe(ss, targetName);
    if (!targetSheet) return ss.toast(`Ошибка: лист «${targetName}» не найден`, "Ошибка");

    const preservedHistory = history.replace(/\[📞 Недозвон:/gi, "[📞 Архивный недозвон:");
    const tag = `[♻️ Возврат из Архива в «${targetName}» - ${nowStr}]`;
    const updatedHistory = prependHistoryLog_(preservedHistory, tag);
    const nextDate = getNextWorkingDateStr_(now, 0, timeZone);

    insertAndFormatRow_(targetSheet, [
      inn, company, city, lpr, phone, email,
      sphere, need, pipeType, application,
      notes, nextDate, "", updatedHistory
    ]);
    sheet.deleteRow(row);
    ss.toast(`«${company || inn}» восстановлен в «${targetName}» со всей историей!`, "♻️ Возврат лида", 4);
    return;
  }

  // === ВЕТКА 3А: ОПЕРАТИВНЫЙ ПЕРЕЗВОН СЕГОДНЯ / ЧЕРЕЗ 2 ЧАСА (Спускает строку вниз таблицы) ===
  if (decision.includes("2 час") || decision.includes("два час") || decision.includes("позже") || decision.includes("перезвон") || decision.includes("сегодня")) {
    const callbackTime = new Date(now.getTime() + 2 * 60 * 60 * 1000);
    const callbackTimeStr = Utilities.formatDate(callbackTime, timeZone, "HH:mm");
    const todayStr = Utilities.formatDate(now, timeZone, "yyyy-MM-dd");

    const tag = `[⏰ Перезвонить после ${callbackTimeStr} - ${nowStr}]`;
    const updatedHistory = prependHistoryLog_(history, tag);

    sheet.getRange(row, 14).setValue(updatedHistory);
    sheet.getRange(row, 12).setValue(todayStr);
    sheet.getRange(row, 13).clearContent();

    moveRowToBottom_(sheet, row);
    ss.toast(`«${company || inn}» -> Перезвон после ${callbackTimeStr}. Строка убрана вниз таблицы!`, "⏰ Через 2 часа", 4);
    return;
  }

  // === ВЕТКА 3Б: ЛИСТ ОЖИДАНИЯ -> Перенос в «3. Лист ожидания» ===
  if (decision.includes("ожидан") || decision.includes("подума") || decision.includes("спящ")) {
    if (sheetName === "3. Лист ожидания") {
      sheet.getRange(row, 13).clearContent();
      moveRowToBottom_(sheet, row);
      ss.toast("Клиент уже в Листе ожидания (перемещен вниз списка)", "⏳ Инфо", 3);
      return;
    }
    const targetSheet = getSheetByNameSafe(ss, "3. Лист ожидания");
    if (!targetSheet) return ss.toast("Ошибка: лист Ожидания не найден", "Ошибка");

    const tag = `[⏳ Перевод в Ожидание (из «${sheetName}») - ${nowStr}]`;
    const updatedHistory = prependHistoryLog_(history, tag);

    insertAndFormatRow_(targetSheet, [
      inn, company, city, lpr, phone, email,
      sphere, need, pipeType, application,
      notes, formattedDeadline, "", updatedHistory
    ]);
    sheet.deleteRow(row);
    highlightDuplicateInnsOnSheet_(sheet);
    ss.toast(`«${company || inn}» перенесён в Лист ожидания!`, "⏳ Ожидание", 4);
    return;
  }

  // === ВЕТКА 4: НЕДОЗВОН (С переносом вниз таблицы на попытках 1/4..3/4 и в Архив на 4/4) ===
  if (decision.includes("недозвон") || decision.includes("не дозвон") || decision.includes("сброс") || decision.includes("не берет")) {
    const match = history.match(/\[📞 Недозвон:\s*(\d+)\/4(?:\s*-\s*[^\]]+)?\]/i);
    let attempts = match ? parseInt(match[1], 10) : 0;
    attempts += 1;

    const callTag = `[📞 Недозвон: ${attempts}/4 - ${nowStr}]`;
    let updatedHistory = match
      ? history.replace(/\[📞 Недозвон:\s*\d+\/4(?:\s*-\s*[^\]]+)?\]/i, callTag)
      : prependHistoryLog_(history, callTag);

    if (attempts < 4) {
      const nextCallStr = getNextWorkingDateStr_(now, 1, timeZone);
      sheet.getRange(row, 14).setValue(updatedHistory);
      sheet.getRange(row, 12).setValue(nextCallStr);
      sheet.getRange(row, 13).clearContent();

      // Опускаем обработанную строку в самый низ таблицы, чтобы не мешалась при обзвоне
      moveRowToBottom_(sheet, row);
      ss.toast(`Недозвон (${attempts}/4). Дедлайн: ${nextCallStr}. Строка опущена вниз!`, "📞 Недозвон", 4);
    } else {
      const targetSheet = getSheetByNameSafe(ss, "4.Архив лидов");
      if (!targetSheet) return ss.toast("Ошибка: лист Архива не найден", "Ошибка");

      const archiveTag = `[📦 В Архив: 4 недозвона (из «${sheetName}») - ${nowStr}]`;
      const finalArchiveHistory = prependHistoryLog_(updatedHistory, archiveTag);

      insertAndFormatRow_(targetSheet, [
        inn, company, city, lpr, phone, email,
        sphere, need, pipeType, application,
        notes, "", "", finalArchiveHistory
      ]);
      sheet.deleteRow(row);
      highlightDuplicateInnsOnSheet_(sheet);
      ss.toast(`«${company || inn}» перенесён в Архив (4/4 недозвона). Все данные сохранены!`, "📦 Архив", 4);
    }
    return;
  }

  // === ВЕТКА 5: ОТКАЗЫ / НЕ ЦЕЛЕВОЙ / ТОКСИЧНЫЙ ДЕМПИНГ -> в «4.Архив лидов» ===
  const isToxic     = decision.includes("токсич") || decision.includes("демпинг") || decision.includes("террор");
  const isNotTarget = decision.includes("не целев") || decision.includes("нецелев");
  const isRefusal   = decision.includes("отказ") || decision.includes("архив") || decision.includes("удалит") || decision.includes("слил");

  if (isToxic || isNotTarget || isRefusal) {
    const targetSheet = getSheetByNameSafe(ss, "4.Архив лидов");
    if (!targetSheet) return ss.toast("Ошибка: лист Архива не найден", "Ошибка");

    let reasonText = "Отказ";
    let icon = "⛔";
    if (isToxic) {
      reasonText = "Токсичный демпинг (Прайс-террорист)";
      icon = "☢️";
    } else if (isNotTarget) {
      reasonText = "Не целевой";
      icon = "🚫";
    }

    const archiveTag = `[${icon} В Архив: ${reasonText} (из «${sheetName}») - ${nowStr}]`;
    const finalArchiveHistory = prependHistoryLog_(history, archiveTag);

    insertAndFormatRow_(targetSheet, [
      inn, company, city, lpr, phone, email,
      sphere, need, pipeType, application,
      notes, "", "", finalArchiveHistory
    ]);
    sheet.deleteRow(row);
    highlightDuplicateInnsOnSheet_(sheet);
    ss.toast(`«${company || inn}» отправлен в Архив (${reasonText}). Все 14 колонок сохранены!`, `${icon} ${reasonText}`, 4);
    return;
  }
}

// =================================================================
// СЛУЖЕБНЫЕ ФУНКЦИИ: ТЕЛЕФОНЫ, ДУБЛИ ИНН И ПЕРЕМЕЩЕНИЕ ВНИЗ
// =================================================================

/**
 * Умное форматирование любых российских номеров (мобильные, городские, 8-800, 7-значные Воронежа)
 */
function formatRussianPhones_(rawPhone) {
  if (!rawPhone) return "";

  // Разбиваем по переводам строки, точкам с запятой, запятым или слешам
  const rawChunks = rawPhone.toString()
    .replace(/(\d)\s*[,;\/]\s*(?=(?:\+7|8|7|9|4|3)\d)/g, '\$1\n')
    .split(/\n+/);

  const resultLines = [];

  rawChunks.forEach(chunk => {
    let line = chunk.trim();
    if (!line) return;

    // Ищем телефонные комбинации (с пробелами, скобками, тире, точками)
    const phoneRegex = /(?:(?:\+7|8|7)[\s\-\.\(\)]*)?(?:\(\s*\d{3,5}\s*\)|\d{3,5})[\s\-\.\(\)]*\d{2,3}[\s\-\.]*\d{2}[\s\-\.]*\d{2}|(?:\+7|8|7)?\d{10}|\b2\d{2}[\s\-]?\d{2}[\s\-]?\d{2}\b/g;

    let matchFound = false;
    const formattedChunk = line.replace(phoneRegex, (matchStr) => {
      const digits = matchStr.replace(/\D/g, '');
      let d10 = "";

      if (digits.length === 11 && (digits.startsWith('7') || digits.startsWith('8'))) {
        d10 = digits.substring(1);
      } else if (digits.length === 10 && /^[3489]/.test(digits)) {
        d10 = digits;
      } else if (digits.length === 7 && digits.startsWith('2')) {
        // 7-значный городской номер Воронежа (2xx-xx-xx) -> код 473
        d10 = '473' + digits;
      }

      if (d10.length === 10) {
        matchFound = true;
        return `+7 (${d10.slice(0, 3)}) ${d10.slice(3, 6)}-${d10.slice(6, 8)}-${d10.slice(8, 10)}`;
      }
      return matchStr;
    });

    // Если в одной строке склеились два номера +7 (...) подряд — переносим второй на новую строку
    const splitMultiple = formattedChunk
      .replace(/(\d{2}-\d{2})\s*[,;]?\s*(\+7\s*\()/g, '\$1\n\$2')
      .trim();

    if (splitMultiple) {
      resultLines.push(splitMultiple);
    }
  });

  return resultLines.join('\n');
}

/**
 * Перемещение отработанной строки в самый низ активной таблицы обзвона
 */
function moveRowToBottom_(sheet, row) {
  const maxRows = sheet.getMaxRows();
  const checkVals = sheet.getRange(1, 1, maxRows, 5).getValues();
  let lastRealRow = 1;

  for (let r = checkVals.length - 1; r >= 1; r--) {
    const hasData = checkVals[r].some(cell => String(cell || '').trim() !== '');
    if (hasData) {
      lastRealRow = r + 1;
      break;
    }
  }

  // Если строка уже последняя — никуда двигать не нужно
  if (row >= lastRealRow) return;

  if (lastRealRow >= maxRows) {
    sheet.insertRowAfter(maxRows);
  }

  sheet.moveRows(sheet.getRange(row, 1, 1, sheet.getLastColumn()), lastRealRow + 1);
}

/**
 * Подсветка дублей ИНН внутри текущего листа (Колонка A)
 */
function highlightDuplicateInnsOnSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const numRows = lastRow - 1;
  const innRange = sheet.getRange(2, 1, numRows, 1);
  const innValues = innRange.getValues();
  const counts = {};

  innValues.forEach(r => {
    const clean = String(r[0] || '').replace(/\D/g, '');
    if (clean.length >= 10) {
      counts[clean] = (counts[clean] || 0) + 1;
    }
  });

  const bgColors = [];
  const fontColors = [];
  const fontWeights = [];
  let dupCount = 0;

  innValues.forEach(r => {
    const clean = String(r[0] || '').replace(/\D/g, '');
    if (clean.length >= 10 && counts[clean] > 1) {
      bgColors.push(['#fee2e2']);   // Ярко-розовый фон дубля
      fontColors.push(['#991b1b']); // Темно-красный текст
      fontWeights.push(['bold']);
      dupCount++;
    } else {
      bgColors.push([null]);
      fontColors.push(['#0f172a']);
      fontWeights.push(['normal']);
    }
  });

  innRange.setBackgrounds(bgColors);
  innRange.setFontColors(fontColors);
  innRange.setFontWeights(fontWeights);
  return dupCount;
}

/**
 * Проверка введенного ИНН по всем 4 базам CRM
 */
function checkInnAcrossAllSheets_(ss, currentSheetName, currentRow, targetInn) {
  const crmSheets = [
    "1. Обзвон (База)",
    "2. Золотой список (Киты)",
    "3. Лист ожидания",
    "4.Архив лидов"
  ];

  const matches = [];

  crmSheets.forEach(sName => {
    const sh = getSheetByNameSafe(ss, sName);
    if (!sh) return;
    const lRow = sh.getLastRow();
    if (lRow < 2) return;

    const data = sh.getRange(2, 1, lRow - 1, 2).getValues();
    data.forEach((r, idx) => {
      const rNum = idx + 2;
      if (sName === currentSheetName && rNum === currentRow) return;
      const cellInn = String(r[0] || '').replace(/\D/g, '');
      if (cellInn === targetInn) {
        const compName = String(r[1] || 'Без названия').trim();
        matches.push(`«${sName}» (стр. ${rNum}: ${compName})`);
      }
    });
  });

  if (matches.length > 0) {
    ss.toast(
      `ИНН ${targetInn} уже есть в базе:\n` + matches.join(';\n'),
      "⚠️ ВНИМАНИЕ: ДУБЛЬ ИНН!",
      8
    );
  }
}

/**
 * Ручной запуск проверки дублей ИНН из меню «📱 Управление»
 */
function highlightAllDuplicateInnsManual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const dupCells = highlightDuplicateInnsOnSheet_(sheet);
  if (dupCells > 0) {
    ss.toast(`Найдено и подсвечено красным ячеек с дублями ИНН: ${dupCells} шт.`, "⚠️ Найдены дубли ИНН", 5);
  } else {
    ss.toast("На текущем листе дублей ИНН нет — все компании уникальны!", "✅ Дублей нет", 4);
  }
}

function prependHistoryLog_(existingHistory, newTag) {
  const cleanHist = String(existingHistory || '').trim();
  if (!cleanHist) return newTag;
  if (cleanHist.includes(newTag)) return cleanHist;
  return `${newTag}\n${cleanHist}`;
}

function getNextWorkingDateStr_(baseDate, addDays, timeZone) {
  const nextDate = new Date(baseDate.getTime());
  nextDate.setDate(nextDate.getDate() + addDays);
  if (nextDate.getDay() === 6) nextDate.setDate(nextDate.getDate() + 2);
  else if (nextDate.getDay() === 0) nextDate.setDate(nextDate.getDate() + 1);
  return Utilities.formatDate(nextDate, timeZone || "Europe/Moscow", "yyyy-MM-dd");
}

function insertAndFormatRow_(targetSheet, rowValues) {
  const maxRows = targetSheet.getMaxRows();
  const checkRange = targetSheet.getRange(1, 1, maxRows, 2).getValues();
  let lastRealRow = 1;

  for (let r = checkRange.length - 1; r >= 1; r--) {
    if (String(checkRange[r][0]).trim() !== "" || String(checkRange[r][1]).trim() !== "") {
      lastRealRow = r + 1;
      break;
    }
  }

  const targetRow = lastRealRow + 1;
  if (targetRow > maxRows) targetSheet.insertRowAfter(maxRows);

  const numCols = rowValues.length;
  const targetRange = targetSheet.getRange(targetRow, 1, 1, numCols);

  if (lastRealRow >= 2) {
    const templateRange = targetSheet.getRange(lastRealRow, 1, 1, numCols);
    templateRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    templateRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  }
  targetSheet.getRange(targetRow, 1).setNumberFormat('@');
  targetSheet.getRange(targetRow, 5).setNumberFormat('@');
  targetRange.setValues([rowValues]);
}