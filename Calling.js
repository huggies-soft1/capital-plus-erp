/**
 * ЕДИНЫЙ ЦЕНТР УПРАВЛЕНИЯ ПРОДАЖАМИ И СКЛАДОМ
 * Файл: Calling.gs — Маршрутизация лидов и авто-формат телефонов (Стандарт A–N)
 */

function handleCallingSheetEdit(e, sheet) {
  const row = e.range.rowStart;
  const col = e.range.columnStart;
  const sheetName = sheet.getName().trim();

  // Игнорируем заголовки (строка 1)
  if (row < 2) return;

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
  // 2. БЕРЕЖНОЕ АВТО-ФОРМАТИРОВАНИЕ ТЕЛЕФОНОВ (Колонка E / 5)
  // =================================================================
  if (col === 5) {
    const phoneVal = e.value !== undefined ? String(e.value) : String(sheet.getRange(row, col).getValue());
    if (phoneVal) {
      const formatted = formatRussianPhones_(phoneVal);
      if (formatted && formatted !== phoneVal.trim()) {
        sheet.getRange(row, col).setValue(formatted);
      }
    }
    return;
  }

  // =================================================================
  // 3. МАРШРУТИЗАЦИЯ ЛИДОВ (Колонка M / 13)
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
      ss.toast(`«${company || inn}» -> Счет зафиксирован! Перенесен в Ожидание (контроль завтра).`, "💰 Выставлен счет", 5);
    } else {
      sheet.getRange(row, 14).setValue(updatedHistory);
      sheet.getRange(row, 12).setValue(nextCheckDate);
      sheet.getRange(row, 13).clearContent();
      ss.toast(`«${company || inn}» -> Счет зафиксирован в Истории! Дедлайн: ${nextCheckDate}`, "💰 Выставлен счет", 4);
    }
    return;
  }

  // === ВЕТКА 2: ВОЗВРАТ ИЗ АРХИВА В РАБОТУ ===
  if (sheetName === "4.Архив лидов" && (decision.includes("обзвон") || decision.includes("ожидан") || decision.includes("вернуть"))) {
    const toWaiting = decision.includes("ожидан");
    const targetName = toWaiting ? "3. Лист ожидания" : "1. Обзвон (База)";
    const targetSheet = getSheetByNameSafe(ss, targetName);
    if (!targetSheet) return ss.toast(`Ошибка: лист «${targetName}» не найден`, "Ошибка");

    // Сохраняем старые недозвоны как архивные (не удаляя их!), чтобы новый счетчик шел с 1/4
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

  // === ВЕТКА 3: ЛИСТ ОЖИДАНИЯ -> в «3. Лист ожидания» ===
  if (decision.includes("ожидан") || decision.includes("перезвон") || decision.includes("подума") || decision.includes("спящ")) {
    if (sheetName === "3. Лист ожидания") {
      sheet.getRange(row, 13).clearContent();
      ss.toast("Клиент уже находится в Листе ожидания", "⏳ Инфо", 3);
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
    ss.toast(`«${company || inn}» перенесён в Лист ожидания!`, "⏳ Ожидание", 4);
    return;
  }

  // === ВЕТКА 4: НЕДОЗВОН (С полной сохранностью логов 4/4 и всех 14 колонок A–N) ===
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
      ss.toast(`Недозвон (${attempts}/4). Дата сдвинута на ${nextCallStr}`, "📞 Недозвон", 4);
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
      ss.toast(`«${company || inn}» перенесён в Архив (4/4 недозвона). Все данные и логи сохранены!`, "📦 Архив", 4);
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
    ss.toast(`«${company || inn}» отправлен в Архив (${reasonText}). Все 14 колонок сохранены!`, `${icon} ${reasonText}`, 4);
    return;
  }
}

// =================================================================
// СЛУЖЕБНЫЕ ФУНКЦИИ
// =================================================================

function formatRussianPhones_(rawPhone) {
  if (!rawPhone) return "";
  const lines = rawPhone.toString().trim().split(/[\n;]+/);
  const formattedLines = lines.map(line => {
    let cleanLine = line.trim();
    if (!cleanLine) return "";

    cleanLine = cleanLine.replace(/(^|[^\d+])[78](\d{3})(\d{3})(\d{2})(\d{2})(?!\d)/g, (m, p, c1, c2, c3, c4) => {
      return `${p}+7 (${c1}) ${c2}-${c3}-${c4}`;
    });
    cleanLine = cleanLine.replace(/(^|[^\d+])([94]\d{2})(\d{3})(\d{2})(\d{2})(?!\d)/g, (m, p, c1, c2, c3, c4) => {
      return `${p}+7 (${c1}) ${c2}-${c3}-${c4}`;
    });
    cleanLine = cleanLine.replace(/(\))\s*[,;]?\s*(\+7\s*\()/g, '\$1\n\$2');
    cleanLine = cleanLine.replace(/(\d)\s*[,;\s]+\s*(\+7\s*\()/g, '\$1\n\$2');

    return cleanLine.trim();
  }).filter(Boolean);

  return formattedLines.join('\n');
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
  targetRange.setValues([rowValues]);
}