/**
 * ЕДИНЫЙ ЦЕНТР УПРАВЛЕНИЯ ПРОДАЖАМИ И СКЛАДОМ (CTRM «Капитал-Плюс» v6.1)
 * Файл: Calling.gs — Маршрутизация лидов, скоростной парсер труб (53010 / 530*10 / 630 8),
 * Авто-Досье клиента (перенос «отсебятины» из K в Заметку карточки B с выжимкой сути),
 * самоочистка зависших потребностей (H -> в Историю N со сбросом в «Нет») и живой мозг колонки O.
 */

const STANDARD_PIPE_DIAMS_ = [
  1420, 1220, 1020, 920, 820, 720, 630, 530, 478, 426, 402, 377, 351,
  325, 299, 273, 245, 219, 194, 168, 159, 152, 148, 146, 133, 127,
  114, 108, 102, 89, 76, 73, 60, 57
];

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
        cell.setNumberFormat('@');
        if (formatted !== phoneVal.trim()) {
          cell.setValue(formatted);
        }
      }
    }
    return;
  }

  // =================================================================
  // 2.5 СКОРОСТНОЙ ВВОД ПОТРЕБНОСТИ ТРУБЫ (Колонка H / 8: 53010 -> 530х10, 630 8 -> 630х8)
  // =================================================================
  if (col === 8) {
    const cell = sheet.getRange(row, col);
    const rawNeed = e.value !== undefined ? String(e.value) : String(cell.getValue());
    if (rawNeed) {
      const formattedNeed = formatPipeNeed_(rawNeed);
      if (formattedNeed && formattedNeed !== rawNeed.trim()) {
        cell.setNumberFormat('@');
        cell.setValue(formattedNeed);
      }
    }
    if (sheetName === "3. Лист ожидания") {
      refreshSingleRowTactic_(sheet, row);
    }
    return;
  }

  // =================================================================
  // 2.8 АВТО-ДОСЬЕ КЛИЕНТА ПРИ ВВОДЕ «ОТСЕБЯТИНЫ» В ПРИМЕЧАНИЕ (Колонка K / 11)
  // =================================================================
  if (col === 11) {
    processRawNoteIntoDossier_(sheet, row);
    if (sheetName === "3. Лист ожидания") {
      refreshSingleRowTactic_(sheet, row);
    }
    return;
  }

  // При изменении Сферы (G), Типа трубы (I), Назначения (J) или Истории (N) на Листе ожидания — мгновенно обновляем тактику в колонке O
  if (sheetName === "3. Лист ожидания" && (col === 7 || col === 9 || col === 10 || col === 14)) {
    refreshSingleRowTactic_(sheet, row);
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

  // Сначала упаковываем свежую отсебятину из K в Досье (если она там есть)
  processRawNoteIntoDossier_(sheet, row);

  // Считываем все 14 базовых колонок (A..N) + полное Досье из заметки на ячейке компании (B)
  const rowData     = sheet.getRange(row, 1, 1, 14).getValues()[0];
  const companyNote = String(sheet.getRange(row, 2).getNote() || '').trim();

  const inn         = String(rowData[0] || '').trim();
  const company     = typeof cleanCompanyName === 'function' ? cleanCompanyName(rowData[1]) : String(rowData[1] || '').trim();
  const city        = String(rowData[2] || '').trim();
  const lpr         = String(rowData[3] || '').trim();
  const phone       = formatRussianPhones_(String(rowData[4] || '').trim());
  const email       = String(rowData[5] || '').trim();
  const sphere      = String(rowData[6] || '').trim(); // G(7): Сфера (постоянный профиль по ОКВЭД)
  const need        = formatPipeNeed_(String(rowData[7] || '').trim()); // H(8): Живая потребность сейчас
  const pipeType    = String(rowData[8] || '').trim(); // I(9): Тип трубы сейчас
  const application = String(rowData[9] || '').trim(); // J(10): Под что берут сейчас
  let notes         = String(rowData[10] || '').trim(); // K(11): Краткая выжимка примечания
  let deadline      = rowData[11];                      // L(12): Дедлайн
  let history       = String(rowData[13] || '').trim(); // N(14): История (системные логи + архив всех прошлых заявок)

  // Формируем слепок текущей живой заявки (H + I + J) для автоматической фиксации в Истории (N)
  const needSnapshot = buildNeedSnapshot_(need, pipeType, application);

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
    const tag = `[💰 Выставлен счет${needSnapshot} - ${nowStr}]`;
    const updatedHistory = prependHistoryLog_(history, tag);
    const nextCheckDate = getNextWorkingDateStr_(now, 1, timeZone);

    if (sheetName === "1. Обзвон (База)") {
      const targetSheet = getSheetByNameSafe(ss, "3. Лист ожидания");
      if (!targetSheet) return ss.toast("Ошибка: лист Ожидания не найден", "Ошибка");

      const mergeRes = insertOrMergeRowByInn_(targetSheet, [
        inn, company, city, lpr, phone, email,
        sphere, need || "В просчете", pipeType, application,
        notes, nextCheckDate, "", updatedHistory
      ], companyNote);
      refreshSingleRowTactic_(targetSheet, mergeRes.row);
      sheet.deleteRow(row);
      highlightDuplicateInnsOnSheet_(sheet);
      const actionWord = mergeRes.merged ? "объединен с карточкой" : "перенесен";
      ss.toast(`«${company || inn}» -> Счет зафиксирован! ${actionWord} в Ожидание (контроль: ${nextCheckDate}).`, "💰 Выставлен счет", 5);
    } else {
      sheet.getRange(row, 8).setValue(need);
      sheet.getRange(row, 14).setValue(updatedHistory);
      sheet.getRange(row, 12).setValue(nextCheckDate);
      sheet.getRange(row, 13).clearContent();
      if (sheetName === "3. Лист ожидания") refreshSingleRowTactic_(sheet, row);
      moveRowToBottom_(sheet, row);
      ss.toast(`«${company || inn}» -> Счет зафиксирован! Тактика переключена на дожим счета (${nextCheckDate}).`, "💰 Выставлен счет", 4);
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
    const tag = `[♻️ Возврат из Архива в «${targetName}»${needSnapshot} - ${nowStr}]`;
    const updatedHistory = prependHistoryLog_(preservedHistory, tag);
    const nextDate = getNextWorkingDateStr_(now, toWaiting ? 7 : 0, timeZone);

    const mergeRes = insertOrMergeRowByInn_(targetSheet, [
      inn, company, city, lpr, phone, email,
      sphere, toWaiting ? "Нет" : need, "", "",
      notes, nextDate, "", updatedHistory
    ], companyNote);
    if (toWaiting) refreshSingleRowTactic_(targetSheet, mergeRes.row);
    sheet.deleteRow(row);
    ss.toast(`«${company || inn}» восстановлен в «${targetName}» со всей историей и Досье!`, "♻️ Возврат лида", 4);
    return;
  }

  // === ВЕТКА 3А: ОПЕРАТИВНЫЙ ПЕРЕЗВОН СЕГОДНЯ / ЧЕРЕЗ 2 ЧАСА ===
  if (decision.includes("2 час") || decision.includes("два час") || decision.includes("перезвон") || decision.includes("сегодня")) {
    const callbackTime = new Date(now.getTime() + 2 * 60 * 60 * 1000);
    const callbackTimeStr = Utilities.formatDate(callbackTime, timeZone, "HH:mm");
    const todayStr = Utilities.formatDate(now, timeZone, "yyyy-MM-dd");

    const tag = `[⏰ Перезвонить после ${callbackTimeStr}${needSnapshot} - ${nowStr}]`;
    const updatedHistory = prependHistoryLog_(history, tag);

    sheet.getRange(row, 8).setValue(need);
    sheet.getRange(row, 14).setValue(updatedHistory);
    sheet.getRange(row, 12).setValue(todayStr);
    sheet.getRange(row, 13).clearContent();
    if (sheetName === "3. Лист ожидания") refreshSingleRowTactic_(sheet, row);

    moveRowToBottom_(sheet, row);
    ss.toast(`«${company || inn}» -> Перезвон после ${callbackTimeStr}. Строка убрана вниз таблицы!`, "⏰ Перезвон сегодня", 4);
    return;
  }

  // === ВЕТКА 3Б: СДВИГ СПРОСА / КУПИЛ У ДРУГИХ (Самоочистка зависшей потребности H -> «Нет» + смена шага тактики O!) ===
  if (decision.includes("сдвиг") || decision.includes("купил") || decision.includes("конкурент") || decision.includes("тендер") || decision.includes("сезон")) {
    const shiftDateStr = getNextWorkingDateStr_(now, 30, timeZone);
    const finalDeadline = (formattedDeadline && formattedDeadline > getNextWorkingDateStr_(now, 2, timeZone))
      ? formattedDeadline
      : shiftDateStr;

    const tag = `[🔄 Сдвиг / Купил у других${needSnapshot} (след. контакт: ${finalDeadline}) - ${nowStr}]`;
    const updatedHistory = prependHistoryLog_(history, tag);

    if (sheetName === "3. Лист ожидания") {
      sheet.getRange(row, 8).setValue("Нет");
      sheet.getRange(row, 9, 1, 2).clearContent();
      sheet.getRange(row, 14).setValue(updatedHistory);
      sheet.getRange(row, 12).setValue(finalDeadline);
      sheet.getRange(row, 13).clearContent();
      refreshSingleRowTactic_(sheet, row);
      moveRowToBottom_(sheet, row);
      ss.toast(`Запрос перенесен в Историю, H сброшена в «Нет», тактика в O переключена на следующий шаг! Дедлайн: ${finalDeadline}.`, "🔄 Сдвиг спроса", 5);
    } else {
      const targetSheet = getSheetByNameSafe(ss, "3. Лист ожидания");
      if (!targetSheet) return ss.toast("Ошибка: лист Ожидания не найден", "Ошибка");
      const mergeRes = insertOrMergeRowByInn_(targetSheet, [
        inn, company, city, lpr, phone, email,
        sphere, "Нет", "", "",
        notes, finalDeadline, "", updatedHistory
      ], companyNote);
      refreshSingleRowTactic_(targetSheet, mergeRes.row);
      sheet.deleteRow(row);
      highlightDuplicateInnsOnSheet_(sheet);
      ss.toast(`«${company || inn}» переведён в Лист ожидания (дедлайн: ${finalDeadline})!`, "🔄 В Ожидание", 5);
    }
    return;
  }

  // === ВЕТКА 3В: ЛИСТ ОЖИДАНИЯ -> Перенос из Обзвона/Китов в «3. Лист ожидания» ===
  if (decision.includes("ожидан") || decision.includes("подума") || decision.includes("спящ")) {
    const defaultWaitDeadline = getNextWorkingDateStr_(now, 14, timeZone);
    const todayStr = Utilities.formatDate(now, timeZone, "yyyy-MM-dd");
    const targetDeadline = (formattedDeadline && formattedDeadline >= todayStr) ? formattedDeadline : defaultWaitDeadline;

    const tag = `[⏳ В Ожидание${needSnapshot} (из «${sheetName}») - ${nowStr}]`;
    const updatedHistory = prependHistoryLog_(history, tag);
    const activeNeedForWait = (need && need.toLowerCase() !== 'нет' && need !== '—') ? need : "Нет";

    if (sheetName === "3. Лист ожидания") {
      sheet.getRange(row, 8).setValue(activeNeedForWait);
      sheet.getRange(row, 14).setValue(updatedHistory);
      sheet.getRange(row, 12).setValue(targetDeadline);
      sheet.getRange(row, 13).clearContent();
      refreshSingleRowTactic_(sheet, row);
      moveRowToBottom_(sheet, row);
      ss.toast(`Запрос зафиксирован в Истории, тактика обновлена (дедлайн: ${targetDeadline})`, "⏳ Ожидание", 3);
      return;
    }
    const targetSheet = getSheetByNameSafe(ss, "3. Лист ожидания");
    if (!targetSheet) return ss.toast("Ошибка: лист Ожидания не найден", "Ошибка");

    const mergeRes = insertOrMergeRowByInn_(targetSheet, [
      inn, company, city, lpr, phone, email,
      sphere, activeNeedForWait, pipeType, application,
      notes, targetDeadline, "", updatedHistory
    ], companyNote);
    refreshSingleRowTactic_(targetSheet, mergeRes.row);
    sheet.deleteRow(row);
    highlightDuplicateInnsOnSheet_(sheet);
    const msgAction = mergeRes.merged ? "объединён с существующей карточкой" : "перенесён";
    ss.toast(`«${company || inn}» ${msgAction} в Лист ожидания (вместе с Досье)! Дедлайн: ${targetDeadline}`, "⏳ Ожидание", 5);
    return;
  }

  // === ВЕТКА 4: НЕДОЗВОН (Каскадный шаг ТЗ v6.1: 1 -> 3 -> 7 -> 14 дней и в Архив на 4/4) ===
  if (decision.includes("недозвон") || decision.includes("не дозвон") || decision.includes("сброс") || decision.includes("не берет")) {
    const match = history.match(/\[(?:📞\s*)?Недозвон:\s*(\d+)\/4(?:\s*-\s*[^\]]+)?\]/i);
    let attempts = match ? parseInt(match[1], 10) : 0;
    attempts += 1;

    const callTag = `[📞 Недозвон: ${attempts}/4 - ${nowStr}]`;
    let updatedHistory = match
      ? history.replace(/\[(?:📞\s*)?Недозвон:\s*\d+\/4(?:\s*-\s*[^\]]+)?\]/i, callTag)
      : prependHistoryLog_(history, callTag);

    if (attempts < 4) {
      const stepDaysMap = { 1: 1, 2: 3, 3: 7 };
      const stepDays = stepDaysMap[attempts] || 1;
      const nextCallStr = getNextWorkingDateStr_(now, stepDays, timeZone);

      sheet.getRange(row, 14).setValue(updatedHistory);
      sheet.getRange(row, 12).setValue(nextCallStr);
      sheet.getRange(row, 13).clearContent();
      if (sheetName === "3. Лист ожидания") refreshSingleRowTactic_(sheet, row);

      moveRowToBottom_(sheet, row);
      ss.toast(`Недозвон (${attempts}/4, шаг +${stepDays} дн.). Дедлайн: ${nextCallStr}. Строка опущена вниз!`, "📞 Недозвон", 4);
    } else {
      const targetSheet = getSheetByNameSafe(ss, "4.Архив лидов");
      if (!targetSheet) return ss.toast("Ошибка: лист Архива не найден", "Ошибка");

      const archiveTag = `[📦 В Архив: 4 недозвона${needSnapshot} (из «${sheetName}») - ${nowStr}]`;
      const finalArchiveHistory = prependHistoryLog_(updatedHistory, archiveTag);

      insertOrMergeRowByInn_(targetSheet, [
        inn, company, city, lpr, phone, email,
        sphere, "Нет", "", "",
        notes, "", "Недозвон 4/4", finalArchiveHistory
      ], companyNote);
      sheet.deleteRow(row);
      highlightDuplicateInnsOnSheet_(sheet);
      ss.toast(`«${company || inn}» перенесён в Архив (4/4 недозвона). Все данные и Досье сохранены!`, "📦 Архив", 4);
    }
    return;
  }

  // === ВЕТКА 5: ФУНДАМЕНТАЛЬНЫЕ ОТКАЗЫ / НЕ ЦЕЛЕВОЙ / ТОКСИЧНЫЙ ДЕМПИНГ -> в «4.Архив лидов» ===
  const isToxic     = decision.includes("токсич") || decision.includes("демпинг") || decision.includes("террор");
  const isNotTarget = decision.includes("не целев") || decision.includes("нецелев");
  const isRefusal   = decision.includes("отказ") || decision.includes("архив") || decision.includes("закрыл") || decision.includes("удалит") || decision.includes("слил");

  if (isToxic || isNotTarget || isRefusal) {
    const targetSheet = getSheetByNameSafe(ss, "4.Архив лидов");
    if (!targetSheet) return ss.toast("Ошибка: лист Архива не найден", "Ошибка");

    let reasonText = "Отказ (Закрылись)";
    let icon = "⛔";
    if (isToxic) {
      reasonText = "Токсичный демпинг (Прайс-террорист)";
      icon = "☢️";
    } else if (isNotTarget) {
      reasonText = "Не целевой";
      icon = "🚫";
    }

    const archiveTag = `[${icon} В Архив: ${reasonText}${needSnapshot} (из «${sheetName}») - ${nowStr}]`;
    const finalArchiveHistory = prependHistoryLog_(history, archiveTag);

    insertOrMergeRowByInn_(targetSheet, [
      inn, company, city, lpr, phone, email,
      sphere, "Нет", "", "",
      notes, "", reasonText, finalArchiveHistory
    ], companyNote);
    sheet.deleteRow(row);
    highlightDuplicateInnsOnSheet_(sheet);
    ss.toast(`«${company || inn}» отправлен в Архив (${reasonText}). Все 14 колонок и Досье сохранены!`, `${icon} ${reasonText}`, 4);
    return;
  }
}

// =================================================================
// ДВУХКОНТУРНОЕ ДОСЬЕ КЛИЕНТА (ОТСЕБЯТИНА В K -> ЗАМЕТКА B + ВЫЖИМКА)
// =================================================================

/**
 * Переносит сырую «отсебятину» из колонки K (Примечание) во всплывающую Заметку-Досье
 * на ячейке Компании (колонка B), а в самой ячейке K оставляет краткую боевую выжимку!
 */
function processRawNoteIntoDossier_(sheet, row) {
  const noteCell = sheet.getRange(row, 11); // K: Примечание
  const compCell = sheet.getRange(row, 2);  // B: Компания (здесь хранится всплывающее Досье)
  const rawText  = String(noteCell.getValue() || '').trim();
  if (!rawText) return;

  // Если это уже аккуратная короткая выжимка с маркером 📌 — не дублируем повторно
  const existingDossier = String(compCell.getNote() || '').trim();
  const timeZone = sheet.getParent().getSpreadsheetTimeZone() || "Europe/Moscow";
  const dateStr  = Utilities.formatDate(new Date(), timeZone, "dd.MM.yyyy HH:mm");

  // Записываем полный текст отсебятины в Заметку-Досье на ячейке Компании (если его там еще нет)
  const cleanRawForCheck = rawText.replace(/^📌\s*/, '').trim();
  if (cleanRawForCheck && !existingDossier.includes(cleanRawForCheck)) {
    const header = existingDossier ? existingDossier + '\n────────────────────\n' : '🗂️ ДОСЬЕ КЛИЕНТА («КАПИТАЛ-ПЛЮС»):\n\n';
    const updatedDossier = `${header}✍️ [${dateStr}]:\n${rawText}`;
    compCell.setNote(updatedDossier);
  }

  // Строим краткую выжимку (самое важное) для самой ячейки K, если текст длинный (> 55 символов или многострочный)
  const summary = extractSmartSummaryFromNote_(rawText);
  if (summary && summary !== rawText && (rawText.length > 55 || rawText.includes('\n'))) {
    noteCell.setValue(summary);
  }
}

/**
 * Извлекает самую суть (ключевые теги + краткий тезис) из сырой «отсебятины» менеджера
 */
function extractSmartSummaryFromNote_(rawText) {
  const lower = rawText.toLowerCase();
  const tags = [];

  if (/тендер|торг|аукцион|закупк|44-фз|223-фз|аст|ртс/i.test(lower)) tags.push('🏛️ Тендеры');
  if (/аккредитац|реестр|сб\b|служб.*безопасн/i.test(lower))          tags.push('📑 Аккредитация');
  if (/весы|взвеш|микрометр|толщиномер|допуск|хитро/i.test(lower))   tags.push('⚖️ Контроль весов/стенки');
  if (/отсрочк|постоплат|аванс|финансир|нет денег|бюджет/i.test(lower)) tags.push('💳 Финансы/Отсрочка');
  if (/прокол|гнб|футляр|гильз/i.test(lower))                        tags.push('🛠️ Прокол/Футляр');
  if (/обсад|скважин|бур/i.test(lower))                              tags.push('💧 Обсадка/Бурение');
  if (/сваи|свай|опор|мост/i.test(lower))                            tags.push('🏗️ Сваи/Опоры');

  // Ищем упомянутые диаметры в тексте
  const diams = Array.from(new Set((rawText.match(/\b\d{2,4}\b/g) || []).map(Number)))
    .filter(d => STANDARD_PIPE_DIAMS_.includes(d));
  if (diams.length > 0) {
    tags.push(`Ø ${diams.slice(0, 5).join(',')}`);
  }

  // Берем первую осмысленную фразу (до 45 символов)
  const firstLine = rawText.replace(/\n+/g, ' ').replace(/^📌\s*/, '').trim();
  const shortPhrase = firstLine.length > 48 ? firstLine.slice(0, 45).trim() + '…' : firstLine;

  if (tags.length > 0) {
    return `📌 ${tags.join(' • ')} | ${shortPhrase}`;
  }
  return `📌 ${shortPhrase}`;
}

/**
 * Пакетная упаковка всех длинных примечаний (K) в Заметки-Досье (B) на активном листе (из меню «📱 Управление»)
 */
function packAllNotesIntoDossierOnActiveSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName().trim();
  const crmSheets = ["1. Обзвон (База)", "2. Золотой список (Киты)", "3. Лист ожидания", "4.Архив лидов"];

  if (!crmSheets.includes(sheetName)) {
    ss.toast('Перейдите на клиентский лист (Обзвон, Ожидание, Киты или Архив).', '⚠️ Внимание', 4);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  let packedCount = 0;
  for (let r = 2; r <= lastRow; r++) {
    const val = String(sheet.getRange(r, 11).getValue() || '').trim();
    if (val) {
      processRawNoteIntoDossier_(sheet, r);
      packedCount++;
    }
  }

  if (sheetName === "3. Лист ожидания") {
    matchWarehouseWithWaitingList();
  }

  ss.toast(`Упаковано примечаний в карточки Досье (наведите мышь на название компании в колонке B): ${packedCount} шт.`, '🗂️ Досье обновлено', 6);
}

// =================================================================
// ЖИВАЯ РОТАЦИЯ ТАКТИКИ ДОЖИМА (КОЛОНКА O В «3. ЛИСТ ОЖИДАНИЯ»)
// =================================================================

function getAvailableStockDiamsMap_(ss) {
  const whSheet = getSheetByNameSafe(ss, 'Склад');
  const stockByDiam = {};
  if (!whSheet || whSheet.getLastRow() < 3) return stockByDiam;

  const whData = whSheet.getRange(3, 1, whSheet.getLastRow() - 2, 12).getValues();
  whData.forEach(r => {
    const status  = String(r[9] || '').trim().toLowerCase();
    const stockTn = typeof safeParseFloat === 'function' ? safeParseFloat(r[10]) : parseFloat(r[10]) || 0;
    const stockM  = typeof safeParseFloat === 'function' ? safeParseFloat(r[11]) : parseFloat(r[11]) || 0;
    const diam    = Math.round(typeof safeParseFloat === 'function' ? safeParseFloat(r[3]) : parseFloat(r[3]) || 0);
    if (status !== 'нет' && (stockTn > 0 || stockM > 0) && diam > 0) {
      stockByDiam[diam] = true;
    }
  });
  return stockByDiam;
}

/**
 * Вычисляет живую тактику дожима (Колонка O / 15), анализируя H (Потребность), N (Историю), K (Выжимку) и B (Полное Досье)
 */
function computeDynamicTacticForRow_(rowVals, stockByDiam, companyDossierNote) {
  const sphere  = String(rowVals[6] || '').trim();  // G: Сфера
  const need    = String(rowVals[7] || '').trim();  // H: Потребность сейчас
  const notes   = String(rowVals[10] || '').trim(); // K: Примечание (выжимка)
  const history = String(rowVals[13] || '').trim(); // N: История (логи)
  const dossier = String(companyDossierNote || '').trim(); // Всплывающее Досье на ячейке B

  const allContext = `${notes} ${dossier} ${history}`;
  const allLower   = allContext.toLowerCase();

  const needLower = need.toLowerCase();
  const hasActiveNeed = need !== '' && needLower !== 'нет' && needLower !== '—' && needLower !== '-';

  // Первая строка в Истории (последнее совершенное действие)
  const firstHistoryLine = (history.split('\n')[0] || '').trim();

  // Флаг сложного клиента, который принимает по весам / меряет стенку (Допуск ГОСТ -12.5%)
  const isScaleClient = /весы|взвеш|микрометр|толщиномер|хитро/i.test(allLower);
  const scaleWarnTag  = isScaleClient ? ' [⚖️ ВЕСЫ: +10% к цене на допуск!]' : '';

  // 1. ВЫСШИЙ ПРИОРИТЕТ: Выставлен счет / активная заявка в H
  if (firstHistoryLine.includes('💰') || firstHistoryLine.toLowerCase().includes('выставлен счет')) {
    const targetPipe = hasActiveNeed ? ` (${need})` : '';
    return `💰 ДОЖИМ СЧЕТА${targetPipe}: узнать проходную цену конкурента / дожать по теоретич. весу${scaleWarnTag}`;
  }

  if (hasActiveNeed) {
    const liveDiams = Array.from(new Set((need.match(/\b\d{2,4}\b/g) || []).map(Number)));
    const inStock = liveDiams.filter(d => stockByDiam[d]);
    if (inStock.length > 0) {
      return `🔥 ГОРЯЩАЯ ЗАЯВКА (${need}): Ø ${inStock.join(', ')} ЕСТЬ НА СКЛАДЕ — дать цену (по теории) и выставить счет!${scaleWarnTag}`;
    }
    return `🔥 ГОРЯЩАЯ ЗАЯВКА (${need}): просчитать транзит + доставку (CPT) за 15 мин и выставить счет${scaleWarnTag}`;
  }

  // 2. ПОСЛЕДНЕЕ ДЕЙСТВИЕ — НЕДОЗВОН: меняем канал связи
  const missMatch = firstHistoryLine.match(/Недозвон:\s*(\d+)\/4/i);
  if (missMatch) {
    const att = parseInt(missMatch[1], 10);
    if (att === 1) return `📲 После недозвона (1/4): набрать повторно в другое время дня`;
    if (att === 2) return `💬 После недозвона (2/4): скинуть короткий пинг в WhatsApp / Telegram ЛПР`;
    return `⚠️ Недозвон (3/4 — последний шанс): пробить через общий номер / почту перед Архивом`;
  }

  // Считаем число прошедших циклов прогрева (сдвигов, пингов), чтобы тактика постоянно вращалась!
  const shiftCount = (history.match(/🔄|Сдвиг|⏳|💰/gi) || []).length;
  const kpSent     = /📧|Прайс|КП/i.test(history) || /кп|прайс/i.test(allLower);

  // 3. СНАЙПЕР СКЛАДА: в H сейчас «Нет», но в Истории (N) или Досье (K/B) есть диаметры, которые лежат на Складе!
  const pastDiams = Array.from(new Set((allContext.match(/\b\d{2,4}\b/g) || []).map(Number)))
    .filter(d => STANDARD_PIPE_DIAMS_.includes(d));
  const stockHits = pastDiams.filter(d => stockByDiam[d]);

  if (stockHits.length > 0) {
    const stockStr = `Ø ${stockHits.slice(0, 4).join(', ')}`;
    const mode = shiftCount % 3;
    if (mode === 0) {
      return `🎯 СНАЙПЕР СКЛАДА: в наличии их калибр (${stockStr}) — предложить живой объем со склада Воронеж${scaleWarnTag}`;
    } else if (mode === 1) {
      return `📉 АРБИТРАЖ СО СКЛАДА (${stockStr}): предложить партию под экономию сметы (восст./лежалая) с доставкой${scaleWarnTag}`;
    } else {
      return `🚛 БРОНЬ СО СКЛАДА (${stockStr}): узнать старт работ и предложить фиксацию цены до подорожания${scaleWarnTag}`;
    }
  }

  // 4. СПЕЦИАЛЬНЫЕ СЦЕНАРИИ ИЗ ДОСЬЕ КЛИЕНТА (На основе вашей «отсебятины» в K / B!)
  if (/аккредитац|реестр|служб.*безопасн/i.test(allLower) && shiftCount % 2 === 0) {
    return `📑 ДОСЬЕ (Аккредитация): проверить статус проверки наших документов в СБ и зайти в реестр поставщиков`;
  }
  if (/тендер|торг|аукцион|44-фз|223-фз|аст/i.test(allLower) && shiftCount % 2 === 0) {
    return `🏛️ ДОСЬЕ (Тендеры): запросить спецификацию ближайшего тендера — посчитать трубу со стенкой в допуске ГОСТ под проходную цену`;
  }
  if (/не финансир|нет денег|ждут аванс|задерж/i.test(allLower) && shiftCount % 2 === 0) {
    return `💳 ДОСЬЕ (Финансирование): проверить, пошел ли транш от заказчика / предложить поставку мелкими партиями`;
  }

  // 5. РОТАЦИЯ СТУПЕНЕЙ ПРОГРЕВА (Никогда не зависает на одном шаге!)
  if (!kpSent) {
    return `📧 ШАГ 1 (Якорь): выслать Прайс + Карточку компании (закрепить контакт снабженца)`;
  }

  if (pastDiams.length === 0 && !sphere) {
    return `🎯 ШАГ 2 (Разведка): проверить получение КП, выбить рабочие Ø и профиль объектов`;
  }

  const rotationStep = shiftCount % 4;
  if (rotationStep === 0) {
    return `📞 ШАГ 2 (Объекты): КП у них — узнать график ближайших объектов и забрать спецификацию на просчет`;
  } else if (rotationStep === 1) {
    return `📉 ШАГ 3 (Ценовой повод): зайти через движение цен / предложить замену новой трубы на лежалую или восст.`;
  } else if (rotationStep === 2) {
    return `🚛 ШАГ 4 (Логистика CPT): предложить расчет транзитной партии с доставкой попутной фурой до их объекта`;
  } else {
    return `🤝 ШАГ 5 (Расширение ЛПР): выйти на главного инженера / ПТО, проверить потребность под сваи/футляры/опоры`;
  }
}

function refreshSingleRowTactic_(sheet, row) {
  if (!sheet || sheet.getName().trim() !== "3. Лист ожидания" || row < 2) return;
  const ss = sheet.getParent();
  const stockByDiam = getAvailableStockDiamsMap_(ss);
  const rowVals = sheet.getRange(row, 1, 1, 14).getValues()[0];
  if (!rowVals[0] && !rowVals[1]) return;
  const dossierNote = sheet.getRange(row, 2).getNote();
  const tactic = computeDynamicTacticForRow_(rowVals, stockByDiam, dossierNote);
  sheet.getRange(row, 15).setValue(tactic);
}

// =================================================================
// СКОРОСТНОЙ ПАРСЕР ПОТРЕБНОСТИ ТРУБЫ И СЛЕПОК ЗАЯВКИ В ИСТОРИЮ
// =================================================================

function formatPipeNeed_(rawInput) {
  if (!rawInput) return "";
  let str = rawInput.toString().trim();
  if (!str) return "";
  if (str.toLowerCase() === 'нет' || str === '—' || str === '-') return "Нет";

  const diamSet = new Set(STANDARD_PIPE_DIAMS_);

  str = str.replace(/(\d{2,4})\s*[\*\/xXхХ\-]\s*(\d{1,2}(?:[.,]\d+)?)/g, '$1х$2');

  str = str.replace(/\b(\d{3,6}(?:[.,]\d+)?)\b/g, (match) => {
    const numVal = parseFloat(match.replace(',', '.'));
    if (diamSet.has(numVal)) return match;

    for (let i = 0; i < STANDARD_PIPE_DIAMS_.length; i++) {
      const d = STANDARD_PIPE_DIAMS_[i];
      const dStr = String(d);
      if (match.startsWith(dStr) && match.length > dStr.length) {
        const wallPart = match.slice(dStr.length).replace(',', '.');
        const wallNum = parseFloat(wallPart);
        if (!wallPart.startsWith('0') && wallNum >= 1.5 && wallNum <= 45) {
          return `${d}х${wallPart}`;
        }
      }
    }
    return match;
  });

  str = str.replace(/\b(\d{2,4})\s+(\d{1,2}(?:[.,]\d+)?)\b(?!\s*(?:т|тн|тонн|м|метр|шт|х))/gi, (full, dStr, wStr) => {
    const d = parseInt(dStr, 10);
    const w = parseFloat(wStr.replace(',', '.'));
    if (diamSet.has(d) && !diamSet.has(w) && w >= 1.5 && w <= 45) {
      return `${d}х${wStr.replace(',', '.')}`;
    }
    return full;
  });

  str = str.replace(/(\dх\d+(?:[.,]\d+)?)\s+(?=(?:Ø\s*)?\d{2,4}(?:х|\b)(?!\s*(?:т|тн|тонн|м|метр|шт)))/gi, '$1, ');
  str = str.replace(/\b(\d{2,4})\s+(?=\d{2,4}\b)/g, (full, d1) => {
    return diamSet.has(parseInt(d1, 10)) ? `Ø ${d1}, ` : full;
  });

  str = str.replace(/(^|[,;]\s*)(\d{2,4})(?=\s*(?:$|[,;]))/g, (full, prefix, dStr) => {
    const d = parseInt(dStr, 10);
    return diamSet.has(d) ? `${prefix}Ø ${d}` : full;
  });

  return str.replace(/\s+/g, ' ').trim();
}

function buildNeedSnapshot_(need, pipeType, application) {
  const cleanNeed = (need && need.toLowerCase() !== 'нет' && need !== '—') ? need : '';
  const details = [pipeType, application].filter(Boolean).join(' • ');
  if (cleanNeed && details) {
    return ` | Запрос: ${cleanNeed} (${details})`;
  } else if (cleanNeed) {
    return ` | Запрос: ${cleanNeed}`;
  } else if (details) {
    return ` | Профиль: ${details}`;
  }
  return '';
}

// =================================================================
// ДВУХУРОВНЕВЫЙ БИРЖЕВОЙ РАДАР + ОБНОВЛЕНИЕ КОЛОНКИ O (ТАКТИКА)
// =================================================================

function matchWarehouseWithWaitingList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stockByDiam = getAvailableStockDiamsMap_(ss);

  if (Object.keys(stockByDiam).length === 0) {
    ss.toast('На складе нет позиций со статусом «В наличии».', '⚠️ Нет остатков', 4);
    return;
  }

  const targetSheets = ["3. Лист ожидания", "2. Золотой список (Киты)"];
  const hotMatches = [];
  const historyMatches = [];

  targetSheets.forEach(sName => {
    const sh = getSheetByNameSafe(ss, sName);
    if (!sh || sh.getLastRow() < 2) return;

    const lRow = sh.getLastRow();
    const data = sh.getRange(2, 1, lRow - 1, 14).getValues(); // A..N
    const compNotes = sh.getRange(2, 2, lRow - 1, 1).getNotes(); // Заметки-Досье на колонке B
    const needBgColors = [];
    const tacticsCol = [];

    data.forEach((r, idx) => {
      const comp    = String(r[1] || '').trim();
      const lpr     = String(r[3] || '').trim();
      const ph      = String(r[4] || '').trim().split('\n')[0];
      const needStr = String(r[7] || '').trim();  // H: Живая потребность сейчас
      const notes   = String(r[10] || '').trim(); // K: Примечание
      const histStr = String(r[13] || '').trim(); // N: История прошлых запросов
      const dossier = String((compNotes[idx] && compNotes[idx][0]) || '').trim();

      if (sName === "3. Лист ожидания") {
        tacticsCol.push([computeDynamicTacticForRow_(r, stockByDiam, dossier)]);
      }

      const isNeedActive = needStr !== '' && needStr.toLowerCase() !== 'нет' && needStr !== '—';

      if (isNeedActive) {
        const liveDiams = Array.from(new Set((needStr.match(/\b\d{2,4}\b/g) || []).map(Number))).filter(d => stockByDiam[d]);
        if (liveDiams.length > 0) {
          needBgColors.push(['#dcfce7']);
          hotMatches.push(
            `🔥 [${sName.replace(/^\d+\.\s*/, '')}] ${comp || 'Клиент'} (${lpr || ph}): сейчас в заявке Ø ${liveDiams.join(', ')}`
          );
          return;
        }
      }

      const combinedPast = `${histStr} ${notes} ${dossier}`;
      const pastDiams = Array.from(new Set((combinedPast.match(/\b\d{2,4}\b/g) || []).map(Number)))
        .filter(d => STANDARD_PIPE_DIAMS_.includes(d) && stockByDiam[d]);

      if (pastDiams.length > 0) {
        needBgColors.push(['#fef9c3']);
        historyMatches.push(
          `💡 [${sName.replace(/^\d+\.\s*/, '')}] ${comp || 'Клиент'} (${lpr || ph}): в истории/досье Ø ${pastDiams.join(', ')}`
        );
        return;
      }

      needBgColors.push([null]);
    });

    sh.getRange(2, 8, lRow - 1, 1).setBackgrounds(needBgColors);
    if (sName === "3. Лист ожидания" && tacticsCol.length > 0) {
      sh.getRange(2, 15, lRow - 1, 1).setValues(tacticsCol);
    }
  });

  const allMatches = hotMatches.concat(historyMatches);
  if (allMatches.length > 0) {
    const preview = allMatches.slice(0, 14).join('\n');
    const more = allMatches.length > 14 ? `\n...и ещё ${allMatches.length - 14} совпадений!` : '';
    SpreadsheetApp.getUi().alert(
      `🎯 Биржевой радар: Горячих (${hotMatches.length}) | Из истории и Досье (${historyMatches.length})`,
      `Колонка O (Чем дожимать) обновлена! Зелёным в H подсвечен активный спрос, жёлтым — совпадения из Истории и Досье:\n\n${preview}${more}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    ss.toast('Колонка O (Чем дожимать) обновлена! Прямых совпадений со складом пока не найдено.', '🔍 Радар обновлен', 5);
  }
}

// =================================================================
// СЛУЖЕБНЫЕ ФУНКЦИИ: ТЕЛЕФОНЫ, ДУБЛИ ИНН И УМНОЕ СЛИЯНИЕ С ДОСЬЕ
// =================================================================

function formatRussianPhones_(rawPhone) {
  if (!rawPhone) return "";

  const rawChunks = rawPhone.toString()
    .replace(/'\+/g, '+')
    .replace(/(\d)\s*[,;\/]\s*(?=(?:\+7|8|7|9|4|3)\d)/g, '$1\n')
    .split(/\n+/);

  const resultLines = [];

  rawChunks.forEach(chunk => {
    let line = chunk.trim();
    if (!line) return;

    const phoneRegex = /(?:(?:\+7|8|7)[\s\-\.\(\)]*)?(?:\(\s*\d{3,5}\s*\)|\d{3,5})[\s\-\.\(\)]*\d{2,3}[\s\-\.]*\d{2}[\s\-\.]*\d{2}|(?:\+7|8|7)?\d{10}|\b2\d{2}[\s\-]?\d{2}[\s\-]?\d{2}\b/g;

    const formattedChunk = line.replace(phoneRegex, (matchStr) => {
      const digits = matchStr.replace(/\D/g, '');
      let d10 = "";

      if (digits.length === 11 && (digits.startsWith('7') || digits.startsWith('8'))) {
        d10 = digits.substring(1);
      } else if (digits.length === 10 && /^[3489]/.test(digits)) {
        d10 = digits;
      } else if (digits.length === 7 && digits.startsWith('2')) {
        d10 = '473' + digits;
      }

      if (d10.length === 10) {
        return `+7 (${d10.slice(0, 3)}) ${d10.slice(3, 6)}-${d10.slice(6, 8)}-${d10.slice(8, 10)}`;
      }
      return matchStr;
    });

    const splitMultiple = formattedChunk
      .replace(/(\d{2}-\d{2})\s*[,;]?\s*(\+7\s*\()/g, '$1\n$2')
      .trim();

    if (splitMultiple && !resultLines.includes(splitMultiple)) {
      resultLines.push(splitMultiple);
    }
  });

  return resultLines.join('\n');
}

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

  if (row >= lastRealRow) return;

  if (lastRealRow >= maxRows) {
    sheet.insertRowAfter(maxRows);
  }

  sheet.moveRows(sheet.getRange(row, 1, 1, sheet.getLastColumn()), lastRealRow + 1);
}

function highlightDuplicateInnsOnSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

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
      bgColors.push(['#fee2e2']);
      fontColors.push(['#991b1b']);
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

function mergeUniqueLines_(existingStr, newStr, separator) {
  const sep = separator || '\n';
  const rawList = [existingStr, newStr]
    .map(s => String(s || '').trim())
    .filter(Boolean)
    .join('\n')
    .split(/\n+/);

  const unique = [];
  rawList.forEach(item => {
    const clean = item.trim();
    if (!clean) return;
    const exists = unique.some(u => u.toLowerCase() === clean.toLowerCase());
    if (!exists) unique.push(clean);
  });
  return unique.join(sep);
}

/**
 * Вставка строки ИЛИ автоматическое слияние с существующей карточкой по ИНН (вместе со всплывающим Досье B!)
 */
function insertOrMergeRowByInn_(targetSheet, rowValues, incomingDossierNote) {
  const incomingInn = String(rowValues[0] || '').replace(/\D/g, '');
  const maxRows = targetSheet.getMaxRows();
  const checkRange = targetSheet.getRange(1, 1, maxRows, 14).getValues();
  let lastRealRow = 1;
  let existingRowIdx = -1;

  for (let r = 1; r < checkRange.length; r++) {
    const rowInn = String(checkRange[r][0] || '').replace(/\D/g, '');
    const rowComp = String(checkRange[r][1] || '').trim();
    if (rowInn !== "" || rowComp !== "") {
      lastRealRow = r + 1;
      if (incomingInn.length >= 10 && rowInn === incomingInn && existingRowIdx === -1) {
        existingRowIdx = r + 1;
      }
    }
  }

  if (existingRowIdx !== -1) {
    const oldRow = checkRange[existingRowIdx - 1];
    const mergedRow = [
      incomingInn || oldRow[0],
      oldRow[1] || rowValues[1],                                     // B: Компания
      oldRow[2] || rowValues[2],                                     // C: Город
      mergeUniqueLines_(oldRow[3], rowValues[3], '\n'),              // D: ЛПР
      formatRussianPhones_(mergeUniqueLines_(oldRow[4], rowValues[4], '\n')), // E: Телефон
      mergeUniqueLines_(oldRow[5], rowValues[5], '\n'),              // F: E-mail
      oldRow[6] || rowValues[6],                                     // G: Сфера
      rowValues[7] || oldRow[7],                                     // H: Потребность
      rowValues[8] || oldRow[8],                                     // I: Тип трубы
      rowValues[9] || oldRow[9],                                     // J: Под что берут
      mergeUniqueLines_(oldRow[10], rowValues[10], ' | '),           // K: Примечание (выжимка)
      rowValues[11] || oldRow[11],                                   // L: Дедлайн
      "",                                                            // M: Решение
      prependHistoryLog_(oldRow[13], rowValues[13])                  // N: История
    ];

    targetSheet.getRange(existingRowIdx, 1).setNumberFormat('@');
    targetSheet.getRange(existingRowIdx, 5).setNumberFormat('@');
    targetSheet.getRange(existingRowIdx, 1, 1, 14).setValues([mergedRow]);

    if (incomingDossierNote) {
      const oldDossier = String(targetSheet.getRange(existingRowIdx, 2).getNote() || '').trim();
      if (!oldDossier) {
        targetSheet.getRange(existingRowIdx, 2).setNote(incomingDossierNote);
      } else if (!oldDossier.includes(incomingDossierNote)) {
        targetSheet.getRange(existingRowIdx, 2).setNote(oldDossier + '\n────────────────────\n' + incomingDossierNote);
      }
    }
    return { merged: true, row: existingRowIdx };
  }

  const newRowIdx = insertAndFormatRow_(targetSheet, rowValues);
  if (incomingDossierNote) {
    targetSheet.getRange(newRowIdx, 2).setNote(incomingDossierNote);
  }
  return { merged: false, row: newRowIdx };
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
    const templateRange = targetSheet.getRange(2, 1, 1, numCols);
    templateRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    templateRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  }
  targetSheet.getRange(targetRow, 1).setNumberFormat('@');
  targetSheet.getRange(targetRow, 5).setNumberFormat('@');
  targetRange.setValues([rowValues]);
  return targetRow;
}