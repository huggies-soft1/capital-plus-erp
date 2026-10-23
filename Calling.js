/**
 * ЕДИНЫЙ ЦЕНТР УПРАВЛЕНИЯ ПРОДАЖАМИ И СКЛАДОМ (CTRM «Капитал-Плюс» v6.5 — HK Trader Desk)
 * Файл: Calling.gs — Маршрутизация лидов, скоростной парсер труб (53010 / 530*10 / 630 8),
 * 100% АВТОМАТИЧЕСКОЕ Досье (K -> Заметка B без обрезки смысла), сброс подсветок на пустых ячейках,
 * авто-архивация H+I+J в Историю (N), живой мозг колонки O (Склад с тоннами/ценами + Транзит/Логистика
 * + Гонконгский детектор блефа «В этом году не берут»), календарь + быстрое время (14 / 1430 / 14 30) в L.
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

  if (row < 2) return;

  if (cleanupRowIfCompletelyEmpty_(sheet, row, sheetName)) {
    if (col === 1) highlightDuplicateInnsOnSheet_(sheet, row);
    return;
  }

  // =================================================================
  // 0. КОНТРОЛЬ И СБРОС ПОДСВЕТКИ ДУБЛЕЙ ИНН (Колонка A / 1)
  // =================================================================
  if (col === 1) {
    const cell = sheet.getRange(row, 1);
    const rawInn = e.value !== undefined ? String(e.value) : String(cell.getValue());
    const cleanInn = rawInn.replace(/\D/g, '');
    if (!cleanInn) {
      cell.setBackground(null).setFontColor('#0f172a').setFontWeight('normal');
      highlightDuplicateInnsOnSheet_(sheet, row);
      return;
    }
    if (cleanInn !== rawInn.trim()) {
      cell.setNumberFormat('@').setValue(cleanInn);
    }
    highlightDuplicateInnsOnSheet_(sheet, row);
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
  // 2.5 ЖИВАЯ ПОТРЕБНОСТЬ ТРУБЫ (Колонка H / 8)
  // =================================================================
  if (col === 8) {
    const cell = sheet.getRange(row, col);
    const rawNeed = e.value !== undefined ? String(e.value).trim() : String(cell.getValue() || '').trim();
    const oldNeed = e.oldValue !== undefined ? String(e.oldValue).trim() : '';
    const formattedNeed = rawNeed ? formatPipeNeed_(rawNeed) : '';

    if (formattedNeed && formattedNeed !== rawNeed) {
      cell.setNumberFormat('@').setValue(formattedNeed);
    }

    const isNowEmptyOrNo = !formattedNeed || formattedNeed.toLowerCase() === 'нет' || formattedNeed === '—';
    const wasActiveNeed  = oldNeed !== '' && oldNeed.toLowerCase() !== 'нет' && oldNeed !== '—' && oldNeed !== '-';

    if (isNowEmptyOrNo && wasActiveNeed) {
      const ss = sheet.getParent();
      const timeZone = ss.getSpreadsheetTimeZone() || "Europe/Moscow";
      const nowStr = Utilities.formatDate(new Date(), timeZone, "dd.MM.yyyy HH:mm");
      const curType = String(sheet.getRange(row, 9).getValue() || '').trim();
      const curApp  = String(sheet.getRange(row, 10).getValue() || '').trim();
      const curHist = String(sheet.getRange(row, 14).getValue() || '').trim();
      const snap    = buildNeedSnapshot_(formatPipeNeed_(oldNeed), curType, curApp);

      if (snap) {
        const logLine = `[📋 Снята/закрыта потребность${snap} - ${nowStr}]`;
        sheet.getRange(row, 14).setValue(prependHistoryLog_(curHist, logLine));
      }
      sheet.getRange(row, 9, 1, 2).clearContent();
      sheet.setActiveSelection(cell);
      ss.toast(`Прежняя заявка (${oldNeed}) сохранена в Историю (N), колонки I и J очищены!`, "📋 Сохранено в Историю", 4);
    }

    updateSingleRowNeedHighlight_(sheet, row, formattedNeed);

    if (sheetName === "3. Лист ожидания") {
      refreshSingleRowTactic_(sheet, row);
    }
    return;
  }

  // =================================================================
  // 2.8 МГНОВЕННОЕ АВТО-ДОСЬЕ КЛИЕНТА ПРИ ВВОДЕ В ПРИМЕЧАНИЕ (Колонка K / 11)
  // =================================================================
  if (col === 11) {
    processRawNoteIntoDossier_(sheet, row, e.oldValue);
    if (sheetName === "3. Лист ожидания") {
      refreshSingleRowTactic_(sheet, row);
    }
    return;
  }

  // =================================================================
  // 2.9 ВЫПАДАЮЩИЙ КАЛЕНДАРЬ ДАТЫ + БЫСТРЫЙ ВВОД ВРЕМЕНИ (Колонка L / 12)
  // =================================================================
  if (col === 12) {
    const cell = sheet.getRange(row, 12);
    const cellVal = cell.getValue();
    const timeZone = sheet.getParent().getSpreadsheetTimeZone() || "Europe/Moscow";

    const dlObj = parseSmartDeadlineInput_(cellVal, e.value, e.oldValue, timeZone);
    ensureDateValidationOnCell_(cell);

    if (dlObj.dateObj) {
      cell.setNumberFormat(dlObj.hasTime ? "dd.MM.yyyy HH:mm" : "dd.MM.yyyy");
      cell.setValue(dlObj.dateObj);
      updateSingleRowDeadlineHighlight_(sheet, row, dlObj);
    } else {
      cell.clearContent();
      cell.setNumberFormat("dd.MM.yyyy");
      cell.setBackground(null).setFontColor('#0f172a').setFontWeight('normal');
    }
    return;
  }

  if (sheetName === "3. Лист ожидания" && (col === 3 || col === 6 || col === 7 || col === 9 || col === 10 || col === 14)) {
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

  processRawNoteIntoDossier_(sheet, row);

  const rowData     = sheet.getRange(row, 1, 1, 14).getValues()[0];
  const companyNote = String(sheet.getRange(row, 2).getNote() || '').trim();

  const inn         = String(rowData[0] || '').trim();
  const company     = typeof cleanCompanyName === 'function' ? cleanCompanyName(rowData[1]) : String(rowData[1] || '').trim();
  const city        = String(rowData[2] || '').trim();
  const lpr         = String(rowData[3] || '').trim();
  const phone       = formatRussianPhones_(String(rowData[4] || '').trim());
  const email       = String(rowData[5] || '').trim();
  const sphere      = String(rowData[6] || '').trim();
  const need        = formatPipeNeed_(String(rowData[7] || '').trim());
  const pipeType    = String(rowData[8] || '').trim();
  const application = String(rowData[9] || '').trim();
  let notes         = String(rowData[10] || '').trim();
  let deadline      = rowData[11];
  let history       = String(rowData[13] || '').trim();

  const needSnapshot = buildNeedSnapshot_(need, pipeType, application);
  const existingDlObj = parseDeadlineObj_(deadline, timeZone);
  const existingIsoDeadline = existingDlObj.isoDate;

  // === ЗАЩИТА ПРАВИЛА №6: ЗАПРЕТ РУЧНОГО ПЕРЕВОДА В КИТЫ ===
  if (decision.includes("кит") || decision.includes("золот")) {
    sheet.getRange(row, 13).clearContent();
    ss.toast(
      "В «Киты» клиент попадает ТОЛЬКО автоматически после оплаты заказа! Выберите «В Ожидание» или «Горящая заявка».",
      "⛔ Правило №6 (Закрытый клуб)",
      6
    );
    return;
  }

  // === ВЕТКА 0.5: 🛡️ ПРОХОД СЕКРЕТАРЯ (ВЫБИТЬ ЛПР) ===
  if (decision.includes("секретар") || decision.includes("выбить") || decision.includes("лпр")) {
    const nextTryDate = getNextWorkingDateObj_(now, 2);
    const nextTryStr  = Utilities.formatDate(nextTryDate, timeZone, "dd.MM.yyyy");
    const secTag = `[🛡️ Проход секретаря — выбить ЛПР${needSnapshot} - ${nowStr}]`;
    const updatedHistory = prependHistoryLog_(history, secTag);

    sheet.getRange(row, 14).setValue(updatedHistory);
    setDeadlineCellDate_(sheet, row, nextTryDate, false);
    sheet.getRange(row, 13).clearContent();
    if (sheetName === "3. Лист ожидания") refreshSingleRowTactic_(sheet, row);

    moveRowToBottom_(sheet, row);
    ss.toast(`«${company || inn}» оставлен в базе на пробив секретаря (контроль: ${nextTryStr}). Строка опущена вниз!`, "🛡️ Выбить ЛПР", 4);
    return;
  }

  // === ВЕТКА 1А: 🔥 ГОРЯЩАЯ ЗАЯВКА (В ПРОСЧЕТ) ===
  if (decision.includes("горящ") || decision.includes("просчет") || decision.includes("просчёт") || decision.includes("заявк")) {
    const nextCheckDateObj = getNextWorkingDateObj_(now, 1);
    const nextCheckDateStr = Utilities.formatDate(nextCheckDateObj, timeZone, "dd.MM.yyyy");
    const hotTag = `[🔥 Горящая заявка${needSnapshot} (из «${sheetName}») - ${nowStr}]`;
    const updatedHistory = prependHistoryLog_(history, hotTag);
    const liveNeed = (need && need.toLowerCase() !== 'нет' && need !== '—') ? need : "В просчете";

    if (sheetName === "1. Обзвон (База)") {
      const targetSheet = getSheetByNameSafe(ss, "3. Лист ожидания");
      if (!targetSheet) return ss.toast("Ошибка: лист Ожидания не найден", "Ошибка");

      const mergeRes = insertOrMergeRowByInn_(targetSheet, [
        inn, company, city, lpr, phone, email,
        sphere, liveNeed, pipeType, application,
        notes, nextCheckDateObj, "", updatedHistory
      ], companyNote);
      setDeadlineCellDate_(targetSheet, mergeRes.row, nextCheckDateObj, false);
      updateSingleRowNeedHighlight_(targetSheet, mergeRes.row, liveNeed);
      refreshSingleRowTactic_(targetSheet, mergeRes.row);
      sheet.deleteRow(row);
      highlightDuplicateInnsOnSheet_(sheet);
      ss.toast(`🔥 «${company || inn}» перенесен в Ожидание с ГОРЯЩЕЙ заявкой (${liveNeed})! Дедлайн: ${nextCheckDateStr}.`, "🔥 Горящая заявка", 5);
    } else {
      sheet.getRange(row, 8).setValue(liveNeed);
      sheet.getRange(row, 14).setValue(updatedHistory);
      setDeadlineCellDate_(sheet, row, nextCheckDateObj, false);
      sheet.getRange(row, 13).clearContent();
      updateSingleRowNeedHighlight_(sheet, row, liveNeed);
      if (sheetName === "3. Лист ожидания") refreshSingleRowTactic_(sheet, row);
      moveRowToBottom_(sheet, row);
      ss.toast(`🔥 Заявка (${liveNeed}) взята в просчет! Контроль: ${nextCheckDateStr}.`, "🔥 В просчете", 4);
    }
    return;
  }

  // === ВЕТКА 1Б: ВЫСТАВЛЕН СЧЕТ ===
  if (decision.includes("выставлен счет") || decision.includes("счет") || decision.includes("счёт") || decision.includes("заказ")) {
    if (sheetName === "1. Обзвон (База)") {
      sheet.getRange(row, 13).clearContent();
      ss.toast(
        "В холодном «Обзвоне» выберите «Горящая заявка (В просчет)» или «В Ожидание (Нет потребности)»!",
        "💡 Маршрутизация Обзвона",
        5
      );
      return;
    }

    const tag = `[💰 Выставлен счет${needSnapshot} - ${nowStr}]`;
    const updatedHistory = prependHistoryLog_(history, tag);
    const nextCheckDateObj = getNextWorkingDateObj_(now, 1);
    const nextCheckDateStr = Utilities.formatDate(nextCheckDateObj, timeZone, "dd.MM.yyyy");
    const liveNeed = (need && need.toLowerCase() !== 'нет' && need !== '—') ? need : "По счету";

    sheet.getRange(row, 8).setValue(liveNeed);
    sheet.getRange(row, 14).setValue(updatedHistory);
    setDeadlineCellDate_(sheet, row, nextCheckDateObj, false);
    sheet.getRange(row, 13).clearContent();
    updateSingleRowNeedHighlight_(sheet, row, liveNeed);
    if (sheetName === "3. Лист ожидания") refreshSingleRowTactic_(sheet, row);
    moveRowToBottom_(sheet, row);
    ss.toast(`«${company || inn}» -> Счет зафиксирован в Истории! Тактика переключена на дожим счета (${nextCheckDateStr}).`, "💰 Выставлен счет", 4);
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
    const nextDateObj = getNextWorkingDateObj_(now, toWaiting ? 7 : 0);

    const mergeRes = insertOrMergeRowByInn_(targetSheet, [
      inn, company, city, lpr, phone, email,
      sphere, toWaiting ? "Нет" : "", "", "",
      notes, nextDateObj, "", updatedHistory
    ], companyNote);
    setDeadlineCellDate_(targetSheet, mergeRes.row, nextDateObj, false);
    updateSingleRowNeedHighlight_(targetSheet, mergeRes.row, toWaiting ? "Нет" : "");
    if (toWaiting) refreshSingleRowTactic_(targetSheet, mergeRes.row);
    sheet.deleteRow(row);
    ss.toast(`«${company || inn}» восстановлен в «${targetName}» со всей историей и Досье!`, "♻️ Возврат лида", 4);
    return;
  }

  // === ВЕТКА 3А: ОПЕРАТИВНЫЙ ПЕРЕЗВОН СЕГОДНЯ / ЧЕРЕЗ 2 ЧАСА ===
  if (decision.includes("2 час") || decision.includes("два час") || decision.includes("перезвон") || decision.includes("сегодня")) {
    let callbackDateObj = new Date(now.getTime() + 2 * 60 * 60 * 1000);
    if (existingDlObj.dateObj && existingDlObj.isoDate === Utilities.formatDate(now, timeZone, "yyyy-MM-dd") && existingDlObj.hasTime) {
      callbackDateObj = existingDlObj.dateObj;
    }
    const chosenTimeStr = Utilities.formatDate(callbackDateObj, timeZone, "HH:mm");
    const fullDispStr   = Utilities.formatDate(callbackDateObj, timeZone, "dd.MM.yyyy HH:mm");

    const tag = `[⏰ Перезвонить после ${chosenTimeStr}${needSnapshot} - ${nowStr}]`;
    const updatedHistory = prependHistoryLog_(history, tag);

    sheet.getRange(row, 8).setValue(need);
    sheet.getRange(row, 14).setValue(updatedHistory);
    setDeadlineCellDate_(sheet, row, callbackDateObj, true);
    sheet.getRange(row, 13).clearContent();
    updateSingleRowNeedHighlight_(sheet, row, need);
    if (sheetName === "3. Лист ожидания") refreshSingleRowTactic_(sheet, row);

    moveRowToBottom_(sheet, row);
    ss.toast(`«${company || inn}» -> Перезвон: ${fullDispStr}. Строка убрана вниз таблицы!`, "⏰ Перезвон сегодня", 4);
    return;
  }

  // === ВЕТКА 3Б: ❄️ СЕЗОН ЗАКРЫТ / В ЭТОМ ГОДУ НЕ БЕРУТ / СДВИГ СПРОСА / КУПИЛ У ДРУГИХ ===
  const isSeasonClosed = decision.includes("сезон") || decision.includes("в этом году") || decision.includes("до весн") || decision.includes("след. год") || decision.includes("замороз");
  if (isSeasonClosed || decision.includes("сдвиг") || decision.includes("купил") || decision.includes("конкурент") || decision.includes("тендер")) {
    const defaultShiftDays = isSeasonClosed ? 25 : 30;
    const shiftDateObj = getNextWorkingDateObj_(now, defaultShiftDays);
    const minWaitIso   = getNextWorkingDateIso_(now, 2, timeZone);
    const useExisting  = (existingDlObj.dateObj && existingIsoDeadline > minWaitIso);
    const finalDateObj = useExisting ? existingDlObj.dateObj : shiftDateObj;
    const finalHasTime = useExisting ? existingDlObj.hasTime : false;
    const finalDisp    = Utilities.formatDate(finalDateObj, timeZone, finalHasTime ? "dd.MM.yyyy HH:mm" : "dd.MM.yyyy");

    const tagPrefix = isSeasonClosed ? `❄️ Сезон закрыт / В этом году всё (Проверка на блеф)` : `🔄 Сдвиг / Купил у других`;
    const tag = `[${tagPrefix}${needSnapshot} (след. контакт: ${finalDisp}) - ${nowStr}]`;
    const updatedHistory = prependHistoryLog_(history, tag);

    if (sheetName === "3. Лист ожидания") {
      sheet.getRange(row, 8).setValue("Нет").setBackground(null);
      sheet.getRange(row, 9, 1, 2).clearContent();
      sheet.getRange(row, 14).setValue(updatedHistory);
      setDeadlineCellDate_(sheet, row, finalDateObj, finalHasTime);
      sheet.getRange(row, 13).clearContent();
      refreshSingleRowTactic_(sheet, row);
      moveRowToBottom_(sheet, row);
      ss.toast(`Запрос H+I+J сохранен в Историю (N), H = «Нет», тактика в O переключена (контроль: ${finalDisp})!`, isSeasonClosed ? "❄️ Фоновый прогрев" : "🔄 Сдвиг спроса", 5);
    } else {
      const targetSheet = getSheetByNameSafe(ss, "3. Лист ожидания");
      if (!targetSheet) return ss.toast("Ошибка: лист Ожидания не найден", "Ошибка");
      const mergeRes = insertOrMergeRowByInn_(targetSheet, [
        inn, company, city, lpr, phone, email,
        sphere, "Нет", "", "",
        notes, finalDateObj, "", updatedHistory
      ], companyNote);
      setDeadlineCellDate_(targetSheet, mergeRes.row, finalDateObj, finalHasTime);
      targetSheet.getRange(mergeRes.row, 8).setBackground(null);
      refreshSingleRowTactic_(targetSheet, mergeRes.row);
      sheet.deleteRow(row);
      highlightDuplicateInnsOnSheet_(sheet);
      ss.toast(`«${company || inn}» переведён в Лист ожидания (дедлайн: ${finalDisp})!`, "🔄 В Ожидание", 5);
    }
    return;
  }

  // === ВЕТКА 3В: В ОЖИДАНИЕ (НЕТ ПОТРЕБНОСТИ / РАЗВЕДКА ПРОФИЛЯ) ===
  if (decision.includes("ожидан") || decision.includes("подума") || decision.includes("спящ") || decision.includes("нет потребн")) {
    const defaultWaitObj = getNextWorkingDateObj_(now, 14);
    const todayIso = Utilities.formatDate(now, timeZone, "yyyy-MM-dd");
    const useExisting = (existingDlObj.dateObj && existingIsoDeadline >= todayIso);
    const targetDateObj = useExisting ? existingDlObj.dateObj : defaultWaitObj;
    const targetHasTime = useExisting ? existingDlObj.hasTime : false;
    const targetDisp    = Utilities.formatDate(targetDateObj, timeZone, targetHasTime ? "dd.MM.yyyy HH:mm" : "dd.MM.yyyy");

    const reconSnap = buildNeedSnapshot_(need, pipeType, application, "Берут в работу");
    const tag = `[⏳ В Ожидание${reconSnap} (из «${sheetName}») - ${nowStr}]`;
    const updatedHistory = prependHistoryLog_(history, tag);

    if (sheetName === "3. Лист ожидания") {
      sheet.getRange(row, 8).setValue("Нет").setBackground(null);
      sheet.getRange(row, 9, 1, 2).clearContent();
      sheet.getRange(row, 14).setValue(updatedHistory);
      setDeadlineCellDate_(sheet, row, targetDateObj, targetHasTime);
      sheet.getRange(row, 13).clearContent();
      refreshSingleRowTactic_(sheet, row);
      moveRowToBottom_(sheet, row);
      ss.toast(`Профиль H+I+J записан в Историю (N), H сброшена в «Нет» (I и J очищены), дедлайн: ${targetDisp}`, "⏳ Ожидание", 4);
      return;
    }

    const targetSheet = getSheetByNameSafe(ss, "3. Лист ожидания");
    if (!targetSheet) return ss.toast("Ошибка: лист Ожидания не найден", "Ошибка");

    const mergeRes = insertOrMergeRowByInn_(targetSheet, [
      inn, company, city, lpr, phone, email,
      sphere, "Нет", "", "",
      notes, targetDateObj, "", updatedHistory
    ], companyNote);
    setDeadlineCellDate_(targetSheet, mergeRes.row, targetDateObj, targetHasTime);
    targetSheet.getRange(mergeRes.row, 8).setBackground(null);
    refreshSingleRowTactic_(targetSheet, mergeRes.row);
    sheet.deleteRow(row);
    highlightDuplicateInnsOnSheet_(sheet);
    const msgAction = mergeRes.merged ? "объединён с карточкой" : "перенесён";
    ss.toast(`«${company || inn}» ${msgAction} в Ожидание! Разведка (H+I+J) сохранена в Историю (N), H = «Нет», I и J очищены.`, "⏳ В Ожидание", 5);
    return;
  }

  // === ВЕТКА 4: НЕДОЗВОН (Каскадный шаг ТЗ v6.4: 1 -> 3 -> 7 дней и в Архив на 4/4) ===
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
      const nextCallDateObj = getNextWorkingDateObj_(now, stepDays);
      const nextCallStr     = Utilities.formatDate(nextCallDateObj, timeZone, "dd.MM.yyyy");

      sheet.getRange(row, 14).setValue(updatedHistory);
      setDeadlineCellDate_(sheet, row, nextCallDateObj, false);
      sheet.getRange(row, 13).clearContent();
      if (sheetName === "3. Лист ожидания") refreshSingleRowTactic_(sheet, row);

      moveRowToBottom_(sheet, row);
      ss.toast(`Недозвон (${attempts}/4, шаг +${stepDays} дн.). Дедлайн: ${nextCallStr}. Строка опущена вниз!`, "📞 Недозвон", 4);
    } else {
      const targetSheet = getSheetByNameSafe(ss, "4.Архив лидов");
      if (!targetSheet) return ss.toast("Ошибка: лист Архива не найден", "Ошибка");

      const archiveTag = `[📦 В Архив: 4 недозвона${needSnapshot} (из «${sheetName}») - ${nowStr}]`;
      const finalArchiveHistory = prependHistoryLog_(updatedHistory, archiveTag);

      const mergeRes = insertOrMergeRowByInn_(targetSheet, [
        inn, company, city, lpr, phone, email,
        sphere, "Нет", "", "",
        notes, "", "Недозвон 4/4", finalArchiveHistory
      ], companyNote);
      targetSheet.getRange(mergeRes.row, 8).setBackground(null);
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

    const mergeRes = insertOrMergeRowByInn_(targetSheet, [
      inn, company, city, lpr, phone, email,
      sphere, "Нет", "", "",
      notes, "", reasonText, finalArchiveHistory
    ], companyNote);
    targetSheet.getRange(mergeRes.row, 8).setBackground(null);
    sheet.deleteRow(row);
    highlightDuplicateInnsOnSheet_(sheet);
    ss.toast(`«${company || inn}» отправлен в Архив (${reasonText}). Все 14 колонок и Досье сохранены!`, `${icon} ${reasonText}`, 4);
    return;
  }
}

// =================================================================
// СБРОС ПОДСВЕТОК ПРИ ПУСТЫХ ЯЧЕЙКАХ И ПАКЕТНОМ УДАЛЕНИИ (DELETE)
// =================================================================

function updateSingleRowNeedHighlight_(sheet, row, needVal) {
  const cell = sheet.getRange(row, 8);
  const str = (needVal !== undefined ? String(needVal) : String(cell.getValue() || '')).trim();
  const isAct = str !== '' && str.toLowerCase() !== 'нет' && str !== '—' && str !== '-';

  if (!isAct) {
    cell.setBackground(null);
    return;
  }

  const stockByDiam = getAvailableStockDiamsMap_(sheet.getParent());
  const liveDiams = Array.from(new Set((str.match(/\b\d{2,4}\b/g) || []).map(Number))).filter(d => stockByDiam[d]);
  cell.setBackground(liveDiams.length > 0 ? '#dcfce7' : null);
}

function cleanupRowIfCompletelyEmpty_(sheet, row, sheetName) {
  if (row < 2) return false;
  const vals = sheet.getRange(row, 1, 1, 14).getValues()[0];
  const hasAnyData = vals.some(v => v !== null && v !== undefined && String(v).trim() !== '');
  if (hasAnyData) return false;

  const maxCol = sheetName === "3. Лист ожидания" ? 15 : 14;
  sheet.getRange(row, 1, 1, maxCol).setBackground(null).setFontColor('#0f172a').setFontWeight('normal');
  sheet.getRange(row, 2).setNote('');
  sheet.getRange(row, 12).setNumberFormat('dd.MM.yyyy');
  if (sheetName === "3. Лист ожидания") {
    sheet.getRange(row, 15).clearContent();
  }
  return true;
}

function handleMultiCellCrmEdit_(e, sheet) {
  const sheetName = sheet.getName().trim();
  const startRow = Math.max(2, e.range.getRow());
  const endRow = e.range.getLastRow();

  for (let r = startRow; r <= endRow; r++) {
    if (!cleanupRowIfCompletelyEmpty_(sheet, r, sheetName)) {
      const hVal = String(sheet.getRange(r, 8).getValue() || '').trim();
      updateSingleRowNeedHighlight_(sheet, r, hVal);
      updateSingleRowDeadlineHighlight_(sheet, r);
      if (sheetName === "3. Лист ожидания") {
        refreshSingleRowTactic_(sheet, r);
      }
    }
  }

  highlightDuplicateInnsOnSheet_(sheet, endRow);
  syncUnpackedNotesOnSheetSilent_(sheet);
}

// =================================================================
// 100% АВТОМАТИЧЕСКОЕ ДОСЬЕ КЛИЕНТА (ВВОД В K -> СРАЗУ В ЗАМЕТКУ B + ВЫЖИМКА В K)
// =================================================================

function processRawNoteIntoDossier_(sheet, row, oldVal) {
  const noteCell = sheet.getRange(row, 11);
  const compCell = sheet.getRange(row, 2);
  const rawText  = String(noteCell.getValue() || '').trim();
  if (!rawText) return;

  const existingDossier = String(compCell.getNote() || '').trim();
  const timeZone = sheet.getParent().getSpreadsheetTimeZone() || "Europe/Moscow";
  const dateStr  = Utilities.formatDate(new Date(), timeZone, "dd.MM.yyyy HH:mm");

  let cleanThought = rawText;
  if (rawText.startsWith('📌')) {
    const pipeIdx = rawText.indexOf('|');
    cleanThought = pipeIdx !== -1 ? rawText.slice(pipeIdx + 1).trim() : rawText.replace(/^📌\s*/, '').trim();
  }

  let updatedDossier = existingDossier;
  if (cleanThought && !existingDossier.includes(cleanThought)) {
    const header = existingDossier ? existingDossier + '\n────────────────────\n' : '🗂️ ДОСЬЕ КЛИЕНТА («КАПИТАЛ-ПЛЮС»):\n\n';
    updatedDossier = `${header}✍️ [${dateStr}]:\n${cleanThought}`;
    compCell.setNote(updatedDossier);
  }

  const summary = extractSmartSummaryFromNote_(cleanThought || rawText, updatedDossier);
  if (summary && summary !== rawText) {
    noteCell.setValue(summary);
  }
}

function extractSmartSummaryFromNote_(latestText, fullDossierText) {
  const combinedSource = `${fullDossierText || ''} ${latestText || ''}`;
  const lower = combinedSource.toLowerCase();
  const tags = [];

  if (/весы|взвеш|микрометр|толщиномер|стенкомер|хитро/i.test(lower))   tags.push('⚖️ ВЕСЫ: +10% на допуск!');
  if (/в этом году|до весн|не сезон|сезон закрыт|след.*год|зимой не/i.test(lower)) tags.push('❄️ Сезон/Блеф?');
  if (/не закуп|отказ|свой поставщ|берут у друг/i.test(lower))          tags.push('🛡️ Барьер/Свой поставщик');
  if (/тендер|торг|аукцион|закупк|44-фз|223-фз|аст|ртс/i.test(lower)) tags.push('🏛️ Тендеры');
  if (/аккредитац|реестр|сб\b|служб.*безопасн/i.test(lower))          tags.push('📑 Аккредитация');
  if (/отсрочк|постоплат|аванс|финансир|нет денег|бюджет/i.test(lower)) tags.push('💳 Финансы/Отсрочка');
  if (/прокол|гнб|футляр|гильз/i.test(lower))                        tags.push('🛠️ Прокол/Футляр');
  if (/обсад|скважин|бур/i.test(lower))                              tags.push('💧 Обсадка/Бурение');
  if (/сваи|свай|опор|мост/i.test(lower))                            tags.push('🏗️ Сваи/Опоры');

  const diams = extractTruePipeDiamsFromText_(combinedSource);
  if (diams.length > 0) {
    tags.push(`Ø ${diams.slice(0, 5).join(',')}`);
  }

  const cleanLine = String(latestText || '').replace(/\n+/g, ' ').replace(/^📌\s*/, '').trim();
  // Увеличен лимит до 110 символов, чтобы не обрезать смысл заметки на полслове!
  const shortPhrase = cleanLine.length > 115 ? cleanLine.slice(0, 112).trim() + '…' : cleanLine;

  if (tags.length > 0) {
    return `📌 ${tags.join(' • ')} | ${shortPhrase}`;
  }
  return `📌 ${shortPhrase}`;
}

function syncUnpackedNotesOnSheetSilent_(sheet) {
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const numRows = lastRow - 1;
  const kVals  = sheet.getRange(2, 11, numRows, 1).getValues();
  const bNotes = sheet.getRange(2, 2, numRows, 1).getNotes();

  for (let i = 0; i < numRows; i++) {
    const rawK = String(kVals[i][0] || '').trim();
    const noteB = String(bNotes[i][0] || '').trim();
    if (rawK && (!noteB || !rawK.startsWith('📌'))) {
      processRawNoteIntoDossier_(sheet, i + 2);
    }
  }
}

// =================================================================
// МОЗГ ТОРГОВОГО ДЕСКА (КОЛОНКА O В «3. ЛИСТ ОЖИДАНИЯ» — СКЛАД + ТРАНЗИТ + БЛЕФ-ЧЕК)
// =================================================================

/**
 * Считывает живой Склад вместе с тоннажем, состоянием и ценой ₽/тн по каждому диаметру
 */
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
    const wall    = typeof safeParseFloat === 'function' ? safeParseFloat(r[4]) : parseFloat(r[4]) || 0;
    const condRaw = String(r[2] || '').trim();
    const priceTn = Math.round(typeof safeParseFloat === 'function' ? safeParseFloat(r[7]) : parseFloat(r[7]) || 0);

    if (status !== 'нет' && (stockTn > 0 || stockM > 0) && diam > 0) {
      if (!stockByDiam[diam]) stockByDiam[diam] = [];
      let condShort = 'Б/у';
      const cl = condRaw.toLowerCase();
      if (cl.includes('нов')) condShort = 'Нов.';
      else if (cl.includes('леж')) condShort = 'Леж.';
      else if (cl.includes('восст')) condShort = 'Восст.';

      const tnStr = stockTn > 0 ? `${ Number(stockTn.toFixed(1)) }т` : `${ Math.round(stockM) }м`;
      const prStr = priceTn > 0 ? ` по ${(priceTn / 1000).toFixed(0)}к` : '';
      stockByDiam[diam].push(`${diam}х${wall} ${condShort} (${tnStr}${prStr})`);
    }
  });
  return stockByDiam;
}

/**
 * Извлекает ТОЛЬКО реальные диаметры труб, игнорируя индексы, дома и случайные числа в тексте
 */
function extractTruePipeDiamsFromText_(text) {
  if (!text) return [];
  const clean = String(text)
    // Убираем даты (19.10.2026) и время (14:26)
    .replace(/\b\d{1,2}[\.\-\/]\d{1,2}[\.\-\/]\d{2,4}\b/g, ' ')
    .replace(/\b\d{1,2}:\d{2}\b/g, ' ');

  const found = [];
  const matches = clean.match(/\b\d{2,4}\b/g) || [];
  matches.forEach(m => {
    const n = parseInt(m, 10);
    if (STANDARD_PIPE_DIAMS_.includes(n) && !found.includes(n)) {
      found.push(n);
    }
  });
  return found;
}

function extractShortCityForTactic_(rawAddr) {
  const s = String(rawAddr || '').trim();
  if (!s) return '';
  const gMatch = s.match(/(?:г\.|город)\s*([А-ЯЁа-яё\-]+(?:\s+[А-ЯЁа-яё\-]+)?)/i);
  if (gMatch && gMatch[1]) return 'г. ' + gMatch[1].trim();
  const oblMatch = s.match(/([А-ЯЁа-яё\-]+)\s+(?:обл|край|респ)/i);
  if (oblMatch && oblMatch[1]) return oblMatch[1].trim();
  const first = s.split(',')[0].replace(/^\d{6}\s*/, '').trim();
  return first.length > 22 ? '' : first;
}

function computeDynamicTacticForRow_(rowVals, stockByDiam, companyDossierNote) {
  const cityRaw = String(rowVals[2] || '').trim();
  const email   = String(rowVals[5] || '').trim();
  const sphere  = String(rowVals[6] || '').trim();
  const need    = String(rowVals[7] || '').trim();
  const pType   = String(rowVals[8] || '').trim();
  const pApp    = String(rowVals[9] || '').trim();
  const notes   = String(rowVals[10] || '').trim();
  const history = String(rowVals[13] || '').trim();
  const dossier = String(companyDossierNote || '').trim();

  const shortCity = extractShortCityForTactic_(cityRaw);
  const cityTag   = shortCity ? ` [${shortCity}]` : '';
  const hasEmail  = email.includes('@');

  const allContext = `${need} ${pType} ${pApp} ${notes} ${dossier} ${history}`;
  const allLower   = allContext.toLowerCase();

  const needLower = need.toLowerCase();
  const hasActiveNeed = need !== '' && needLower !== 'нет' && needLower !== '—' && needLower !== '-';

  const historyLines = history.split('\n').map(s => s.trim()).filter(Boolean);
  const firstHistoryLine = historyLines[0] || '';
  const touchCount = historyLines.length;

  const isScaleClient = /весы|взвеш|микрометр|толщиномер|стенкомер|хитро/i.test(allLower);
  const scaleWarnTag  = isScaleClient ? ' [⚖️ ВЕСЫ: +10% на допуск!]' : '';

  // 1. ПРИОРИТЕТ №1: ДОЖИМ ВЫСТАВЛЕННОГО СЧЕТА
  if (firstHistoryLine.includes('💰') || firstHistoryLine.toLowerCase().includes('выставлен счет')) {
    const targetPipe = hasActiveNeed ? ` (${need})` : '';
    return `💰 ДОЖИМ СЧЕТА${targetPipe}${cityTag}: узнать проходную цену конкурентов / подтвердить дату оплаты${scaleWarnTag}`;
  }

  // 2. ПРИОРИТЕТ №2: ЖИВАЯ ЗАЯВКА В КОЛОНКЕ H (Склад vs Транзит + Логистика)
  if (hasActiveNeed) {
    const liveDiams = extractTruePipeDiamsFromText_(need);
    const inStockDetails = [];
    liveDiams.forEach(d => {
      if (stockByDiam[d] && stockByDiam[d].length > 0) {
        inStockDetails.push(stockByDiam[d].slice(0, 2).join(', '));
      }
    });

    if (inStockDetails.length > 0) {
      return `🔥 НА СКЛАДЕ ПОД ЗАЯВКУ (${need})${cityTag}: есть ${inStockDetails.join(' | ')} — дать цену + расчет машины и закрыть в Счет!${scaleWarnTag}`;
    }
    return `🌐 ТРАНЗИТ В ПРОСЧЕТ (${need})${cityTag}: нет на базе Воронеж — пробить Закупку у поставщиков + попутную фуру (CPT) за 15 мин!${scaleWarnTag}`;
  }

  // 3. ПРИОРИТЕТ №3: ПРОБИВ НЕДОЗВОНА ИЛИ СЕКРЕТАРЯ
  const missMatch = firstHistoryLine.match(/Недозвон:\s*(\d+)\/4/i);
  if (missMatch) {
    const att = parseInt(missMatch[1], 10);
    if (att === 1) return `📲 После недозвона (1/4): набрать в другое время дня (утро до планерки / после 15:00)`;
    if (att === 2) return `💬 После недозвона (2/4): скинуть короткий пинг в WhatsApp / Telegram на мобильный ЛПР`;
    return `⚠️ Недозвон (3/4 — последний шанс): пробить через ПТО / снабжение перед Архивом`;
  }

  // 4. ПРИОРИТЕТ №4: ГОНКОНГСКИЙ ПРОТОКОЛ «В ЭТОМ ГОДУ УЖЕ НИЧЕГО НЕ БУДЕТ / СЕЗОН ЗАКРЫТ» (ПРОВЕРКА НА БЛЕФ vs ФОН)
  const isOffSeasonOrBluff = /❄️|в этом году|до весн|не сезон|сезон закрыт|след.*год|зимой не|объектов нет до/i.test(allLower);
  if (isOffSeasonOrBluff) {
    const bluffStep = touchCount % 3;
    if (bluffStep === 0) {
      return `🕵️ ПРОВЕРКА НА БЛЕФ («В этом году всё»)${cityTag}: пробить по ИНН активные тендеры + скинуть в WA 1 позицию ниже рынка (если врут и берут у других — клюнут!)`;
    } else if (bluffStep === 1) {
      return `💰 КОНЕЦ ГОДА / БАРТЕР${cityTag}: предложить освоить остаток годового бюджета (счет сейчас + зимнее хранение у нас до весны) ИЛИ выкупить их остатки трубы с базы!`;
    } else {
      return `📐 ЗАХОД В ВЕСЕННЮЮ СМЕТУ${cityTag}: забрать спецификацию будущих объектов сейчас — заложить нашу лежалую/восст. вместо новой трубы (-25% к смете)!`;
    }
  }

  // 5. ПРИОРИТЕТ №5: СНАЙПЕР СКЛАДА (По прошлым диаметрам из Истории N и Досье B/K)
  const pastDiams = extractTruePipeDiamsFromText_(`${notes} ${dossier} ${history}`);
  const matchedStockList = [];
  pastDiams.forEach(d => {
    if (stockByDiam[d] && stockByDiam[d].length > 0) {
      matchedStockList.push(stockByDiam[d][0]);
    }
  });

  if (matchedStockList.length > 0) {
    const stockSummary = matchedStockList.slice(0, 2).join('; ');
    const sniperMode = touchCount % 3;
    if (sniperMode === 0) {
      return `🎯 СНАЙПЕР СКЛАДА${cityTag}: под их профиль в наличии ${stockSummary} — предложить живой объем с базы Воронеж${scaleWarnTag}`;
    } else if (sniperMode === 1) {
      return `🚛 ПОПУТКА / ДОГРУЗ СО СКЛАДА${cityTag}: идет машина в их сторону — предложить закинуть ${stockSummary} по цене попутки${scaleWarnTag}`;
    } else {
      return `📉 АРБИТРАЖ СМЕТЫ СО СКЛАДА${cityTag}: предложить партию ${stockSummary} под экономию бюджета или выкупить их деловые остатки${scaleWarnTag}`;
    }
  }

  // 6. ПРИОРИТЕТ №6: СПЕЦ-СИГНАЛЫ ИЗ ДОСЬЕ КЛИЕНТА (Отказ снабженца, Тендеры, Аккредитация, Финансы)
  if (/не закуп|жесткий отказ|свой поставщ|берут у друг/i.test(allLower)) {
    return `🛡️ ОБХОД СНАБЖЕНЦА${cityTag}: снабженец блокирует — выйти напрямую на гл. инженера / начальника участка или дать проходную цену как 2-й резервный поставщик`;
  }
  if (/аккредитац|реестр|служб.*безопасн/i.test(allLower) && touchCount % 2 === 0) {
    return `📑 АККРЕДИТАЦИЯ В СБ: проверить статус проверки наших уставных документов и зайти в закрытый реестр поставщиков`;
  }
  if (/тендер|торг|аукцион|44-фз|223-фз|аст/i.test(allLower) && touchCount % 2 === 0) {
    return `🏛️ ТЕНДЕРНЫЙ ДЕСК: запросить ТЗ ближайшего тендера — дать проходную цену за счет стенки в минусовом допуске ГОСТ или восстановки`;
  }
  if (/не финансир|нет денег|ждут аванс|задерж/i.test(allLower) && touchCount % 2 === 0) {
    return `💳 ФИНАНСЫ / ТРАНШ: проверить приход аванса от заказчика / предложить отгрузку частями (по 5–10 тн) или бартер на их лежалые остатки`;
  }

  // 7. ПРИОРИТЕТ №7: ОТРАСЛЕВОЙ ЗАХОД ПО СФЕРЕ (Колонка G) + АКТИВНАЯ РОТАЦИЯ (Без вечного зависания на Шаге 1!)
  const kpSent = /📧|💬|📄|Прайс|КП/i.test(history) || /кп|прайс/i.test(allLower);

  // Шаг 1 показываем ТОЛЬКО если с клиентом еще вообще не было касаний и не отправлялся прайс
  if (!kpSent && touchCount === 0 && !notes) {
    if (hasEmail) {
      return `📧 ШАГ 1 (Якорь на Email)${cityTag}: выслать Прайс + Карточку компании и узнать рабочие Ø`;
    }
    return `💬 ШАГ 1 (Якорь в WhatsApp/TG)${cityTag}: нет почты в F — скинуть визитку и наличие в мессенджер ЛПР + выбить рабочий Email`;
  }

  // Если известна Сфера (G), но еще не выбиты точные диаметры — бьем калибрами со склада под их отрасль!
  if (pastDiams.length === 0 && sphere) {
    const sphLower = sphere.toLowerCase();
    if (sphLower.includes('бур') || sphLower.includes('гнб')) {
      return `💧 ПОД БУРЕНИЕ / ГНБ${cityTag}: предложить со склада НКТ 73, 89 и обсадную/футляры 168–325 мм (узнать текущие буровые установки)`;
    }
    if (sphLower.includes('дорож') || sphLower.includes('мост')) {
      return `🏗️ ПОД ДРСУ / МОСТЫ${cityTag}: предложить трубу под водопропускные гильзы, сваи и проколы Ø 325–820 мм + выкуп их остатков с базы`;
    }
    if (sphLower.includes('строит') || sphLower.includes('пгс')) {
      return `🏢 ПОД СТРОЙКУ (ПГС)${cityTag}: предложить трубу под сваи, футляры коммуникаций и распорки котлованов (лежалая/восст. со склада)`;
    }
  }

  const rot = touchCount % 4;
  if (rot === 0) {
    return `🎯 РАЗВЕДКА КАЛИБРОВ${cityTag}: выбить рабочие Ø, допуск по б/у-восстановке и график ближайших объектов`;
  } else if (rot === 1) {
    return `📉 ЗАМЕНА НОВОЙ НА ЛЕЖАК/ВОССТ.${cityTag}: предложить пересогласовать трубу в смете на нашу лежалую или восстановленную (-25% к бюджету)`;
  } else if (rot === 2) {
    return `♻️ ОБРАТНЫЙ ВЫКУП / ДОГРУЗ${cityTag}: спросить, нет ли на их базе невостребованных остатков трубы на выкуп/обмен + предложить попутную доставку`;
  } else {
    return `🤝 ВЫХОД НА ПТО / ГЛ. ИНЖЕНЕРА${cityTag}: узнать контакт ПТО или механика, забрать спецификацию под сваи/футляры/опоры на просчет`;
  }
}

function refreshSingleRowTactic_(sheet, row) {
  if (!sheet || sheet.getName().trim() !== "3. Лист ожидания" || row < 2) return;
  const rowVals = sheet.getRange(row, 1, 1, 14).getValues()[0];
  const hasAny = rowVals.some(v => v !== null && v !== undefined && String(v).trim() !== '');
  if (!hasAny) {
    sheet.getRange(row, 15).clearContent().setBackground(null);
    return;
  }
  const ss = sheet.getParent();
  const stockByDiam = getAvailableStockDiamsMap_(ss);
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

function buildNeedSnapshot_(need, pipeType, application, customPrefix) {
  const cleanNeed = (need && need.toLowerCase() !== 'нет' && need !== '—' && need !== '-') ? need : '';
  const details = [pipeType, application].filter(Boolean).join(' • ');
  const label = customPrefix || 'Запрос';

  if (cleanNeed && details) {
    return ` | ${label}: ${cleanNeed} (${details})`;
  } else if (cleanNeed) {
    return ` | ${label}: ${cleanNeed}`;
  } else if (details) {
    return ` | Профиль: ${details}`;
  }
  return '';
}

// =================================================================
// ДВУХУРОВНЕВЫЙ БИРЖЕВОЙ РАДАР + СБРОС ЗАЛИПШИХ ПОДСВЕТОК В H
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

    syncUnpackedNotesOnSheetSilent_(sh);

    const lRow = sh.getLastRow();
    const data = sh.getRange(2, 1, lRow - 1, 14).getValues();
    const compNotes = sh.getRange(2, 2, lRow - 1, 1).getNotes();
    const needBgColors = [];
    const tacticsCol = [];

    data.forEach((r, idx) => {
      const comp    = String(r[1] || '').trim();
      const lpr     = String(r[3] || '').trim();
      const ph      = String(r[4] || '').trim().split('\n')[0];
      const needStr = String(r[7] || '').trim();
      const notes   = String(r[10] || '').trim();
      const histStr = String(r[13] || '').trim();
      const dossier = String((compNotes[idx] && compNotes[idx][0]) || '').trim();

      if (sName === "3. Лист ожидания") {
        tacticsCol.push([computeDynamicTacticForRow_(r, stockByDiam, dossier)]);
      }

      const isNeedActive = needStr !== '' && needStr.toLowerCase() !== 'нет' && needStr !== '—' && needStr !== '-';

      if (isNeedActive) {
        const liveDiams = extractTruePipeDiamsFromText_(needStr).filter(d => stockByDiam[d]);
        if (liveDiams.length > 0) {
          needBgColors.push(['#dcfce7']);
          hotMatches.push(
            `🔥 [${sName.replace(/^\d+\.\s*/, '')}] ${comp || 'Клиент'} (${lpr || ph}): сейчас в заявке Ø ${liveDiams.join(', ')}`
          );
          return;
        }
      }

      needBgColors.push([null]);

      const combinedPast = `${histStr} ${notes} ${dossier}`;
      const pastDiams = extractTruePipeDiamsFromText_(combinedPast).filter(d => stockByDiam[d]);

      if (pastDiams.length > 0) {
        historyMatches.push(
          `💡 [${sName.replace(/^\d+\.\s*/, '')}] ${comp || 'Клиент'} (${lpr || ph}): в истории/досье Ø ${pastDiams.join(', ')}`
        );
      }
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
      `🎯 Биржевой радар: Горячих в H (${hotMatches.length}) | По Истории и Досье (${historyMatches.length})`,
      `Колонка O (Чем дожимать) обновлена со свежими тоннами и ценами со Склада!\n\n${preview}${more}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    ss.toast('Колонка O обновлена со свежими данными Склада!', '🔍 Радар обновлен', 5);
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

function highlightDuplicateInnsOnSheet_(sheet, minRowToCover) {
  const lastRow = Math.min(sheet.getMaxRows(), Math.max(sheet.getLastRow() + 5, minRowToCover || 2));
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

function getNextWorkingDateObj_(baseDate, addDays) {
  const nextDate = new Date(baseDate.getTime());
  nextDate.setHours(0, 0, 0, 0);
  nextDate.setDate(nextDate.getDate() + addDays);
  if (nextDate.getDay() === 6) nextDate.setDate(nextDate.getDate() + 2);
  else if (nextDate.getDay() === 0) nextDate.setDate(nextDate.getDate() + 1);
  return nextDate;
}

function getNextWorkingDateIso_(baseDate, addDays, timeZone) {
  const nextDate = getNextWorkingDateObj_(baseDate, addDays);
  return Utilities.formatDate(nextDate, timeZone || "Europe/Moscow", "yyyy-MM-dd");
}

function getNextWorkingDateStr_(baseDate, addDays, timeZone) {
  const nextDate = getNextWorkingDateObj_(baseDate, addDays);
  return Utilities.formatDate(nextDate, timeZone || "Europe/Moscow", "dd.MM.yyyy");
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
    const newNeed = (rowValues[7] && String(rowValues[7]).toLowerCase() !== 'нет') ? rowValues[7] : (oldRow[7] || "Нет");
    const mergedRow = [
      incomingInn || oldRow[0],
      oldRow[1] || rowValues[1],
      oldRow[2] || rowValues[2],
      mergeUniqueLines_(oldRow[3], rowValues[3], '\n'),
      formatRussianPhones_(mergeUniqueLines_(oldRow[4], rowValues[4], '\n')),
      mergeUniqueLines_(oldRow[5], rowValues[5], '\n'),
      oldRow[6] || rowValues[6],
      newNeed,
      rowValues[8] || oldRow[8],
      rowValues[9] || oldRow[9],
      mergeUniqueLines_(oldRow[10], rowValues[10], ' | '),
      rowValues[11] || oldRow[11],
      "",
      prependHistoryLog_(oldRow[13], rowValues[13])
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
  targetRange.setBackground(null);
  targetSheet.getRange(targetRow, 1).setNumberFormat('@');
  targetSheet.getRange(targetRow, 5).setNumberFormat('@');
  ensureDateValidationOnCell_(targetSheet.getRange(targetRow, 12));
  targetRange.setValues([rowValues]);
  return targetRow;
}

// =================================================================
// НАСТОЯЩИЕ ДАТЫ (ВЫПАДАЮЩИЙ КАЛЕНДАРЬ) + БЫСТРОЕ ВРЕМЯ (14 / 1430 / 14 30) + АВТО-ПОДЪЕМ
// =================================================================

function ensureDateValidationOnCell_(range) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(true)
    .build();
  range.setDataValidation(rule);
}

function setDeadlineCellDate_(sheet, row, dateObj, hasTime) {
  const cell = sheet.getRange(row, 12);
  ensureDateValidationOnCell_(cell);
  if (!dateObj) {
    cell.clearContent();
    cell.setNumberFormat("dd.MM.yyyy");
    cell.setBackground(null).setFontColor('#0f172a').setFontWeight('normal');
    return;
  }
  cell.setNumberFormat(hasTime ? "dd.MM.yyyy HH:mm" : "dd.MM.yyyy");
  cell.setValue(dateObj);
  updateSingleRowDeadlineHighlight_(sheet, row);
}

function parseSmartDeadlineInput_(cellVal, rawEventVal, rawOldVal, timeZone) {
  const rawStr = rawEventVal !== undefined ? String(rawEventVal).trim() : "";
  const now = new Date();
  const tz = timeZone || "Europe/Moscow";

  // 1. Введено 1-2 цифры часа ("9", "14") -> Сегодня в 09:00 / 14:00
  const hourOnlyMatch = rawStr.match(/^([01]?\d|2[0-3])$/);
  if (hourOnlyMatch) {
    const hh = parseInt(hourOnlyMatch[1], 10);
    return parseDeadlineObj_(new Date(now.getFullYear(), now.getMonth(), now.getDate(), hh, 0, 0), tz);
  }

  // 2. Введено 3-4 цифры времени слитно ("930", "1430") -> Сегодня в 09:30 / 14:30
  const fourDigitTime = rawStr.match(/^([01]?\d|2[0-3])([0-5]\d)$/);
  if (fourDigitTime) {
    const hh = parseInt(fourDigitTime[1], 10);
    const mm = parseInt(fourDigitTime[2], 10);
    return parseDeadlineObj_(new Date(now.getFullYear(), now.getMonth(), now.getDate(), hh, mm, 0), tz);
  }

  // 3. Введено время через пробел, двоеточие или точку ("14 30", "14:30", "14.30")
  const timeMatch = rawStr.match(/^([01]?\d|2[0-3])[:\.\s]+([0-5]\d)$/);
  if (timeMatch) {
    const hh = parseInt(timeMatch[1], 10);
    const mm = parseInt(timeMatch[2], 10);
    return parseDeadlineObj_(new Date(now.getFullYear(), now.getMonth(), now.getDate(), hh, mm, 0), tz);
  }

  // 4. Если Google Таблицы превратили ввод времени в дату 1899–1910 гг.
  if (cellVal instanceof Date && !isNaN(cellVal.getTime()) && cellVal.getFullYear() <= 1910) {
    const combined = new Date(now.getFullYear(), now.getMonth(), now.getDate(), cellVal.getHours(), cellVal.getMinutes(), 0);
    return parseDeadlineObj_(combined, tz);
  }

  return parseDeadlineObj_(cellVal !== undefined && cellVal !== "" ? cellVal : rawStr, tz);
}

function parseDeadlineObj_(val, timeZone) {
  const empty = { dateObj: null, isoDate: "", timeStr: "", hasTime: false, displayStr: "" };
  if (val === null || val === undefined || val === "") return empty;
  const tz = timeZone || "Europe/Moscow";
  const now = new Date();

  if (val instanceof Date && !isNaN(val.getTime())) {
    if (val.getFullYear() <= 1910) {
      const combined = new Date(now.getFullYear(), now.getMonth(), now.getDate(), val.getHours(), val.getMinutes(), 0);
      return parseDeadlineObj_(combined, tz);
    }
    const isoDate = Utilities.formatDate(val, tz, "yyyy-MM-dd");
    const ruDate  = Utilities.formatDate(val, tz, "dd.MM.yyyy");
    const hhmm    = Utilities.formatDate(val, tz, "HH:mm");
    const hasTime = (hhmm !== "00:00");
    return {
      dateObj: val,
      isoDate: isoDate,
      timeStr: hasTime ? hhmm : "",
      hasTime: hasTime,
      displayStr: hasTime ? `${ruDate} ${hhmm}` : ruDate
    };
  }

  const s = String(val).trim();
  if (!s) return empty;

  const hourOnlyMatch = s.match(/^([01]?\d|2[0-3])$/);
  if (hourOnlyMatch) {
    const dObj = new Date(now.getFullYear(), now.getMonth(), now.getDate(), parseInt(hourOnlyMatch[1], 10), 0, 0);
    return parseDeadlineObj_(dObj, tz);
  }

  const fourDigitTime = s.match(/^([01]?\d|2[0-3])([0-5]\d)$/);
  if (fourDigitTime) {
    const dObj = new Date(now.getFullYear(), now.getMonth(), now.getDate(), parseInt(fourDigitTime[1], 10), parseInt(fourDigitTime[2], 10), 0);
    return parseDeadlineObj_(dObj, tz);
  }

  const timeOnlyMatch = s.match(/^([01]?\d|2[0-3])[:\.\s]+([0-5]\d)$/);
  if (timeOnlyMatch) {
    const dObj = new Date(now.getFullYear(), now.getMonth(), now.getDate(), parseInt(timeOnlyMatch[1], 10), parseInt(timeOnlyMatch[2], 10), 0);
    return parseDeadlineObj_(dObj, tz);
  }

  const isoMatch = s.match(/^(\d{4})[\-\.](\d{1,2})[\-\.](\d{1,2})(?:[\sT,]+([01]?\d|2[0-3])[:\.\s]([0-5]\d))?/);
  if (isoMatch) {
    const y  = parseInt(isoMatch[1], 10);
    const m  = parseInt(isoMatch[2], 10) - 1;
    const d  = parseInt(isoMatch[3], 10);
    const hh = isoMatch[4] !== undefined ? parseInt(isoMatch[4], 10) : 0;
    const mm = isoMatch[5] !== undefined ? parseInt(isoMatch[5], 10) : 0;
    return parseDeadlineObj_(new Date(y, m, d, hh, mm, 0), tz);
  }

  const ruMatch = s.match(/^(\d{1,2})[\.\-\/](\d{1,2})[\.\-\/](\d{2,4})(?:[\s,]+([01]?\d|2[0-3])[:\.\s]([0-5]\d))?/);
  if (ruMatch) {
    const d  = parseInt(ruMatch[1], 10);
    const m  = parseInt(ruMatch[2], 10) - 1;
    let y    = parseInt(ruMatch[3], 10);
    if (y < 100) y += 2000;
    const hh = ruMatch[4] !== undefined ? parseInt(ruMatch[4], 10) : 0;
    const mm = ruMatch[5] !== undefined ? parseInt(ruMatch[5], 10) : 0;
    return parseDeadlineObj_(new Date(y, m, d, hh, mm, 0), tz);
  }

  const shortRuMatch = s.match(/^(\d{1,2})[\.\-\/](\d{1,2})(?:[\s,]+([01]?\d|2[0-3])[:\.\s]([0-5]\d))?$/);
  if (shortRuMatch) {
    const d  = parseInt(shortRuMatch[1], 10);
    const m  = parseInt(shortRuMatch[2], 10) - 1;
    const y  = now.getFullYear();
    const hh = shortRuMatch[3] !== undefined ? parseInt(shortRuMatch[3], 10) : 0;
    const mm = shortRuMatch[4] !== undefined ? parseInt(shortRuMatch[4], 10) : 0;
    return parseDeadlineObj_(new Date(y, m, d, hh, mm, 0), tz);
  }

  return empty;
}

function extractHistoryCallbackTime_(firstHistLine) {
  if (!firstHistLine || !firstHistLine.includes('⏰')) return "";
  const m = firstHistLine.match(/Перезвонить после\s*(\d{1,2}:\d{2})/i);
  return m ? m[1].padStart(5, '0') : "";
}

function computeDeadlineStyle_(dlObj, todayStr, nowHHMM) {
  if (!dlObj || !dlObj.isoDate) {
    return { bg: null, color: '#0f172a', weight: 'normal' };
  }
  if (dlObj.isoDate < todayStr) {
    return { bg: '#fee2e2', color: '#991b1b', weight: 'bold' };
  }
  if (dlObj.isoDate === todayStr) {
    if (dlObj.timeStr && nowHHMM >= dlObj.timeStr) {
      return { bg: '#fee2e2', color: '#991b1b', weight: 'bold' };
    }
    return { bg: '#fef9c3', color: '#854d0e', weight: 'bold' };
  }
  return { bg: '#dcfce7', color: '#166534', weight: 'normal' };
}

function updateSingleRowDeadlineHighlight_(sheet, row, preParsedObj) {
  if (!sheet || row < 2) return;
  const ss = sheet.getParent();
  const timeZone = ss.getSpreadsheetTimeZone() || "Europe/Moscow";
  const now = new Date();
  const todayStr = Utilities.formatDate(now, timeZone, "yyyy-MM-dd");
  const nowHHMM  = Utilities.formatDate(now, timeZone, "HH:mm");

  const cell = sheet.getRange(row, 12);
  const dlObj = preParsedObj || parseDeadlineObj_(cell.getValue(), timeZone);
  const st = computeDeadlineStyle_(dlObj, todayStr, nowHHMM);

  cell.setBackground(st.bg).setFontColor(st.color).setFontWeight(st.weight);
}

function sortAndHighlightDeadlinesOnSheet_(sheet) {
  if (!sheet) return;
  const sheetName = sheet.getName().trim();
  const allowedSheets = ["1. Обзвон (База)", "2. Золотой список (Киты)", "3. Лист ожидания"];
  if (!allowedSheets.includes(sheetName)) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const ss = sheet.getParent();
  const timeZone = ss.getSpreadsheetTimeZone() || "Europe/Moscow";
  const now = new Date();
  const todayStr = Utilities.formatDate(now, timeZone, "yyyy-MM-dd");
  const nowHHMM  = Utilities.formatDate(now, timeZone, "HH:mm");
  const numCols = sheetName === "3. Лист ожидания" ? 15 : 14;
  const numRows = lastRow - 1;

  const fullRange = sheet.getRange(2, 1, numRows, numCols);
  const allValues = fullRange.getValues();
  const allNotes  = fullRange.getNotes();
  const stockByDiam = getAvailableStockDiamsMap_(ss);

  const items = [];
  let hasStringDatesToConvert = false;
  let tacticsChanged = false;

  for (let i = 0; i < numRows; i++) {
    const r = allValues[i];
    const n = allNotes[i];
    const hasData = r.slice(0, 14).some(c => c !== null && c !== undefined && String(c).trim() !== '');

    if (!hasData) {
      items.push({
        origIdx: i,
        isEmpty: true,
        bucket: 99,
        dStr: "9999-99-99",
        tStr: "99:99",
        dlObj: { dateObj: null, isoDate: "", timeStr: "", hasTime: false },
        values: new Array(numCols).fill(""),
        notes: new Array(numCols).fill("")
      });
      continue;
    }

    const rawDlCell = r[11];
    const dlObj = parseDeadlineObj_(rawDlCell, timeZone);
    if (rawDlCell && !(rawDlCell instanceof Date) && dlObj.dateObj) {
      hasStringDatesToConvert = true;
    }
    r[11] = dlObj.dateObj ? dlObj.dateObj : "";

    const needStr = String(r[7] || '').trim();
    const isNeedActive = needStr !== '' && needStr.toLowerCase() !== 'нет' && needStr !== '—' && needStr !== '-';
    const firstHist = String(r[13] || '').split('\n')[0] || '';
    const isHotOrInvoice = isNeedActive || firstHist.includes('💰') || firstHist.includes('🔥');

    const effectiveTime = dlObj.timeStr || extractHistoryCallbackTime_(firstHist);
    const waitingForTimeToday = (dlObj.isoDate === todayStr) && effectiveTime && (nowHHMM < effectiveTime);

    let bucket = 3;
    if (dlObj.isoDate) {
      if (dlObj.isoDate < todayStr) {
        bucket = isHotOrInvoice ? 0 : 1;
      } else if (dlObj.isoDate === todayStr) {
        if (waitingForTimeToday) {
          bucket = 4;
        } else {
          bucket = isHotOrInvoice ? 0 : 2;
        }
      } else {
        bucket = 5;
      }
    } else if (isHotOrInvoice) {
      bucket = 2;
    }

    if (sheetName === "3. Лист ожидания") {
      const newTactic = computeDynamicTacticForRow_(r, stockByDiam, n[1]);
      if (String(r[14] || '') !== newTactic) tacticsChanged = true;
      r[14] = newTactic;
    }

    items.push({
      origIdx: i,
      isEmpty: false,
      bucket: bucket,
      dStr: dlObj.isoDate || "",
      tStr: effectiveTime || "",
      dlObj: dlObj,
      values: r,
      notes: n
    });
  }

  items.sort((a, b) => {
    if (a.bucket !== b.bucket) return a.bucket - b.bucket;
    if (a.bucket === 1 || a.bucket === 5) {
      if (a.dStr !== b.dStr) return a.dStr.localeCompare(b.dStr);
    }
    if ((a.bucket === 2 || a.bucket === 4) && a.tStr !== b.tStr) {
      return (a.tStr || "00:00").localeCompare(b.tStr || "00:00");
    }
    return a.origIdx - b.origIdx;
  });

  const orderChanged = items.some((it, idx) => it.origIdx !== idx);
  const dlRange = sheet.getRange(2, 12, numRows, 1);

  if (orderChanged) {
    ensureDateValidationOnCell_(dlRange);
    const dlNumberFormats = items.map(it => [it.dlObj && it.dlObj.hasTime ? "dd.MM.yyyy HH:mm" : "dd.MM.yyyy"]);
    sheet.getRange(2, 1, numRows, 1).setNumberFormat('@');
    sheet.getRange(2, 5, numRows, 1).setNumberFormat('@');
    sheet.getRange(2, 8, numRows, 1).setNumberFormat('@');
    dlRange.setNumberFormats(dlNumberFormats);

    fullRange.setValues(items.map(it => it.values));
    fullRange.setNotes(items.map(it => it.notes));
  } else {
    if (hasStringDatesToConvert) {
      ensureDateValidationOnCell_(dlRange);
      const dlNumberFormats = items.map(it => [it.dlObj && it.dlObj.hasTime ? "dd.MM.yyyy HH:mm" : "dd.MM.yyyy"]);
      dlRange.setNumberFormats(dlNumberFormats);
      dlRange.setValues(items.map(it => [it.values[11] || ""]));
    }
    if (sheetName === "3. Лист ожидания" && tacticsChanged) {
      sheet.getRange(2, 15, numRows, 1).setValues(items.map(it => [it.values[14] || ""]));
    }
  }

  const needBgs = [];
  const dlBgs = [];
  const dlColors = [];
  const dlWeights = [];

  items.forEach(it => {
    if (it.isEmpty) {
      needBgs.push([null]);
      dlBgs.push([null]);
      dlColors.push(['#0f172a']);
      dlWeights.push(['normal']);
      return;
    }

    const needStr = String(it.values[7] || '').trim();
    const isAct = needStr !== '' && needStr.toLowerCase() !== 'нет' && needStr !== '—' && needStr !== '-';
    if (isAct) {
      const liveDiams = extractTruePipeDiamsFromText_(needStr).filter(d => stockByDiam[d]);
      needBgs.push([liveDiams.length > 0 ? '#dcfce7' : null]);
    } else {
      needBgs.push([null]);
    }

    const st = computeDeadlineStyle_(it.dlObj, todayStr, nowHHMM);
    dlBgs.push([st.bg]);
    dlColors.push([st.color]);
    dlWeights.push([st.weight]);
  });

  sheet.getRange(2, 8, numRows, 1).setBackgrounds(needBgs);
  dlRange.setBackgrounds(dlBgs);
  dlRange.setFontColors(dlColors);
  dlRange.setFontWeights(dlWeights);

  highlightDuplicateInnsOnSheet_(sheet);
}