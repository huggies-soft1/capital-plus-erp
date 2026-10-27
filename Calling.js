/**
 * ЕДИНЫЙ ЦЕНТР УПРАВЛЕНИЯ ПРОДАЖАМИ И СКЛАДОМ (CTRM «Капитал-Плюс» v6.6)
 * Файл: Calling.js — Полный боевой движок для листов:
 * «1. Обзвон (База)», «2. Золотой список (Киты)», «3. Лист ожидания», «4.Архив лидов».
 * - Монитор Склада в колонке O (Вариант 1 — только сухие факты остатков и цен Склада) для Ожидания и Китов.
 * - VIP-логика для «2. Золотой список (Киты)»: Киты не вылетают в Ожидание при сбросе потребности и не улетают в Архив на 4/4 недозвоне.
 * - Быстрый ввод времени в L (14 / 1430 / 14 30) + выпадающий календарь + авто-подъем очереди при обновлении.
 */

const STANDARD_PIPE_DIAMS_ = [
  1420, 1220, 1020, 920, 820, 720, 630, 530, 478, 426, 402, 377, 351,
  325, 299, 273, 245, 219, 194, 168, 159, 152, 148, 146, 133, 127,
  114, 108, 102, 89, 76, 73, 60, 57
];

// Максимальное число плановых контактов на 1 рабочий день (для умного распределения дат без свалки в один день)
const MAX_CALLS_PER_DAY_ = 10;

function isSheetWithStockMonitor_(sheetName) {
  const n = String(sheetName || '').trim();
  return n === "3. Лист ожидания" || n === "2. Золотой список (Киты)";
}

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

    if (isSheetWithStockMonitor_(sheetName)) {
      refreshSingleRowTactic_(sheet, row);
    }
    return;
  }

  // =================================================================
  // 2.8 МГНОВЕННОЕ АВТО-ДОСЬЕ КЛИЕНТА ПРИ ВВОДЕ В ПРИМЕЧАНИЕ (Колонка K / 11)
  // =================================================================
  if (col === 11) {
    processRawNoteIntoDossier_(sheet, row, e.oldValue);
    if (isSheetWithStockMonitor_(sheetName)) {
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
      const safeDateObj = dlObj.hasTime
        ? shiftWeekendToMonday_(dlObj.dateObj)
        : getSmartBalancedDeadlineObj_(sheet, dlObj.dateObj, timeZone, row);
      const finalDlObj  = (safeDateObj.getTime() !== dlObj.dateObj.getTime())
        ? parseDeadlineObj_(safeDateObj, timeZone)
        : dlObj;
      cell.setNumberFormat(finalDlObj.hasTime ? "dd.MM.yyyy HH:mm" : "dd.MM.yyyy");
      cell.setValue(finalDlObj.dateObj);
      updateSingleRowDeadlineHighlight_(sheet, row, finalDlObj);
      if (!dlObj.hasTime && finalDlObj.isoDate !== dlObj.isoDate) {
        sheet.getParent().toast(
          `Чтобы не копить все звонки в один день или выходной, дата распределена на ${finalDlObj.displayStr}`,
          "📅 Умное распределение",
          4
        );
      }
    } else {
      cell.clearContent();
      cell.setNumberFormat("dd.MM.yyyy");
      cell.setBackground(null).setFontColor('#0f172a').setFontWeight('normal');
    }
    return;
  }

  if (isSheetWithStockMonitor_(sheetName) && (col === 7 || col === 9 || col === 10 || col === 14)) {
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
  const isWhaleSheet = (sheetName === "2. Золотой список (Киты)");

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
    if (isSheetWithStockMonitor_(sheetName)) refreshSingleRowTactic_(sheet, row);

    moveRowToBottom_(sheet, row);
    ss.toast(`«${company || inn}» оставлен на пробив секретаря (контроль: ${nextTryStr}). Строка опущена вниз!`, "🛡️ Выбить ЛПР", 4);
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
      if (isSheetWithStockMonitor_(sheetName)) refreshSingleRowTactic_(sheet, row);
      moveRowToBottom_(sheet, row);
      ss.toast(`🔥 Заявка (${liveNeed}) взята в просчет! Контроль: ${nextCheckDateStr}.`, "🔥 В просчете", 4);
    }
    return;
  }

  // === ВЕТКА 1Б: ВЫСТАВЛЕН СЧЕТ (Автоматический статус модуля «Заказы») ===
  if (decision.includes("выставлен счет") || decision.includes("счет") || decision.includes("счёт") || decision.includes("заказ")) {
    if (sheetName === "1. Обзвон (База)" || sheetName === "3. Лист ожидания") {
      sheet.getRange(row, 13).clearContent();
      ss.toast(
        "Статус «Выставлен счет» ставится и снимается автоматически через модуль «Заказы»! Для живого запроса используйте «Горящая заявка (В просчет)».",
        "💡 Авто-статус модуля Заказы",
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
    if (isSheetWithStockMonitor_(sheetName)) refreshSingleRowTactic_(sheet, row);
    moveRowToBottom_(sheet, row);
    ss.toast(`«${company || inn}» -> Счет зафиксирован в Истории! Контроль оплаты: ${nextCheckDateStr}.`, "💰 Выставлен счет", 4);
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
    if (isSheetWithStockMonitor_(sheetName)) refreshSingleRowTactic_(sheet, row);

    moveRowToBottom_(sheet, row);
    ss.toast(`«${company || inn}» -> Перезвон: ${fullDispStr}. Строка убрана вниз таблицы!`, "⏰ Перезвон сегодня", 4);
    return;
  }

  // === ВЕТКА 3Б: СДВИГ СПРОСА / КУПИЛ У ДРУГИХ / СЕЗОН (ВЕСНА, ЛЕТО, ОСЕНЬ, ЗИМА) ===
  // Разделены: «Сдвиг спроса», «Купил у других» и 4 сезона («Сезон: Весна», «Сезон: Лето», «Сезон: Осень», «Сезон: Зима»).
  // Если в L вручную выставлена будущая дата — сохраняем её (с защитой от Сб/Вс и свалки в один день)!
  const isBoughtOther = decision.includes("купил") || decision.includes("конкурент");
  const isSeason      = decision.includes("сезон") || decision.includes("весн") || decision.includes("лет") || decision.includes("осен") || decision.includes("зим") || decision.includes("след. год") || decision.includes("в этом году");
  const isShiftDemand = decision.includes("сдвиг") || isBoughtOther || isSeason;

  if (isShiftDemand) {
    const todayIso = Utilities.formatDate(now, timeZone, "yyyy-MM-dd");
    const useExisting = (existingDlObj.dateObj && existingIsoDeadline > todayIso);

    let baseTargetDateObj;
    let actionHeader = "🔄 Сдвиг спроса";

    if (isSeason) {
      const seasonInfo = getSeasonalStartDateObj_(now, decision);
      baseTargetDateObj = seasonInfo.dateObj;
      actionHeader = seasonInfo.label;
    } else if (isBoughtOther) {
      const boughtDays = isWhaleSheet ? 14 : 30;
      baseTargetDateObj = getNextWorkingDateObj_(now, boughtDays);
      actionHeader = "🛒 Купил у других";
    } else {
      const shiftDays = isWhaleSheet ? 10 : 14;
      baseTargetDateObj = getNextWorkingDateObj_(now, shiftDays);
      actionHeader = "🔄 Сдвиг спроса";
    }

    const destSheetForBalance = (sheetName === "3. Лист ожидания" || isWhaleSheet)
      ? sheet
      : (getSheetByNameSafe(ss, "3. Лист ожидания") || sheet);

    const rawCandidateDate = useExisting ? existingDlObj.dateObj : baseTargetDateObj;
    const finalHasTime     = useExisting ? existingDlObj.hasTime : false;
    const finalDateObj     = finalHasTime
      ? shiftWeekendToMonday_(rawCandidateDate)
      : getSmartBalancedDeadlineObj_(destSheetForBalance, rawCandidateDate, timeZone, destSheetForBalance === sheet ? row : -1);
    const finalDisp        = Utilities.formatDate(finalDateObj, timeZone, finalHasTime ? "dd.MM.yyyy HH:mm" : "dd.MM.yyyy");

    const tag = `[${actionHeader}${needSnapshot} (след. контакт: ${finalDisp}) - ${nowStr}]`;
    const updatedHistory = prependHistoryLog_(history, tag);

    if (sheetName === "3. Лист ожидания" || isWhaleSheet) {
      sheet.getRange(row, 8).setValue("Нет").setBackground(null);
      sheet.getRange(row, 9, 1, 2).clearContent();
      sheet.getRange(row, 14).setValue(updatedHistory);
      setDeadlineCellDate_(sheet, row, finalDateObj, finalHasTime);
      sheet.getRange(row, 13).clearContent();
      refreshSingleRowTactic_(sheet, row);
      moveRowToBottom_(sheet, row);
      ss.toast(`${actionHeader}: запрос H+I+J сохранен в Историю (N), H = «Нет», след. контакт: ${finalDisp}!`, actionHeader, 5);
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
      ss.toast(`«${company || inn}» переведён в Лист ожидания (${actionHeader}, дедлайн: ${finalDisp})!`, actionHeader, 5);
    }
    return;
  }

  // === ВЕТКА 3В: В ОЖИДАНИЕ / НЕТ ПОТРЕБНОСТИ ===
  // Кит остается в Китах (VIP-пульс +7 раб. дней), Ожидание остается в Ожидании (+14 дн.), из Обзвона переходит в Ожидание!
  if (decision.includes("ожидан") || decision.includes("подума") || decision.includes("спящ") || decision.includes("нет потребн")) {
    const defaultWaitDays = isWhaleSheet ? 7 : 14;
    const defaultWaitObj = getNextWorkingDateObj_(now, defaultWaitDays);
    const todayIso = Utilities.formatDate(now, timeZone, "yyyy-MM-dd");
    const useExisting = (existingDlObj.dateObj && existingIsoDeadline > todayIso);

    const destSheetForBalance = (sheetName === "3. Лист ожидания" || isWhaleSheet)
      ? sheet
      : (getSheetByNameSafe(ss, "3. Лист ожидания") || sheet);

    const rawCandidateDate = useExisting ? existingDlObj.dateObj : defaultWaitObj;
    const targetHasTime    = useExisting ? existingDlObj.hasTime : false;
    const targetDateObj    = targetHasTime
      ? shiftWeekendToMonday_(rawCandidateDate)
      : getSmartBalancedDeadlineObj_(destSheetForBalance, rawCandidateDate, timeZone, destSheetForBalance === sheet ? row : -1);
    const targetDisp       = Utilities.formatDate(targetDateObj, timeZone, targetHasTime ? "dd.MM.yyyy HH:mm" : "dd.MM.yyyy");

    const reconSnap = buildNeedSnapshot_(need, pipeType, application, "Берут в работу");
    const actionLabel = isWhaleSheet ? `🐋 Плановый контакт с Китом (Нет текущей потребности)` : `⏳ В Ожидание`;
    const tag = `[${actionLabel}${reconSnap} - ${nowStr}]`;
    const updatedHistory = prependHistoryLog_(history, tag);

    if (sheetName === "3. Лист ожидания" || isWhaleSheet) {
      sheet.getRange(row, 8).setValue("Нет").setBackground(null);
      sheet.getRange(row, 9, 1, 2).clearContent();
      sheet.getRange(row, 14).setValue(updatedHistory);
      setDeadlineCellDate_(sheet, row, targetDateObj, targetHasTime);
      sheet.getRange(row, 13).clearContent();
      refreshSingleRowTactic_(sheet, row);
      moveRowToBottom_(sheet, row);
      ss.toast(`Профиль H+I+J записан в Историю (N), H = «Нет», следующий контакт: ${targetDisp}`, isWhaleSheet ? "🐋 Кит на контроле" : "⏳ Ожидание", 4);
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
    ss.toast(`«${company || inn}» ${msgAction} в Ожидание! Разведка (H+I+J) сохранена в Историю (N), H = «Нет», дедлайн: ${targetDisp}.`, "⏳ В Ожидание", 5);
    return;
  }

  // === ВЕТКА 4: НЕДОЗВОН (Каскад 1 -> 3 -> 7 дней; Ожидание и Киты защищены от авто-сброса в Архив на 4/4!) ===
  if (decision.includes("недозвон") || decision.includes("не дозвон") || decision.includes("сброс") || decision.includes("не берет")) {
    const isProtectedFromAutoArchive = isWhaleSheet || (sheetName === "3. Лист ожидания");
    const match = history.match(/\[(?:📞\s*)?Недозвон:\s*(\d+)\/4(?:\s*-\s*[^\]]+)?\]/i);
    let attempts = match ? parseInt(match[1], 10) : 0;
    attempts += 1;

    const callTag = `[📞 Недозвон: ${attempts}/4 - ${nowStr}]`;
    let updatedHistory = match
      ? history.replace(/\[(?:📞\s*)?Недозвон:\s*\d+\/4(?:\s*-\s*[^\]]+)?\]/i, callTag)
      : prependHistoryLog_(history, callTag);

    if (attempts < 4 || isProtectedFromAutoArchive) {
      const stepDaysMap = { 1: 1, 2: 3, 3: 7 };
      const stepDays = stepDaysMap[attempts] || (sheetName === "3. Лист ожидания" ? 14 : 7);
      const rawNextCallDateObj = getNextWorkingDateObj_(now, stepDays);
      const nextCallDateObj    = getSmartBalancedDeadlineObj_(sheet, rawNextCallDateObj, timeZone, row);
      const nextCallStr        = Utilities.formatDate(nextCallDateObj, timeZone, "dd.MM.yyyy");

      sheet.getRange(row, 14).setValue(updatedHistory);
      setDeadlineCellDate_(sheet, row, nextCallDateObj, false);
      sheet.getRange(row, 13).clearContent();
      if (isSheetWithStockMonitor_(sheetName)) refreshSingleRowTactic_(sheet, row);

      moveRowToBottom_(sheet, row);
      if (isProtectedFromAutoArchive && attempts >= 4) {
        const protectTitle = isWhaleSheet ? "🐋 Защита Кита" : "⏳ Защита Ожидания";
        ss.toast(`Клиент НЕ удаляется в Архив при недозвоне (${attempts}/4)! Сдвинут вниз на ${nextCallStr} (шаг +${stepDays} дн.) — прогревочная база сохранена.`, protectTitle, 6);
      } else {
        ss.toast(`Недозвон (${attempts}/4, шаг +${stepDays} дн.). Дедлайн: ${nextCallStr}. Строка опущена вниз!`, "📞 Недозвон", 4);
      }
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

  // === ВЕТКА 5: ФУНДАМЕНТАЛЬНЫЕ ОТКАЗЫ / ЗАКРЫЛИСЬ / НЕ ЦЕЛЕВОЙ / ТОКСИЧНЫЙ ДЕМПИНГ -> в «4.Архив лидов» ===
  const isToxic     = decision.includes("токсич") || decision.includes("демпинг") || decision.includes("террор");
  const isNotTarget = decision.includes("не целев") || decision.includes("нецелев");
  const isClosed    = decision.includes("закрыл") || decision.includes("ликвидир") || decision.includes("банкрот");
  const isRefusal   = isClosed || decision.includes("отказ") || decision.includes("архив") || decision.includes("удалит") || decision.includes("слил");

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
    } else if (isClosed) {
      reasonText = "Отказ (Закрылись)";
      icon = "⛔";
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
  const liveDiams = extractCleanPipeDiams_(str).filter(d => stockByDiam[d]);
  cell.setBackground(liveDiams.length > 0 ? '#dcfce7' : null);
}

function cleanupRowIfCompletelyEmpty_(sheet, row, sheetName) {
  if (row < 2) return false;
  const vals = sheet.getRange(row, 1, 1, 14).getValues()[0];
  const hasAnyData = vals.some(v => v !== null && v !== undefined && String(v).trim() !== '');
  if (hasAnyData) return false;

  const hasMonitor = isSheetWithStockMonitor_(sheetName);
  const maxCol = hasMonitor ? 15 : 14;
  sheet.getRange(row, 1, 1, maxCol).setBackground(null).setFontColor('#0f172a').setFontWeight('normal');
  sheet.getRange(row, 2).setNote('');
  sheet.getRange(row, 12).setNumberFormat('dd.MM.yyyy');
  if (hasMonitor && sheet.getMaxColumns() >= 15) {
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
      if (isSheetWithStockMonitor_(sheetName)) {
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
  const cleanLine = String(latestText || '')
    .replace(/\n+/g, ' ')
    .replace(/^📌\s*(?:[^|]+\|\s*)?/, '')
    .trim();
  const shortPhrase = cleanLine.length > 95 ? cleanLine.slice(0, 92).trim() + '…' : cleanLine;
  return shortPhrase ? `📌 ${shortPhrase}` : '';
}


// =================================================================
// ЕДИНЫЕ ЧИСТЫЕ ВЫПАДАЮЩИЕ СПИСКИ (G: Сфера, I: Тип трубы, J: Под что берут, M: Решение)
// Строго по одному чистому названию на пункт (без склеек через «/»).
// ВНИМАНИЕ: Если вы добавите или удалите пункты прямо в Google Таблице (Данные -> Настроить проверку данных),
// скрипт НЕ будет затирать ваши новые пункты, а автоматически подхватит и сохранит их!
// =================================================================

const CRM_SPHERES_LIST_ = [
  "Строительство",
  "Бурение",
  "ГНБ",
  "Водоканал",
  "ЖКХ",
  "Теплосети",
  "Энергетика",
  "Газификация",
  "Дорожное строительство",
  "Мостостроение",
  "Гидротехника",
  "Фундаментостроение",
  "Карьер",
  "ГОК",
  "Дноуглубление",
  "Металлоконструкции",
  "Промышленность",
  "Машиностроение",
  "Агропром",
  "Нефтегаз",
  "Металлобаза",
  "Комплектация",
  "Другое"
];

const CRM_PIPE_TYPES_LIST_ = [
  "Новая",
  "Лежалая",
  "Восстановленная",
  "Б/у",
  "НКТ"
];

const CRM_APPLICATIONS_LIST_ = [
  "Сваи",
  "Шпунт",
  "Распорки",
  "Опоры",
  "Колонны",
  "Фермы",
  "Металлоконструкции",
  "Мачты",
  "Прокол",
  "Футляр",
  "Гильза",
  "Обсадка",
  "Кондуктор",
  "Водопровод",
  "Канализация",
  "Водоотведение",
  "Теплотрасса",
  "Паропровод",
  "Газопровод",
  "Пульпопровод",
  "Шламопровод",
  "Земснаряд",
  "Понтоны",
  "Причалы",
  "Водопропуск",
  "Дренаж",
  "Котлы",
  "Печи",
  "Ролики",
  "Барабаны",
  "Забор",
  "Ограждения",
  "Стойки",
  "Перепродажа",
  "Тендер"
];

const WAITING_DECISIONS_LIST_ = [
  "Горящая заявка (В просчет)",
  "В Ожидание (Нет потребности)",
  "Сдвиг спроса",
  "Купил у других",
  "Сезон: Весна",
  "Сезон: Лето",
  "Сезон: Осень",
  "Сезон: Зима",
  "Перезвонить сегодня",
  "Недозвон",
  "Проход секретаря (Выбить ЛПР)",
  "Отказ (Закрылись)",
  "Не целевой",
  "Токсичный демпинг"
];

/**
 * Аккуратно обновляет выпадающие списки, сохраняя любые новые пункты без «/», которые вы добавили вручную в Таблице!
 */
function updateListValidationPreservingStyle_(range, baseItems, preserveCustomCleanItems) {
  const existingRule = range.getCell(1, 1).getDataValidation();
  let finalItems = baseItems.slice();

  if (existingRule && existingRule.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    const args = existingRule.getCriteriaValues();
    const curList = (args && args[0]) ? args[0] : [];

    // Если вы сами добавили в Таблице новую чистую сферу или назначение (без старых склеек через «/») — бережно сохраняем её!
    if (preserveCustomCleanItems) {
      curList.forEach(item => {
        const clean = String(item || '').trim();
        if (clean && !clean.includes('/') && !clean.includes('(') && !finalItems.includes(clean)) {
          finalItems.push(clean);
        }
      });
    }

    const isSame = curList.length === finalItems.length && curList.every((v, idx) => v === finalItems[idx]);
    if (isSame) return;

    const showDrp = args && args.length > 1 ? Boolean(args[1]) : true;
    const rebuilt = existingRule.copy()
      .requireValueInList(finalItems, showDrp)
      .setAllowInvalid(true)
      .build();
    range.setDataValidation(rebuilt);
    return;
  }

  const freshRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(finalItems, true)
    .setAllowInvalid(true)
    .build();
  range.setDataValidation(freshRule);
}

function setupCleanDropdownsOnSheetSilent_(sheet) {
  if (!sheet) return;
  const sheetName = sheet.getName().trim();
  const crmSheets = ["1. Обзвон (База)", "2. Золотой список (Киты)", "3. Лист ожидания"];
  if (!crmSheets.includes(sheetName)) return;

  const maxRows = Math.max(sheet.getLastRow(), 50);
  const numRows = maxRows - 1;
  if (numRows < 1) return;

  updateListValidationPreservingStyle_(sheet.getRange(2, 7, numRows, 1), CRM_SPHERES_LIST_, true);
  updateListValidationPreservingStyle_(sheet.getRange(2, 9, numRows, 1), CRM_PIPE_TYPES_LIST_, true);
  updateListValidationPreservingStyle_(sheet.getRange(2, 10, numRows, 1), CRM_APPLICATIONS_LIST_, true);

  // В «1. Обзвон (База)» колонку M (Решение) НЕ трогаем (там уже идеальный список)!
  if (sheetName === "3. Лист ожидания") {
    updateListValidationPreservingStyle_(sheet.getRange(2, 13, numRows, 1), WAITING_DECISIONS_LIST_, false);
  }
}

function syncUnpackedNotesOnSheetSilent_(sheet) {
  try { setupCleanDropdownsOnSheetSilent_(sheet); } catch (e) {}
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const numRows = lastRow - 1;
  const kVals  = sheet.getRange(2, 11, numRows, 1).getValues();
  const bNotes = sheet.getRange(2, 2, numRows, 1).getNotes();

  for (let i = 0; i < numRows; i++) {
    const rawK = String(kVals[i][0] || '').trim();
    const noteB = String(bNotes[i][0] || '').trim();
    if (rawK && (!noteB || !rawK.startsWith('📌') || rawK.includes('|'))) {
      processRawNoteIntoDossier_(sheet, i + 2);
    }
  }
}

// =================================================================
// ЧИСТЫЙ МОНИТОР СКЛАДА И ПРАЙСА (КОЛОНКА O В «3. ЛИСТ ОЖИДАНИЯ» И «2. КИТЫ»)
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

      const qtyStr = stockTn > 0 ? `${Number(stockTn.toFixed(1))}т` : `${Math.round(stockM)}м`;
      const prStr  = priceTn > 0 ? ` — ${priceTn.toLocaleString('ru-RU')} ₽` : '';
      stockByDiam[diam].push(`${diam}х${wall} ${condShort} (${qtyStr}${prStr})`);
    }
  });
  return stockByDiam;
}

function extractCleanPipeDiams_(text) {
  if (!text) return [];
  const withoutDates = String(text)
    .replace(/\b\d{1,2}[\.\-\/]\d{1,2}[\.\-\/]\d{2,4}\b/g, ' ')
    .replace(/\b\d{1,2}:\d{2}\b/g, ' ');
  return Array.from(new Set((withoutDates.match(/\b\d{2,4}\b/g) || []).map(Number)))
    .filter(d => STANDARD_PIPE_DIAMS_.includes(d));
}

function computeDynamicTacticForRow_(rowVals, stockByDiam, companyDossierNote) {
  const need    = String(rowVals[7] || '').trim();
  const notes   = String(rowVals[10] || '').trim();
  const history = String(rowVals[13] || '').trim();
  const dossier = String(companyDossierNote || '').trim();

  const allContext = `${notes} ${dossier} ${history}`;
  const allLower   = allContext.toLowerCase();

  const needLower = need.toLowerCase();
  const hasActiveNeed = need !== '' && needLower !== 'нет' && needLower !== '—' && needLower !== '-';
  const firstHistoryLine = (history.split('\n')[0] || '').trim();

  const isScaleClient = /весы|взвеш|микрометр|толщиномер|стенкомер|хитро/i.test(allLower);
  const scaleTag = isScaleClient ? ' [⚖️ Весы]' : '';

  // 1. Выставлен счет
  if (firstHistoryLine.includes('💰') || firstHistoryLine.toLowerCase().includes('выставлен счет')) {
    const targetPipe = hasActiveNeed ? `: ${need}` : '';
    return `💰 Выставлен счет${targetPipe}${scaleTag}`;
  }

  // 2. Живая заявка в H — сверка со Складом
  if (hasActiveNeed) {
    const liveDiams = extractCleanPipeDiams_(need);
    const matches = [];
    liveDiams.forEach(d => {
      if (stockByDiam[d]) matches.push(stockByDiam[d].join(', '));
    });

    if (matches.length > 0) {
      return `🟢 Под заявку на складе: ${matches.join(' | ')}${scaleTag}`;
    }
    return `⚪ Нет на складе (заявка: ${need})${scaleTag}`;
  }

  // 3. В H «Нет» — сверка прошлых диаметров из Истории (N) и Досье (B/K) со Складом
  const pastDiams = extractCleanPipeDiams_(allContext);
  if (pastDiams.length > 0) {
    const stockHits = [];
    pastDiams.forEach(d => {
      if (stockByDiam[d]) stockHits.push(stockByDiam[d][0]);
    });

    if (stockHits.length > 0) {
      return `🟢 На складе их Ø: ${stockHits.slice(0, 3).join('; ')}${scaleTag}`;
    }
    return `⚪ На складе нет (профиль: Ø ${pastDiams.slice(0, 5).join(', ')})${scaleTag}`;
  }

  // 4. Диаметры еще не известны — сухой статус по Прайсу
  const kpSent = /📧|💬|📄|Прайс|КП/i.test(history);
  return (kpSent ? `✅ Прайс отправлен` : `📧 Прайс не высылался`) + scaleTag;
}

function refreshSingleRowTactic_(sheet, row) {
  if (!sheet || !isSheetWithStockMonitor_(sheet.getName()) || row < 2) return;
  if (sheet.getMaxColumns() < 15) return;
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
// ДВУХУРОВНЕВЫЙ БИРЖЕВОЙ РАДАР (ДЛЯ «3. ЛИСТ ОЖИДАНИЯ» И «2. КИТЫ»)
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

      tacticsCol.push([computeDynamicTacticForRow_(r, stockByDiam, dossier)]);

      const isNeedActive = needStr !== '' && needStr.toLowerCase() !== 'нет' && needStr !== '—' && needStr !== '-';

      if (isNeedActive) {
        const liveDiams = extractCleanPipeDiams_(needStr).filter(d => stockByDiam[d]);
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
      const pastDiams = extractCleanPipeDiams_(combinedPast).filter(d => stockByDiam[d]);

      if (pastDiams.length > 0) {
        historyMatches.push(
          `💡 [${sName.replace(/^\d+\.\s*/, '')}] ${comp || 'Клиент'} (${lpr || ph}): в профиле Ø ${pastDiams.join(', ')}`
        );
      }
    });

    sh.getRange(2, 8, lRow - 1, 1).setBackgrounds(needBgColors);
    if (sh.getMaxColumns() >= 15 && tacticsCol.length > 0) {
      sh.getRange(2, 15, lRow - 1, 1).setValues(tacticsCol);
    }
  });

  const allMatches = hotMatches.concat(historyMatches);
  if (allMatches.length > 0) {
    const preview = allMatches.slice(0, 14).join('\n');
    const more = allMatches.length > 14 ? `\n...и ещё ${allMatches.length - 14} совпадений!` : '';
    SpreadsheetApp.getUi().alert(
      `🎯 Биржевой радар: Горячих в H (${hotMatches.length}) | По профилю (${historyMatches.length})`,
      `Монитор Склада (Колонка O) обновлен в Ожидании и Китах:\n\n${preview}${more}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    ss.toast('Монитор Склада (Колонка O) обновлен!', '🔍 Радар обновлен', 5);
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

function shiftWeekendToMonday_(dateObj) {
  if (!(dateObj instanceof Date) || isNaN(dateObj.getTime())) return dateObj;
  const d = new Date(dateObj.getTime());
  const day = d.getDay();
  if (day === 6) d.setDate(d.getDate() + 2);
  else if (day === 0) d.setDate(d.getDate() + 1);
  return d;
}

function getNextWorkingDateObj_(baseDate, addDays) {
  const nextDate = new Date(baseDate.getTime());
  nextDate.setHours(0, 0, 0, 0);
  nextDate.setDate(nextDate.getDate() + addDays);
  return shiftWeekendToMonday_(nextDate);
}

/**
 * Определяет стартовую рабочую дату для сезонов: Весна (1 марта), Лето (1 июня), Осень (1 сентября), Зима (1 декабря)
 */
function getSeasonalStartDateObj_(baseDate, decisionLower) {
  const y = baseDate.getFullYear();
  const m = baseDate.getMonth(); // 0..11
  let targetMonth = 2; // По умолчанию Весна (Март = 2)
  let label = "🌱 Сезон: Весна";

  if (decisionLower.includes("лет")) {
    targetMonth = 5; // Июнь
    label = "☀️ Сезон: Лето";
  } else if (decisionLower.includes("осен")) {
    targetMonth = 8; // Сентябрь
    label = "🍂 Сезон: Осень";
  } else if (decisionLower.includes("зим")) {
    targetMonth = 11; // Декабрь
    label = "❄️ Сезон: Зима";
  } else if (decisionLower.includes("весн")) {
    targetMonth = 2; // Март
    label = "🌱 Сезон: Весна";
  } else {
    // Если просто «Не сезон / В след. году» — ставим на 1 марта следующего сезона
    targetMonth = 2;
    label = "🌱 Не сезон (Весна)";
  }

  let targetYear = y;
  const candidate = new Date(targetYear, targetMonth, 1, 0, 0, 0);
  // Если 1-е число выбранного сезона уже прошло или наступит меньше чем через 10 дней — берем следующий год
  if (candidate.getTime() <= baseDate.getTime() + 10 * 24 * 60 * 60 * 1000) {
    targetYear += 1;
  }

  const finalSeasonDate = shiftWeekendToMonday_(new Date(targetYear, targetMonth, 1, 0, 0, 0));
  return { dateObj: finalSeasonDate, label: label };
}

/**
 * Умный балансировщик дедлайнов:
 * 1) Сдвигает Сб/Вс на Понедельник.
 * 2) Проверяет загрузку выбранной даты на листе: если на этот день уже назначено >= MAX_CALLS_PER_DAY_ (10) компаний,
 *    автоматически шагает на следующий рабочий день, чтобы не сваливать всех клиентов в один день!
 */
function getSmartBalancedDeadlineObj_(sheet, targetDateObj, timeZone, excludeRow, customMaxPerDay) {
  let candidate = shiftWeekendToMonday_(targetDateObj);
  if (!sheet || !(candidate instanceof Date) || isNaN(candidate.getTime())) return candidate;

  const tz = timeZone || "Europe/Moscow";
  const maxPerDay = customMaxPerDay || MAX_CALLS_PER_DAY_;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return candidate;

  const dlValues = sheet.getRange(2, 12, lastRow - 1, 1).getValues();
  const countByIso = {};

  for (let i = 0; i < dlValues.length; i++) {
    const rNum = i + 2;
    if (excludeRow && rNum === excludeRow) continue;
    const cellVal = dlValues[i][0];
    if (!cellVal) continue;
    const parsed = parseDeadlineObj_(cellVal, tz);
    if (parsed.isoDate) {
      countByIso[parsed.isoDate] = (countByIso[parsed.isoDate] || 0) + 1;
    }
  }

  let safety = 0;
  while (safety < 60) {
    const iso = Utilities.formatDate(candidate, tz, "yyyy-MM-dd");
    if ((countByIso[iso] || 0) < maxPerDay) {
      return candidate;
    }
    candidate = new Date(candidate.getTime());
    candidate.setDate(candidate.getDate() + 1);
    candidate = shiftWeekendToMonday_(candidate);
    safety++;
  }
  return candidate;
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
      shiftWeekendToMonday_(rowValues[11] || oldRow[11]),
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
  const safeDateObj = shiftWeekendToMonday_(dateObj);
  cell.setNumberFormat(hasTime ? "dd.MM.yyyy HH:mm" : "dd.MM.yyyy");
  cell.setValue(safeDateObj);
  updateSingleRowDeadlineHighlight_(sheet, row);
}

function parseSmartDeadlineInput_(cellVal, rawEventVal, rawOldVal, timeZone) {
  const rawStr = rawEventVal !== undefined ? String(rawEventVal).trim() : "";
  const now = new Date();
  const tz = timeZone || "Europe/Moscow";

  const hourOnlyMatch = rawStr.match(/^([01]?\d|2[0-3])$/);
  if (hourOnlyMatch) {
    const hh = parseInt(hourOnlyMatch[1], 10);
    return parseDeadlineObj_(new Date(now.getFullYear(), now.getMonth(), now.getDate(), hh, 0, 0), tz);
  }

  const fourDigitTime = rawStr.match(/^([01]?\d|2[0-3])([0-5]\d)$/);
  if (fourDigitTime) {
    const hh = parseInt(fourDigitTime[1], 10);
    const mm = parseInt(fourDigitTime[2], 10);
    return parseDeadlineObj_(new Date(now.getFullYear(), now.getMonth(), now.getDate(), hh, mm, 0), tz);
  }

  const timeMatch = rawStr.match(/^([01]?\d|2[0-3])[:\.\s]+([0-5]\d)$/);
  if (timeMatch) {
    const hh = parseInt(timeMatch[1], 10);
    const mm = parseInt(timeMatch[2], 10);
    return parseDeadlineObj_(new Date(now.getFullYear(), now.getMonth(), now.getDate(), hh, mm, 0), tz);
  }

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
  const hasMonitor = isSheetWithStockMonitor_(sheetName) && sheet.getMaxColumns() >= 15;
  const numCols = hasMonitor ? 15 : 14;
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

    if (hasMonitor) {
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
    if (hasMonitor && tacticsChanged) {
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
      const liveDiams = extractCleanPipeDiams_(needStr).filter(d => stockByDiam[d]);
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