/**
 * ЕДИНЫЙ ЦЕНТР УПРАВЛЕНИЯ ПРОДАЖАМИ И СКЛАДОМ (CTRM «Капитал-Плюс» v6.6)
 * Файл: Calling.js — Полный боевой движок для листов:
 * «1. Обзвон (База)», «2. Золотой список (Киты)», «3. Лист ожидания», «4.Архив лидов».
 * - Монитор Склада в колонке O (Вариант 1 — только сухие факты остатков и цен Склада) для Ожидания и Китов.
 * - VIP-логика и 2-этажная архитектура для «2. Золотой список (Киты)»:
 *   Этаж 1 (с телефонами) всегда наверху в активной ротации, Этаж 2 (из 1С без телефонов) ниже по году (2026->2015) и выручке.
 * - Колонка S «📦 Что брали в 1С (КП / РЛ)» хранит историю УПД обеих фирм, не засоряя Историю звонков (N) и Примечания (K).
 * - Быстрый ввод времени в L (14 / 1430 / 14 30) + выпадающий календарь + авто-подъем очереди при обновлении.
 */

/**
 * 🚀 ОДНОКРАТНАЯ АВТО-СБОРКА И ОЧИСТКА БАЗЫ КИТОВ ИЗ ДВУХ 1С («КАПИТАЛ-ПЛЮС» + «РЕГИОН-ЛИДЕР»)
 * Запускается АВТОМАТИЧЕСКИ при обновлении страницы Таблицы (onOpen) или кнопкой «Выполнить» в редакторе Apps Script!
 */
function rebuildWhalesAndCleanCrmFrom1C() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  autoMigrateAndCleanWhalesV67_(ss, true);
}

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

function isWhaleSheetName_(sheetName) {
  return String(sheetName || '').trim() === "2. Золотой список (Киты)";
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
  // 2.95 АВТО-РАСЧЕТ СЕГМЕНТА A / B / C ПРИ ВВОДЕ ТОННАЖА (Q / 17) ИЛИ ВЫРУЧКИ (R / 18) В КИТАХ
  // =================================================================
  if (isWhaleSheetName_(sheetName) && (col === 16 || col === 17 || col === 18)) {
    if (col === 16) {
      sheet.getParent().toast(
        "Сегмент A / B / C рассчитывается автоматически от средней выручки и тоннажа по листу «Киты»! Заполните колонки Q (Объем, тн) или R (Выручка, ₽).",
        "🤖 Авто-сегментация ABC",
        5
      );
    } else {
      const numCell = sheet.getRange(row, col);
      const parsedNum = parseWhaleMetricNumber_(numCell.getValue());
      if (parsedNum > 0) {
        numCell.setValue(parsedNum);
        numCell.setNumberFormat(col === 18 ? '#,##0 "₽"' : '#,##0.##');
      } else {
        numCell.clearContent();
      }
    }
    const stats = recalculateWhaleSegmentsOnSheet_(sheet);
    if (stats && stats.filledCount > 0) {
      const rowSeg = String(sheet.getRange(row, 16).getValue() || '—');
      const avgRevStr = stats.avgRev > 0 ? `${Math.round(stats.avgRev).toLocaleString('ru-RU')} ₽` : '—';
      sheet.getParent().toast(
        `Присвоен сегмент: ${rowSeg} | Средняя выручка по базе Китов (${stats.filledCount} комп.): ${avgRevStr}`,
        "🏆 Сегмент ABC обновлен",
        5
      );
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
    sheet.getRange(row, 13).clearContent();
    ss.toast(
      "Статус «Выставлен счет» ставится и снимается автоматически через модуль «Заказы»! Для живого запроса используйте «Горящая заявка (В просчет)».",
      "💡 Авто-статус модуля Заказы",
      5
    );
    return;
  }

  // === ВЕТКА 1В (ИСКЛЮЧЕНИЕ ДЛЯ КИТОВ): РАЗОВАЯ ПОКУПКА -> РАЗЖАЛОВАТЬ В «3. ЛИСТ ОЖИДАНИЯ» ===
  if (decision.includes("разов")) {
    const targetSheet = getSheetByNameSafe(ss, "3. Лист ожидания");
    if (!targetSheet) return ss.toast("Ошибка: лист Ожидания не найден", "Ошибка");

    const todayIso = Utilities.formatDate(now, timeZone, "yyyy-MM-dd");
    const useExisting = (existingDlObj.dateObj && existingIsoDeadline > todayIso);
    const rawCandidateDate = useExisting ? existingDlObj.dateObj : getNextWorkingDateObj_(now, 30);
    const targetHasTime    = useExisting ? existingDlObj.hasTime : false;
    const targetDateObj    = targetHasTime
      ? shiftWeekendToMonday_(rawCandidateDate)
      : getSmartBalancedDeadlineObj_(targetSheet, rawCandidateDate, timeZone, -1);
    const targetDisp       = Utilities.formatDate(targetDateObj, timeZone, targetHasTime ? "dd.MM.yyyy HH:mm" : "dd.MM.yyyy");

    const demoteTag = `[📉 Из Китов в Ожидание: Разовая покупка${needSnapshot} (след. контакт: ${targetDisp}) - ${nowStr}]`;
    const updatedHistory = prependHistoryLog_(history, demoteTag);

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
    if (isWhaleSheet) recalculateWhaleSegmentsOnSheet_(sheet);
    ss.toast(
      `«${company || inn}» переведён из Китов в «3. Лист ожидания» (Разовая покупка). Следующий контакт: ${targetDisp}.`,
      "📉 Перевод в Ожидание",
      5
    );
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

  // === ВЕТКА 3В: В ОЖИДАНИЕ / НА КОНТРОЛЕ (НЕТ ПОТРЕБНОСТИ) ===
  // Кит остается в Китах (шаг по сегменту: A = +7 дн., B = +14 дн., C = +21 дн.), Ожидание остается в Ожидании (+14 дн.), из Обзвона переходит в Ожидание!
  if (decision.includes("ожидан") || decision.includes("на контроле") || decision.includes("контрол") || decision.includes("подума") || decision.includes("спящ") || decision.includes("нет потребн")) {
    const defaultWaitDays = isWhaleSheet ? getWhaleFollowUpDays_(sheet, row) : 14;
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
    if (isWhaleSheet) recalculateWhaleSegmentsOnSheet_(sheet);
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

  const isWhale = isWhaleSheetName_(sheetName);
  const hasMonitor = isSheetWithStockMonitor_(sheetName);
  const maxCol = isWhale ? Math.min(sheet.getMaxColumns(), 19) : (hasMonitor ? 15 : 14);
  sheet.getRange(row, 1, 1, maxCol).setBackground(null).setFontColor('#0f172a').setFontWeight('normal');
  sheet.getRange(row, 2).setNote('');
  sheet.getRange(row, 12).setNumberFormat('dd.MM.yyyy');
  if (hasMonitor && sheet.getMaxColumns() >= 15) {
    sheet.getRange(row, 15).clearContent();
  }
  if (isWhale && sheet.getMaxColumns() >= 18) {
    sheet.getRange(row, 16, 1, Math.min(4, sheet.getMaxColumns() - 15)).clearContent();
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

  if (isWhaleSheetName_(sheetName)) {
    recalculateWhaleSegmentsOnSheet_(sheet);
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

const WHALES_DECISIONS_LIST_ = [
  "Горящая заявка (В просчет)",
  "На контроле (Нет потребности)",
  "Сдвиг спроса",
  "Купил у других",
  "Сезон: Весна",
  "Сезон: Лето",
  "Сезон: Осень",
  "Сезон: Зима",
  "Перезвонить сегодня",
  "Недозвон",
  "Разовая покупка (В Ожидание)",
  "Отказ (Закрылись)",
  "Не целевой"
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
  } else if (sheetName === "2. Золотой список (Киты)") {
    updateListValidationPreservingStyle_(sheet.getRange(2, 13, numRows, 1), WHALES_DECISIONS_LIST_, false);
    ensureWhaleColumnsAndSegments_(sheet);
  }
}

// =================================================================
// СПЕЦ-КОЛОНКИ КИТОВ (P: Сегмент, Q: Объем тн, R: Выручка ₽) И АВТО-РАСЧЕТ ABC ОТ СРЕДНЕГО
// Почва под будущий лист «6. Заказы» + мгновенная сегментация без ручного гадания
// =================================================================

function parseWhaleMetricNumber_(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  const cleaned = String(val)
    .replace(/\s+/g, '')
    .replace(/₽|руб\.?|тн|тонн|т/gi, '')
    .replace(',', '.')
    .replace(/[^0-9.\-]/g, '');
  const n = parseFloat(cleaned);
  return isNaN(n) || n <= 0 ? 0 : n;
}

function ensureWhaleColumnsAndSegments_(sheet) {
  if (!sheet || !isWhaleSheetName_(sheet.getName())) return;
  const maxCols = sheet.getMaxColumns();
  if (maxCols < 19) {
    sheet.insertColumnsAfter(maxCols, 19 - maxCols);
  }

  const headers = sheet.getRange(1, 16, 1, 4).getValues()[0];
  const expectedHeaders = ["🏆 Сегмент (A/B/C)", "⚖️ Объем (тн)", "💰 Выручка (₽)", "📦 Что брали в 1С (КП / РЛ)"];
  const needsHeaderInit = headers.some((h, idx) => String(h || '').trim() !== expectedHeaders[idx]);

  if (needsHeaderInit) {
    const styleSource = sheet.getRange(1, Math.min(15, maxCols), 1, 1);
    const targetHeaderRange = sheet.getRange(1, 16, 1, 4);
    styleSource.copyTo(targetHeaderRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    targetHeaderRange.setValues([expectedHeaders]);
    targetHeaderRange.setHorizontalAlignment("center").setVerticalAlignment("middle");
    sheet.setColumnWidth(16, 155);
    sheet.setColumnWidth(17, 130);
    sheet.setColumnWidth(18, 155);
    sheet.setColumnWidth(19, 360);
  }

  autoMigrateAndCleanWhalesV67_(sheet.getParent(), false);
  recalculateWhaleSegmentsOnSheet_(sheet);
}

function recalculateWhaleSegmentsOnSheet_(sheet) {
  if (!sheet || !isWhaleSheetName_(sheet.getName())) return null;
  if (sheet.getMaxColumns() < 18) return null;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { filledCount: 0, avgRev: 0, avgTn: 0 };

  const numRows = lastRow - 1;
  const data = sheet.getRange(2, 1, numRows, 18).getValues();

  let sumRev = 0;
  let cntRev = 0;
  let sumTn  = 0;
  let cntTn  = 0;

  const parsedRows = data.map(r => {
    const hasCompany = String(r[0] || '').trim() !== '' || String(r[1] || '').trim() !== '';
    const tn  = parseWhaleMetricNumber_(r[16]);
    const rev = parseWhaleMetricNumber_(r[17]);
    if (hasCompany && rev > 0) {
      sumRev += rev;
      cntRev++;
    }
    if (hasCompany && tn > 0) {
      sumTn += tn;
      cntTn++;
    }
    return { hasCompany, tn, rev, curSeg: String(r[15] || '').trim() };
  });

  const avgRev = cntRev > 0 ? (sumRev / cntRev) : 0;
  const avgTn  = cntTn > 0 ? (sumTn / cntTn) : 0;

  const segValues  = [];
  const segBgs     = [];
  const segColors  = [];
  const segWeights = [];
  let changed = false;

  parsedRows.forEach(item => {
    if (!item.hasCompany) {
      segValues.push([""]);
      segBgs.push([null]);
      segColors.push(["#0f172a"]);
      segWeights.push(["normal"]);
      if (item.curSeg !== "") changed = true;
      return;
    }
    if (item.rev <= 0 && item.tn <= 0) {
      const noShipLabel = "⚪ Без отгрузок";
      segValues.push([noShipLabel]);
      segBgs.push(["#f8fafc"]);
      segColors.push(["#64748b"]);
      segWeights.push(["normal"]);
      if (item.curSeg !== noShipLabel) changed = true;
      return;
    }

    // Отношение к среднему по заполненным Китам (приоритет у Выручки ₽, если она заполнена; иначе по Тоннажу)
    let ratio = 1;
    if (item.rev > 0 && avgRev > 0) {
      ratio = cntRev === 1 ? (item.rev >= 1500000 ? 1.3 : (item.rev >= 500000 ? 0.8 : 0.3)) : (item.rev / avgRev);
    } else if (item.tn > 0 && avgTn > 0) {
      ratio = cntTn === 1 ? (item.tn >= 20 ? 1.3 : (item.tn >= 5 ? 0.8 : 0.3)) : (item.tn / avgTn);
    }

    let segLabel = "⭐ B (Постоянный)";
    let bg = "#e0f2fe";
    let fg = "#075985";
    let fw = "bold";

    if (ratio >= 1.20) {
      segLabel = "🔥 A (VIP)";
      bg = "#fef3c7";
      fg = "#92400e";
      fw = "bold";
    } else if (ratio < 0.45) {
      segLabel = "🔹 C (Малый)";
      bg = "#f1f5f9";
      fg = "#334155";
      fw = "normal";
    }

    if (item.curSeg !== segLabel) changed = true;
    segValues.push([segLabel]);
    segBgs.push([bg]);
    segColors.push([fg]);
    segWeights.push([fw]);
  });

  const segRange = sheet.getRange(2, 16, numRows, 1);
  if (changed) {
    segRange.setValues(segValues);
  }
  segRange.setBackgrounds(segBgs).setFontColors(segColors).setFontWeights(segWeights).setHorizontalAlignment("center");
  sheet.getRange(2, 17, numRows, 1).setNumberFormat('#,##0.##');
  sheet.getRange(2, 18, numRows, 1).setNumberFormat('#,##0 "₽"');

  return { filledCount: Math.max(cntRev, cntTn), avgRev, avgTn };
}

function getWhaleFollowUpDays_(sheet, row) {
  if (!sheet || sheet.getMaxColumns() < 16) return 10;
  const seg = String(sheet.getRange(row, 16).getValue() || '').toUpperCase();
  if (seg.includes('A')) return 7;
  if (seg.includes('B')) return 14;
  if (seg.includes('C')) return 21;
  return 10;
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

  // 3. Если в H пусто или «Нет» — показываем только чистый статус по Прайсу (без свалки прошлых труб в O)
  // Все исторические трубы лежат в Колонке S (1С) и в Истории (N)
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
  const isWhale = isWhaleSheetName_(sheetName);
  if (isWhale) {
    try { ensureWhaleColumnsAndSegments_(sheet); } catch (e) {}
  }
  const hasMonitor = isSheetWithStockMonitor_(sheetName) && sheet.getMaxColumns() >= 15;
  const numCols = isWhale ? Math.max(19, sheet.getLastColumn()) : (hasMonitor ? 15 : 14);
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

    const hasPhone = String(r[4] || "").replace(/\D/g, "").length >= 6;
    const sInfo = String(r[18] || "");
    let lastYear = 0;
    const posMatch = sInfo.match(/посл\.\s*\d{2}\.\d{2}\.(\d{4})/);
    if (posMatch) {
      lastYear = parseInt(posMatch[1], 10);
    } else {
      const yMatches = sInfo.match(/20(?:1\d|2\d)/g);
      if (yMatches && yMatches.length > 0) {
        lastYear = Math.max(...yMatches.map(Number));
      }
    }

    items.push({
      origIdx: i,
      isEmpty: false,
      hasPhone: hasPhone,
      lastYear: lastYear,
      bucket: bucket,
      dStr: dlObj.isoDate || "",
      tStr: effectiveTime || "",
      dlObj: dlObj,
      values: r,
      notes: n
    });
  }

  items.sort((a, b) => {
    if (a.isEmpty !== b.isEmpty) return a.isEmpty ? 1 : -1;
    if (a.isEmpty && b.isEmpty) return a.origIdx - b.origIdx;

    // 2-этажная архитектура на листе «2. Золотой список (Киты)»:
    // Этаж 1 (с телефонами) всегда выше Этажа 2 (без телефонов)
    if (isWhale) {
      if (a.hasPhone !== b.hasPhone) return a.hasPhone ? -1 : 1;
      if (!a.hasPhone && !b.hasPhone) {
        if (a.lastYear !== b.lastYear) return b.lastYear - a.lastYear;
        const revA = parseWhaleMetricNumber_(a.values[17]);
        const revB = parseWhaleMetricNumber_(b.values[17]);
        if (revA !== revB) return revB - revA;
        return a.origIdx - b.origIdx;
      }
    }

    if (a.bucket !== b.bucket) return a.bucket - b.bucket;
    if (a.bucket === 1 || a.bucket === 5) {
      if (a.dStr !== b.dStr) return a.dStr.localeCompare(b.dStr);
    }
    if ((a.bucket === 2 || a.bucket === 4) && a.tStr !== b.tStr) {
      return (a.tStr || "00:00").localeCompare(b.tStr || "00:00");
    }
    if (isWhale) {
      const freshA = a.lastYear >= 2021 ? 1 : (a.lastYear > 0 ? 0 : -1);
      const freshB = b.lastYear >= 2021 ? 1 : (b.lastYear > 0 ? 0 : -1);
      if (freshA !== freshB) return freshB - freshA;
      const revA = parseWhaleMetricNumber_(a.values[17]);
      const revB = parseWhaleMetricNumber_(b.values[17]);
      if (revA !== revB) return revB - revA;
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

  if (isWhale) {
    try { recalculateWhaleSegmentsOnSheet_(sheet); } catch (e) {}
  }

  highlightDuplicateInnsOnSheet_(sheet);
}

// =================================================================
// АВТО-СБОРКА И ОЧИСТКА «2. ЗОЛОТОЙ СПИСОК (КИТЫ)» (ВАРИАНТ 1 + КОЛОНКА S, 2 ФИРМЫ: КП + РЛ)
// =================================================================

const WHALES_2FIRMS_MASTER_ROWS_ = [["3123101212", "ООО \"АВТОМОСТ\"", "308009, Белгородская область, г. Белгород, ул. Рабочая, д. 10", "Вадим Александрович\nЕвгений", "+7 (915) 520-95-00\n+7 (960) 635-40-20", "avtomostpto@mail.ru\navtomostbel@gmail.com", "Мостостроение", "Нет", "", "", "📌 не финансируют как 3 месяца на23.07.26", "09.11.2026", "", "", "📧 Прайс не высылался", "🔥 A (VIP)", 780.17, 67708357, "КП+РЛ (2018–2023, посл. 21.12.2023, 41 УПД): 1220х14 (207.8т), 530х9 (143.7т), 530х8-9 (81.0т), Шпунт Ларсена GU22N б/у (77.3т), 530х8 (70.1т), 630х10 (63.7т)"], ["2308087914", "ООО МТУ \"ЮКС\"", "119454, г. Москва, вн. тер. г. муниципальный округ Проспект Вернадского, ул. Удальцова, д. 85А, кв. 116", "Алексей Николаевич", "+7 (952) 820-04-46", "uks@yugkomstroy.ru, cheftelova@yugkomstroy.ru", "", "Нет", "", "", "", "12.11.2026", "", "[⏳ В Ожидание - 29.10.2026 15:12]", "📧 Прайс не высылался", "🔥 A (VIP)", 1029.07, 47191193, "КП+РЛ (2017–2022, посл. 07.07.2022, 59 УПД): 530х12 (258.4т), 720х10 (199.4т), 1020х12 (122.4т), 820х11 (67.2т), 530х10 (66.5т), 720х9,5 (63.3т)"], ["3663033415", "ООО ТСП \"ВОРОНЕЖ СТРОЙ КОМПЛЕКС\"", "396310, Воронежская область, м. р-н Новоусманский, с. п. Усманское 1-Е, с. Новая Усмань, ул. Ленина, д. 275, помещ. 15", "Сергей\nЕлена\nВиктория", "+7 (908) 142-71-65\n+7 (905) 051-37-37\n+7 (915) 547-50-93", "snab@vsk36.ru", "Строительство", "", "", "", "", "21.10.2026", "", "", "📧 Прайс не высылался", "🔥 A (VIP)", 399.64, 41682161, "КП (2022–2026, посл. 01.07.2026, 31 УПД): 920х10 (144.2т), 820х10 (52.2т), 720х10 (50.3т), 630х10 (45.6т), 920х9 (35.7т), 325х7 (21.9т)"], ["3663101785", "ООО \"ГИДРОСПЕЦФУНДАМЕНТСТРОЙ\"", "394028, Воронежская область, г. Воронеж, ул. Димитрова, д. 148, неж.пом. 1 а", "Сергей\nАлексей Анатольевич\nГлавный инженер Сергей Петрович", "+7 (908) 135-11-87", "gsfs@yandex.ru", "Фундаментостроение", "377х6 200 м", "Восстановленная", "Сваи", "", "", "", "", "🟢 Под заявку на складе: 377х9 Леж. (2т — 68 000 ₽), 377х11 Леж. (1.2т — 68 000 ₽)", "🔥 A (VIP)", 1408.8, 32860987, "КП+РЛ (2015–2024, посл. 18.04.2024, 41 УПД): 630х8 (589.1т), Шпилька ст М 24*600 (212.0т), 820х10 (86.5т), 168х12-6000 б/у (86.0т), 530х8 (74.3т), 1020х10 (73.2т)"], ["2372022625", "ООО \"КУБАНЬ СТРОЙ МОНТАЖ\"", "352272, Краснодарский край, Отрадненский район, ст-ца Удобная, ул. Родниковская, д. 16А, офис 1", "Андрей Александрович", "+7 (928) 444-42-01\n+7 (967) 532-73-98", "karpov.a.a@bk.ru", "Другое", "", "", "", "📌 Под вышки билайна,мегафона ,т2", "02.11.2026", "", "", "📧 Прайс не высылался", "🔥 A (VIP)", 317.31, 27034262, "РЛ (2022–2025, посл. 28.03.2025, 18 УПД): 720х10 (163.4т), 720х9 (67.9т), 530х10,0-11 (37.0т), 530х11 (17.9т), 530х8 (11.6т), 820х11 (4.9т)"], ["3616017631", "ООО \"АКВАСТРОЙ\"", "396311, Воронежская область, Новоусманский район, с. Новая Усмань, ул. Ленина, д. 60/2, офис 3", "Сергей Сергеевич", "+7 (910) 749-12-10", "T57666@yandex.ru", "Бурение на воду", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔥 A (VIP)", 431.09, 23464969, "КП+РЛ (2015–2026, посл. 06.05.2026, 88 УПД): 530х8 (89.1т), 630х7 (54.5т), 426х8 (50.4т), 426х9 (34.7т), 219х7 (31.1т), 630х8 (23.1т)"], ["3662168794", "ООО \"СТРОЙ ИНВЕСТ\"", "394026, Воронежская область, г. о. город Воронеж, г. Воронеж, ул. Текстильщиков, д. 2А, офис 203", "Олег Иванович\nВиктория Викторовна", "+7 (906) 670-22-33\n+7 (960) 101-21-20", "stroyinvest23@mail.ru", "Водоканал / ЖКХ, Бурение / ГНБ", "Нет", "", "", "📌 Пластик", "12.11.2026", "", "[⏳ В Ожидание - 29.10.2026 13:55]", "📧 Прайс не высылался", "🔥 A (VIP)", 371.9, 17576579, "КП+РЛ (2016–2022, посл. 25.11.2022, 32 УПД): 426х10 (52.8т), 219х8 (38.3т), 820х9,5 (23.6т), 530х9 (23.3т), 530х8 (23.0т), 530х12 (20.8т)"], ["4629045050", "АО \"КОРПОРАЦИЯ \"ГРИНН\"", "241037, Брянская область, г. о. город Брянск, г. Брянск, ул. Брянского Фронта, д. 2", "Царьков Александр Фёдорович", "+7 (910) 740-11-52", "Zarkov.AF@grinn-corp.ru", "Агропром", "Нет", "", "", "📌 Потребности сейчас нет, материалы используют.", "09.11.2026", "", "", "📧 Прайс не высылался", "🔥 A (VIP)", 228.23, 10971227, "КП (2015–2021, посл. 22.10.2021, 13 УПД): 325х16 (77.2т), 426х18 (35.3т), 426х17 (31.0т), 530х16 (29.6т), 530х12 (23.7т), 426х7 (15.7т)"], ["3123461261", "ООО \"БЕЛГЕОСТРОЙ\"", "308017, Белгородская область, г. Белгород, ул. Константина Заслонова, д. 167А, офис 110", "Владимир Васильевич\nВиктор\nЮрий Владимирович\nВладимир Владимирович", "+7 (910) 741-89-68\n+7 (910) 322-60-50\n+7 (919) 438-99-99", "igsbelgorod@mail.ru", "Бурение на воду", "", "", "", "📌 ООО \"Инжгидрострой\" 3123121025 должник", "", "", "", "📧 Прайс не высылался", "🔥 A (VIP)", 99.5, 6999722, "КП+РЛ (2019–2025, посл. 24.09.2025, 28 УПД): 219х8 (28.9т), 273х8 (23.5т), 219х7 (14.7т), 273х7 (10.2т), 325х8 (5.0т), 426х9 (3.6т)"], ["366602553970", "ИП Швецов Николай Иванович", "Воронежская область, г. Воронеж", "Павел Белогубов", "+7 (904) 213-88-44", "pssv.vrn@rambler.ru", "Бурение на воду", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔥 A (VIP)", 100.62, 6636639, "КП+РЛ (2022–2025, посл. 06.11.2025, 15 УПД): 426х9 (20.5т), 630х9 (15.3т), 530х8 (14.9т), 630х7 (11.7т), 720х10 (9.7т), 530х7,0-9 (7.3т)"], ["3122510364", "ООО \"АЛОКС\"", "309857, Белгородская область, м. о. Алексеевский, г. Алексеевка, ул. Производственная, д. 33", "", "+7 (472) 343-02-70\n+7 (472) 344-53-87", "aloks7777@mail.ru", "Другое", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔥 A (VIP)", 129.23, 5771845, "КП (2021, посл. 12.04.2021, 7 УПД): 325х9 (76.0т), 325х8 (53.3т)"], ["3111005578", "АО \"САМАРИНСКОЕ\"", "", "Виктор Иванович", "+7 (919) 280-25-00", "samarinskoe@yandex.ru", "", "Нет", "", "", "", "", "", "", "📧 Прайс не высылался", "🔥 A (VIP)", 137.59, 5759330, "РЛ (2020–2021, посл. 21.05.2021, 7 УПД): 60х5 (86.3т), 73х5,5 (51.3т)"], ["2372033507", "ООО \"ЮСМ\"", "352272, Краснодарский край, м. р-н Отрадненский, с. п. Удобненское, ст-ца Удобная, ул. Международная, д. 219А", "Игорь Александрович", "+7 (967) 665-51-65", "yusm-yug@bk.ru\nkarpov.i.a@bk.ru", "Другое", "720х8- 2 шт 530х7- 3 шт 720х10- 1 шт 530х8- 1 шт", "Новая", "Металлоконструкции", "", "", "", "", "🟢 Под заявку на складе: 720х10 Нов. (2.2т — 83 000 ₽) | 530х6 Нов. (3.1т — 55 000 ₽), 530х8 Нов. (1т — 60 000 ₽), 530х9 Нов. (1.3т — 60 000 ₽), 530х9 Леж. (3.8т — 70 000 ₽), 530х10 Нов. (16.5т — 62 000 ₽), 530х10 Нов. (22.9т — 74 000 ₽)", "🔥 A (VIP)", 73.59, 5747242, "КП (2024–2026, посл. 18.09.2026, 4 УПД): 530х10 (17.8т), 720х10 (17.6т), 720х9 (13.3т), 426х8 (10.6т), 720х8 (5.0т), 530х8 (3.2т)"], ["4821014339", "АО \"ОЛЬШАНСКИЙ КАРЬЕР\"", "399776, Липецкая область, г. Елец, пер. Томский, зд. 17", "Татьяна Леонидовна", "+7 (919) 257-87-29", "snab@olshkar.ru", "Карьер", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔥 A (VIP)", 1030.36, 5273807, "КП (2015–2022, посл. 25.01.2022, 10 УПД): 325х9 (599.6т), 325х8 (339.6т), 273х8 (72.2т), 273х8,0-9 (18.6т), 630х8 (0.3т)"], ["3128093452", "ООО \"СК ЭТАЛОН\"", "309512, Белгородская область, г. Старый Оскол, ул. Прядченко, д. 87, офис 01", "Алексей Юрьевич", "+7 (910) 229-32-42", "etalon-bs@yandex.ru", "", "Нет", "", "", "", "12.11.2026", "", "[⏳ В Ожидание - 29.10.2026 15:08]", "📧 Прайс не высылался", "🔥 A (VIP)", 53.08, 5092126, "КП+РЛ (2022–2023, посл. 07.09.2023, 4 УПД): 325х8 (40.0т), 219х6 (9.3т), 159х5 (3.8т)"], ["3663130377", "ООО \"ЭНЕРГОСТАЛЬСТРОЙ\"", "394074, Воронежская область, г. о. город Воронеж, г. Воронеж, ул. Новосибирская, д. 13А", "Антон", "+7 (920) 465-94-13", "energstr21@mail.ru", "Металлобаза", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔥 A (VIP)", 61.7, 4665406, "КП+РЛ (2019–2026, посл. 09.10.2026, 33 УПД): 159х8 (19.3т), 325х8 (8.9т), 426х10 (6.1т), 1020х12 (4.0т), 426х8 (3.6т), 325х10 (3.5т)"], ["6165150175", "ООО \"НАУЧНО-ПРОИЗВОДСТВЕННОЕ ПРЕДПРИЯТИЕ \"РОСТОВСКАЯ БУРОВАЯ КОМПАНИЯ\"", "346720, Ростовская область, м. р-н Аксайский, г. п. Аксайское, г. Аксай, ул. Промышленная, влд. 6А", "Александр", "+7 (989) 711-83-33", "rostovburenie@mail.ru", "Бурение на воду", "Нет", "", "", "", "", "", "", "📧 Прайс не высылался", "🔥 A (VIP)", 55.68, 3877665, "КП (2022–2025, посл. 29.07.2025, 5 УПД): 219х8 (20.9т), 273х5,5-6 (10.9т), 325х6 (10.7т), 273х6 (8.3т), 168х7 (4.9т)"], ["3457004070", "ООО \"КФХ ГУБИНА А.А.\"", "403901, Волгоградская область, Новониколаевский район, рп. Новониколаевский, ул. Первомайская, д. 47, офис 1", "Марина", "+7 (844) 446-10-01", "ag23@list.ru", "Агропром", "Нет", "", "", "", "28.10.2026", "", "[📧 Прайс: Полный + PDF (для: Марина) - 26.10.2026 11:19]", "✅ Прайс отправлен", "🔥 A (VIP)", 151.03, 3865462, "КП+РЛ (2015–2024, посл. 07.11.2024, 12 УПД): 73х5,5 НКТ (35.3т), 273х5,0-8 (20.0т), 273х6 (20.0т), 273х6,0-8 (11.0т), 48х4 (10.0т), 73х5,5 (8.9т)"], ["3123174732", "ООО \"АККРЕДИТИВ\"", "308023, Белгородская область, г. Белгород, ул. Студенческая, д. 25, помещ. 1", "Илья \nСергей Николаевич", "+7 (910) 321-93-83\n+7 (919) 282-81-02", "btk31@bk.ru\nmail@sb31.ru", "Бурение на воду", "", "", "", "📌 Смежные юрлица для оплат: ТД \"БТК\".(ИНН 3102047405) ООО \"СОЮЗБУРВОД\".(ИНН 3123174718) ООО \"БЕЛИ…", "", "", "[📋 Снята/закрыта потребность | Запрос: Ø 168, Ø 219, Ø 273, Ø 325, Ø 426, Ø 530, Ø 630, Ø 820 (Новая • Обсадка / Скважины) - 19.10.2026 13:51]", "📧 Прайс не высылался", "🔥 A (VIP)", 52.55, 3571344, "КП (2022–2026, посл. 02.09.2026, 4 УПД): 89х6,5 НКТ (20.3т), 219х10 (11.3т), 325х8 (10.1т), 426х8 (8.5т), 219х8 (2.3т)"], ["3616020144", "ООО \"БУРВИГ\"", "396333, Воронежская область, Новоусманский район, п. 1-Го Отделения Совхоза \"Масловский\", ул. Школьная, зд. 20/1, офис 1", "Гончаров Владимир Иванович", "+7 (920) 227-84-99", "vgoncharov@bk.ru", "Бурение на воду", "Нет", "", "", "", "", "", "", "📧 Прайс не высылался", "🔥 A (VIP)", 47.61, 2776112, "КП+РЛ (2020–2022, посл. 14.06.2022, 16 УПД): 219х8 (31.2т), 219х7 (9.8т), 60х5 (4.8т), 219х6 (1.5т), 168х8 (0.3т)"], ["3661172815", "ООО \"ВОРОНЕЖСТРОЙМОНТАЖ\"", "394030, Воронежская область, г. о. город Воронеж, г. Воронеж, ул. Карла Маркса, д. 116А, помещ. 1/1, офис 1", "Иван Николаевич", "+7 (920) 443-62-07", "oarghunova@mail.ru", "Теплосети", "", "", "", "", "", "", "", "📧 Прайс не высылался", "⭐ B (Постоянный)", 37.33, 2473585, "КП (2025–2026, посл. 17.07.2026, 6 УПД): 325х6 (11.6т), 325х8 (11.5т), 159х6 (11.2т), 426х9 (3.0т)"], ["3663164626", "ООО \"ПРОМБУРВОД\"", "394072, Воронежская область, г. о. город Воронеж, г. Воронеж, ул. Иркутская, д. 19, офис 1, помещ. 5/3", "Павел Константинович", "+7 (920) 414-24-50", "pbv-vrn36@mail.ru\nsbv-vrn@mail.ru", "Бурение на воду", "", "", "", "📌 Смежные юрлица для оплат: ООО \"СТРОЙБУРВОД-П…", "", "", "[📋 Снята/закрыта потребность | Запрос: Ø 89, Ø 159, Ø 168, Ø 219, Ø 273, Ø 325, Ø 73, Ø 426 (Новая, Восстановленная, Лежалая, Б/у, НКТ, ЭСВ • Обсадка / Скважины) - 19.10.2026 13:52]", "📧 Прайс не высылался", "⭐ B (Постоянный)", 36.9, 2466432, "КП (2024–2026, посл. 04.09.2026, 17 УПД): 325х8 (6.5т), 273х7 (5.8т), 219х6 (5.3т), 273х6 (4.4т), 219х8 (4.3т), 325х6 (3.8т)"], ["6686110245", "ООО \"АЛЬЯНС\"", "123007, г. Москва, вн. тер. г. муниципальный округ Хорошевский, ул. 1-я Магистральная, д. 16, стр. 1", "Евгений Петрович", "+7 (912) 656-00-07", "aliance.18@mail.ru", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "⭐ B (Постоянный)", 65.17, 2272700, "РЛ (2021, посл. 15.04.2021, 5 УПД): 1020х12 (46.3т), 1020х11 (18.9т)"], ["3664066318", "ООО \"ВОРОНЕЖСТРОЙРЕКОНСТРУКЦИЯ\"", "394033, Воронежская область, г. Воронеж, ул. Остужева, д. 21, кв. 6", "Алесандр Алексеевич", "+7 (473) 260-65-00", "vrn-srk@mail.ru", "", "Нет", "", "", "📌 нет заказов ,берут под проколы разного диамет… | 📌 [19.10.2026 11:19]: на воду отопление полтара года в ,э/с от 100", "09.11.2026", "", "", "📧 Прайс не высылался", "⭐ B (Постоянный)", 27.05, 1716968, "КП (2019–2021, посл. 17.11.2021, 4 УПД): 219х8 (13.0т), 530х6 (9.9т), 426х6,0-7 (2.2т), 426х10 (1.1т), 426х8 (0.8т)"], ["7328063364", "ООО \"СМР-2020\"", "432072, Ульяновская область, г. о. город Ульяновск, г. Ульяновск, бульвар Киевский, д. 24, помещ. 4", "Сергей Камалов", "+7 (986) 736-25-17", "sergeykamalov@mail.ru", "Металлоконструкции", "", "", "", "", "", "", "", "📧 Прайс не высылался", "⭐ B (Постоянный)", 19.15, 1685024, "КП (2023, посл. 31.05.2023, 1 УПД): 426х9 (19.1т)"], ["3604082795", "ООО \"КВАДРО\"", "397165, Воронежская область, г. о. Борисоглебский, г. Борисоглебск, ул. Матросовская, зд. 111Д", "Виталий Калинин", "+7 (908) 137-12-22", "kvadro_isk@mail.ru", "Строительство (ПГС)", "Нет", "", "", "📌 Жесткий отказ. Говорит, не закупают (хотя зак…", "09.11.2026", "", "", "📧 Прайс не высылался", "⭐ B (Постоянный)", 20.34, 1400027, "КП (2021, посл. 08.10.2021, 3 УПД): 219х8 (5.3т), 426х10 (4.2т), 426х8,0-9 (4.1т), 219х7 (3.7т), 426х8 (1.9т), 426х9 (1.1т)"], ["3661047395", "ООО \"БУРСЕРВИС\"", "394010, Воронежская область, г. Воронеж, ул. Артамонова, д. 30А, кв. 153", "Сергей Данковцев", "+7 (920) 223-85-55", "dankovtseva.elena@yandex.ru", "Бурение на воду", "", "", "", "📌 Смежные юрлица для оплат : ООО \"СТРОЙРЕМБУРВОД\".(ИНН 3663128579)", "", "", "", "📧 Прайс не высылался", "⭐ B (Постоянный)", 22.3, 1291709, "КП+РЛ (2017–2026, посл. 30.03.2026, 11 УПД): 530х8 (7.1т), 219х8 (2.5т), 273х7,5 (2.2т), 273х8 (1.9т), 273х6 (1.8т), 426х7 (1.6т)"], ["3123221277", "ООО ПМК-1 \"БЕЛГОРОДБУРВОДСТРОЙ\"", "308023, Белгородская область, г. Белгород, ул. Студенческая, д. 17Г, офис 27", "Юрий", "+7 (903) 642-03-03", "", "", "Нет", "", "", "", "09.11.2026", "", "", "📧 Прайс не высылался", "⭐ B (Постоянный)", 39.12, 1248923, "КП (2016–2026, посл. 23.03.2026, 7 УПД): 530х8 (9.3т), 273х6 (9.0т), 426х8 (7.5т), 630х9 (5.5т), 426х7 (4.0т), 159х6 (2.6т)"], ["3662287255", "ООО \"АКТИВЭНЕРГО\"", "394026, Воронежская область, г. о. город Воронеж, г. Воронеж, ул. Газовая, д. 26, офис 2", "Дмитрий Васильевич", "+7 (952) 958-48-15", "aktivenergo@list.ru", "", "Нет", "", "", "", "12.11.2026", "", "[⏳ В Ожидание - 29.10.2026 14:14]", "📧 Прайс не высылался", "⭐ B (Постоянный)", 19.8, 1227662, "КП (2022, посл. 01.08.2022, 1 УПД): 273х8 (15.0т), 273х7 (4.8т)"], ["3128030928", "ООО \"СТРОЙТЕХМОНТАЖ\"", "309517, Белгородская область, г. о. Старооскольский, г. Старый Оскол, промзона. Юз П/Р, пл-ка Машиностроительная, зд. 30", "Сергей Анатольевич", "+7 (472) 544-19-07", "stm6@rambler.ru, sedneva.54@mail.ru, mku_uho@mail.ru, rubnikov3584@gmail.com", "", "Нет", "", "", "", "12.11.2026", "", "[⏳ В Ожидание - 29.10.2026 14:22]", "📧 Прайс не высылался", "⭐ B (Постоянный)", 35.74, 1095778, "КП+РЛ (2015–2021, посл. 23.07.2021, 9 УПД): 325х6 (8.7т), 168х7 (6.0т), 325х8 (5.5т), 168х6 (4.8т), 426х6 (3.6т), 530х8-9 (3.3т)"], ["3420002239", "ООО \"КХ АРЬЕВА А.В.\"", "403901, Волгоградская область, Новониколаевский район, рп. Новониколаевский, ул. Советская, д. 33", "Алексей Владимирович", "+7 (937) 741-44-67\n+7 (905) 064-57-45(основ)", "vika.24.12@yandex.ru", "", "Нет", "", "", "", "", "", "", "📧 Прайс не высылался", "⭐ B (Постоянный)", 29.99, 1080191, "КП (2017–2025, посл. 29.01.2025, 3 УПД): 159х6 (27.6т), 273х6 (1.1т), 57х4 (0.6т), 57х7 (0.5т), 57х5 (0.3т)"], ["3123409857", "ООО \"ЭЛИТКОМПАНИ\"", "308017, Белгородская область, г. о. город Белгород, г. Белгород, ул. Константина Заслонова, д. 181, помещ. 7", "Татьяна Гребенькова\nВладимир Васильевич", "+7 (910) 222-11-88\n+7 (910) 222-99-88", "promburvod@list.ru\nvlasova.75@list.ru", "", "", "", "", "📌 3123388660 ООО \"ПРОМБУРВОД\" Белгород…", "", "", "", "📧 Прайс не высылался", "⭐ B (Постоянный)", 18.66, 1074187, "КП+РЛ (2018–2021, посл. 08.04.2021, 5 УПД): 325х8 (13.7т), 426х7 (3.0т), 273х8 (1.2т), 426х10-11 (0.8т)"], ["3663176283", "ООО \"ВОДБУРСТРОЙ\"", "394094, Воронежская область, г. о. город Воронеж, г. Воронеж, ул. Пятилетка, д. 54", "Николай Гиндин", "+7 (920) 211-79-74", "ip.gindin@gmail.com, buh_gp2018@mail.ru", "Бурение на воду", "", "", "", "📌 Смежные юрлица для оплат: ООО \"СТРОЙРЕМБУРВО…", "", "", "[📋 Снята/закрыта потребность | Запрос: Ø 159, Ø 168, Ø 219, Ø 273, Ø 325, Ø 426, Ø 530, Ø 630, Ø 720, Ø 377 (Новая, Лежалая, Восстановленная, Б/у, Спиральношовная, ЭСВ • Прокол / ГНБ, Обсадка / Скважины) - 19.10.2026 13:51]", "📧 Прайс не высылался", "⭐ B (Постоянный)", 15.11, 1026921, "КП (2026, посл. 02.11.2026, 6 УПД): 219х8 (8.1т), 219х7 (3.4т), 273х8 (1.8т), 530х7 (0.9т), 273х6 (0.9т)"], ["3604011804", "ООО \"БОРМАШ\"", "397340, Воронежская область, Поворинский район, с. Пески, ул. Пролетарская, зд. 59", "Сергей Сусоров", "+7 (473) 763-29-90(209)", "omts@bormash.ru", "", "Нет", "", "", "📌 Берут в основном бесшовную трубу ,электросвар…", "09.11.2026", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 13.13, 921450, "КП (2021, посл. 08.10.2021, 1 УПД): 325х9,0-10 (8.1т), 219х8 (3.8т), 426х10 (1.2т)"], ["6234147257", "ООО \"СОЮЗАЛЬП\"", "390026, Рязанская область, г. Рязань, ул. Стройкова, д. 92, кв. 64", "Юрий Валерьевич", "+7 (910) 639-64-64", "alpsouz@yandex.ru", "Строительство", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 9.76, 796641, "РЛ (2022–2023, посл. 26.06.2023, 2 УПД): 426х11 (4.2т), 630х9 (3.2т), 426х10 (2.4т)"], ["342000081487", "Глава КФХ Самоходкин Геннадий Павлович", "Волгоградская область, х. Галушкинский", "Геннадий Павлович", "+7 (937) 714-32-35", "ldpz@yandex.ru", "", "Нет", "", "", "", "12.11.2026", "", "[⏳ В Ожидание - 29.10.2026 14:35]", "📧 Прайс не высылался", "🔹 C (Малый)", 10.0, 650000, "КП (2021, посл. 03.12.2021, 1 УПД): 159х6 (10.0т)"], ["361916372929", "ИП Кузнецов Евгений Леонтьевич", "Воронежская область, с. Сторожевое 1-е", "Евгений", "+7 (920) 416-23-07", "remselvod-ostro@mail.ru", "Бурение на воду", "Нет", "", "", "📌 Берет 530,630,720,426 Под обсадку", "23.10.2026", "", "[📧 Прайс: Полный - 14.10.2026 12:44]", "✅ Прайс отправлен", "🔹 C (Малый)", 9.95, 649032, "КП (2025, посл. 05.11.2025, 2 УПД): 530х9 (3.9т), 219х7 (2.1т), 630х9 (1.6т), 530х8 (1.2т), 630х7 (1.1т)"], ["3123412232", "ООО \"КВАЛИТЕТ-СЕРВИС\"", "308033, Белгородская область, г. о. город Белгород, г. Белгород, ул. Королева, д. 20, кв. 46", "Сергей\nАлександр", "+7 (999) 700-33-54\n+7 (910) 320-17-11", "Kvalitet031@mail.ru", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 8.19, 639110, "КП (2024–2025, посл. 06.10.2025, 2 УПД): 273х7 (3.7т), 219х6 (2.9т), 325х8,0-9 (1.6т)"], ["4632172315", "ООО \"РЕМСТРОЙМОНТАЖ\"", "305047, Курская область, г. Курск, ул. Энгельса, д. 171А, офис 211", "Андрей Иванович", "+7 (919) 174-36-24", "rsm46@mail.ru", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 14.74, 601276, "РЛ (2021, посл. 10.02.2021, 1 УПД): 273х7 (10.0т), 325х7,0-8 (3.9т), 426х8 (0.9т)"], ["3665076132", "ООО \"ФИРМА СМУ-5\"", "121596, г. Москва, вн. тер. г. муниципальный округ Можайский, ул. Горбунова, д. 2, стр. 3, помещ. 35", "Павел Николаевич", "+7 (495) 374-61-20\n+7 (800) 700-31-15\n+7 (473) 280-12-25", "790089@mail.ru\n2801225@bk.ru\nsmu-700@mail.ru", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 12.25, 600348, "КП (2021, посл. 11.06.2021, 1 УПД): 530х8 (9.2т), 530х7 (3.1т)"], ["3620014376", "ООО \"ПАВЛОВСКАЯ ПМК № 18 \"ВОДСТРОЙ\"", "396420, Воронежская область, Павловский район, г. Павловск, ул. Строительная, д. 50", "Вячеслав Владимирович", "+7 (920) 413-10-05\n+7 (473) 622-11-77", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 12.55, 577051, "КП+РЛ (2018–2025, посл. 07.03.2025, 4 УПД): 273х8 (6.6т), 273х6 (4.7т), 273х8,5 (1.2т)"], ["4808005127", "АО \"ХМЕЛИНЕЦКИЙ КАРЬЕР\"", "399240, Липецкая область, Задонский район, с. Донское, ул. Свердлова, д. 201", "Дмитрий Николаевич", "+7 (980) 261-04-36\n+7 (474) 228-84-26\n+7 (474) 712-23-50\n+7 (474) 713-32-14\n+7 (074) 712-23-50", "sbyt@hmelkar.ru\nhmelinecd@mail.ru\nrogdkar@mail.ru", "Другое", "Нет", "", "", "", "12.11.2026", "", "[⏳ В Ожидание - 29.10.2026 14:41]", "📧 Прайс не высылался", "🔹 C (Малый)", 26.63, 536030, "КП (2015–2026, посл. 02.06.2026, 4 УПД): 325х8 (12.7т), 530х8 (11.6т), 630х8 (1.4т), 530х6 (0.9т)"], ["3665145675", "ООО \"СТАЛЬМОСТКОНСТРУКЦИЯ\"", "394033, Воронежская область, г. о. город Воронеж, г. Воронеж, пр-кт Ленинский, д. 119, стр. 1, помещ. 1", "Антон Сергеевич", "+7 (908) 141-48-88", "smk@ppka36.ru", "", "Нет", "", "", "", "13.11.2026", "", "[⏳ В Ожидание - 29.10.2026 14:44]", "📧 Прайс не высылался", "🔹 C (Малый)", 7.11, 493167, "КП+РЛ (2019–2022, посл. 18.11.2022, 2 УПД): 630х9 (5.4т), 530х12 (1.7т)"], ["3663127649", "ООО \"ПРОММОНТАЖ-ВОРОНЕЖ\"", "394090, Воронежская область, г. о. город Воронеж, г. Воронеж, ул. Ростовская, д. 45, офис 1", "Юлия", "+7 (908) 148-50-30", "pragmo@yandex.ru", "", "", "", "", "📌 Низкая валидность", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 8.73, 479985, "РЛ (2023, посл. 13.10.2023, 1 УПД): Труба ст.530 (8.7т)"], ["3123329714", "ООО \"СК ЧЕРНОЗЕМЬЯ\"", "308015, Белгородская область, г. о. город Белгород, г. Белгород, ул. Сумская, д. 66А, офис 203", "Владислав \nВячеслав Александрович", "+7 (919) 280-85-38\n+7 (920) 594-94-27", "vladik588@yandex.ru\nvd.servis31@yandex.ru", "", "Нет", "", "", "📌 сейчас потребности нет", "09.11.2026", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 3.98, 352491, "КП (2024, посл. 10.10.2024, 1 УПД): 219х6 (2.5т), 325х8 (1.4т)"], ["3128124372", "ООО \"РЕСУРС-КОМПЛЕКТ\"", "Старый Оскол", "Виктор Евгеньевич", "+7 (915) 565-01-92", "rkomplekt31@gmail.com", "Металлоторговля / Трейдер", "Нет", "", "", "", "05.11.2026", "", "[🚫 В Архив: Не целевой (из «1. Обзвон (База)») - 14.10.2026 15:23]", "📧 Прайс не высылался", "🔹 C (Малый)", 5.22, 340254, "КП (2024, посл. 28.11.2024, 1 УПД): 325х9 (2.9т), 630х7 (2.3т)"], ["3616027069", "ООО \"ПРОМ РЕСУРС\"", "396310, Воронежская область, м. р-н Новоусманский, с. п. Усманское 1-Е, с. Новая Усмань, ул. Кулибина, д. 131", "Анна\nЛилия", "+7 (905) 652-31-43\n+7 (950) 755-27-98\n+7 (910) 344-99-40", "prom-resurs-vrn@mail.ru\npromresurs_vrn@mail.ru", "Металлоторговля / Трейдер", "Нет", "", "", "", "09.11.2026", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 4.93, 305970, "КП (2024, посл. 17.10.2024, 1 УПД): 219х8 (4.9т)"], ["3662303884", "ООО СК \"ИНЖЕНЕРНЫЕ СИСТЕМЫ\"", "394016, Воронежская область, г. о. город Воронеж, г. Воронеж, ул. 45 Стрелковой Дивизии, д. 251Д, помещ. 7/1", "Дмитрий Николаевич", "+7 (960) 139-26-06", "ck.sistem@yandex.ru", "Строительство", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 3.75, 279673, "РЛ (2023, посл. 14.07.2023, 1 УПД): 273х7 (3.8т)"], ["3128098877", "ООО \"Сигма\"", "", "Влад", "+7 (960) 630-40-50", "vgm050@gmail.com", "", "Нет", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 3.14, 263784, "КП (2024, посл. 26.04.2024, 1 УПД): 426х10 (2.4т), 426х11 (0.8т)"], ["3664252794", "ООО \"ПАС\"", "394036, Воронежская область, г. о. город Воронеж, г. Воронеж, ул. Пушкинская, д. 1", "Артём Сергеевич", "+7 (908) 135-11-87", "pas-136@mail.ru", "Теплосети / Энергетика, Бурение / ГНБ", "Нет", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 3.67, 259878, "РЛ (2023, посл. 24.08.2023, 2 УПД): 159х5 (1.3т), 273х5,0-6 (0.7т), 426х8-9 (0.7т), 426х8 (0.7т), 159х6 (0.3т)"], ["5612075114", "ООО \"УК - РЦ \"ЮЖНЫЙ УРАЛ\"", "460051, Оренбургская область, г. о. город Оренбург, г. Оренбург, пр-кт Гагарина, д. 43/3, помещ. 5", "Юля", "+7 (987) 844-96-66", "", "Строительство", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 3.24, 212706, "КП (2025–2026, посл. 27.01.2026, 3 УПД): 426х7 (1.5т), 426х9 (1.1т), 377х8 леж. (0.3т), 273х5 (0.3т)"], ["3620005854", "АО \"ПАВЛОВСКАГРОПРОДУКТ\"", "396420, Воронежская область, Павловский район, г. Павловск, ул. Строительная, д. 8", "Владимир", "+7 (473) 622-35-60\n+7 (473) 622-04-47", "pavlovsk-agroprodukt@yandex.ru\nmail@aprotek.ru\noms@kalibr.vrn.ru\nzam@aprotek.ru\n200072@nalog.ktkt.ru\nmail@agroprodukt.vrn.ru", "", "Нет", "", "", "", "13.11.2026", "", "[⏳ В Ожидание - 29.10.2026 14:53]", "📧 Прайс не высылался", "🔹 C (Малый)", 3.34, 212693, "КП (2023, посл. 12.10.2023, 2 УПД): 273х5 (1.5т), 426х8-9 (1.0т), 426х7 (0.8т)"], ["3662196576", "ООО \"ВОРОНЕЖСКАЯ СТРОИТЕЛЬНАЯ КОМПАНИЯ\"", "394026, Воронежская область, г. Воронеж, ул. Антонова-Овсеенко, д. 36Д, офис 12", "Андрей Александрович Добрин", "+7 (900) 304-17-27\n+7 (900) 304-17-39", "andrey_vrn@rambler.ru", "Строительство", "Нет", "", "", "", "12.11.2026", "", "[⏳ В Ожидание - 29.10.2026 14:43]\n[📞 Недозвон: 1/4 - 21.10.2026 10:26]", "📧 Прайс не высылался", "🔹 C (Малый)", 3.86, 201066, "КП+РЛ (2019–2024, посл. 18.01.2024, 4 УПД): 630х8 (1.4т), 325х6 (0.9т), 325х6,5 (0.8т), 325х8 (0.7т)"], ["3500005210", "ООО \"ЭТПС\"", "160014, Вологодская область, г. о. город Вологда, г. Вологда, ул. Некрасова, д. 60А, помещ. 15", "Сергей Сергеевич", "+7 (985) 115-60-00\n+7 (927) 975-75-59", "tsstnctop@yandex.ru", "", "Нет", "", "", "", "13.11.2026", "", "[⏳ В Ожидание - 29.10.2026 14:55]", "📧 Прайс не высылался", "🔹 C (Малый)", 2.53, 186276, "КП (2025, посл. 26.05.2025, 1 УПД): 426х8 (1.8т), 273х7 (0.5т), 219х6 (0.2т)"], ["3128065293", "АО \"МАВР-БЕЛГОРОД\"", "309500, Белгородская область, г. Старый Оскол, ст. Котел Промузел, пл-ка Транспортная Проезд Ш-4 (Станция Котел Промузел Тер), площадка транспортная, проезд ш-4 строение 29", "Павел Станиславович", "+7 (472) 523-00-16\n+7 (472) 523-00-15", "os_snab@mail.ru", "Строительство", "Нет", "", "", "📌 используют редко под насыпные переправы брал…", "14.01.2027", "", "[📧 Прайс: Полный + PDF (для: Павел Станиславович) - 21.10.2026 13:50]\n[⏰ Перезвонить после 12:24 - 21.10.2026 10:24]\n[Недозвон: 1/4 - 24.09.2026 03:32]", "✅ Прайс отправлен", "🔹 C (Малый)", 2.61, 173228, "КП (2025, посл. 28.03.2025, 1 УПД): 325х6 (2.3т), 89х6,5 НКТ (0.3т)"], ["3662186722", "ООО \"СК АТОМ СТРОЙ\"", "394036, Воронежская область, г. о. город Воронеж, г. Воронеж, ул. 20-Летия Влксм, д. 35, помещ. 3", "Дмитрий Вадимович \nАндрей Анисимов", "+7 (919) 180-77-17\n+7 (915) 544-33-96", "info@atomstroi.ru", "", "Нет", "", "", "", "13.11.2026", "", "[⏳ В Ожидание - 29.10.2026 14:57]", "📧 Прайс не высылался", "🔹 C (Малый)", 1.92, 163716, "КП (2024, посл. 29.10.2024, 1 УПД): 426х8 (1.9т)"], ["360401430374", "ИП Андросов Иван Михайлович", "Воронежская область, г. Борисоглебск", "Иван Михайлович", "+7 (910) 348-17-79", "uni12x15@mail.ru", "", "Нет", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 1.79, 116675, "КП (2025, посл. 29.01.2025, 1 УПД): 273х7 (1.8т)"], ["3663072816", "ООО \"РОСТЕХЭНЕРГО-ВОРОНЕЖ\"", "", "Анастасия", "+7 (909) 212-11-55", "info@rte-vrn.ru, sav.rte-vrn@yandex.ru", "", "Нет", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 0.66, 56819, "РЛ (2023, посл. 21.08.2023, 1 УПД): 426х10 (0.7т)"], ["3663133201", "ООО ТД \"СТАЛЬИНВЕСТ\"", "394036, Воронежская область, г. о. город Воронеж, г. Воронеж, ул. Степана Разина, д. 5А, офис 304", "Склад Кирилл", "+7 (903) 654-50-99", "steel_vrn@bk.ru", "Металлобаза", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 0.83, 52416, "КП (2021, посл. 02.07.2021, 1 УПД): 426х10 (0.8т)"], ["772490250004", "ИП Руднев Павел Николаевич", "", "Павел", "+7 (473) 228-65-91", "rudnev_pavel83@mail.ru", "", "Нет", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 0.55, 42978, "КП (2023, посл. 31.10.2023, 1 УПД): 325х6 (0.6т)"], ["3664242771", "ООО \"СПЕЦТАЙЛ\"", "394036, Воронежская область, г. о. город Воронеж, г. Воронеж, пр-кт Революции, д. 39, офис 9", "Дмитрий Алексеевич", "+7 (960) 132-92-61\n+7 (920) 440-26-26", "79601247533@mail.ru", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 0.68, 37125, "РЛ (2023, посл. 12.05.2023, 1 УПД): 89х6,5 (0.7т)"], ["3663165965", "ООО \"АЛЬЯНС\"", "396335, Воронежская область, м. р-н Новоусманский, с. п. Отрадненское, п. Отрадное, ул. Родниковая, д. 5, кв. 100", "Александр Игоревич", "+7 (980) 537-58-82", "aleksandrbazhkov@mail.ru", "", "Нет", "", "", "", "16.11.2026", "", "[⏳ В Ожидание - 29.10.2026 15:04]", "📧 Прайс не высылался", "🔹 C (Малый)", 0.46, 27360, "КП (2025, посл. 03.06.2025, 2 УПД): 530х8 (0.2т), 426х8 (0.2т), 219х8 (0.1т)"], ["3662136721", "ООО РПП \"БЛИКФАНГ\"", "394026, Воронежская область, г. Воронеж, пр-кт Труда, д. 48, помещ. 2", "Екатерина (127)", "+7 (473) 204-51-59\n+7 (473) 250-23-87", "kro2@blikfang.ru\ntula1@blikfang.ru\nbank@blikfang.ru", "Другое", "Нет", "", "", "", "16.11.2026", "", "[⏳ В Ожидание - 29.10.2026 15:05]", "📧 Прайс не высылался", "🔹 C (Малый)", 0.39, 25225, "КП (2024, посл. 13.05.2024, 1 УПД): 273х5,5-6 (0.4т)"], ["3665818828", "ООО \"А-МЕДИА ПРО\"", "394065, Воронежская область, г. о. город Воронеж, г. Воронеж, пр-кт Патриотов, д. 49Б, помещ. 4", "Алексей Анатольевич", "+7 (952) 950-58-05", "mmps@mail.ru", "", "Нет", "", "", "📌 Сказал набрать через полтора месяца с сегодне…", "10.11.2026", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 0.28, 19670, "КП (2022, посл. 26.05.2022, 1 УПД): 325х8 (0.3т)"], ["3662253129", "ООО \"АГРОЛАКС\"", "394088, Воронежская область, г. Воронеж, бульвар Победы, д. 49, кв. 67", "Максим Михайлович", "+7 (908) 141-16-01", "info@laksgrup.com", "Другое", "Нет", "", "", "📌 Берут под пожарку в основном электросварную", "02.11.2026", "", "[📋 Снята/закрыта потребность | Запрос: Ø 159, Ø 108, 56, Ø 76 (Новая) - 19.10.2026 11:20]\n[⏳ В Ожидание | Запрос: Ø 159, Ø 108, 50 76 (Новая) (из «1. Обзвон (База)») - 19.10.2026 11:06]\n[📞 Недозвон: 1/4 - 14.10.2026 15:59]", "📧 Прайс не высылался", "🔹 C (Малый)", 0.1, 5512, "КП (2023, посл. 10.08.2023, 1 УПД): 219х7 (0.1т), 273х5 (0.0т)"], ["3123377073", "ООО \"ПП БАРС\"", "308027, Белгородская область, г. о. город Белгород, г. Белгород, ул. Лермонтова, д. 49А, кв. 54", "Евгений Викторович", "+7 (908) 789-92-30\n+7 (919) 222-90-85", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔥 A (VIP)", 250.75, 8490607, "КП (2016–2019, посл. 02.12.2019, 21 УПД): 273х8 (54.9т), 426х8 (33.5т), 273х6 (26.4т), 273х8,5 (17.4т), 426х7 (13.8т), 426х8,0-10 (9.8т)"], ["4628006724", "ООО \"ТЕХНОЛОГИЯ\"", "305022, Курская область, г. Курск, ул. Республиканская, д. 50А, офис 1", "Юрий Николаевич", "+7 (910) 210-92-51", "a2404@mail.ru\ntmtehno@gmail.com\na200212@mail.ru", "Металлоконструкции", "Нет", "", "", "📌 в основном много профиля покупаю,говорит по 200-400 тн в месяц металла", "28.10.2026", "", "[⏳ В Ожидание - 26.10.2026 12:10]\n[📧 Прайс: Полный + PDF (для: Юрий Николаевич) - 26.10.2026 12:08]", "✅ Прайс отправлен", "🔥 A (VIP)", 114.71, 3665908, "КП (2015–2018, посл. 27.06.2018, 6 УПД): 273х5,5-6 (33.9т), 273х6 (23.6т), 200*200*8,0 Труба профильная (17.1т), 60х40 (13.9т), 120*3 (10.0т), 114х5 (5.5т)"], ["3666150847", "ООО \"ЭНЕРГОКОМПЛЕКС\"", "394036, Воронежская область, г. Воронеж, ул. Студенческая, д. 26А, офис 2", "Сергей Иванович", "+7 (473) 253-23-64\n+7 (473) 235-63-90\n+7 (473) 229-98-86\n+7 (800) 444-60-94\n+7 (473) 279-33-39", "ek.vrn@mail.ru\nprizs@mail.ru", "Водоканал", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔥 A (VIP)", 557.58, 3467568, "КП (2015–2018, посл. 20.08.2018, 13 УПД): 820х9,5 ВУС (208.2т), 630х8 ВУС (151.8т), 530х8 ВУС (108.7т), 325х6 ВУС (50.0т), 219х8 (36.7т), 530х8,5 (1.3т)"], ["3123121025", "ООО \"ИНЖГИДРОСТРОЙ\"", "308017, Белгородская область, г. о. город Белгород, г. Белгород, ул. Константина Заслонова, д. 167А, офис 111", "Владимир Васильевич", "+7 (910) 322-60-50\n+7 (910) 221-15-36", "igsbelgorod@mail.ru", "Бурение на воду", "", "", "", "", "", "", "", "📧 Прайс не высылался", "⭐ B (Постоянный)", 80.86, 2666615, "КП+РЛ (2017–2020, посл. 04.06.2020, 12 УПД): 325х10 (25.3т), 530х8 (22.0т), 273х8 (10.5т), 426х7 (9.2т), 219х8 (3.5т), 630х8 (2.9т)"], ["3663115812", "ООО \"АКВАБУРСТРОЙ\"", "394028, Воронежская область, г. Воронеж, ул. Димитрова, д. 148, офис 3", "Татьяна Александровна", "+7 (930) 856-46-46\n+7 (905) 644-12-79", "", "Водоканал", "", "", "", "", "", "", "", "📧 Прайс не высылался", "⭐ B (Постоянный)", 80.2, 2549482, "КП (2016–2018, посл. 09.10.2018, 12 УПД): 273х8 (32.8т), 630х8 (12.2т), 426х5,5 (5.8т), 720х9 (4.9т), 426х7 (4.9т), 377х9,5 (4.1т)"], ["4608005793", "ООО \"ОЛЫМСКИЙ САХАРНЫЙ ЗАВОД\"", "306716, Курская область, Касторенский район, пгт Олымский, ул. 20 Лет Победы, д. 1", "", "+7 (905) 158-34-64\n+7 (471) 576-65-72\n+7 (383) 355-98-82\n+7 (471) 332-15-50\n+7 (471) 576-68-41", "smotrovaep@olsah.ru", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "⭐ B (Постоянный)", 138.49, 2495521, "КП (2016, посл. 08.11.2016, 8 УПД): 530х9 (74.7т), 325х8,0-9 (19.0т), 530х8 (18.3т), 159х5,0-6 (14.9т), 273х7,0-8 (6.7т), 1020х12 (3.2т)"], ["3662198277", "ООО \"ТД \"ВОРОНЕЖТРУБКОМПЛЕКТ\"", "394077, Воронежская область, г. о. город Воронеж, г. Воронеж, бульвар Победы, д. 42, помещ. 5/4/1", "Андрей Вячеславович", "+7 (473) 220-29-68\n+7 (473) 280-18-28", "p.shukin@vtkvrn.ru\nvtkvrn@mail.ru", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "⭐ B (Постоянный)", 98.41, 2015291, "КП (2018, посл. 22.08.2018, 6 УПД): 530х8 (46.8т), 630х8 (33.4т), 530х8,3 (18.2т)"], ["3612007551", "ООО \"СХП \"НОВОМАРКОВСКОЕ\"", "396702, Воронежская область, Кантемировский район, с. Новомарковка, ул. Советская, зд. 14", "Владимир Александрович", "+7 (473) 674-93-47\n+7 (473) 674-91-89\n+7 (473) 674-93-10\n+7 (473) 674-91-30", "novomark36@mail.ru", "Агропром", "", "", "", "", "", "", "", "📧 Прайс не высылался", "⭐ B (Постоянный)", 79.0, 1936703, "КП (2018, посл. 25.04.2018, 4 УПД): 73х5,5 НКТ (79.0т)"], ["3123286450", "ООО \"СТРОИТЕЛЬНЫЕ ТЕХНОЛОГИИ\"", "308024, Белгородская область, г. Белгород, ул. Мокроусова, д. 23А, офис 200", "Оксененко Валерий Иванович", "+7 (472) 223-12-61\n+7 (472) 290-17-61\n+7 (980) 329-28-19\n+7 (800) 700-03-30", "", "", "Нет", "", "", "", "12.11.2026", "", "[⏳ В Ожидание - 29.10.2026 14:15]", "📧 Прайс не высылался", "⭐ B (Постоянный)", 50.19, 1037910, "КП (2015–2018, посл. 31.01.2018, 6 УПД): 426х10 (18.5т), 426х8 (9.8т), 426х8,0-9 (7.7т), 426х11 (7.0т), 426х12 (2.0т), 273х10 (1.6т)"], ["3123338684", "ЗАО \"ГИДРОИНЖСТРОЙ-ЮГ\"", "308017, Белгородская область, г. Белгород, ул. Константина Заслонова, д. 161", "Алексей Геннадьевич", "+7 (472) 221-15-43\n+7 (472) 220-56-43\n+7 (910) 329-22-06\n+7 (916) 184-49-04\n+7 (472) 221-16-19", "gisbelug@mail.ru", "Бурение на воду", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 34.31, 983525, "КП (2016–2019, посл. 14.06.2019, 7 УПД): 325х6 (15.4т), 273х6 (8.9т), 325х8 (4.1т), 426х6,5 (3.1т), 273х6,5 (2.8т)"], ["4632095371", "ООО \"ЖИЛСТРОЙСЕРВИС\"", "", "Андрей Малышев", "+7 (919) 174-36-24", "rsm46@mail.ru", "", "Нет", "", "", "📌 набрать через пару месяцев 28.10.26", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 13.9, 685388, "РЛ (2020, посл. 23.09.2020, 1 УПД): 426х8-9 (7.5т), 219х6 (6.4т)"], ["3661017898", "ООО \"ЖИЛСТРОЙИНВЕСТ\"", "394042, Воронежская область, г. Воронеж, пер. Серафимовича, д. 4", "Роман Сергеевич", "+7 (473) 237-96-84\n+7 (473) 226-25-64\n+7 (473) 290-38-19\n+7 (473) 223-44-97", "ejss01@mail.ru", "Строительство", "Нет", "", "", "", "12.11.2026", "", "[⏳ В Ожидание - 29.10.2026 13:40]", "📧 Прайс не высылался", "🔹 C (Малый)", 352.79, 604443, "КП (2015–2018, посл. 07.06.2018, 5 УПД): Кольца рез. САМ (Д300) (132.0т), ВТ6 300*5000 (х\\ц Трубы+муфт (78.0т), ВМ6 300 САМ (Муфты) (66.0т), ВТ6300*5000 (х\\ц трубы) (66.0т), ВТ6 400*5000 (х\\ц Трубы+муфт (10.0т), 530х8-9 (0.5т)"], ["4810002282", "АО \"РОЖДЕСТВЕНСКИЙ КАРЬЕР\"", "399675, Липецкая область, м. о. Краснинский, п. Лески, ул. Торговая, д. 23", "Дмитрий Николаевич", "+7 (474) 676-77-80\n+7 (474) 696-42-46\n+7 (474) 696-42-10", "rogdkar@mail.ru\nkrsnrczn@admlr.lipetsk.ru", "Другое", "Нет", "", "", "", "12.11.2026", "", "[⏳ В Ожидание - 29.10.2026 14:38]", "📧 Прайс не высылался", "🔹 C (Малый)", 31.3, 547959, "КП (2015, посл. 12.03.2015, 2 УПД): 273х5 (28.1т), 219х6 (3.2т)"], ["3625012654", "ООО \"АГРОТЕХ-ГАРАНТ\" ЗАДОНЬЕ", "396037, Воронежская область, Рамонский район, с. Скляево, ул. Центральная, д. 58, офис 5", "Григорий Викторович", "+7 (473) 404-91-17\n+7 (910) 732-13-72", "atg-zadonye@mail.ru", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 16.15, 532950, "КП (2018, посл. 06.08.2018, 2 УПД): 168х7-8 (12.4т), 168х7,5 (3.7т)"], ["3663120481", "ООО \"РЭС-36\"", "396336, Воронежская область, Новоусманский район, п. Отрадное, пер. Гололобовский, зд. 10, помещ. 1", "Тимур\nАртём", "+7 (920) 225-34-82\n+7 (980) 548-11-52", "timur.vrn@list.ru\ninfo@res-36.ru", "Теплосети / Энергетика", "Нет", "", "", "📌 По закупкам в основном Артём Берут трубу под…", "23.10.2026", "", "[⏳ В Ожидание | Профиль: Восстановленная, Лежалая • Конструкции / Опоры, Сваи (из «1. Обзвон (База)») - 21.10.2026 14:23]\n[💬 КП в Мессенджер: Полный (для: Артём) - 21.10.2026 14:20]\n[📧 Прайс: Полный + PDF (для: Артём) - 21.10.2026 14:19]", "✅ Прайс отправлен", "🔹 C (Малый)", 11.01, 430049, "КП+РЛ (2018–2019, посл. 04.07.2019, 3 УПД): 273х6 (4.2т), 273х8 (3.9т), 219х7 (2.2т), 325х8 (0.7т)"], ["3663105116", "ООО УК \"ФНК\"", "394018, Воронежская область, г. Воронеж, ул. Свободы, д. 37", "Марина Николаевна", "+7 (473) 261-18-15\n+7 (473) 296-08-18", "ooo.uk.fnk@mail.ru", "", "", "", "", "📌 ООО \"СМУ- 77\" 3663105116 Лаврентьев Сергей Николаевич бывший кит", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 295.88, 307164, "КП (2015, посл. 16.06.2015, 1 УПД): 325х6,5 (295.9т)"], ["3666142099", "ООО \"ВОРОНЕЖГЕОЛОГИЯ\"", "394077, Воронежская область, г. Воронеж, ул. Генерала Лизюкова, д. 39, помещ. 1", "Копейкин Юрий Михайлович", "+7 (473) 224-50-90\n+7 (473) 279-33-57\n+7 (473) 279-01-76\n+7 (473) 234-72-60\n+7 (473) 242-50-21", "vorongeo@mail.ru", "Водоканал", "Нет", "", "", "", "13.11.2026", "", "[⏳ В Ожидание - 29.10.2026 14:48]", "📧 Прайс не высылался", "🔹 C (Малый)", 13.04, 286579, "КП (2015–2018, посл. 24.01.2018, 3 УПД): 530х9 (6.5т), 630х8 (4.2т), 426х6,5 (2.3т)"], ["3664139830", "ООО \"СТРОЙАГРОИНВЕСТ\"", "394053, Воронежская область, г. о. город Воронеж, г. Воронеж, пр-кт Московский, д. 90, к. 1, кв. 606", "Игорь Николаевич", "+7 (903) 650-36-15", "31051958@mail.ru", "", "Нет", "", "", "", "13.11.2026", "", "[⏳ В Ожидание - 29.10.2026 14:56]", "📧 Прайс не высылался", "🔹 C (Малый)", 8.89, 246055, "КП+РЛ (2016–2019, посл. 24.06.2019, 3 УПД): 159х6 (6.4т), 159х4-5 (2.5т)"], ["3233010335", "ООО \"БРЯНСКПРОМБУРВОД-КОМПЛЕКТ\"", "241022, Брянская область, г. Брянск, бульвар Щорса, д. 31", "Михаил Юрьевич", "+7 (910) 333-26-68", "pto@bpbv.ru\nkomplekt@bpbv.ru", "", "Нет", "", "", "", "09.11.2026", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 5.97, 218502, "КП (2018, посл. 11.05.2018, 1 УПД): 273х6 (6.0т)"], ["3612006090", "ООО \"КАНТЕМИРОВСКОЕ ДРСУ\"", "396792, Воронежская область, м. р-н Богучарский, г. п. город Богучар, г. Богучар, нп. Военный Городок, д. 55, помещ. 25", "Алексей Николаевич / Артём", "+7 (960) 127-00-36\n+7 (903) 654-45-79", "kdrsu36@mail.ru\nyancen_artem@mail.ru\nsea@dorsu.ru\ninfo@dorsu.ru\ndrsu.ooo@mail.ru", "Дорожное / Мостостроение", "Нет", "", "", "", "", "", "[⏳ Перевод в Ожидание (из «1. Обзвон (База)») - 14.10.2026 15:00]", "📧 Прайс не высылался", "🔹 C (Малый)", 5.53, 209435, "КП (2020, посл. 09.10.2020, 1 УПД): 426х9 (4.0т), 630х9 (1.5т)"], ["3620007636", "ЗАО \"АГРОФИРМА ПАВЛОВСКАЯ НИВА\"", "396422, Воронежская область, м. р-н Павловский, г. п. город Павловск, г. Павловск, ул. Покровская, д. 11", "Владимир", "+7 (919) 185-64-73\n+7 (473) 622-44-28", "dobrsnab@niva.vrn.ru\nagrofirma@land.ru\nniva@niva.vrn.ru\nbelousova@niva.vrn.ru\nit@niva.vrn.ru", "", "Нет", "", "", "", "09.11.2026", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 7.92, 156918, "КП (2015–2020, посл. 19.05.2020, 2 УПД): 114х5 (6.7т), 325х7 (1.2т)"], ["3113001427", "ООО \"БЕЛЗНАК\"", "309300, Белгородская область, Ракитянский район, п. Пролетарский, ул. Железнодорожная, д. 50", "Сергей", "+7 (472) 238-08-32", "omts@belznak.com", "", "Нет", "", "", "", "12.11.2026", "", "[⏳ В Ожидание - 29.10.2026 15:09]", "📧 Прайс не высылался", "🔹 C (Малый)", 3.76, 139223, "КП (2018, посл. 02.08.2018, 1 УПД): 377х9,5 (2.4т), 630х8 (1.4т)"], ["3602009038", "ООО \"ВОДМОНТАЖСТРОЙ\"", "397700, Воронежская область, Бобровский район, г. Бобров, ул. Авдеева, д. 58", "Александр Александрович", "+7 (473) 504-34-84\n+7 (951) 559-89-27\n+7 (473) 504-14-06", "vodokanal75@mail.ru\nvodokanal70@mail.ru", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 2.94, 138086, "РЛ (2019, посл. 10.04.2019, 1 УПД): 219х6 (2.9т)"], ["3128092466", "ООО \"ТЕХНОСТРОЙ\"", "309508, Белгородская область, г. Старый Оскол, ул. Молодежная, д. 1А, кв. 11", "Пинской Роман Валериевич", "+7 (472) 547-00-15\n+7 (472) 526-37-90", "", "", "Нет", "", "", "", "13.11.2026", "", "[⏳ В Ожидание - 29.10.2026 14:59]", "📧 Прайс не высылался", "🔹 C (Малый)", 46.0, 135271, "КП (2015, посл. 22.06.2015, 1 УПД): 426х8 (46.0т)"], ["7743944097", "ООО \"ТЕХСТРОЙ\"", "117420, г. Москва, вн. тер. г. муниципальный округ Черемушки, ул. Намёткина, д. 11, помещ. I, ком. 22", "Алексей Владимирович", "+7 (499) 519-03-43\n+7 (987) 401-15-99\n+7 (843) 230-04-30\n+7 (843) 230-07-30", "info@tehstroi.ru\nolga@tehstroi.ru", "Металлоконструкции / Завод", "Нет", "", "", "📌 Завод по производсту полимерных труб", "05.11.2026", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 5.3, 105491, "КП (2015, посл. 17.02.2015, 1 УПД): 630х7,0-8 (5.3т)"], ["3624003329", "ЗАО \"АГРОФИРМА АПРОТЕК - ПОДГОРЕНСКАЯ\"", "396560, Воронежская область, Подгоренский район, пгт Подгоренский, ул. Первомайская, д. 60", "Владимир Сергеевич", "+7 (473) 626-12-31\n+7 (473) 300-32-09", "mail@niva.vrn.ru\nbelousova@niva.vrn.ru\npodgornoe@niva.vrn.ru\nzubkov@niva.vrn.ru\nkadri@niva.vrn.ru\nniva@mail.ru", "", "Нет", "", "", "", "13.11.2026", "", "[⏳ В Ожидание - 29.10.2026 15:02]", "📧 Прайс не высылался", "🔹 C (Малый)", 3.27, 69237, "КП (2016, посл. 15.03.2016, 1 УПД): 325х9 (3.3т)"], ["4825137518", "ООО \"МНОГОПРОФИЛЬНОЕ ПРЕДПРИЯТИЕ СТРОЙСОЮЗ\"", "398042, Липецкая область, г. о. город Липецк, г. Липецк, Универсальный проезд, влд. 14, помещ. 9-2, офис 9/7", "Дмитрий Анатольевич", "+7 (980) 250-04-52", "mpc048@yandex.ru", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 1.85, 65739, "РЛ (2020, посл. 13.05.2020, 1 УПД): 820х9 (1.9т)"], ["7701045732", "ПАО \"МОСТОТРЕСТ\"", "121087, г. Москва, ул. Барклая, д. 6, стр. 5", "Евгений Владимирович", "+7 (499) 178-10-77\n+7 (495) 669-79-99\n+7 (495) 626-80-34\n+7 (495) 915-54-83\n+7 (863) 245-28-60", "efedoseeva@mostro.ru\nmostro@mostro.ru\necp@mostro.ru\nelalekseeva@mostro.ru\nnashkatova@mostro.ru", "Мостостроение", "Нет", "", "", "📌 Работают через НПС снабжение", "12.11.2026", "", "[⏳ В Ожидание - 29.10.2026 15:08]", "📧 Прайс не высылался", "🔹 C (Малый)", 3.21, 62548, "КП (2015, посл. 04.09.2015, 1 УПД): 630х8 (3.2т)"], ["3651005974", "ООО \"ЭЛЕКТРОМОНТАЖ\"", "396070, Воронежская область, г. Нововоронеж, промзона Южная, д. 7А", "Виктор Николаевич", "+7 (473) 642-86-87\n+7 (473) 645-09-21\n+7 (952) 958-37-16", "\"elektromontazhnv@mail.ru", "Энергетика", "Нет", "", "", "", "13.11.2026", "", "[⏳ В Ожидание - 29.10.2026 15:02]", "📧 Прайс не высылался", "🔹 C (Малый)", 2.16, 41598, "КП (2015, посл. 26.03.2015, 1 УПД): 219х7 (2.2т)"], ["3661080829", "ООО \"САНВЕНТЭЛ\"", "394033, Воронежская область, г. Воронеж, пр-кт Ленинский, д. 119А, офис 210", "Безносько Юлия Артуровна", "+7 (473) 252-76-15\n+7 (915) 589-94-90\n+7 (473) 260-56-59", "", "Строительство", "Нет", "", "", "", "16.11.2026", "", "[⏳ В Ожидание - 29.10.2026 15:04]", "📧 Прайс не высылался", "🔹 C (Малый)", 1.46, 38450, "КП (2015–2020, посл. 04.02.2020, 3 УПД): 530х8 (0.7т), 325х8 (0.7т), 273х8 (0.1т)"], ["3662204227", "ООО ТПК \"МАКСИМУМ\"", "394019, Воронежская область, г. Воронеж, ул. Малаховского, д. 52, помещ. 10", "Дмитрий", "+7 (473) 300-30-53", "snab2@tpk36.ru", "", "Нет", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 0.26, 10218, "РЛ (2018, посл. 28.11.2018, 1 УПД): 273х8 (0.3т)"], ["3661018891", "ЗАО СМП \"ЭЛЕКТРОНЖИЛСОЦСТРОЙ\"", "394042, Воронежская область, г. Воронеж, пер. Серафимовича, д. 4", "Гайдай Юрий Федосиевич", "+7 (473) 226-88-27\n+7 (473) 226-28-08\n+7 (473) 237-96-60\n+7 (473) 226-25-64\n+7 (473) 256-40-15", "enjs.s.ok@gmail.com\n610262@nalog.ktkt.ru\ninfo@ejss.ru\nejss@vmail.ru\nofea-ejss@yandex.ru", "Строительство", "", "", "", "📌 Кит", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 0.51, 9415, "КП (2015, посл. 29.06.2015, 1 УПД): 325х6 (0.5т)"], ["4800012450", "ООО \"ТЕХНОВЕК\"", "398016, Липецкая область, г. о. город Липецк, г. Липецк, ул. Нестерова, влд. 17, помещ. 10", "Бурцев Дмитрий Александрович", "+7 (952) 591-81-28", "technovek48@mail.ru", "Металлоконструкции", "", "", "", "", "", "", "", "📧 Прайс не высылался", "⚪ Без отгрузок", "", "", "В 1С (КП/РЛ) отгрузок по этому ИНН нет"], ["2312274610", "ООО \"МОСТОБУРСТРОЙ\"", "350020, Краснодарский край, г. о. город Краснодар, г. Краснодар, ул. Одесская, д. 41, этаж 3, помещ. 57", "Алексей", "+7 (918) 930-96-10\n+7 (928) 408-87-29", "mostmbur_stroi@mail.ru\nmostmbur_kdr@mail.ru\nmbs.snab@yandex.ru", "Бурение", "", "", "", "", "", "", "", "📧 Прайс не высылался", "⚪ Без отгрузок", "", "", "В 1С (КП/РЛ) отгрузок по этому ИНН нет"], ["4028068969", "ООО \"КАЛУГАСПЕЦБУРЕНИЕ\"", "Калужская область, г. о. город Калуга, г. Калуга, ул. Черновская, зд. 54, стр. 1", "Александр Викторович", "+7 (910) 914-81-84", "viktorovav1982@gmail.com", "Бурение на воду", "", "", "", "", "", "", "", "📧 Прайс не высылался", "⚪ Без отгрузок", "", "", "В 1С (КП/РЛ) отгрузок по этому ИНН нет"], ["3663128579", "ООО \"СТРОЙРЕМБУРВОД\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔥 A (VIP)", 463.15, 30457425, "КП+РЛ (2017–2026, посл. 21.07.2026, 83 УПД): 219х8 (83.9т), 630х9 (44.1т), 530х8 (42.3т), 325х8 (39.8т), 426х8 (39.2т), 426х9 (30.3т)"], ["3664231385", "ООО \"ТрубСтройИнновация\"", "394090, Воронежская область, г. о. город Воронеж, г. Воронеж, ул. Ростовская, д. 80А, кв. 6", "Валерий Ч.", "", "", "Металлобаза", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔥 A (VIP)", 102.18, 4980967, "КП+РЛ (2017–2026, посл. 06.10.2026, 16 УПД): 530х8 (28.2т), 325х10 (20.0т), 1020х10,0-13,5 (12.8т), 530х7 (12.1т), 630х8 (8.0т), 1220х20 (6.3т)"], ["3662184475", "ООО \"Акватехпласт\"", "394019, Воронежская область, г. Воронеж, ул. Подклетенская, д. 34", "", "", "", "Металлобаза", "", "", "", "", "", "", "", "📧 Прайс не высылался", "⭐ B (Постоянный)", 19.35, 1276816, "КП+РЛ (2025–2026, посл. 15.06.2026, 3 УПД): 426х8 (12.2т), 325х8 (6.5т), 325х12 (0.7т)"], ["3665834210", "ООО «ГЕОБУРВОД»", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 5.1, 355424, "КП (2026, посл. 21.08.2026, 2 УПД): 325х6 (3.3т), 273х8 (1.8т)"], ["3662068486", "ООО «Металлопрокат»", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 2.33, 170166, "КП (2026, посл. 12.08.2026, 2 УПД): 426х8 (1.9т), 219х8 (0.4т)"], ["366510342577", "ИП Гущин Сергей Викторович", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 1.55, 68551, "РЛ (2026, посл. 06.04.2026, 1 УПД): 820х9 (1.5т)"], ["3102047405", "ООО ТД \"БТК\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔥 A (VIP)", 785.74, 57938858, "КП+РЛ (2020–2025, посл. 25.07.2025, 52 УПД): 273х8 (157.7т), 630х10 (100.1т), 219х7 (69.0т), 325х8,0-9 (47.6т), 325х8 (47.2т), 820х10 (45.6т)"], ["3666262004", "ООО \"АНТЕЙ СПЕЦКОМПЛЕКТ\"", "394036, Воронежская область, г. о. город Воронеж, г. Воронеж, ул. 20-Летия Влксм, д. 43А, кв. 1", "Чубина Анна Анатольевна", "", "", "Металлобаза", "", "", "", "", "", "", "", "📧 Прайс не высылался", "⭐ B (Постоянный)", 32.6, 1973138, "КП+РЛ (2022–2025, посл. 17.10.2025, 4 УПД): 530х8 (18.1т), 630х8,0-9 (9.8т), 273х5 (4.8т)"], ["3119007707", "ООО \"ТЕХМЕХСЕРВИС\"", "309560, Белгородская область, Чернянский район, п. Чернянка, ул. Южная, д. 15", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "⭐ B (Постоянный)", 25.58, 1797259, "КП (2024–2025, посл. 06.02.2025, 2 УПД): 273х8 (5.6т), 530х9 (5.3т), 159х5 (3.5т), 219х8 (2.4т), 820х9 (2.2т), 426х8 (2.0т)"], ["3662280820", "ООО \"СТРОЙБУРВОД-ПЛЮС\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 8.29, 471600, "КП+РЛ (2021–2025, посл. 05.06.2025, 6 УПД): 60х5 (4.4т), 273х6 (1.8т), 89х6,5 НКТ (1.5т), 60х5 НКТ (0.6т)"], ["6150099423", "ООО \"Промстрой\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 6.92, 430285, "КП (2025, посл. 09.10.2025, 1 УПД): 377х9 леж. (6.2т), Труба НКТ 73 (0.7т)"], ["3665128662", "ООО \"Металл36\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 17.69, 393544, "КП (2024–2025, посл. 17.06.2025, 2 УПД): 426х8 (12.6т), 630х9 (5.1т)"], ["6829160832", "Общество с ограниченной ответственностью \"РЕЯ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 3.05, 210841, "КП (2025, посл. 02.10.2025, 2 УПД): 159х6 (2.2т), 219х7 (0.9т)"], ["3666261272", "Общество с ограниченной ответсвенностью \"РОМБ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 1.13, 81578, "РЛ (2025, посл. 22.10.2025, 2 УПД): 530х8 (0.9т), 530х6 (0.2т)"], ["3604021440", "ООО \"БСМ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 1.09, 60115, "КП (2025, посл. 08.10.2025, 1 УПД): 273х7 (1.1т)"], ["262814043770", "ИП Ярошенко Даниил Олегович", "Ставропольский край, г. Кисловодск", "Даниил Олегович", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔥 A (VIP)", 34.21, 2962643, "КП (2024, посл. 21.11.2024, 2 УПД): 720х10 (33.5т), 325х10 (0.7т)"], ["4807028636", "ООО \"Достояние\"", "399780, Липецкая область, Елецкий район, с. Лавы, ул. Кирова, д. 29", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "⭐ B (Постоянный)", 28.24, 1628654, "КП (2024, посл. 14.05.2024, 2 УПД): 1020х12 (28.2т)"], ["3128149867", "ООО \"СК РЕСУРС\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "⭐ B (Постоянный)", 96.0, 1622880, "КП (2024, посл. 12.09.2024, 1 УПД): 530х10 ВУС (96.0т)"], ["3628013416", "ООО \"Стройполимер\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 1.91, 134234, "КП (2024, посл. 29.07.2024, 1 УПД): Труба 426 (1.9т)"], ["7727282030", "ООО СК  \"Экотех\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 1.77, 114855, "КП (2024, посл. 02.05.2024, 1 УПД): 530х8 (1.8т)"], ["3420000810", "Крестьянское(Фермерское)Хозяйство \"Губина Андрея Андреевича\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 10.0, 538500, "РЛ (2023, посл. 28.08.2023, 1 УПД): 73х5,5 (10.0т)"], ["3661173632", "ООО \"ЭПМ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 5.04, 312418, "КП (2023, посл. 07.12.2023, 1 УПД): 219х6 (5.0т)"], ["342001596760", "ИП Глава КФХ Губин Алескей Андреевич", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 5.2, 312240, "РЛ (2023, посл. 03.07.2023, 1 УПД): 219х6 (3.2т), 273х5,0-6 (2.0т)"], ["362003810250", "Индивидуальный предприниматель Харцызов Роман Александрович", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 0.99, 82170, "КП (2023, посл. 29.03.2023, 1 УПД): 530х7 (1.0т)"], ["3665818183", "ООО \"ТТС\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 0.96, 62882, "РЛ (2023, посл. 18.01.2023, 1 УПД): 530х8 (0.7т), 820х11 (0.2т)"], ["3663045354", "ООО \"СтройПолимерМонтаж\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔥 A (VIP)", 86.06, 3779360, "КП (2020–2022, посл. 01.08.2022, 6 УПД): 1020х14 (46.2т), 1220х14 (25.0т), 219х8 (13.1т), 530х12 (1.8т)"], ["3666189442", "ООО \"ТПК СТАЛЬСОЮЗ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "⭐ B (Постоянный)", 20.14, 1319498, "РЛ (2022, посл. 14.03.2022, 1 УПД): 325х8 (20.1т)"], ["6828008320", "ООО \"МеталлСтрой\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 7.25, 504518, "КП (2022, посл. 29.07.2022, 1 УПД): 530х9 (2.6т), 820х9 (2.0т), 720х10 (2.0т), 325х7 (0.6т)"], ["3664113254", "ООО \"АСД-Строй\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 2.66, 194271, "КП (2022, посл. 04.04.2022, 1 УПД): 273х7 (2.7т)"], ["3623006937", "ООО ИТЦ \"Нефтемаш-инжиниринг\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "⭐ B (Постоянный)", 27.76, 1038872, "РЛ (2021, посл. 20.05.2021, 2 УПД): 1220х15,7 (16.0т), 1220х13 (11.8т)"], ["1512017993", "ООО \"ТРУБОТОРГ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 18.89, 968175, "РЛ (2021, посл. 07.07.2021, 1 УПД): 530х10 (18.9т)"], ["366100490074", "ИП ГОНЧАРОВ ВЛАДИМИР ИВАНОВИЧ", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 11.63, 904109, "РЛ (2021, посл. 12.07.2021, 2 УПД): 219х7 (7.5т), 219х8 (4.2т)"], ["3127507380", "ООО \"ДСПИ ГРУП\"", "309185, Белгородская область, г. Губкин, ул. Вокзальная, д. 2Б", "Давид Георгиевич", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 16.61, 780923, "КП (2018–2021, посл. 04.02.2021, 2 УПД): 426х11 (12.9т), 530х8,5 (2.5т), 530х8 (1.2т)"], ["3616023480", "ООО \"ЛИОТ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 17.2, 774000, "РЛ (2021, посл. 23.09.2021, 1 УПД): 1220х12 (17.2т)"], ["3663115097", "ООО \"Стройбурвод\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 14.42, 524306, "РЛ (2018–2021, посл. 19.05.2021, 7 УПД): 377х10 (5.6т), 630х7 (3.7т), 426х5,5 (2.2т), 60х5 (2.1т), 530х8 (0.7т), 325х6 (0.2т)"], ["4825025620", "ООО \"ДСУ №7\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 7.86, 455880, "КП (2021, посл. 30.06.2021, 1 УПД): 325х8 (7.9т)"], ["3604018292", "ООО \"СтарКом плюс\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 7.0, 419820, "КП (2021, посл. 20.08.2021, 1 УПД): 426х10 (7.0т)"], ["3665806501", "ООО ТОРГОВЫЙ ДОМ \"ТРУБТЕХСЕРВИС\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 5.9, 272451, "РЛ (2021, посл. 23.06.2021, 1 УПД): 820х9 (5.9т)"], ["3666226020", "ООО \"ТРУБОСТАЛЬ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 4.0, 180000, "КП (2021, посл. 01.06.2021, 1 УПД): 820х9 (4.0т)"], ["3123099482", "ЗАО \"ПКФ\"СТРОЙСНАБ-БЕЛОГОРЬЕ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 2.18, 151669, "КП (2021, посл. 30.06.2021, 1 УПД): 426х10 (2.2т)"], ["3665089117", "ООО \"МЕТАЛЛТРУБКОМПЛЕКТ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 2.22, 100000, "КП (2021, посл. 08.11.2021, 1 УПД): 1020х11 (2.2т)"], ["4826030207", "ООО \"Липецксантехмонтаж-1\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 1.0, 45816, "РЛ (2021, посл. 02.09.2021, 1 УПД): 325х6 (1.0т)"], ["3661085425", "ООО \"ПРОМСТРОЙ-ИНВЕСТ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 0.5, 20205, "КП (2020–2021, посл. 14.01.2021, 3 УПД): 630х7 (0.2т), 530х8 (0.2т), 219х7 (0.2т)"], ["3662129770", "ООО \"Строй Ресурс\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 0.59, 19371, "КП (2021, посл. 02.04.2021, 1 УПД): 630х8 (0.3т), 630х7 (0.3т)"], ["3123174718", "ООО \"СоюзБурвод\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔥 A (VIP)", 1351.75, 51353177, "КП+РЛ (2015–2020, посл. 16.10.2020, 85 УПД): 273х8,5 (175.9т), 530х12 (147.6т), 273х8 (116.1т), 325х8 (96.5т), 273х6 (80.4т), 530х8 (61.1т)"], ["3123372572", "ООО \"БелИнфор\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔥 A (VIP)", 377.81, 14976185, "КП+РЛ (2018–2020, посл. 20.03.2020, 23 УПД): 273х6 (76.5т), 530х12 (40.1т), 820х10 (39.0т), 325х10 (37.9т), 273х8 (37.7т), 426х8,5 (20.3т)"], ["360900892582", "Глава КФХ Чепилевич Александр Петрович", "Воронежская область, с. Новомакарово", "Александр Петрович", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "⭐ B (Постоянный)", 69.59, 2220234, "КП (2016–2020, посл. 07.08.2020, 6 УПД): 159х6 (35.9т), 219х6 (33.7т)"], ["3665088882", "ООО \"ТРУБСТРОЙИННОВАЦИЯ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "⭐ B (Постоянный)", 48.8, 1304404, "РЛ (2019–2020, посл. 02.12.2020, 2 УПД): 630х7,0-8 (40.0т), 630х8 (8.8т)"], ["4826108622", "ООО Система", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 19.2, 708780, "КП+РЛ (2020, посл. 15.04.2020, 2 УПД): 273х8 (11.1т), 1020х14 (4.6т), 426х10 (3.5т)"], ["3663002079", "ЗАО \"Ремсельбурвод\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 11.25, 361718, "КП (2018–2020, посл. 09.09.2020, 6 УПД): 720х9,2 (5.6т), 720х9,5 (1.9т), 630х9 (1.7т), 630х8 (1.5т), 377х10 (0.6т)"], ["3123123248", "ООО \"МЕЛИОВОДПРОЕКТ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 7.63, 301142, "КП (2020, посл. 27.08.2020, 2 УПД): 530х8 (4.6т), 219х8 (3.0т)"], ["3128039399", "ООО \"ДАЛС\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 7.02, 244662, "РЛ (2020, посл. 22.06.2020, 1 УПД): 530х8 (7.0т)"], ["5262224374", "ООО \"ВОЛГАСПЕЦСТРОЙ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 3.87, 159023, "КП (2020, посл. 19.11.2020, 1 УПД): 219х7 (3.9т)"], ["3662243755", "ООО \"ИНСИКО\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 3.67, 151134, "КП (2020, посл. 10.08.2020, 1 УПД): 630х7 (2.3т), 630х8 (1.4т)"], ["3604018704", "ООО \"Стройком\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 4.05, 139622, "КП (2020, посл. 30.04.2020, 1 УПД): 820х9 (4.0т)"], ["3420006138", "ООО \"КХ ТРОФИМОВА В.Г\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 3.4, 122060, "КП (2020, посл. 11.12.2020, 1 УПД): 159х6 (3.4т)"], ["3628015420", "ООО \"Сферастек\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 2.58, 100503, "РЛ (2020, посл. 25.08.2020, 1 УПД): 530х7 (2.0т), 273х7 (0.5т)"], ["3664118502", "ООО \"ВОРОНЕЖПРОМТЕХНОЛОГИИ\"", "394007, Воронежская область, г. о. город Воронеж, г. Воронеж, ул. Ленинградская, д. 2, ком. 411А", "Наталия", "", "efr-online@yandex.ru\nvrnpromtech@mail.ru\ninfo@prom-texno.ru\nnataliya.k@fiting-v.ru", "Строительство", "", "", "", "", "26.10.2026", "", "[📞 Недозвон: 2/4 - 21.10.2026 10:32]", "📧 Прайс не высылался", "🔹 C (Малый)", 2.69, 82117, "РЛ (2019–2020, посл. 25.03.2020, 2 УПД): 530х8 (2.3т), 273х5 (0.4т)"], ["3661051507", "ООО \"Фарадей Групп\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 1.41, 58155, "КП (2020, посл. 15.09.2020, 1 УПД): 273х6 (1.4т)"], ["3610007669", "ЗАО ПОБЕДА", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 1.18, 44726, "КП (2020, посл. 17.11.2020, 1 УПД): 325х6 (1.0т), 325х8 (0.1т)"], ["4632077887", "ООО \"ПромКомплект\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 0.77, 30720, "КП (2020, посл. 10.08.2020, 1 УПД): 219х8 (0.8т)"], ["4806016437", "ООО \"СЕМЬСОТ ДОРОГ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 0.39, 15366, "КП (2020, посл. 12.05.2020, 1 УПД): 325х8 (0.4т)"], ["4634013286", "ООО \"ВПМ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 0.03, 1292, "КП (2020, посл. 19.11.2020, 1 УПД): 426х8 (0.0т)"], ["6155065562", "ООО \"МТМ\"", "346513, Ростовская область, г. Шахты, пер. Киквидзе, д. 4", "Любовь Георгиевна", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "⭐ B (Постоянный)", 84.23, 2227804, "КП (2019, посл. 18.02.2019, 4 УПД): 630х7,0-8 (70.7т), 630х8 (13.5т)"], ["3662226781", "ООО \"ВОРОНЕЖГЕОЛОГИЯ К\"", "394077, Воронежская область, г. Воронеж, ул. Генерала Лизюкова, д. 39, помещ. 1", "Юрий Михайлович", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "⭐ B (Постоянный)", 47.95, 1652314, "КП+РЛ (2018–2019, посл. 15.03.2019, 12 УПД): 325х8 (18.8т), 530х8 (6.0т), 530х8-9 (5.7т), 630х8 (5.3т), 630х7,0-8 (3.8т), 426х6 (2.8т)"], ["3123401953", "ООО \"Бурводстрой 31\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 14.91, 616832, "РЛ (2019, посл. 25.06.2019, 2 УПД): 325х6-7 (7.1т), 1020х12 (5.1т), 426х7-8 (2.7т)"], ["4823060161", "ООО \"ХЕЛФЕР\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 16.31, 498562, "КП+РЛ (2017–2019, посл. 30.08.2019, 8 УПД): 325х7 (7.9т), 530х7 (5.3т), 530х12 (1.8т), 426х7 (1.2т), 273х8 (0.1т)"], ["462000100343", "ИП ГРЯЗНОВ АЛЕКСАНДР НИКОЛАЕВИЧ", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 11.17, 394167, "РЛ (2019, посл. 24.04.2019, 1 УПД): 530х12 (6.9т), 1020х11 (3.0т), 325х8 (1.3т)"], ["3128074379", "ООО ТД \"СКОЛОТ\"", "309512, Белгородская область, г. Старый Оскол, мкр Конева, д. 17, кв. 314", "Алексей Алексеевич", "", "461507@mail.ru", "Металлоконструкции", "Нет", "", "", "", "13.11.2026", "", "[⏳ В Ожидание - 29.10.2026 14:44]", "📧 Прайс не высылался", "🔹 C (Малый)", 14.32, 370592, "КП (2015–2019, посл. 16.09.2019, 5 УПД): 325х7 (3.6т), 426х8 (3.4т), 1020х11 (3.3т), 426х7 (1.5т), 426х10 (0.9т), 426х9 (0.8т)"], ["3123400212", "ООО \"Геоцентр \"Черноземье\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 2.68, 112686, "РЛ (2019, посл. 25.10.2019, 1 УПД): 325х9 (1.4т), 325х8 (1.3т)"], ["3119007880", "ООО \"Масло Потаповское\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 2.86, 105600, "КП (2017–2019, посл. 03.10.2019, 2 УПД): 630х8 (2.9т)"], ["3662252125", "ООО \"СТРОЙПРОМСНАБ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 2.8, 100656, "КП (2019, посл. 15.07.2019, 1 УПД): 325х8 (2.8т)"], ["3121070417", "ЗАО \"КРАСНЕНСКОЕ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 2.33, 82290, "КП (2018–2019, посл. 19.09.2019, 2 УПД): 325х8 (1.7т), 325х7 (0.7т)"], ["3662219713", "ООО \"ПЛАСТПОЛИМЕР\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 1.39, 48545, "РЛ (2019, посл. 26.08.2019, 1 УПД): 630х8 (1.4т)"], ["360101886335", "СВИСТОВ ГЕННАДИЙ АЛЕКСАНДРОВИЧ (ИП ГКФХ)", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 1.49, 46252, "КП (2019, посл. 27.08.2019, 1 УПД): 114х5 (1.5т)"], ["3666221400", "ООО \"Амулет-Гидро\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 0.5, 25578, "КП (2019, посл. 23.08.2019, 1 УПД): 219х8 (0.5т)"], ["3123404520", "ООО \"ТехСтройМонтаж\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 0.7, 22968, "КП (2019, посл. 07.06.2019, 1 УПД): 426х6 (0.7т)"], ["7448151460", "ООО \"ТРУБОСТАЛЬ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 1.0, 15000, "РЛ (2019, посл. 22.05.2019, 1 УПД): 426х7 (1.0т)"], ["3662182365", "ООО \"СКС\"", "394016, Воронежская область, г. Воронеж, ул. 45 Стрелковой Дивизии, д. 236А, офис 7", "Алексей Егорович", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 51.5, 964583, "КП (2018, посл. 12.02.2018, 3 УПД): 530х8 (35.7т), 630х8 (15.8т)"], ["3663113815", "ООО \"ЭНЕРГОСТАЛЬ ИНВЕСТ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 21.82, 635162, "КП+РЛ (2016–2018, посл. 04.07.2018, 6 УПД): 630х8 (6.3т), 1020х14 (5.3т), 530х9 (3.8т), 325х10 (3.5т), 273х6 (2.2т), 219х8 (0.7т)"], ["3664208918", "ООО \"Альтернатива\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 17.26, 351071, "КП (2018, посл. 23.01.2018, 1 УПД): 325х8 (17.3т)"], ["6150078617", "ООО \"Агрегатка-Сервис\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 7.92, 273309, "КП (2018, посл. 31.05.2018, 1 УПД): 530х9 (7.9т)"], ["3666166861", "ООО \"Виктория\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 61.95, 256951, "КП (2017–2018, посл. 21.06.2018, 6 УПД): 325х6 ВУС (28.7т), 325х8 ВУС (20.6т), 325х6 (8.0т), 325х8 (3.3т), 114х5 (1.0т), 426х7 (0.4т)"], ["3123430070", "ООО \"КОМПАНИЯ ТЕРМИКА\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 7.88, 231279, "КП (2018, посл. 06.04.2018, 1 УПД): 530х8-9 (4.9т), 273х8 (2.4т), 114х5 (0.5т)"], ["6122018736", "ООО \"ПРОМСТРОЙ\"", "346815, Ростовская область, Мясниковский район, с. Султан-Салы, ул. Налбандяна, д. 8", "Виталий Александрович", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 5.34, 198010, "КП (2018, посл. 18.06.2018, 2 УПД): 530х9 (2.7т), 720х9,5 (1.9т), 325х8 (0.7т)"], ["615000439262", "Карастан Андрей Павлович(ИП)", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 5.15, 172659, "КП (2018, посл. 25.05.2018, 1 УПД): 630х8 (5.2т)"], ["3666112873", "ООО \"Спецпромстрой\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 3.46, 136844, "КП (2016–2018, посл. 28.09.2018, 2 УПД): 325х8 (1.9т), 273х8 (1.6т)"], ["3663126892", "ООО \"МИР\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 3.63, 109915, "КП (2017–2018, посл. 02.02.2018, 3 УПД): 426х7 (2.2т), 426х8 (1.4т)"], ["3601008419", "ООО \"Агротех-Гарант Пугачевский\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 3.31, 109395, "КП (2018, посл. 16.07.2018, 1 УПД): 168х7-8 (3.3т)"], ["3629007648", "ООО \"АГРОТЕХ-ГАРАНТ СЛАВЯНСКИЙ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 2.93, 96558, "КП (2018, посл. 01.08.2018, 1 УПД): 168х7,5 (2.9т)"], ["3666225108", "ООО \"СПЕЦСНАББИЗНЕС\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 2.54, 74812, "РЛ (2018, посл. 03.10.2018, 1 УПД): 530х8,5 (2.5т)"], ["3662215268", "ООО \"АСТК\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 2.35, 65940, "КП (2018, посл. 30.07.2018, 1 УПД): 159х6 (2.4т)"], ["4825099686", "ООО \"КАПИТАЛCТРОЙ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 1.2, 35518, "РЛ (2018, посл. 02.11.2018, 1 УПД): 530х8 (1.2т)"], ["3662233764", "ООО \"ПОЛИГРУПП\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 0.96, 31548, "КП (2018, посл. 19.10.2018, 1 УПД): 530х8 (1.0т)"], ["3664225984", "ООО \"СТРОЙТЕХ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 0.99, 29988, "КП (2018, посл. 26.02.2018, 1 УПД): 273х8 (1.0т)"], ["4826120605", "ООО \"ПК \"ЛПС\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 0.8, 29121, "КП (2018, посл. 30.05.2018, 1 УПД): 530х8 (0.8т)"], ["3663070400", "ООО \"Стройтранс\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 0.83, 24900, "КП (2018, посл. 09.06.2018, 1 УПД): 20х2,8 (0.8т)"], ["7841026793", "ООО Санти", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 4.45, 14752, "КП (2018, посл. 11.10.2018, 1 УПД): 530х8 (4.5т)"], ["366101059287", "ИП Бганцев Михаил Алексеевич", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 0.31, 11196, "КП (2018, посл. 11.05.2018, 1 УПД): 426х7 (0.3т)"], ["3123399648", "ООО \"КОНТИНЕНТ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔥 A (VIP)", 160.0, 3941736, "КП (2017, посл. 16.11.2017, 11 УПД): 273х8 (85.4т), 426х8 (18.7т), 426х7 (18.5т), 273х6 (16.6т), 273х8,5 (13.9т), 426х6,5 (3.8т)"], ["3663111173", "ООО \"Стройбаза № 1\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "⭐ B (Постоянный)", 109.75, 2629095, "РЛ (2017, посл. 14.09.2017, 1 УПД): 630х8 (109.8т)"], ["3602005717", "ООО ВПФ \"ПССВ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "⭐ B (Постоянный)", 81.7, 1666867, "КП+РЛ (2015–2017, посл. 11.12.2017, 12 УПД): 273х8 (21.4т), 630х8 (16.6т), 720х9 (15.0т), 530х8 (11.9т), 530х8,5 (8.5т), 325х6,5 (4.1т)"], ["6829092903", "ООО \"Союз Агро\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "⭐ B (Постоянный)", 50.81, 1227167, "КП (2017, посл. 15.08.2017, 3 УПД): 273х8 (39.8т), 273х6 (9.1т), 426х8 (1.9т)"], ["342001010802", "НЕМУХИН ВИТАЛИЙ АЛЕКСАНДРОВИЧ (ИП ГКФХ)", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 74.48, 606209, "КП (2017, посл. 07.06.2017, 5 УПД): 720х10 (56.6т), 45х2,5-3 (10.2т), 57х4 (5.4т), 60х3,5 (2.3т)"], ["3662184387", "ООО \"РЕПЕР\"", "394029, Воронежская область, г. Воронеж, пр-кт Ленинский, д. 25/1, помещ. 18", "Александр Евгеньевич", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 179.67, 565383, "КП (2017, посл. 27.04.2017, 3 УПД): 273х6 (145.2т), 325х6 ВУС (34.5т)"], ["6383005763", "ООО \"ПРОМЕТЕЙ\"", "446031, Самарская область, г. Сызрань, пр-кт 50 Лет Октября, д. 59, кв. 5", "Павел Игоревич", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 15.98, 451076, "КП (2017, посл. 29.09.2017, 7 УПД): 426х6 (12.9т), 426х7 (3.1т)"], ["7726730916", "ООО \"УК АПОФЕМА\"", "127521, г. Москва, 17-й Марьиной Рощи проезд, д. 4, корп. 1, помещ. XVI, к. 19, оф. 1, эт. 6", "Дмитрий Олегович", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 20.79, 449340, "КП (2017, посл. 08.08.2017, 2 УПД): 630х8 (20.8т)"], ["3650002018", "ООО \"ФОРММАТЕРИАЛЫ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 18.1, 427659, "КП (2017, посл. 12.07.2017, 2 УПД): 530х10 (5.0т), 720х10 (5.0т), 720х9 (4.7т), 530х9 (2.7т), 630х8 (0.8т)"], ["3666074868", "ООО \"Центрчерноземгаз\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 260.0, 420627, "КП (2017, посл. 08.06.2017, 1 УПД): ПЭ 100 SDR 11 225*20.5 (260.0т)"], ["3663108050", "ООО \"КОРПОРАЦИЯ МЕТАЛЛИНВЕСТ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 15.53, 348767, "КП (2017, посл. 12.04.2017, 2 УПД): 325х8 (15.5т)"], ["6154129397", "ООО \"СЕВЕРО-КАВКАЗСКАЯ ТЕРРИТОРИАЛЬНАЯ КОМПАНИЯ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 16.0, 325525, "КП (2017, посл. 13.07.2017, 1 УПД): 299х4 (16.0т)"], ["3664133557", "ООО \"СТРОЙАГРОМАШ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 140.23, 254584, "КП (2017, посл. 11.07.2017, 3 УПД): 273х6 (92.0т), 273х8,5 (43.4т), 273х8 (4.8т)"], ["3665107140", "ООО \"ВОРОНЕЖСКАЯ СТРОИТЕЛЬНАЯ КОМПАНИЯ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 10.47, 225302, "КП (2016–2017, посл. 16.03.2017, 3 УПД): 325х8 (6.2т), 273х6,5 (2.0т), 273х8 (1.8т), 630х8 (0.4т)"], ["3665138607", "ООО \"ТрубТехСервис\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 10.24, 211038, "КП (2017, посл. 13.09.2017, 2 УПД): Цепь для Auto Picle S и Picl (4.0т), 530х9 (3.9т), 530х8 (2.3т)"], ["3662172705", "ООО \"ЭЛКОМ ПЛЮС\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 12.11, 200156, "КП (2017, посл. 09.06.2017, 1 УПД): 1220х12 (12.1т)"], ["3120101983", "ООО \"СПЕЦКОМПЛЕКТ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 5.88, 168194, "КП (2017, посл. 11.07.2017, 1 УПД): 219х8 (3.5т), 325х7 (1.9т), 325х6 ВУС (0.5т)"], ["3662248584", "ООО \"Талисман\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 8.08, 150625, "КП (2017, посл. 04.08.2017, 1 УПД): 273х8 (8.1т)"], ["3661034043", "ООО \"Спецжилуправление\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 4.36, 130901, "КП (2017, посл. 30.01.2017, 1 УПД): 1220х10 (3.5т), 1220х12 (0.9т)"], ["3662138990", "ООО \"НАВАСТРОЙ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 4.96, 107100, "КП (2017, посл. 13.09.2017, 1 УПД): 273х8 (5.0т)"], ["4823029468", "ООО \"ПРОМТЕХНИКА-2\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 35.31, 102853, "КП (2017, посл. 19.07.2017, 2 УПД): 325х6 ВУС (33.9т), 530х9 (1.3т), 530х6 (0.1т)"], ["6452953112", "ООО \"ЛОГИСТ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 3.51, 97495, "КП (2017, посл. 19.10.2017, 1 УПД): 325х7 (2.5т), 426х7 (0.9т), 273х8 (0.1т)"], ["3123331329", "ООО \"БЕЛГЕОЦЕНТР\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 3.14, 94346, "КП (2017, посл. 08.09.2017, 1 УПД): 325х7 (3.1т)"], ["3614009875", "ООО \"ИМЕНИ ТЕЛЬМАНА\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 3.65, 88157, "КП (2017, посл. 19.09.2017, 1 УПД): 530х6 (2.6т), 530х7 (1.0т)"], ["3663032683", "ООО \"СМУ-22в\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 57.05, 88119, "КП (2017, посл. 15.09.2017, 1 УПД): 325х6,5-7 (57.0т)"], ["6166151446", "ООО УК \"ЖЕЛДОРСТРОЙ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 4.06, 86894, "КП (2017, посл. 18.05.2017, 1 УПД): 1020х14 (4.1т)"], ["3665057179", "ООО \"Фосфорель\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 2.69, 71317, "КП (2017, посл. 09.11.2017, 1 УПД): 325х10 (2.7т)"], ["3016051334", "ООО \"Аннинский элеватор\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 2.66, 64354, "КП (2017, посл. 01.08.2017, 1 УПД): 530х6 (2.7т)"], ["3662016336", "ЮВСК ООО ПСКП", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 2.71, 58622, "КП (2016–2017, посл. 13.07.2017, 2 УПД): 720х9 (1.8т), 530х6 (0.9т)"], ["3666216600", "Городская строительная компания ООО (ГСК)", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 2.07, 56027, "КП (2017, посл. 24.10.2017, 1 УПД): 273х8 (2.1т)"], ["3663062945", "ООО \"Завод АрБет\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 2.35, 54077, "КП (2017, посл. 12.09.2017, 1 УПД): 530х8 (2.3т)"], ["6829044258", "ООО \"СПЕЦТЕХНОЛОГИИ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 1.83, 52622, "КП (2017, посл. 31.10.2017, 2 УПД): 426х8 (0.9т), 426х6 (0.9т)"], ["482608649267", "ХОЛМЕЦКИЙ ИГОРЬ СЕРГЕЕВИЧ (ИП)", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 2.44, 52535, "КП (2017, посл. 28.11.2017, 1 УПД): 530х8 (2.3т), 426х10 (0.2т)"], ["3123335490", "ООО \"АЛЬФАСТРОЙ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 57.5, 51019, "КП (2017, посл. 25.07.2017, 1 УПД): 273х6 (57.5т)"], ["3663119831", "ООО \"ЮВСМ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 1.93, 40114, "КП (2017, посл. 04.05.2017, 1 УПД): 273х6,5 (1.9т)"], ["2632080580", "ООО \"ТСК \"Гросс-строй\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 1.84, 39719, "КП (2017, посл. 31.07.2017, 1 УПД): 720х9 (1.8т)"], ["3666221576", "ООО \"ВоронежСнабКомплекс\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 2.08, 36964, "КП (2017, посл. 30.10.2017, 1 УПД): 530х7 (2.1т)"], ["3661057820", "ООО \"Технопромрегион\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 1.64, 35506, "КП (2017, посл. 07.02.2017, 1 УПД): 377х10 (1.0т), 273х9 (0.6т)"], ["3662075797", "ООО \"ПроектИнвестСтрой\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 1.42, 34775, "КП (2017, посл. 19.10.2017, 1 УПД): 630х8 (1.4т)"], ["3662128689", "ООО \"РСК \"Панорама\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 1.62, 33550, "КП (2016–2017, посл. 07.02.2017, 2 УПД): 426х6 (0.9т), 325х8 (0.7т)"], ["4807011826", "ООО \"ДОРСПЕЦСТРОЙ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 1.0, 18663, "КП (2017, посл. 15.06.2017, 1 УПД): 325х8 (1.0т)"], ["3620009111", "ООО \"ККЗ \"ЗОЛОТОЙ ПОЧАТОК\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 0.64, 13350, "КП (2016–2017, посл. 27.04.2017, 2 УПД): 426х10 (0.6т)"], ["3128105330", "ООО \"КОМПАНИЯ ИНВЕСТПРОМ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 5.51, 11196, "КП (2017, посл. 22.05.2017, 2 УПД): 325х8 (5.4т), 273х9 (0.1т)"], ["3663090999", "ООО \"СтальИнвест\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 0.4, 8559, "КП (2017, посл. 29.03.2017, 1 УПД): 426х10 (0.4т)"], ["3662995071", "ООО \"КОМИЛЬФО\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 0.31, 6499, "КП (2017, посл. 24.05.2017, 1 УПД): 57х4 (0.3т)"], ["3128083662", "ООО \"ОСКОЛ-БИЛД\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔥 A (VIP)", 1040.57, 4840780, "КП (2016, посл. 07.10.2016, 12 УПД): ПЭ 100 sdr 17 D 400*23.7 (10 (756.0т), 500*(29,7) труба ПЭ100 SDR17 (168.0т), 325х8 (23.9т), Втулка сегм.ПЭ100 sdr17 Д315 (8.0т), Задвижка чуг 30ч6бр,Ду 300,Р (8.0т), 273х7 (8.0т)"], ["3616016081", "ООО \"МостСервис\"", "396302, Воронежская область, Новоусманский район, с. Новая Усмань, ул. Ростовская, д. 4А, офис 1", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔥 A (VIP)", 112.12, 2714041, "КП+РЛ (2016, посл. 25.07.2016, 4 УПД): 720х9 (42.5т), Шпунт Л-5у 12,0*31 (42.3т), Шпунт Л-5у 8,0*30 (27.3т)"], ["6816004674", "ООО \"Агрострой\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 20.96, 413018, "КП (2016, посл. 30.08.2016, 1 УПД): 219х8 (17.5т), 159х6 (3.5т)"], ["6165180765", "ООО \"СМАРТ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 34.54, 349791, "КП (2016, посл. 01.09.2016, 1 УПД): 720х11 ВУС (34.5т)"], ["3663044199", "ООО \"Разноторг\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 15.21, 264294, "КП (2016, посл. 01.08.2016, 1 УПД): 720х9 (15.2т)"], ["6829048284", "ООО \"Электромир\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 11.21, 242142, "КП (2016, посл. 14.07.2016, 2 УПД): 273х8 (11.2т)"], ["3128061789", "ООО \"Комплект Водстрой\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 14.49, 241605, "КП (2016, посл. 19.02.2016, 1 УПД): 530х6 (7.1т), 530х8 (4.9т), 720х8 (2.5т)"], ["2365020899", "ООО \"ЮРСК \"Артель\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 9.51, 236471, "КП (2016, посл. 11.11.2016, 3 УПД): 219х7 (4.9т), 325х8 (4.3т), 325х6 (0.3т)"], ["312303602002", "Селин Сергей Михайлович ИП", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 7.78, 184815, "КП (2016, посл. 16.09.2016, 2 УПД): 273х6 (3.3т), 219х8 (3.1т), 325х8 (1.4т)"], ["6316135943", "ООО \"Диамант\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 14.61, 177511, "КП (2016, посл. 06.12.2016, 1 УПД): 325х8 (9.6т), 426х8 (3.8т), 530х8 (1.2т)"], ["3662194931", "ООО \"ИНВЕСТСТРОЙ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 7.17, 169492, "КП (2016, посл. 30.09.2016, 1 УПД): 426х8 (3.6т), 630х8 (3.6т)"], ["342000396293", "КОЛОБАНОВ АНДРЕЙ ВАСИЛЬЕВИЧ (ИП ГКФХ)", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 7.69, 162619, "КП (2016, посл. 12.12.2016, 1 УПД): 159х6 (7.7т)"], ["3666186473", "ООО \"МОСТ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 7.41, 160067, "КП (2016, посл. 01.09.2016, 1 УПД): 273х8 (7.4т)"], ["3663101697", "ООО \"СТАЛЬПОЛИМЕРСЕРВИС\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 6.54, 133528, "КП (2016, посл. 20.10.2016, 2 УПД): 219х8 (4.2т), 820х10 (2.3т)"], ["3662047623", "ООО \"ВЫБОР\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 7.06, 131924, "КП (2016, посл. 23.05.2016, 1 УПД): 377х10 (3.1т), 530х8 (2.1т), 325х8 (1.0т), 426х6 (0.7т), 273х8 (0.2т)"], ["3662187613", "ООО \"ПРОФИСТРОЙ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 5.86, 121651, "КП (2016, посл. 02.12.2016, 1 УПД): 60х3,5 (5.9т)"], ["3662177132", "ООО \"ВоронежТехСтрой\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 2.76, 58496, "КП (2016, посл. 10.06.2016, 1 УПД): 426х7 (2.1т), 426х6 (0.7т)"], ["3665083242", "ООО ФК \"Стандарт\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 3.13, 46625, "КП (2016, посл. 22.07.2016, 1 УПД): 1020х12 (3.1т)"], ["361704245904", "ИП ХУДОВА НАТАЛИЯ ЗАХАРОВНА", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 1.97, 44152, "КП (2016, посл. 04.04.2016, 1 УПД): 820х9 (2.0т)"], ["3662190060", "ООО \"СтройПолимер\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 1.98, 41845, "КП (2016, посл. 19.09.2016, 2 УПД): 273х8 (1.2т), 273х5 (0.8т)"], ["362903579484", "Маслак Игорь Александрович ИП КФХ", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 1.77, 34442, "КП (2016, посл. 24.06.2016, 1 УПД): 720х9 (1.8т)"], ["3663067140", "ООО \"СТРОЙМОНТАЖ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 1.34, 32834, "КП (2016, посл. 26.05.2016, 1 УПД): 273х6 (1.3т)"], ["3665104799", "ООО \"СТРОЙМАШСЕРВИС-ВОРОНЕЖ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 1.42, 31234, "КП (2016, посл. 22.03.2016, 1 УПД): 426х6 (1.4т)"], ["3665089501", "ООО ПК \"ОСТПРОД\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 0.55, 11671, "КП (2016, посл. 05.04.2016, 1 УПД): 720х9 (0.6т)"], ["3664061447", "ООО \"ДМ-ЭНЕРГОПРОФИЛЬ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 0.59, 11307, "КП (2016, посл. 20.04.2016, 1 УПД): 273х6 (0.6т)"], ["3602008789", "ООО \"СПЕЦИАЛИЗИРОВАННОЕ ХОЗЯЙСТВО \"МОСКОВСКОЕ\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 0.18, 3505, "КП (2016, посл. 11.08.2016, 1 УПД): 219х8 (0.2т)"], ["3128048040", "ООО \"ОМСМ\"", "309500, Белгородская область, г. Старый Оскол, ст. Котел Промузел, пл-ка Монтажная Проезд Ш-6, стр. 13К", "Юрий Борисович", "", "gorstroymoskva@gmail.com", "", "", "", "", "📌 ООО \"ОМСМ\" инн 3128048040 , Белгородская область, г. Старый Оскол, ст. Котел Промузел,…", "", "", "", "📧 Прайс не высылался", "🔥 A (VIP)", 120.72, 6494709, "КП (2015, посл. 06.04.2015, 5 УПД): 820х12 (61.0т), 630х10 (52.5т), 159х5 (7.2т)"], ["3664096200", "ООО \"Пром-Ресурс\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "⭐ B (Постоянный)", 152.73, 1245161, "РЛ (2015, посл. 30.09.2015, 1 УПД): 426х10 (58.8т), ВТ6 300*5000 (х\\ц Трубы+муфт (50.0т), 630х8 (29.6т), ВТ6 400*5000 (х\\ц Трубы+муфт (10.0т), 219х6 (3.0т), 273х6 (1.4т)"], ["3666195326", "ООО \"ПРОЕКТ ИНВЕСТ СТРОЙ\"", "394036, Воронежская область, г. Воронеж, ул. 25 Октября, д. 45, кв. 211", "Николай Алексеевич", "", "", "", "Нет", "", "", "", "12.11.2026", "", "[⏳ В Ожидание - 29.10.2026 14:42]", "📧 Прайс не высылался", "🔹 C (Малый)", 14.98, 419072, "КП (2015, посл. 21.05.2015, 1 УПД): 630х10 (15.0т)"], ["4823014373", "АО \"СК\"ЛМС\"", "398017, Липецкая область, г. о. город Липецк, г. Липецк, ул. Осипенко, д. 15, офис 4", "", "", "lms-perf2013@yandex.ru\nlmsku@lipetsk.ru", "Строительство (ПГС)", "Нет", "", "", "", "05.11.2026", "", "[⛔ В Архив: Отказ (Закрылись) (из «1. Обзвон (База)») - 26.10.2026 11:54]", "📧 Прайс не высылался", "🔹 C (Малый)", 9.77, 194592, "КП (2015, посл. 25.02.2015, 1 УПД): 325х6 (9.8т)"], ["361601119506", "Болотский Алексей Валерьевич ИП", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 5.23, 161775, "КП (2015, посл. 02.12.2015, 1 УПД): Двутавр С255;18Б1х12000;МД;С (5.2т)"], ["3663036060", "ООО \"Стальторг\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 7.95, 153718, "КП (2015, посл. 28.07.2015, 2 УПД): 1020х13 (6.9т), 377х10 (1.0т)"], ["3663094665", "ООО \"Энергосталь\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 6.81, 136436, "КП (2015, посл. 09.06.2015, 4 УПД): 530х8 (2.4т), 325х6 (2.1т), 168х7 (1.2т), 325х7 (1.1т)"], ["3616012721", "ООО \"ПромРегион\"", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 7.88, 103959, "КП (2015, посл. 15.01.2015, 1 УПД): 114х5 (3.0т), 273х8 (2.9т), 273х7 (2.0т)"], ["360900149049", "Башлыков Иван Иванвич ИП Глава КФХ", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 4.28, 93456, "КП (2015, посл. 16.10.2015, 1 УПД): 426х6 (4.3т)"], ["3666180432", "ВСК+ООО", "", "", "", "", "", "", "", "", "", "", "", "", "📧 Прайс не высылался", "🔹 C (Малый)", 2.22, 36206, "КП (2015, посл. 29.07.2015, 1 УПД): 530х6 (2.2т)"]];

function autoMigrateAndCleanWhalesV67_(ss, forceRun) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const whaleSheet = getSheetByNameSafe(ss, "2. Золотой список (Киты)");
  if (!whaleSheet) return;

  // Проверяем, выполнена ли уже миграция на Вариант 1 + Колонку S
  if (!forceRun && whaleSheet.getMaxColumns() >= 19 && whaleSheet.getLastRow() >= 281) {
    const s1 = String(whaleSheet.getRange(1, 19).getValue() || "").trim();
    const s2 = String(whaleSheet.getRange(2, 19).getValue() || "").trim();
    const n2 = String(whaleSheet.getRange(2, 14).getValue() || "").trim();
    if (s1.includes("1С") && s2 !== "" && !n2.includes("Брали по 1С")) {
      return;
    }
  }

  const callSheet = getSheetByNameSafe(ss, "1. Обзвон (База)");
  const waitSheet = getSheetByNameSafe(ss, "3. Лист ожидания");
  const archSheet = getSheetByNameSafe(ss, "4.Архив лидов");

  // 1. Собираем существующие заметки Досье (на ячейках колонки B) со всех листов по ИНН, чтобы ни одно Досье не пропало
  const dossierByInn = {};
  [whaleSheet, callSheet, waitSheet].forEach(sh => {
    if (!sh || sh.getLastRow() < 2) return;
    const lRow = sh.getLastRow();
    const inns = sh.getRange(2, 1, lRow - 1, 1).getValues();
    const notes = sh.getRange(2, 2, lRow - 1, 1).getNotes();
    for (let i = 0; i < inns.length; i++) {
      const cleanInn = String(inns[i][0] || "").replace(/\D/g, "");
      const noteText = String(notes[i][0] || "").trim();
      if (cleanInn && noteText) {
        if (!dossierByInn[cleanInn]) {
          dossierByInn[cleanInn] = noteText;
        } else if (!dossierByInn[cleanInn].includes(noteText)) {
          dossierByInn[cleanInn] += "\n────────────────────\n" + noteText;
        }
      }
    }
  });

  const masterInnSet = new Set(WHALES_2FIRMS_MASTER_ROWS_.map(r => String(r[0] || "").replace(/\D/g, "")).filter(Boolean));

  // 2. Удаляем перенесенных в Киты покупателей 1С с листов «1. Обзвон (База)» и «3. Лист ожидания» (снизу вверх)
  [callSheet, waitSheet].forEach(sh => {
    if (!sh || sh.getLastRow() < 2) return;
    const lRow = sh.getLastRow();
    const inns = sh.getRange(2, 1, lRow - 1, 1).getValues();
    for (let i = inns.length - 1; i >= 0; i--) {
      const cleanInn = String(inns[i][0] || "").replace(/\D/g, "");
      if (cleanInn && masterInnSet.has(cleanInn)) {
        sh.deleteRow(i + 2);
      }
    }
  });

  // 3. Возвращаем закрытую компанию ООО "ЖИЛПРОМСТРОЙ ПЛЮС" (3128089417) в «4.Архив лидов», если её там ещё нет
  if (archSheet) {
    const aLast = archSheet.getLastRow();
    let hasJps = false;
    if (aLast >= 2) {
      const aInns = archSheet.getRange(2, 1, aLast - 1, 1).getValues();
      hasJps = aInns.some(r => String(r[0] || "").replace(/\D/g, "") === "3128089417");
    }
    if (!hasJps) {
      insertAndFormatRow_(archSheet, [
        "3128089417",
        'ООО "ЖИЛПРОМСТРОЙ ПЛЮС"',
        "309514, Белгородская область, г. Старый Оскол, ул. Стадионная, д. 11А",
        "",
        "+7 (472) 522-52-07\n+7 (472) 543-97-04\n+7 (472) 522-04-89",
        "jps_plus@mail.ru\ngps_plus@mail.ru\noskol.avtodor@gmail.com",
        "Строительство",
        "Нет",
        "",
        "",
        "",
        "",
        "Отказ (Закрылись)",
        "[⛔ В Архив: Отказ (Закрылись) (из «1. Обзвон (База)») - 21.10.2026 11:37]"
      ]);
    }
  }

  // 4. Подготавливаем 19 колонок (A..S) на листе «2. Золотой список (Киты)»
  const maxCols = whaleSheet.getMaxColumns();
  if (maxCols < 19) {
    whaleSheet.insertColumnsAfter(maxCols, 19 - maxCols);
  }
  const styleSource = whaleSheet.getRange(1, 15, 1, 1);
  const headerTarget = whaleSheet.getRange(1, 16, 1, 4);
  styleSource.copyTo(headerTarget, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  headerTarget.setValues([["🏆 Сегмент (A/B/C)", "⚖️ Объем (тн)", "💰 Выручка (₽)", "📦 Что брали в 1С (КП / РЛ)"]]);
  headerTarget.setHorizontalAlignment("center").setVerticalAlignment("middle");
  whaleSheet.setColumnWidth(16, 155);
  whaleSheet.setColumnWidth(17, 130);
  whaleSheet.setColumnWidth(18, 155);
  whaleSheet.setColumnWidth(19, 360);

  // 5. Записываем 280 строк (Этаж 1: строки 2..101 с телефонами; Этаж 2: строки 102..281 без телефонов)
  const totalMaster = WHALES_2FIRMS_MASTER_ROWS_.length;
  const neededMaxRow = totalMaster + 1;
  if (whaleSheet.getMaxRows() < neededMaxRow) {
    whaleSheet.insertRowsAfter(whaleSheet.getMaxRows(), neededMaxRow - whaleSheet.getMaxRows());
  }

  const timeZone = ss.getSpreadsheetTimeZone() || "Europe/Moscow";
  const outValues = [];
  const outNotesB = [];
  const dlFormats = [];

  for (let i = 0; i < totalMaster; i++) {
    const r = WHALES_2FIRMS_MASTER_ROWS_[i].slice();
    const cleanInn = String(r[0] || "").replace(/\D/g, "");
    const dlParsed = parseDeadlineObj_(r[11], timeZone);
    r[11] = dlParsed.dateObj ? dlParsed.dateObj : "";
    dlFormats.push([dlParsed.hasTime ? "dd.MM.yyyy HH:mm" : "dd.MM.yyyy"]);
    outValues.push(r);
    outNotesB.push([dossierByInn[cleanInn] || ""]);
  }

  // Очищаем старое содержимое и заметки на листе Китов
  const oldLastRow = Math.max(whaleSheet.getLastRow(), neededMaxRow);
  whaleSheet.getRange(2, 1, oldLastRow - 1, 19).clearContent().setBackground(null).setFontColor("#0f172a").setFontWeight("normal");
  whaleSheet.getRange(2, 2, oldLastRow - 1, 1).clearNote();
  if ( oldLastRow > neededMaxRow ) {
    whaleSheet.getRange(neededMaxRow + 1, 1, oldLastRow - neededMaxRow, 19).clearContent().setBackground(null);
  }

  // Форматы колонок A (ИНН), E (Телефон), H (Потребность), L (Дедлайн), Q (Объем), R (Выручка), S (1С)
  whaleSheet.getRange(2, 1, totalMaster, 1).setNumberFormat("@");
  whaleSheet.getRange(2, 5, totalMaster, 1).setNumberFormat("@");
  whaleSheet.getRange(2, 8, totalMaster, 1).setNumberFormat("@");
  const dlRange = whaleSheet.getRange(2, 12, totalMaster, 1);
  ensureDateValidationOnCell_(dlRange);
  dlRange.setNumberFormats(dlFormats);

  whaleSheet.getRange(2, 1, totalMaster, 19).setValues(outValues);
  whaleSheet.getRange(2, 2, totalMaster, 1).setNotes(outNotesB);
  whaleSheet.getRange(2, 17, totalMaster, 1).setNumberFormat("#,##0.##");
  whaleSheet.getRange(2, 18, totalMaster, 1).setNumberFormat('#,##0 "₽"');
  whaleSheet.getRange(2, 19, totalMaster, 1).setHorizontalAlignment("left");

  setupCleanDropdownsOnSheetSilent_(whaleSheet);
  recalculateWhaleSegmentsOnSheet_(whaleSheet);
  ss.toast(
    "База 2 фирм (КП + РЛ) собрана: Этаж 1 (стр. 2–101, с телефонами) и Этаж 2 (стр. 102–281, без телефонов по годам). Колонки K, N, O очищены, 1С вынесена в колонку S!",
    "✅ Киты 2.0 (Вариант 1 + Колонка S)",
    8
  );
}