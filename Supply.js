/**
 * ЕДИНЫЙ ЦЕНТР УПРАВЛЕНИЯ ПРОДАЖАМИ И СКЛАДОМ (CTRM «Капитал-Плюс» v7.4)
 * Файл: Supply.js — Модуль Закупок (Book 3):
 * - Лист «7. Закупки» (17 колонок A–Q): Стакан дефицита CRM + котировки заводов + MTC Gate + Оприходование на Склад.
 * - Лист «7.1 Заводы и поставщики» (13 колонок A–M): Матрица заводов, допусков стенки, диаметров и контактов.
 * - Автопарсер входящих цен и остатков из Gmail (ОМК/ВМЗ гр Д + Попутчики, КТЗ, ЗТЗ, ТД Уральская Сталь).
 * - Генератор черновиков запросов цены (RFQ) подходящим заводам по диапазону Ø мин–макс.
 * - Вся цветовая индикация управляется строго кодом (без Условного форматирования Таблиц).
 */

const SUPPLY_SHEET_NAME    = '7. Закупки';
const SUPPLIERS_SHEET_NAME = '7.1 Заводы и поставщики';
const SUPPLY_MIN_MARGIN_   = typeof MIN_MARGIN_PER_TON !== 'undefined' ? MIN_MARGIN_PER_TON : 3000;

const SUPPLIERS_HEADERS = [
  'Поставщик / Завод',             // A (1)
  'Тип',                           // B (2)
  'Приоритет',                     // C (3)
  'Сортамент',                     // D (4)
  'Ø мин (мм)',                    // E (5)
  'Ø макс (мм)',                   // F (6)
  'Допуск стенки',                 // G (7)
  'Качество (Грат / Фаска / Шов)', // H (8)
  'Базис (Город)',                 // I (9)
  'Менеджер (ФИО)',                // J (10)
  'Телефон',                       // K (11)
  'E-mail',                        // L (12)
  'Аргумент для клиента / Заметка' // M (13)
];

const DEFAULT_SUPPLIERS_SEED = [
  [
    'ВМЗ (Выкса / ОМК)', 'Завод', '⭐ Основной (Белый список)',
    'Круглая ЭСВ (ГОСТ, гр Д, Попутчики)', 159, 1420,
    'Номинал (ровная стенка)', 'Без поперечки, грат снят, фаска нарезана',
    'г. Выкса', 'Телегина Ирина Юрьевна', '+7 (986) 751-52-02', 'telegina_iiu@omk.ru',
    'Эталонная геометрия, без поперечного шва, проходит технадзор и микрометр; гр Д: 219-325=50к, 377-426=55к, 530=59к'
  ],
  [
    'АО «ЗТЗ» (Загорский ТЗ)', 'Завод', '⭐ Основной (Белый список)',
    'ТБД ЭСВ (530–1420+, сваи, отводы)', 530, 1420,
    'ГОСТ 10706 / 20295', 'От 1 трубы со склада, возможна изоляция',
    'г. Пересвет (МО)', 'Лясковский Кирилл / Батуревич Виктор', '+7 (961) 070-25-25', 'k.lyaskovskiy@uralsteel.com, v.baturevich@ztz.ru',
    'Отгрузка ТБД от 1 трубы со склада в МО; дают цену без НДС и с НДС 22%'
  ],
  [
    'ТД «Уральская Сталь»', 'Завод', '🟡 Резерв',
    'ТБД ЭСВ (530–1420)', 530, 1420,
    'ГОСТ 20295 (К52, 17Г1С-У)', 'Складские партии от 20 тн',
    'г. Москва / завод', 'Морозов Ростислав Сергеевич', '+7 (925) 700-44-43', 'r.morozov@uralsteel.trade',
    'Резервный канал под 530–820+; часто дороже Выксы — сверять вход'
  ],
  [
    'КТЗ (Королёвский ТЗ)', 'Завод', '🟡 Резерв',
    'Круглая ЭСВ до 325 + Профиль', 57, 325,
    'Макс. минус ГОСТ (-10…-12.5%)', 'Часто без фаски (фаска отдельно); сверять цену с ВМЗ',
    'Воронеж / МО', 'Отдел продаж КТЗ (Воронеж)', '+7 (473) 200-04-29', 'subscribe@ktzholding.ru',
    'Подхват в Воронеже/МО до Ø 325; проверять допуск стенки и входную цену против ВМЗ'
  ],
  [
    'ТМК', 'Завод', '🟡 Резерв',
    'Бесшовная / ЭСВ / НКТ', 57, 1420,
    'ГОСТ (уточнять)', 'Заводское качество, MTC',
    'Уточнять', '', '', '',
    'Основной канал под бесшовную и спец-стали'
  ],
  [
    'Темпо (НТЗ «ТЭМ-ПО»)', 'Завод', '🟡 Резерв',
    'ЭСВ / Профиль', 57, 530,
    'ГОСТ (уточнять)', 'Уточнять',
    'г. Набережные Челны', '', '', '',
    'Дополнительный завод под средние диаметры'
  ]
];

const SUPPLY_HEADERS = [
  'ID лота',                            // A (1)
  'Дата / Источник',                    // B (2)
  'Завод / Поставщик',                  // C (3)
  'Размер (Ø×Стенка)',                  // D (4)
  'Тип трубы',                          // E (5)
  'Объём (тн)',                         // F (6)
  'Базис отгрузки',                     // G (7)
  'Цена входа (₽/тн)',                  // H (8)
  'Логистика (₽/тн)',                   // I (9)
  'Себестоимость с дост. (₽/тн)',       // J (10)
  'Целевая продажа (₽/тн)',             // K (11)
  'НДС / Оплата',                       // L (12)
  'Актуально до / Бронь',               // M (13)
  'Статус лота',                        // N (14)
  'Сертификат / Плавка / ГОСТ (MTC)',   // O (15)
  'Кому продать / Заводы RFQ',          // P (16)
  'Примечание (Качество / Цех / 20т)'   // Q (17)
];

const SUPPLY_STATUS_LIST = [
  'Поиск (Дефицит CRM)',
  'Котировка завода',
  'В рынке (Предложение)',
  'Под бронью',
  'Карантин (Нет MTC)',
  'Выкуплено (В транзите)',
  'Оприходовано на Склад',
  'Архив / Ушло'
];

// Базовая сетка цен ОМК/ВМЗ на гр Д (с НДС 22%) — обновляется автоматом, если в письме приходит новая таблица цен
const OMK_GRD_BASE_PRICES_ = [
  { min: 159, max: 325, price: 50000 },
  { min: 351, max: 426, price: 55000 },
  { min: 478, max: 530, price: 59000 }
];

// =================================================================
// 1. ИНИЦИАЛИЗАЦИЯ ОБОИХ ЛИСТОВ («7.1 Заводы» И «7. Закупки»)
// =================================================================

function setupSupplySystem() {
  setupSuppliersSheet();
  setupSupplySheet();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Созданы и настроены листы «7. Закупки» (A–Q) и «7.1 Заводы и поставщики» (A–M)!',
    'Модуль Закупок v7.4',
    5
  );
}

function setupSuppliersSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = typeof getSheetByNameSafe === 'function'
    ? getSheetByNameSafe(ss, SUPPLIERS_SHEET_NAME)
    : ss.getSheetByName(SUPPLIERS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SUPPLIERS_SHEET_NAME);

  const maxCols = SUPPLIERS_HEADERS.length;
  if (sheet.getMaxColumns() < maxCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), maxCols - sheet.getMaxColumns());
  }

  sheet.getRange(1, 1, 1, maxCols)
    .setValues([SUPPLIERS_HEADERS])
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
  sheet.clearConditionalFormatRules();

  const widths = [190, 110, 185, 220, 80, 85, 195, 240, 145, 210, 155, 230, 320];
  widths.forEach((w, idx) => sheet.setColumnWidth(idx + 1, w));

  // Предзаполняем реальными заводами из почты только если лист пустой
  if (sheet.getLastRow() < 2) {
    sheet.getRange(2, 1, DEFAULT_SUPPLIERS_SEED.length, maxCols).setValues(DEFAULT_SUPPLIERS_SEED);
  }

  const numRows = Math.max(sheet.getMaxRows() - 1, 50);
  sheet.getRange(2, 1, numRows, maxCols)
    .setFontFamily('Arial')
    .setFontSize(10)
    .setVerticalAlignment('middle');
  sheet.getRange(2, 1, numRows, 12).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  sheet.getRange(2, 13, numRows, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.getRange(2, 5, numRows, 2).setHorizontalAlignment('center').setNumberFormat('0');
  sheet.getRange(2, 11, numRows, 1).setNumberFormat('@');

  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Завод', 'Трейдер / Б/у', 'Металлобаза'], true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 2, numRows, 1).setDataValidation(typeRule);

  const prioRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['⭐ Основной (Белый список)', '🟡 Резерв', '⛔ Чёрный список'], true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 3, numRows, 1).setDataValidation(prioRule);

  refreshSuppliersStyling_(sheet);
}

function setupSupplySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = typeof getSheetByNameSafe === 'function'
    ? getSheetByNameSafe(ss, SUPPLY_SHEET_NAME)
    : ss.getSheetByName(SUPPLY_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SUPPLY_SHEET_NAME);

  const maxCols = SUPPLY_HEADERS.length;
  if (sheet.getMaxColumns() < maxCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), maxCols - sheet.getMaxColumns());
  }

  sheet.getRange(1, 1, 1, maxCols)
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
  sheet.clearConditionalFormatRules(); // По ТЗ п. 4.11: 0 правил Условного форматирования!

  const widths = [105, 135, 190, 120, 120, 85, 145, 120, 110, 135, 130, 125, 115, 175, 240, 310, 260];
  widths.forEach((w, idx) => sheet.setColumnWidth(idx + 1, w));

  const numRows = Math.max(sheet.getMaxRows() - 1, 150);
  sheet.getRange(2, 1, numRows, maxCols)
    .setFontFamily('Arial')
    .setFontSize(10)
    .setVerticalAlignment('middle');

  sheet.getRange(2, 1, numRows, 14).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  sheet.getRange(2, 15, numRows, 3).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  sheet.getRange(2, 1, numRows, 2).setNumberFormat('@');
  sheet.getRange(2, 4, numRows, 1).setNumberFormat('@');
  sheet.getRange(2, 6, numRows, 1).setNumberFormat('0.00').setHorizontalAlignment('center');
  sheet.getRange(2, 8, numRows, 4).setNumberFormat('#,##0 "₽"');
  sheet.getRange(2, 13, numRows, 1).setNumberFormat('dd.MM.yyyy').setHorizontalAlignment('center');

  // Валидация C (Завод / Поставщик) из листа «7.1 Заводы и поставщики»
  const supDirSheet = typeof getSheetByNameSafe === 'function'
    ? getSheetByNameSafe(ss, SUPPLIERS_SHEET_NAME)
    : ss.getSheetByName(SUPPLIERS_SHEET_NAME);
  if (supDirSheet) {
    const supRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(supDirSheet.getRange('A2:A100'), true)
      .setAllowInvalid(true)
      .build();
    sheet.getRange(2, 3, numRows, 1).setDataValidation(supRule);
  }

  const pipeTypeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Новая', 'Лежалая', 'Восстановленная', 'Б/у', 'НКТ'], true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 5, numRows, 1).setDataValidation(pipeTypeRule);

  const vatRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['С НДС 22%', 'Без НДС', 'Нал / Карта', 'Отсрочка / Бронь'], true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 12, numRows, 1).setDataValidation(vatRule);

  const dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build();
  sheet.getRange(2, 13, numRows, 1).setDataValidation(dateRule);

  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(SUPPLY_STATUS_LIST, true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 14, numRows, 1).setDataValidation(statusRule);

  refreshAllSupplyStyling_(sheet);
}

// =================================================================
// 2. СПРАВОЧНИК ЗАВОДОВ И ПОДБОР ЗАВОДОВ ПО ДИАМЕТРУ (RFQ-МАТРИЦА)
// =================================================================

function getSuppliersDirectory_(ss) {
  const sheet = typeof getSheetByNameSafe === 'function'
    ? getSheetByNameSafe(ss, SUPPLIERS_SHEET_NAME)
    : ss.getSheetByName(SUPPLIERS_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, SUPPLIERS_HEADERS.length).getValues();
  return data
    .map(r => {
      const name = String(r[0] || '').trim();
      if (!name) return null;
      const emailRaw = String(r[11] || '').trim();
      const emails = emailRaw.split(/[,;\s\n]+/).map(x => x.trim()).filter(x => x && x.includes('@'));
      return {
        name: name,
        type: String(r[1] || '').trim(),
        priority: String(r[2] || '').trim(),
        isBlacklisted: String(r[2] || '').toLowerCase().includes('чёрн') || String(r[2] || '').toLowerCase().includes('черн'),
        isPrimary: String(r[2] || '').toLowerCase().includes('основн') || String(r[2] || '').toLowerCase().includes('бел'),
        assortment: String(r[3] || '').trim(),
        minDiam: safeParseFloat(r[4]) || 0,
        maxDiam: safeParseFloat(r[5]) || 9999,
        tolerance: String(r[6] || '').trim(),
        quality: String(r[7] || '').trim(),
        basis: String(r[8] || '').trim(),
        manager: String(r[9] || '').trim(),
        phone: String(r[10] || '').trim(),
        emails: emails,
        argument: String(r[12] || '').trim()
      };
    })
    .filter(Boolean);
}

function findMatchingFactoriesForPipe_(diam, pipeType, suppliers) {
  const isUsedOnly = /б\/у|восст/i.test(String(pipeType || ''));
  return (suppliers || []).filter(s => {
    if (s.isBlacklisted) return false;
    if (isUsedOnly && s.type.toLowerCase() === 'завод') return false;
    if (diam > 0 && (diam < s.minDiam || diam > s.maxDiam)) return false;
    return true;
  }).sort((a, b) => (b.isPrimary ? 1 : 0) - (a.isPrimary ? 1 : 0));
}

// =================================================================
// 3. УМНЫЙ СБОР ДЕФИЦИТА ИЗ CRM (КИТЫ + ОЖИДАНИЕ)
// Ключ = Размер + Тип трубы, фильтр хвостов < 1,0 тн, авто-подбор заводов!
// =================================================================

function collectDeficitFromCRM() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let supplySheet = typeof getSheetByNameSafe === 'function'
    ? getSheetByNameSafe(ss, SUPPLY_SHEET_NAME)
    : ss.getSheetByName(SUPPLY_SHEET_NAME);
  if (!supplySheet) {
    setupSupplySystem();
    supplySheet = ss.getSheetByName(SUPPLY_SHEET_NAME);
  }

  const stockEntriesMap = typeof getAvailableStockEntriesMap_ === 'function'
    ? getAvailableStockEntriesMap_(ss)
    : {};
  const suppliers = getSuppliersDirectory_(ss);

  // Карта существующих строк дефицита по ключу "размер||тип"
  const existingDeficits = new Map();
  if (supplySheet.getLastRow() > 1) {
    const supData = supplySheet.getRange(2, 1, supplySheet.getLastRow() - 1, SUPPLY_HEADERS.length).getValues();
    supData.forEach((r, idx) => {
      const size = String(r[3] || '').trim();
      const pType = String(r[4] || '').trim();
      const status = String(r[13] || '').trim();
      if (size && status === 'Поиск (Дефицит CRM)') {
        existingDeficits.set(`${size.toLowerCase()}||${pType.toLowerCase()}`, idx + 2);
      }
    });
  }

  const crmSheets = ['2. Золотой список (Киты)', '3. Лист ожидания'];
  const deficitMap = new Map();

  crmSheets.forEach(sheetName => {
    const sh = typeof getSheetByNameSafe === 'function'
      ? getSheetByNameSafe(ss, sheetName)
      : ss.getSheetByName(sheetName);
    if (!sh || sh.getLastRow() < 2) return;

    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();
    const tag = sheetName.indexOf('Киты') !== -1 ? '[КИТ]' : '[ОЖИД]';

    rows.forEach(row => {
      const inn = String(row[0] || '').trim();
      const company = String(row[1] || '').trim();
      const city = typeof extractShortCity_ === 'function' ? extractShortCity_(row[2]) : String(row[2] || '').trim();
      const need = String(row[7] || '').trim();
      const pipeType = String(row[8] || '').trim() || 'Новая';

      if (!company || !need || need.toLowerCase() === 'нет' || need === '—' || need === '-') return;

      const items = need.split(/[,;]+/).map(s => s.trim()).filter(Boolean);
      items.forEach(item => {
        const diams = typeof extractCleanPipeDiams_ === 'function' ? extractCleanPipeDiams_(item) : [];
        if (diams.length === 0) return;
        const d = diams[0];

        // Проверяем: есть ли на Складе полноценная позиция (>= 1 тн) нужного типа трубы
        const hasFullStock = (stockEntriesMap[d] || []).some(it =>
          !it.isTail &&
          (typeof doesStockMatchClientPipeType_ === 'function'
            ? doesStockMatchClientPipeType_(it.condRaw, it.category, pipeType)
            : true)
        );
        if (hasFullStock) return;

        const cleanSize = typeof formatPipeNeed_ === 'function' ? formatPipeNeed_(item) : item;
        const key = `${cleanSize.toLowerCase()}||${pipeType.toLowerCase()}`;
        if (!deficitMap.has(key)) {
          deficitMap.set(key, { size: cleanSize, diam: d, pipeType: pipeType, clients: [] });
        }
        const entry = deficitMap.get(key);
        const clientLine = `${tag} ${company}${city ? ' (' + city + ')' : (inn ? ' (ИНН ' + inn + ')' : '')}`;
        if (!entry.clients.includes(clientLine)) entry.clients.push(clientLine);
      });
    });
  });

  const tz = ss.getSpreadsheetTimeZone() || 'Europe/Moscow';
  const now = new Date();
  const stampId = Utilities.formatDate(now, tz, 'ddMM');
  const todayDisp = Utilities.formatDate(now, tz, 'dd.MM.yyyy');

  let addedCount = 0;
  let updatedCount = 0;
  const newRows = [];

  deficitMap.forEach((info, key) => {
    const matchedFactories = findMatchingFactoriesForPipe_(info.diam, info.pipeType, suppliers);
    const rfqLine = matchedFactories.length > 0
      ? 'ЗАПРОСИТЬ: ' + matchedFactories.map(f => f.name).join(', ')
      : 'Подходящие заводы не найдены (проверить 7.1)';
    const colPText = `КЛИЕНТЫ:\n${info.clients.join('\n')}\n${rfqLine}`;

    if (existingDeficits.has(key)) {
      const rowIndex = existingDeficits.get(key);
      supplySheet.getRange(rowIndex, 16).setValue(colPText);
      updatedCount++;
    } else {
      addedCount++;
      const lotId = `DEF-${stampId}-${addedCount}`;
      newRows.push([
        lotId,
        `${todayDisp} • CRM`,
        'ДЕФИЦИТ ПОД ЗАЯВКУ CRM',
        info.size,
        info.pipeType,
        '',
        '',
        '',
        '',
        '',
        '',
        'С НДС 22%',
        '',
        'Поиск (Дефицит CRM)',
        info.pipeType.toLowerCase().includes('нов') ? 'Требуется MTC (сертификат)' : '—',
        colPText,
        'Авто-сбор из колонки H (с учётом типа трубы и отсечки <1т)'
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

  refreshAllSupplyStyling_(supplySheet);
  ss.toast(`Новых позиций дефицита: ${addedCount} | Обновлено: ${updatedCount}`, 'Сбор дефицита CRM', 5);
}

// =================================================================
// 4. АВТОПАРСЕР ВХОДЯЩИХ ПРАЙСОВ И ОТВЕТОВ ЗАВОДОВ ИЗ GMAIL
// Поддерживает: ВМЗ/ОМК (Попутчики + гр Д), КТЗ (круглые ЭСПШ), ЗТЗ, ТД Уральская Сталь
// =================================================================

function syncFactoryPricesFromGmail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let supplySheet = typeof getSheetByNameSafe === 'function'
    ? getSheetByNameSafe(ss, SUPPLY_SHEET_NAME)
    : ss.getSheetByName(SUPPLY_SHEET_NAME);
  if (!supplySheet) {
    setupSupplySystem();
    supplySheet = ss.getSheetByName(SUPPLY_SHEET_NAME);
  }

  const tz = ss.getSpreadsheetTimeZone() || 'Europe/Moscow';
  const query = '(from:omk.ru -from:auction.omk.ru) OR from:ktzholding.ru OR from:uralsteel.com OR from:ztz.ru OR from:uralsteel.trade newer_than:21d';
  const threads = GmailApp.search(query, 0, 35);

  if (!threads || threads.length === 0) {
    ss.toast('За последние 21 день новых писем с ценами от заводов не найдено.', 'Парсер Gmail', 5);
    return;
  }

  const parsedLots = [];
  const omkPriceGrid = OMK_GRD_BASE_PRICES_.map(x => ({ min: x.min, max: x.max, price: x.price }));
  let latestKtzProcessed = false;
  let latestOmkGrdProcessed = false;

  // Сортируем цепочки от самых свежих к старым
  threads.sort((a, b) => b.getLastMessageDate().getTime() - a.getLastMessageDate().getTime());

  threads.forEach(thread => {
    const msgs = thread.getMessages().slice().reverse(); // от новых сообщений внутри ветки к старым
    msgs.forEach(msg => {
      const from = String(msg.getFrom() || '').toLowerCase();
      if (from.includes('kapital-p.ru') || from.includes('auction.omk.ru')) return;

      const msgDate = msg.getDate();
      const dateShort = Utilities.formatDate(msgDate, tz, 'dd.MM');
      const validUntil = new Date(msgDate.getTime() + 7 * 86400000);
      const rawHtml = msg.getBody() || '';
      const freshHtml = stripQuotedReplyHtml_(rawHtml);
      const tables = extractHtmlTables_(freshHtml);
      const plainFresh = htmlToPlainLines_(freshHtml);

      // --- 1. ВМЗ / ОМК МАРКЕТ ---
      if (from.includes('omk.ru')) {
        // Проверяем, нет ли в письме обновлённой таблицы базовых цен гр Д
        tables.forEach(rows => updateOmkGrdPriceGridFromTable_(rows, omkPriceGrid));

        // 1A. Свободные лоты в тексте: "15 тн 530х8 ... по цене 59 тр/тн"
        const freeMatch = plainFresh.match(/(\d+(?:[.,]\d+)?)\s*тн\s+(\d{2,4})\s*[хx×*]\s*(\d+(?:[.,]\d+)?)(?:\s*\(([^)]+)\))?[^.\n]{0,60}?(\d{2,3})\s*(?:т\.?\s*р|тр)\/?тн/i);
        if (freeMatch) {
          const tons = safeParseFloat(freeMatch[1]);
          const diam = parseInt(freeMatch[2], 10);
          const wall = cleanWallStr_(freeMatch[3]);
          const extra = freeMatch[4] ? freeMatch[4].trim() : '';
          const price = parseInt(freeMatch[5], 10) * 1000;
          if (diam >= 57 && price >= 25000) {
            parsedLots.push({
              sourcePrefix: 'OMK',
              dateSource: `${dateShort} • ОМК текст`,
              supplier: 'ВМЗ (Выкса / ОМК)',
              size: `${diam}х${wall}`,
              pipeType: 'Новая',
              tons: tons,
              basis: 'г. Выкса',
              priceIn: price,
              vat: 'С НДС 22%',
              validUntil: validUntil,
              mtc: 'ВМЗ Выкса • гр Д (уточнить MTC)',
              note: `гр Д (свободный объём${extra ? ', ' + extra : ''}); без поперечки, ровная стенка`,
              dedupKey: `ВМЗ||${diam}х${wall}||грД-своб`
            });
          }
        }

        // 1B. HTML-таблицы ОМК: Попутчики с ценой ИЛИ еженедельная гр Д
        tables.forEach(rows => {
          if (rows.length < 2) return;
          const headerText = rows[0].join(' ').toLowerCase();

          // Вариант 1: Попутчики (ответ с ценой, плавкой, партией и ГОСТ)
          if (headerText.includes('диаметр') && headerText.includes('цена')) {
            const hMap = mapHeaderColumns_(rows[0]);
            for (let i = 1; i < rows.length; i++) {
              const r = rows[i];
              const diam = Math.round(safeParseFloat(r[hMap.diam]));
              const wall = cleanWallStr_(r[hMap.wall]);
              const price = Math.round(extractPriceRub_(r[hMap.price]));
              const tons = hMap.tons >= 0 ? safeParseFloat(r[hMap.tons]) : 0;
              if (diam < 57 || !wall || price < 25000) continue;

              const batch = hMap.batch >= 0 ? String(r[hMap.batch] || '').trim() : '';
              const heat  = hMap.heat >= 0 ? String(r[hMap.heat] || '').trim() : '';
              const steel = hMap.steel >= 0 ? String(r[hMap.steel] || '').trim() : '';
              const gost  = hMap.gost >= 0 ? String(r[hMap.gost] || '').trim() : '';
              const shop  = hMap.shop >= 0 ? String(r[hMap.shop] || '').trim() : '';

              const mtcParts = [
                gost,
                steel ? `Ст ${steel}` : '',
                batch ? `Парт. ${batch}` : '',
                heat ? `Плавка ${heat}` : ''
              ].filter(Boolean);

              parsedLots.push({
                sourcePrefix: 'OMK',
                dateSource: `${dateShort} • ОМК Попутчик`,
                supplier: 'ВМЗ (Выкса / ОМК)',
                size: `${diam}х${wall}`,
                pipeType: 'Новая',
                tons: tons,
                basis: 'г. Выкса',
                priceIn: price,
                vat: 'С НДС 22%',
                validUntil: validUntil,
                mtc: mtcParts.join(' | ') || 'ВМЗ Выкса',
                note: `Попутчик ${shop} • без поперечки, ровная стенка, грат снят, фаска`.trim(),
                dedupKey: `ВМЗ||${diam}х${wall}||${gost}||${batch || price}`
              });
            }
            return;
          }

          // Вариант 2: Таблица остатков гр Д (берём только самое свежее письмо гр Д, чтобы не дублировать старые остатки)
          if (!latestOmkGrdProcessed && headerText.includes('диаметр') && headerText.includes('толщин')) {
            let foundAnyGrd = false;
            const hasShopCol = rows[0][0].toLowerCase().includes('цех');
            const idxD = hasShopCol ? 1 : 0;
            const idxW = hasShopCol ? 2 : 1;
            const idxT = hasShopCol ? 3 : 2;

            for (let i = 1; i < rows.length; i++) {
              const r = rows[i];
              if (r.length <= idxT) continue;
              const diam = Math.round(safeParseFloat(r[idxD]));
              const wallRaw = String(r[idxW] || '').replace(/итог/gi, '').trim();
              const wall = cleanWallStr_(wallRaw);
              const tons = safeParseFloat(r[idxT]);
              if (diam < 57 || !wall || tons < 1.0) continue; // Отсечка хвостов < 1 тн

              const basePrice = lookupOmkGrdPrice_(diam, omkPriceGrid);
              foundAnyGrd = true;
              parsedLots.push({
                sourcePrefix: 'OMK',
                dateSource: `${dateShort} • ОМК гр Д`,
                supplier: 'ВМЗ (Выкса / ОМК)',
                size: `${diam}х${wall}`,
                pipeType: 'Новая',
                tons: tons,
                basis: 'г. Выкса',
                priceIn: basePrice,
                vat: 'С НДС 22%',
                validUntil: validUntil,
                mtc: 'ВМЗ Выкса • гр Д (ТЭСЦ-3)',
                note: 'гр Д (остаток завода) • без поперечки, ровная стенка',
                dedupKey: `ВМЗ||${diam}х${wall}||грД`
              });
            }
            if (foundAnyGrd) latestOmkGrdProcessed = true;
          }
        });
        return;
      }

      // --- 2. КТЗ (КОРОЛЁВСКИЙ ТРУБНЫЙ ЗАВОД) — берём только самое свежее КП ---
      if (from.includes('ktzholding.ru') && !latestKtzProcessed) {
        const allowedRegions = ['воронеж', 'москов', 'волгоград', 'ростов'];
        let foundKtz = false;

        tables.forEach(rows => {
          rows.forEach(r => {
            if (r.length < 3) return;
            const reg = String(r[0] || '').trim();
            const name = String(r[1] || '').trim();
            const price = Math.round(extractPriceRub_(r[2]));

            const rl = reg.toLowerCase();
            const nl = name.toLowerCase();
            if (!allowedRegions.some(ar => rl.includes(ar))) return;
            if (!nl.startsWith('труба')) return;
            if (/проф|оц\b|х\/к|вгп/i.test(nl)) return; // Только круглые ЭСПШ!
            if (price < 25000 || price > 300000) return;

            const m = name.match(/(\d{2,4})\s*[хx×*]\s*(\d+(?:[.,]\d+)?)/);
            if (!m) return;
            const diam = parseInt(m[1], 10);
            const wall = cleanWallStr_(m[2]);
            if (diam < 57) return;

            const hasChamfer = nl.includes('фаск');
            const isDiscount = nl.includes('уценк') || /\(с\)/i.test(name);
            const shortReg = rl.includes('воронеж') ? 'Воронежская обл.'
              : (rl.includes('москов') ? 'Московская обл.' : reg);

            foundKtz = true;
            parsedLots.push({
              sourcePrefix: 'KTZ',
              dateSource: `${dateShort} • КП КТЗ`,
              supplier: 'КТЗ (Королёвский ТЗ)',
              size: `${diam}х${wall}`,
              pipeType: 'Новая',
              tons: '',
              basis: shortReg,
              priceIn: price,
              vat: 'С НДС 22%',
              validUntil: validUntil,
              mtc: 'КТЗ • ЭСПШ (запросить MTC)',
              note: `${name} • макс. минус ГОСТ${hasChamfer ? ', с фаской' : ', без фаски'}${isDiscount ? ', уценка/С' : ''}`,
              dedupKey: `КТЗ||${diam}х${wall}||${shortReg}||${hasChamfer ? 'ф' : 'бф'}${isDiscount ? '-уц' : ''}`
            });
          });
        });
        if (foundKtz) latestKtzProcessed = true;
        return;
      }

      // --- 3. АО «ЗТЗ» И ТД «УРАЛЬСКАЯ СТАЛЬ» ---
      if (from.includes('uralsteel.com') || from.includes('ztz.ru') || from.includes('uralsteel.trade')) {
        const isTradeHouse = from.includes('uralsteel.trade');
        const supName = isTradeHouse ? 'ТД «Уральская Сталь»' : 'АО «ЗТЗ» (Загорский ТЗ)';
        const basis = isTradeHouse ? 'г. Москва / завод' : 'г. Пересвет (МО)';

        // Вытаскиваем тоннажи из текстовых описаний над таблицами (например "720х8 ... 15,05тн" или "530х10 ... 20 т.")
        const tonsBySize = {};
        const tRegex = /(\d{2,4})\s*[хx×*]\s*(\d+(?:[.,]\d+)?)[^\n]{0,70}?(\d+(?:[.,]\d+)?)\s*(?:тн|т\.)/gi;
        let tm;
        while ((tm = tRegex.exec(plainFresh)) !== null) {
          const keySz = `${parseInt(tm[1], 10)}х${cleanWallStr_(tm[2])}`;
          tonsBySize[keySz] = safeParseFloat(tm[3]);
        }

        tables.forEach(rows => {
          rows.forEach(r => {
            if (r.length < 4) return;
            const diam = Math.round(safeParseFloat(r[0]));
            const wall = cleanWallStr_(r[1]);
            if (diam < 57 || diam > 2520 || !wall) return;

            let priceNds = 0;
            let priceNoNds = 0;
            let pricePlain = 0;
            let gost = '';
            let steelParts = [];

            for (let c = 2; c < r.length; c++) {
              const cell = String(r[c] || '').trim();
              const cl = cell.toLowerCase();
              if (!cell || cell === '-') continue;

              const numVal = extractPriceRub_(cell);
              if (numVal >= 25000 && numVal <= 500000) {
                if (cl.includes('без ндс')) priceNoNds = numVal;
                else if (cl.includes('с ндс')) priceNds = numVal;
                else if (cl.includes('руб') || cl.includes('₽') || /^\d[\d\s.,]+$/.test(cell)) pricePlain = numVal;
              } else if (/гост|ту|тс/i.test(cell)) {
                gost = gost ? `${gost}, ${cell}` : cell;
              } else if (cell.length <= 24) {
                steelParts.push(cell);
              }
            }

            const finalPrice = Math.round(priceNds || pricePlain || (priceNoNds > 0 ? priceNoNds * 1.22 : 0));
            if (finalPrice < 25000) return;

            const sizeKey = `${diam}х${wall}`;
            parsedLots.push({
              sourcePrefix: isTradeHouse ? 'UST' : 'ZTZ',
              dateSource: `${dateShort} • ${isTradeHouse ? 'УралСталь' : 'ЗТЗ'}`,
              supplier: supName,
              size: sizeKey,
              pipeType: 'Новая',
              tons: tonsBySize[sizeKey] || '',
              basis: basis,
              priceIn: finalPrice,
              vat: 'С НДС 22%',
              validUntil: validUntil,
              mtc: [gost, steelParts.slice(0, 2).join(' ')].filter(Boolean).join(' | ') || supName,
              note: `${isTradeHouse ? 'Партия со склада завода' : 'От 1 трубы со склада ЗТЗ'}${priceNoNds > 0 ? ' (без НДС: ' + Math.round(priceNoNds).toLocaleString('ru-RU') + ' ₽)' : ''}`,
              dedupKey: `${supName}||${sizeKey}||${gost}`
            });
          });
        });
      }
    });
  });

  if (parsedLots.length === 0) {
    ss.toast('В письмах за 21 день не распознано новых товарных позиций.', 'Парсер Gmail', 5);
    return;
  }

  // Дедупликация внутри текущей выгрузки (оставляем самую свежую котировку по dedupKey)
  const uniqueLots = [];
  const seenKeys = new Set();
  parsedLots.forEach(lot => {
    if (!seenKeys.has(lot.dedupKey)) {
      seenKeys.add(lot.dedupKey);
      uniqueLots.push(lot);
    }
  });

  // Сверяем с существующими строками на листе «7. Закупки» (Upsert: обновляем цену/тоннаж или добавляем)
  const existingMap = new Map();
  const lastRow = supplySheet.getLastRow();
  if (lastRow > 1) {
    const curData = supplySheet.getRange(2, 1, lastRow - 1, SUPPLY_HEADERS.length).getValues();
    curData.forEach((r, idx) => {
      const sup = String(r[2] || '').trim();
      const sz  = String(r[3] || '').trim();
      const bas = String(r[6] || '').trim();
      const mtc = String(r[14] || '').trim();
      const st  = String(r[13] || '').trim();
      if (sup && sz && st !== 'Оприходовано на Склад' && st !== 'Поиск (Дефицит CRM)') {
        const key1 = `${sup}||${sz}||${bas}||${mtc.slice(0, 24)}`.toLowerCase();
        existingMap.set(key1, { row: idx + 2, logistics: safeParseFloat(r[8]) });
      }
    });
  }

  const rowsToAppend = [];
  let updated = 0;
  let added = 0;
  const stampId = Utilities.formatDate(new Date(), tz, 'ddMM');

  uniqueLots.forEach(lot => {
    const matchKey = `${lot.supplier}||${lot.size}||${lot.basis}||${String(lot.mtc || '').slice(0, 24)}`.toLowerCase();
    if (existingMap.has(matchKey)) {
      const ex = existingMap.get(matchKey);
      const totalCost = lot.priceIn + ex.logistics;
      supplySheet.getRange(ex.row, 2).setValue(lot.dateSource);
      if (lot.tons !== '') supplySheet.getRange(ex.row, 6).setValue(lot.tons);
      supplySheet.getRange(ex.row, 8).setValue(lot.priceIn);
      supplySheet.getRange(ex.row, 10).setValue(totalCost);
      supplySheet.getRange(ex.row, 13).setValue(lot.validUntil);
      supplySheet.getRange(ex.row, 15).setValue(lot.mtc);
      supplySheet.getRange(ex.row, 17).setValue(lot.note);
      updated++;
    } else {
      added++;
      rowsToAppend.push([
        `${lot.sourcePrefix}-${stampId}-${added}`,
        lot.dateSource,
        lot.supplier,
        lot.size,
        lot.pipeType,
        lot.tons,
        lot.basis,
        lot.priceIn,
        '',
        lot.priceIn,
        '',
        lot.vat,
        lot.validUntil,
        'Котировка завода',
        lot.mtc,
        '',
        lot.note
      ]);
    }
  });

  if (rowsToAppend.length > 0) {
    const startRow = supplySheet.getLastRow() + 1;
    if (supplySheet.getMaxRows() < startRow + rowsToAppend.length) {
      supplySheet.insertRowsAfter(supplySheet.getMaxRows(), rowsToAppend.length + 15);
    }
    supplySheet.getRange(startRow, 1, rowsToAppend.length, SUPPLY_HEADERS.length).setValues(rowsToAppend);
  }

  matchSupplyWithDemand();
  refreshAllSupplyStyling_(supplySheet);

  ss.toast(
    `Парсер Gmail готов! Новых котировок: ${added} | Обновлено: ${updated} (ОМК/ВМЗ, КТЗ, ЗТЗ, УралСталь).`,
    'Автопарсер заводов',
    6
  );
}

// =================================================================
// 5. ГЕНЕРАТОР ЧЕРНОВИКОВ ЗАПРОСОВ (RFQ) ПОДХОДЯЩИМ ЗАВОДАМ
// =================================================================

function createRfqDraftsForDeficit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const supplySheet = typeof getSheetByNameSafe === 'function'
    ? getSheetByNameSafe(ss, SUPPLY_SHEET_NAME)
    : ss.getSheetByName(SUPPLY_SHEET_NAME);
  if (!supplySheet || supplySheet.getLastRow() < 2) {
    ss.toast('На листе «7. Закупки» нет позиций дефицита.', 'Запрос заводам (RFQ)', 4);
    return;
  }

  const suppliers = getSuppliersDirectory_(ss).filter(s =>
    !s.isBlacklisted &&
    s.emails.some(em => !em.toLowerCase().startsWith('subscribe@'))
  );
  if (suppliers.length === 0) {
    SpreadsheetApp.getUi().alert('На листе «7.1 Заводы и поставщики» нет активных заводов с рабочим Email.');
    return;
  }

  const data = supplySheet.getRange(2, 1, supplySheet.getLastRow() - 1, SUPPLY_HEADERS.length).getValues();
  const requestsBySupplier = {};

  data.forEach(r => {
    const status = String(r[13] || '').trim();
    if (status !== 'Поиск (Дефицит CRM)') return;

    const size = String(r[3] || '').trim();
    const pipeType = String(r[4] || '').trim();
    const tons = safeParseFloat(r[5]);
    const diams = typeof extractCleanPipeDiams_ === 'function' ? extractCleanPipeDiams_(size) : [];
    if (diams.length === 0) return;

    const matched = findMatchingFactoriesForPipe_(diams[0], pipeType, suppliers);
    matched.forEach(sup => {
      const workEmails = sup.emails.filter(em => !em.toLowerCase().startsWith('subscribe@'));
      if (workEmails.length === 0) return;
      if (!requestsBySupplier[sup.name]) {
        requestsBySupplier[sup.name] = { supplier: sup, emails: workEmails, lines: [] };
      }
      const line = `• Труба ${size} (${pipeType})${tons > 0 ? ' — ' + tons + ' тн' : ''}`;
      if (!requestsBySupplier[sup.name].lines.includes(line)) {
        requestsBySupplier[sup.name].lines.push(line);
      }
    });
  });

  const keys = Object.keys(requestsBySupplier);
  if (keys.length === 0) {
    ss.toast('Нет строк со статусом «Поиск (Дефицит CRM)» под заводы из «7.1».', 'Запрос заводам (RFQ)', 5);
    return;
  }

  const aliases = GmailApp.getAliases();
  const senderAlias = aliases.find(a => String(a).toLowerCase().includes('kapital-p.ru')) || (aliases[0] || null);
  const sig = typeof getSenderSignature === 'function' ? getSenderSignature() : { html: '' };
  const dateStr = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone() || 'Europe/Moscow', 'dd.MM.yyyy');

  let draftsCreated = 0;
  keys.forEach(supName => {
    const pack = requestsBySupplier[supName];
    const subject = `Запрос наличия и цены (ООО «Капитал-Плюс») от ${dateStr}`;
    const plainBody =
      `Добрый день!\n\n` +
      `Прошу сориентировать по наличию и актуальной цене с НДС по следующим позициям:\n` +
      `${pack.lines.join('\n')}\n\n` +
      `Подскажите также базис отгрузки и возможность комплектации 20-тонной машины.\n\n` +
      `С уважением, Кирилл (ООО «Капитал-Плюс», +7 906 588-37-77)`;

    const htmlBody =
      `<div style="font-family:Arial,sans-serif;font-size:14px;color:#0f172a;line-height:1.5;">` +
      `<p>Добрый день!</p>` +
      `<p>Прошу сориентировать по наличию и актуальной цене с НДС по следующим позициям:</p>` +
      `<p>${pack.lines.join('<br>')}</p>` +
      `<p>Подскажите также базис отгрузки и возможность комплектации 20-тонной машины.</p>` +
      `${sig.html || ''}` +
      `</div>`;

    const opts = { htmlBody: htmlBody, name: 'Кирилл Манюхин' };
    if (senderAlias) {
      opts.from = senderAlias;
      opts.replyTo = senderAlias;
    }
    GmailApp.createDraft(pack.emails.join(', '), subject, plainBody, opts);
    draftsCreated++;
  });

  SpreadsheetApp.getUi().alert(
    'Черновики запросов (RFQ) созданы!',
    `В папке «Черновики» Gmail создано ${draftsCreated} писем для заводов:\n• ${keys.join('\n• ')}\n\nПроверьте черновики и нажмите «Отправить».`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// =================================================================
// 6. УМНЫЙ ОБРАТНЫЙ РАДАР (С ФИЛЬТРОМ ПО ТИПУ ТРУБЫ + VIP-СОРТИРОВКА)
// =================================================================

function matchSupplyWithDemand() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const supplySheet = typeof getSheetByNameSafe === 'function'
    ? getSheetByNameSafe(ss, SUPPLY_SHEET_NAME)
    : ss.getSheetByName(SUPPLY_SHEET_NAME);
  if (!supplySheet || supplySheet.getLastRow() < 2) {
    ss.toast('На листе «7. Закупки» пока нет позиций.', '7. Закупки', 4);
    return;
  }

  const demandByDiam = {};
  const ensureDiam_ = (d) => {
    if (!demandByDiam[d]) demandByDiam[d] = { hot: [], profile: [] };
  };

  const whalesSheet = typeof getSheetByNameSafe === 'function'
    ? getSheetByNameSafe(ss, '2. Золотой список (Киты)')
    : ss.getSheetByName('2. Золотой список (Киты)');
  if (whalesSheet && whalesSheet.getLastRow() > 1) {
    const wCols = Math.min(whalesSheet.getLastColumn(), 19);
    const wData = whalesSheet.getRange(2, 1, whalesSheet.getLastRow() - 1, wCols).getValues();
    wData.forEach(r => {
      const company = String(r[1] || '').trim();
      const needH   = String(r[7] || '').trim();
      const typeI   = String(r[8] || '').trim();
      const tonsQ   = wCols >= 17 ? safeParseFloat(r[16]) : 0;
      const revR    = wCols >= 18 ? safeParseFloat(r[17]) : 0;
      const hist1C  = wCols >= 19 ? String(r[18] || '').trim() : '';
      if (!company) return;

      if (needH && needH.toLowerCase() !== 'нет' && needH !== '—') {
        extractCleanPipeDiams_(needH).forEach(d => {
          ensureDiam_(d);
          demandByDiam[d].hot.push({ label: `[ЗАЯВКА КИТ] ${company} (${needH})`, pipeType: typeI, rev: revR });
        });
      }
      if (hist1C) {
        extractCleanPipeDiams_(hist1C).forEach(d => {
          ensureDiam_(d);
          demandByDiam[d].profile.push({
            name: company,
            label: `${company}${tonsQ > 0 ? ' ' + Math.round(tonsQ) + 'т' : ''}`,
            rev: revR
          });
        });
      }
    });
  }

  const waitSheet = typeof getSheetByNameSafe === 'function'
    ? getSheetByNameSafe(ss, '3. Лист ожидания')
    : ss.getSheetByName('3. Лист ожидания');
  if (waitSheet && waitSheet.getLastRow() > 1) {
    const waitData = waitSheet.getRange(2, 1, waitSheet.getLastRow() - 1, 14).getValues();
    waitData.forEach(r => {
      const company = String(r[1] || '').trim();
      const needH   = String(r[7] || '').trim();
      const typeI   = String(r[8] || '').trim();
      const histN   = String(r[13] || '').trim();
      if (!company) return;

      if (needH && needH.toLowerCase() !== 'нет' && needH !== '—') {
        extractCleanPipeDiams_(needH).forEach(d => {
          ensureDiam_(d);
          demandByDiam[d].hot.push({ label: `[ЗАЯВКА ОЖИД] ${company} (${needH})`, pipeType: typeI, rev: 0 });
        });
      } else if (histN) {
        extractCleanPipeDiams_(histN).forEach(d => {
          ensureDiam_(d);
          demandByDiam[d].profile.push({ name: company, label: company, rev: 0 });
        });
      }
    });
  }

  Object.keys(demandByDiam).forEach(d => {
    demandByDiam[d].hot.sort((a, b) => b.rev - a.rev);
    demandByDiam[d].profile.sort((a, b) => b.rev - a.rev);
  });

  const numRows = supplySheet.getLastRow() - 1;
  const supData = supplySheet.getRange(2, 1, numRows, SUPPLY_HEADERS.length).getValues();
  const outColP = [];

  supData.forEach(r => {
    const status = String(r[13] || '').trim();
    const curColP = String(r[15] || '').trim();
    // Не затираем список заказчиков и заводов RFQ у строк «Поиск (Дефицит CRM)»!
    if (status === 'Поиск (Дефицит CRM)') {
      outColP.push([curColP]);
      return;
    }

    const sizeD = String(r[3] || '').trim();
    const lotType = String(r[4] || '').trim();
    const diams = typeof extractCleanPipeDiams_ === 'function' ? extractCleanPipeDiams_(sizeD) : [];
    if (diams.length === 0) {
      outColP.push([curColP]);
      return;
    }

    const matched = demandByDiam[diams[0]];
    if (!matched) {
      outColP.push(['Спрос в базе не найден']);
      return;
    }

    const compatHot = matched.hot.filter(h =>
      typeof doesStockMatchClientPipeType_ === 'function'
        ? doesStockMatchClientPipeType_(lotType, '', h.pipeType)
        : true
    );
    const seenProf = new Set();
    const topProf = [];
    matched.profile.forEach(p => {
      if (!seenProf.has(p.name) && topProf.length < 5) {
        seenProf.add(p.name);
        topProf.push(p.label);
      }
    });

    const lines = [];
    if (compatHot.length > 0) lines.push(compatHot.slice(0, 4).map(x => x.label).join(' | '));
    if (topProf.length > 0) lines.push('VIP/Профиль: ' + topProf.join(', '));
    outColP.push([lines.length > 0 ? lines.join('\n') : 'Спрос по этому типу не найден']);
  });

  supplySheet.getRange(2, 16, numRows, 1).setValues(outColP);
  ss.toast('Обратный Радар по «7. Закупки» обновлён (с фильтром по типу трубы и VIP)!', 'Обратный радар', 4);
}

// =================================================================
// 7. МОСТ «7. ЗАКУПКИ → СКЛАД» (ОПРИХОДОВАНИЕ С MTC GATE И ЗАКУПКОЙ G)
// =================================================================

function receiveSupplyLotToWarehouse() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getActiveSheet();

  if (sheet.getName().trim() !== SUPPLY_SHEET_NAME) {
    ui.alert('Перейдите на лист «7. Закупки» и выделите строку лота для оприходования на Склад.');
    return;
  }
  const range = sheet.getActiveRange();
  const row = range ? range.getRow() : 0;
  if (row < 2) {
    ui.alert('Выделите строку с товарным лотом (ниже заголовка).');
    return;
  }

  const r = sheet.getRange(row, 1, 1, SUPPLY_HEADERS.length).getValues()[0];
  const supplier = String(r[2] || '').trim();
  const sizeStr  = String(r[3] || '').trim();
  const pipeType = String(r[4] || '').trim() || 'Новая';
  const tons     = safeParseFloat(r[5]);
  const priceIn  = safeParseFloat(r[7]);
  const logist   = safeParseFloat(r[8]);
  const costTot  = safeParseFloat(r[9]) || (priceIn + logist);
  const sellGoal = safeParseFloat(r[10]);
  const status   = String(r[13] || '').trim();
  const mtcVal   = String(r[14] || '').trim();
  const noteVal  = String(r[16] || '').trim();

  if (status === 'Оприходовано на Склад') {
    ui.alert('Этот лот уже оприходован на Склад!');
    return;
  }
  if (tons <= 0) {
    ui.alert('Укажите фактический объём (тн) в колонке F перед оприходованием на Склад.');
    return;
  }

  const m = sizeStr.match(/(\d{2,4})\s*[хx×*]\s*(\d+(?:[.,]\d+)?)/);
  if (!m) {
    ui.alert(`Не удалось распознать диаметр и стенку в «${sizeStr}» (нужен формат 530х8).`);
    return;
  }
  const diam = parseFloat(m[1]);
  const wall = parseFloat(m[2].replace(',', '.'));

  // MTC Gate для новой трубы (п. 4.10 ТЗ)
  const isNewPipe = pipeType.toLowerCase().includes('нов');
  const isPlaceholderMtc = /требуется mtc|уточнить mtc|запросить mtc/i.test(mtcVal);
  let allowQuarantine = false;

  if (isNewPipe && (!mtcVal || isPlaceholderMtc)) {
    const resp = ui.alert(
      'MTC GATE — Нет сертификата качества',
      `Для НОВОЙ трубы ${sizeStr} в колонке O не указан сертификат качества.\n\n` +
      `Нажмите «ДА», чтобы оприходовать на Склад в статус «Карантин (Нет MTC)» (будет скрыто из Прайса).\n` +
      `Нажмите «НЕТ», чтобы отменить и сначала вписать сертификат в колонку O.`,
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) {
      sheet.getRange(row, 14).setValue('Карантин (Нет MTC)');
      styleSupplyRow_(sheet, row);
      return;
    }
    allowQuarantine = true;
  }

  if (typeof executeBatchAddition !== 'function') {
    ui.alert('Функция executeBatchAddition из Warehouse.js не найдена.');
    return;
  }

  const gostMatch = mtcVal.match(/(?:ГОСТ|ТУ|ТС)\s*[\d.\-]+/i);
  const gostStr = gostMatch ? gostMatch[0] : (isNewPipe ? 'ГОСТ 10704-91' : '—');
  const category = pipeType.toLowerCase().includes('нкт') ? 'НКТ' : 'ЭСВ';

  const msg = executeBatchAddition([{
    isNew: true,
    name: `Труба ${diam}х${wall}`,
    diam: diam,
    wall: wall,
    category: category,
    condition: pipeType,
    gost: gostStr,
    buyPriceTn: costTot,
    priceTn: sellGoal > 0 ? sellGoal : (costTot > 0 ? costTot + SUPPLY_MIN_MARGIN_ : 0),
    tons: tons,
    mtc: isPlaceholderMtc ? '' : `${supplier}: ${mtcVal}`,
    allowQuarantine: allowQuarantine,
    notes: [supplier, noteVal].filter(Boolean).join(' • ')
  }]);

  sheet.getRange(row, 14).setValue('Оприходовано на Склад');
  styleSupplyRow_(sheet, row);
  ui.alert('Оприходовано на Склад!', msg, ui.ButtonSet.OK);
}

// =================================================================
// 8. ОБРАБОТЧИК ИЗМЕНЕНИЙ (onEdit) И 100% КОДОВАЯ ПОКРАСКА
// =================================================================

function handleSupplySheetEdit(e, sheet) {
  if (!e || !e.range) return;
  const sheetName = sheet.getName().trim();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < 2) return;

  if (sheetName === SUPPLIERS_SHEET_NAME) {
    refreshSuppliersStyling_(sheet, row);
    return;
  }

  const tz = sheet.getParent().getSpreadsheetTimeZone() || 'Europe/Moscow';

  // Очистка пустой строки
  if (col === 3 || col === 4) {
    const rVals = sheet.getRange(row, 1, 1, SUPPLY_HEADERS.length).getValues()[0];
    if (!String(rVals[2] || '').trim() && !String(rVals[3] || '').trim()) {
      sheet.getRange(row, 1, 1, SUPPLY_HEADERS.length).clearContent().setBackground(null);
      return;
    }
  }

  // Выбор завода в C (3) → автоподстановка Базиса (G) и Качества (Q) из «7.1 Заводы»
  if (col === 3) {
    const chosenSup = String(e.value !== undefined ? e.value : sheet.getRange(row, 3).getValue() || '').trim();
    if (chosenSup) {
      const supObj = getSuppliersDirectory_(sheet.getParent()).find(s => s.name.toLowerCase() === chosenSup.toLowerCase());
      if (supObj) {
        if (!String(sheet.getRange(row, 7).getValue() || '').trim() && supObj.basis) {
          sheet.getRange(row, 7).setValue(supObj.basis);
        }
        if (!String(sheet.getRange(row, 17).getValue() || '').trim() && supObj.quality) {
          sheet.getRange(row, 17).setValue(`${supObj.tolerance} • ${supObj.quality}`);
        }
      }
    }
  }

  // Автоформат размера в D (4) + автогенерация ID (A), даты (B) и статуса (N)
  if (col === 4) {
    const rawSize = String(e.value !== undefined ? e.value : sheet.getRange(row, 4).getValue() || '').trim();
    if (rawSize && typeof formatPipeNeed_ === 'function') {
      const formatted = formatPipeNeed_(rawSize);
      if (formatted && formatted !== rawSize) {
        sheet.getRange(row, 4).setNumberFormat('@').setValue(formatted);
      }
    }
    const now = new Date();
    if (!sheet.getRange(row, 1).getValue()) {
      sheet.getRange(row, 1).setValue(`SUP-${Utilities.formatDate(now, tz, 'ddMM')}-${row - 1}`);
    }
    if (!sheet.getRange(row, 2).getValue()) {
      sheet.getRange(row, 2).setValue(`${Utilities.formatDate(now, tz, 'dd.MM')} • Ручной`);
    }
    if (!sheet.getRange(row, 14).getValue()) {
      sheet.getRange(row, 14).setValue('В рынке (Предложение)');
    }
  }

  // Авторасчёт Себестоимости с доставкой: J (10) = H (8) + I (9)
  if (col === 8 || col === 9) {
    const priceIn = safeParseFloat(sheet.getRange(row, 8).getValue());
    const logistics = safeParseFloat(sheet.getRange(row, 9).getValue());
    const totalCell = sheet.getRange(row, 10);
    if (priceIn > 0 || logistics > 0) totalCell.setValue(priceIn + logistics);
    else totalCell.clearContent();
  }

  // MTC Gate при ручном переводе новой трубы в «Выкуплено» без сертификата в O
  if (col === 14) {
    const newStatus = String(sheet.getRange(row, 14).getValue() || '').trim();
    const pipeType = String(sheet.getRange(row, 5).getValue() || '').toLowerCase();
    const mtcVal = String(sheet.getRange(row, 15).getValue() || '').trim();
    const isMissingMtc = !mtcVal || /требуется mtc|уточнить mtc|запросить mtc/i.test(mtcVal);
    if (pipeType.includes('нов') && newStatus === 'Выкуплено (В транзите)' && isMissingMtc) {
      sheet.getRange(row, 14).setValue('Карантин (Нет MTC)');
      sheet.getParent().toast(
        'MTC Gate: для выкупа НОВОЙ трубы внесите реквизиты сертификата/плавки в колонку O — пока поставлен «Карантин (Нет MTC)»!',
        'Контроль сертификата',
        6
      );
    }
  }

  styleSupplyRow_(sheet, row);
}

function styleSupplyRow_(sheet, row) {
  if (!sheet || row < 2) return;
  const r = sheet.getRange(row, 1, 1, SUPPLY_HEADERS.length).getValues()[0];
  const tz = sheet.getParent().getSpreadsheetTimeZone() || 'Europe/Moscow';
  const todayIso = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  // 1. Контроль Стоп-Лосса по Целевой продаже K (11) против Себестоимости J (10)
  const costJ = safeParseFloat(r[9]);
  const sellK = safeParseFloat(r[10]);
  const cellK = sheet.getRange(row, 11);
  if (costJ > 0 && sellK > 0) {
    const margin = sellK - costJ;
    if (margin < SUPPLY_MIN_MARGIN_) {
      cellK.setBackground('#fee2e2').setFontColor('#991b1b').setFontWeight('bold');
    } else {
      cellK.setBackground('#dcfce7').setFontColor('#166534').setFontWeight('bold');
    }
  } else {
    cellK.setBackground(null).setFontColor('#0f172a').setFontWeight('normal');
  }

  // 2. Светофор актуальности / брони M (13)
  const cellM = sheet.getRange(row, 13);
  const dlObj = typeof parseDeadlineObj_ === 'function' ? parseDeadlineObj_(r[12], tz) : { isoDate: '' };
  const status = String(r[13] || '').trim();
  if (!dlObj.isoDate || status === 'Архив / Ушло' || status === 'Оприходовано на Склад') {
    cellM.setBackground(null).setFontColor('#0f172a').setFontWeight('normal');
  } else if (dlObj.isoDate < todayIso) {
    cellM.setBackground('#fee2e2').setFontColor('#991b1b').setFontWeight('bold');
  } else if (dlObj.isoDate === todayIso) {
    cellM.setBackground('#dcfce7').setFontColor('#166534').setFontWeight('bold');
  } else {
    cellM.setBackground('#fef9c3').setFontColor('#854d0e').setFontWeight('normal');
  }

  // 3. Кодовая раскраска статуса лота N (14)
  const stStyle = getSupplyStatusStyle_(status);
  sheet.getRange(row, 14)
    .setBackground(stStyle.bg)
    .setFontColor(stStyle.color)
    .setFontWeight(stStyle.weight);
}

function getSupplyStatusStyle_(status) {
  switch (String(status || '').trim()) {
    case 'Поиск (Дефицит CRM)':    return { bg: '#fee2e2', color: '#991b1b', weight: 'bold' };
    case 'Котировка завода':       return { bg: '#e0f2fe', color: '#0369a1', weight: 'bold' };
    case 'В рынке (Предложение)':  return { bg: '#dbeafe', color: '#1e40af', weight: 'normal' };
    case 'Под бронью':             return { bg: '#fef3c7', color: '#92400e', weight: 'bold' };
    case 'Карантин (Нет MTC)':     return { bg: '#ffedd5', color: '#9a3412', weight: 'bold' };
    case 'Выкуплено (В транзите)': return { bg: '#dcfce7', color: '#166534', weight: 'bold' };
    case 'Оприходовано на Склад':  return { bg: '#f0fdf4', color: '#15803d', weight: 'normal' };
    case 'Архив / Ушло':           return { bg: '#f1f5f9', color: '#64748b', weight: 'normal' };
    default:                       return { bg: null, color: '#0f172a', weight: 'normal' };
  }
}

function refreshAllSupplyStyling_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return;
  sheet.clearConditionalFormatRules();
  const lastRow = sheet.getLastRow();
  for (let r = 2; r <= lastRow; r++) {
    styleSupplyRow_(sheet, r);
  }
}

function refreshSuppliersStyling_(sheet, singleRow) {
  if (!sheet || sheet.getLastRow() < 2) return;
  const start = singleRow || 2;
  const end = singleRow || sheet.getLastRow();
  for (let r = start; r <= end; r++) {
    const val = String(sheet.getRange(r, 3).getValue() || '').toLowerCase();
    const cell = sheet.getRange(r, 3);
    if (val.includes('основн') || val.includes('бел')) {
      cell.setBackground('#dcfce7').setFontColor('#166534').setFontWeight('bold');
    } else if (val.includes('чёрн') || val.includes('черн')) {
      cell.setBackground('#fee2e2').setFontColor('#991b1b').setFontWeight('bold');
    } else if (val.includes('резерв')) {
      cell.setBackground('#fef3c7').setFontColor('#92400e').setFontWeight('normal');
    } else {
      cell.setBackground(null).setFontColor('#0f172a').setFontWeight('normal');
    }
  }
}

// =================================================================
// 9. УТИЛИТЫ HTML-ПАРСИНГА ПИСЕМ ЗАВОДОВ
// =================================================================

function stripQuotedReplyHtml_(html) {
  let s = String(html || '');
  const cutMarkers = [
    /<div[^>]*class=["']?gmail_quote/i,
    /<blockquote\b/i,
    /<div[^>]*border-top:\s*solid\s*#(?:E1E1E1|B5C4DF)/i,
    /From:\s*(?:"|")?Кирилл|От:\s*Кирилл/i
  ];
  cutMarkers.forEach(rx => {
    const idx = s.search(rx);
    if (idx > 0) s = s.slice(0, idx);
  });
  return s;
}

function decodeHtmlEntities_(str) {
  return String(str || '')
    .replace(/ | /gi, ' ')
    .replace(/&/gi, '&')
    .replace(/</gi, '<')
    .replace(/>/gi, '>')
    .replace(/"/gi, '"')
    .replace(/&#(\d+);/g, (_, code) => String.fromCharCode(parseInt(code, 10)))
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function htmlToPlainLines_(html) {
  return decodeHtmlEntities_(
    String(html || '')
      .replace(/<(?:br|\/p|\/div|\/tr|\/li)[^>]*>/gi, '\n')
      .replace(/<[^>]+>/g, ' ')
  );
}

function extractHtmlTables_(html) {
  const tables = [];
  const tblRx = /<table\b[^>]*>([\s\S]*?)<\/table>/gi;
  let tm;
  while ((tm = tblRx.exec(String(html || ''))) !== null) {
    const rows = [];
    const trRx = /<tr\b[^>]*>([\s\S]*?)<\/tr>/gi;
    let rm;
    while ((rm = trRx.exec(tm[1])) !== null) {
      const cells = [];
      const tdRx = /<t[dh]\b[^>]*>([\s\S]*?)<\/t[dh]>/gi;
      let cm;
      while ((cm = tdRx.exec(rm[1])) !== null) {
        cells.push(decodeHtmlEntities_(cm[1].replace(/<[^>]+>/g, ' ')));
      }
      if (cells.some(Boolean)) rows.push(cells);
    }
    if (rows.length > 0) tables.push(rows);
  }
  return tables;
}

function mapHeaderColumns_(headerRow) {
  const norm = headerRow.map(h => String(h || '').toLowerCase());
  return {
    shop:  norm.findIndex(h => h.includes('цех')),
    batch: norm.findIndex(h => h.includes('парт')),
    heat:  norm.findIndex(h => h.includes('плавк')),
    diam:  norm.findIndex(h => h.includes('диаметр')),
    wall:  norm.findIndex(h => h.includes('толщин') || h.includes('стенк')),
    steel: norm.findIndex(h => h.includes('марка') || h.includes('стал')),
    gost:  norm.findIndex(h => h.includes('гост') || h.includes('ту')),
    price: norm.findIndex(h => h.includes('цена')),
    tons:  norm.findIndex(h => h.includes('тн') || h.includes('тонн'))
  };
}

function cleanWallStr_(raw) {
  const m = String(raw || '').replace(',', '.').match(/\d+(?:\.\d+)?/);
  if (!m) return '';
  return String(parseFloat(m[0]));
}

function extractPriceRub_(val) {
  const m = String(val || '').replace(/\u00A0/g, ' ').match(/(\d{2,3}[\s ]?\d{3}(?:[.,]\d{1,2})?)/);
  if (!m) return 0;
  return parseFloat(m[1].replace(/\s+/g, '').replace(',', '.')) || 0;
}

function updateOmkGrdPriceGridFromTable_(rows, grid) {
  if (!rows || rows.length < 3) return;
  const h0 = rows[0].join(' ').toLowerCase();
  if (!h0.includes('цена') || !h0.includes('ндс')) return;
  rows.slice(1).forEach(r => {
    if (r.length < 2) return;
    const pr = Math.round(extractPriceRub_(r[1]));
    if (pr < 25000 || pr > 200000) return;
    const m = String(r[0] || '').match(/(\d{2,4})(?:\s*[-–]\s*(\d{2,4}))?/);
    if (!m) return;
    const d1 = parseInt(m[1], 10);
    const d2 = m[2] ? parseInt(m[2], 10) : d1;
    const found = grid.find(g => d1 >= g.min && d2 <= g.max);
    if (found) found.price = pr;
    else grid.push({ min: d1, max: d2, price: pr });
  });
}

function lookupOmkGrdPrice_(diam, grid) {
  const hit = (grid || []).find(g => diam >= g.min && diam <= g.max);
  return hit ? hit.price : '';
}