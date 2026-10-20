/**
 * ЕДИНЫЙ ЦЕНТР УПРАВЛЕНИЯ ПРОДАЖАМИ И СКЛАДОМ (CTRM «Капитал-Плюс» v7.4)
 * Файл: PriceList.gs — Генератор прайс-листа и КП:
 * 1. На ПК (Gmail и Яндекс.Почта) возвращена широкая, солидная карточка 660px (1-в-1 как на ваших эталонных скриншотах «как было»):
 *    - Оригинальный крупный размер логотипа в подписи (не порезанный!).
 *    - Строка реквизитов «ООО «Капитал-Плюс» | ИНН ... | ОГРН ... | г. Воронеж...» на ПК помещается в ОДНУ ровную строку.
 *    - Просторная темно-синяя шапка #0f172a с плашкой «База: г. Воронеж» и гладко скругленными углами 12px.
 * 2. На телефоне полностью устранено обрезание ОГРН, сайта и Конфиденциальности:
 *    - На саму таблицу подписи ставится table-layout: fixed; width: 100%; max-width: 100%, а неразрывные пробелы (&nbsp;) из редактора Gmail заменяются на обычные, чтобы ОГРН, сайт и Конфиденциальность переносились ровно по границе экрана телефона без обрезки справа.
 *    - Через @media (max-width: 600px) и аккуратные пропорции контактов на телефоне масштаб подписи чуть компактнее, чтобы всё влезало с запасом.
 * 3. Сохранены все 4 обязательных правила: подпись из Gmail API без fallback-дубля, группировка позиций склада, рандомизация тем и анти-спам пауза.
 */

const PRICE_STD_DIAMS_ = [
  1420, 1220, 1020, 920, 820, 720, 630, 530, 478, 426, 402, 377, 351,
  325, 299, 273, 245, 219, 194, 168, 159, 152, 148, 146, 133, 127,
  114, 108, 102, 89, 76, 73, 60, 57
];

const PIPE_GROUP_ORDER_ = [
  "Трубы новые",
  "Трубы лежалые",
  "Трубы восстановленные",
  "Трубы б/у"
];

function openPriceListModal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clientsState = getClientsAndActiveSelection_(ss);
  const warehouseState = getWarehouseGroupedData_(ss);

  if (warehouseState.error) {
    SpreadsheetApp.getUi().alert(warehouseState.error);
    return;
  }

  const template = HtmlService.createTemplateFromFile('PriceListModal');
  template.groupedData = warehouseState.groupedData;
  template.groupedDataJson = JSON.stringify(warehouseState.groupedData);
  template.generatedDate = warehouseState.generatedDate;
  template.prefillClientKey = clientsState.prefillClientKey;
  template.clientsJson = JSON.stringify(clientsState.allClientsList);

  const html = template.evaluate()
    .setWidth(1220)
    .setHeight(880)
    .setTitle('Прайс-лист и КП • КАПИТАЛ-ПЛЮС');

  SpreadsheetApp.getUi().showModelessDialog(html, 'Прайс-лист и КП • КАПИТАЛ-ПЛЮС');
}

/**
 * Позволяет окну прайса получить свежий список всех клиентов и активную строку
 */
function getActiveClientFromSheetLive() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const state = getClientsAndActiveSelection_(ss);
  return {
    activeKey: state.prefillClientKey,
    clients: state.allClientsList
  };
}

/**
 * Извлекает короткое название города из полного юридического адреса (Колонка C)
 */
function extractShortCity_(rawAddr) {
  const s = String(rawAddr || '').trim();
  if (!s) return '';
  const gMatch = s.match(/(?:г\.|город)\s*([А-ЯЁа-яё\-]+(?:\s+[А-ЯЁа-яё\-]+)?)/i);
  if (gMatch && gMatch[1]) return 'г. ' + gMatch[1].trim();
  const oblMatch = s.match(/([А-ЯЁа-яё\-]+)\s+(?:обл|край|респ)/i);
  if (oblMatch && oblMatch[1]) return oblMatch[1].trim();
  const firstPart = s.split(',')[0].replace(/^\d{6}\s*/, '').trim();
  return firstPart.length > 28 ? firstPart.slice(0, 26) + '…' : firstPart;
}

/**
 * Разбивает ячейку D (ЛПР) на массив отдельных контактных лиц
 */
function parseLprList_(lprRaw) {
  const clean = String(lprRaw || '').trim();
  if (!clean) return [];
  const lines = clean
    .split(/[\n\/;,]+/)
    .map(s => s.replace(/\(.*?\)/g, '').trim())
    .filter(Boolean);
  return Array.from(new Set(lines));
}

/**
 * Собирает всех клиентов с 3 листов («3. Лист ожидания», «2. Золотой список (Киты)», «1. Обзвон (База)»)
 */
function getClientsAndActiveSelection_(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const activeSheetName = activeSheet ? activeSheet.getName().trim() : "";
  const activeRange = activeSheet ? activeSheet.getActiveRange() : null;

  const clientSheets = ["3. Лист ожидания", "2. Золотой список (Киты)", "1. Обзвон (База)"];
  const allClientsList = [];
  let prefillClientKey = "";

  clientSheets.forEach(sName => {
    const sh = typeof getSheetByNameSafe === 'function' ? getSheetByNameSafe(ss, sName) : ss.getSheetByName(sName);
    if (!sh) return;
    const lRow = sh.getLastRow();
    const lCol = sh.getLastColumn();
    if (lRow < 2 || lCol < 2) return;

    const numCols = Math.min(15, lCol);
    const data = sh.getRange(2, 1, lRow - 1, numCols).getValues();
    data.forEach((r, idx) => {
      const rowNum  = idx + 2;
      const inn     = String(r[0] || '').trim();
      const comp    = String(r[1] || '').trim();
      const cityRaw = String(r[2] || '').trim();
      const lprRaw  = String(r[3] || '').trim();
      const phRaw   = String(r[4] || '').trim();
      const rawEm   = String(r[5] || '').trim();
      const sphere  = String(r[6] || '').trim();
      const need    = String(r[7] || '').trim();
      const pType   = String(r[8] || '').trim();
      const pApp    = String(r[9] || '').trim();
      const notes   = String(r[10] || '').trim();
      const history = String(r[13] || '').trim();
      const tactic  = numCols >= 15 ? String(r[14] || '').trim() : '';

      const emList = Array.from(new Set(
        rawEm.split(/[,;\s\n]+/).map(x => x.trim()).filter(x => x && x.includes('@'))
      ));
      const lprList = parseLprList_(lprRaw);

      const combinedContext = `${need} ${notes} ${history}`;
      const pastDiams = Array.from(new Set((combinedContext.match(/\b\d{2,4}\b/g) || []).map(Number)))
        .filter(d => PRICE_STD_DIAMS_.includes(d));

      if (comp || emList.length > 0 || phRaw) {
        const actualSheetName = sh.getName().trim();
        const clientKey = sName + '_' + rowNum;
        allClientsList.push({
          key: clientKey,
          sheet: sName,
          row: rowNum,
          inn: inn,
          company: comp || 'Без названия',
          city: extractShortCity_(cityRaw),
          lpr: lprRaw,
          lprList: lprList,
          phone: phRaw.split('\n')[0] || '',
          emailStr: emList.join(', '),
          emailList: emList,
          sphere: sphere,
          need: (need && need.toLowerCase() !== 'нет' && need !== '—') ? need : '',
          pipeType: pType,
          application: pApp,
          tactic: tactic,
          diams: pastDiams
        });

        if (activeRange && (activeSheetName === sName || activeSheetName === actualSheetName) && activeRange.getRow() === rowNum) {
          prefillClientKey = clientKey;
        }
      }
    });
  });

  return { allClientsList: allClientsList, prefillClientKey: prefillClientKey };
}

/**
 * Считывает все позиции Склада в наличии и группирует их по категориям:
 * Трубы новые -> Трубы лежалые -> Трубы восстановленные -> Трубы б/у
 */
function getWarehouseGroupedData_(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = typeof getSheetByNameSafe === 'function' ? getSheetByNameSafe(ss, 'Склад') : ss.getSheetByName('Склад');
  if (!sheet) return { error: 'Ошибка: лист «Склад» не найден!' };

  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(14, sheet.getLastColumn());
  if (lastRow < 3) return { error: 'На листе «Склад» нет данных!' };

  const headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0].map(h => String(h).trim().toLowerCase());
  const colCategory  = headers.indexOf('категория') !== -1 ? headers.indexOf('категория') : 0;
  const colName      = headers.indexOf('позиция') !== -1 ? headers.indexOf('позиция') : 1;
  const colCondition = headers.indexOf('состояние') !== -1 ? headers.indexOf('состояние') : 2;
  const colDiam      = headers.findIndex(h => h.includes('ø') || h.includes('диаметр'));
  const colWall      = headers.findIndex(h => h.includes('стенка'));
  const colGost      = headers.findIndex(h => h.includes('гост') || h.includes('сырье') || h.includes('сталь'));

  let colPriceTn = headers.findIndex(h => h.includes('продаж') && h.includes('тн'));
  if (colPriceTn === -1) colPriceTn = headers.findIndex(h => h.includes('цена') && h.includes('тн'));

  let colPriceM = headers.findIndex(h => h.includes('продаж') && h.includes('м'));
  if (colPriceM === -1) colPriceM = headers.findIndex(h => h.includes('цена') && h.includes('м'));

  const colStatus  = headers.findIndex(h => h.includes('наличи')) !== -1 ? headers.findIndex(h => h.includes('наличи')) : 9;
  const colStockTn = headers.findIndex(h => h.includes('остаток') && h.includes('тн'));
  const colStockM  = headers.findIndex(h => h.includes('остаток') && h.includes('м'));

  const values = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();
  const rawItems = [];

  values.forEach((row, idx) => {
    const rawName = (row[colName] || '').toString().trim();
    if (!rawName || rawName.toUpperCase().startsWith('ИТОГО')) return;

    const rawStatus = (row[colStatus] || '').toString().trim().toLowerCase();
    const stockTn = safeParseFloat(row[colStockTn]);
    const stockM = safeParseFloat(row[colStockM]);

    if (rawStatus === 'нет') return;
    if (stockTn <= 0 && stockM <= 0) return;

    const rawCategory = (row[colCategory] || 'ЭСВ').toString().trim();
    const condition = (row[colCondition] || 'Б/у').toString().trim();
    const groupTitle = resolvePipeGroupTitle_(condition);

    const diamNum = safeParseFloat(row[colDiam]);
    const wallNum = safeParseFloat(row[colWall]);
    const diamStr = diamNum > 0 ? String(diamNum) : (row[colDiam] || '—').toString().trim();
    const wallStr = wallNum > 0 ? String(wallNum) : (row[colWall] || '—').toString().trim();

    let gostRaw = (row[colGost] || '').toString().trim();
    if (!gostRaw || gostRaw === '-') gostRaw = '—';

    const priceTnNum = safeParseFloat(row[colPriceTn]);
    const priceMNum = safeParseFloat(row[colPriceM]);

    const priceTnStr = priceTnNum > 0
      ? Math.round(priceTnNum).toString().replace(/\B(?=(\d{3})+(?!\d))/g, " ") + " ₽"
      : "По запросу";

    const priceMStr = priceMNum > 0
      ? Math.round(priceMNum).toString().replace(/\B(?=(\d{3})+(?!\d))/g, " ") + " ₽"
      : "—";

    const stockTnStr = stockTn > 0 ? stockTn + " тн" : "";
    const stockMStr = stockM > 0 ? stockM + " м" : "";

    let stockText = "В наличии";
    if (stockTnStr && stockMStr) stockText = stockTnStr + " / " + stockMStr;
    else if (stockTnStr) stockText = stockTnStr;
    else if (stockMStr) stockText = stockMStr;

    rawItems.push({
      id: 'pipe_' + idx,
      name: rawName,
      category: rawCategory,
      condition: condition,
      group: groupTitle,
      diam: diamStr,
      diamNum: diamNum || 0,
      wall: wallStr,
      wallNum: wallNum || 0,
      std: gostRaw,
      gost: gostRaw,
      stockTnStr: stockTnStr,
      stockMStr: stockMStr,
      stockText: stockText,
      stock: stockText,
      priceM: priceMStr,
      priceTn: priceTnStr
    });
  });

  const groupedData = {};
  PIPE_GROUP_ORDER_.forEach(grp => {
    const items = rawItems.filter(x => x.group === grp);
    if (items.length > 0) {
      items.sort((a, b) => (a.diamNum - b.diamNum) || (a.wallNum - b.wallNum));
      groupedData[grp] = items;
    }
  });

  const months = [
    "января", "февраля", "марта", "апреля", "мая", "июня",
    "июля", "августа", "сентября", "октября", "ноября", "декабря"
  ];
  const today = new Date();
  const dateStr = today.getDate() + " " + months[today.getMonth()] + " " + today.getFullYear() + " г.";

  return {
    groupedData: groupedData,
    allItems: rawItems,
    generatedDate: dateStr
  };
}

function resolvePipeGroupTitle_(conditionStr) {
  const condLower = String(conditionStr || '').toLowerCase();
  if (condLower.includes('восстанов') || condLower.includes('восст')) return "Трубы восстановленные";
  if (condLower.includes('лежал')) return "Трубы лежалые";
  if (condLower.includes('нов')) return "Трубы новые";
  return "Трубы б/у";
}

function groupItemsByConditionOrder_(items) {
  const map = {};
  PIPE_GROUP_ORDER_.forEach(g => { map[g] = []; });

  (items || []).forEach(it => {
    const grp = (it.group && PIPE_GROUP_ORDER_.includes(it.group))
      ? it.group
      : resolvePipeGroupTitle_(it.condition || it.group);
    if (!map[grp]) map[grp] = [];
    map[grp].push(it);
  });

  return PIPE_GROUP_ORDER_
    .filter(g => map[g] && map[g].length > 0)
    .map(g => ({
      groupTitle: g,
      items: map[g]
    }));
}

/**
 * Форматирует ГОСТ / сырье для подстроки в письме (ГОСТ 10704-91, из-под пара, из-под нефти, Металл)
 */
function formatGostClean_(rawGost) {
  const s = String(rawGost || '').trim();
  if (!s || s === '—' || s === '-') return '';
  const lower = s.toLowerCase();
  if (lower.startsWith('гост') || lower.startsWith('ту')) return s;
  if (lower.includes('нефт')) return 'из-под нефти';
  if (lower.includes('газ')) return 'из-под газа';
  if (lower.includes('вод')) return 'из-под воды';
  if (lower.includes('пар')) return 'из-под пара';
  return s;
}

/**
 * Воздушный PDF-файл ПОЛНОГО прайса со строгим разделением на группы:
 * «Трубы новые» -> «Трубы лежалые» -> «Трубы восстановленные» -> «Трубы б/у»
 */
function buildFullPricePdfBlob_(ss, warehouseState) {
  if (!warehouseState) warehouseState = getWarehouseGroupedData_(ss);
  const groupedData = warehouseState.groupedData || {};
  const generatedDate = warehouseState.generatedDate || '';

  let tablesHtml = '';
  PIPE_GROUP_ORDER_.forEach(grp => {
    const grpItems = groupedData[grp] || [];
    if (grpItems.length === 0) return;

    tablesHtml += `
      <tr>
        <td colspan="8" style="padding: 24px 8px 10px 8px; border-bottom: 1.5px solid #cbd5e1; background-color: #ffffff;">
          <span style="font-size: 14.5px; font-weight: 800; color: #0f172a; letter-spacing: -0.2px;">${grp}</span>
          <span style="font-size: 11px; font-weight: 500; color: #64748b; margin-left: 8px;">(${grpItems.length} поз.)</span>
        </td>
      </tr>
    `;

    grpItems.forEach((it) => {
      const gostDisplay = (it.std && it.std !== '-') ? it.std : '—';
      const tnPart = it.stockTnStr
        ? `<div style="font-weight: 800; color: #0f172a; font-size: 11.5px; line-height: 1.25;">${it.stockTnStr}</div>`
        : '';
      const mPart = it.stockMStr
        ? `<div style="color: #64748b; font-size: 10px; font-weight: 500; margin-top: 2px; line-height: 1.2;">${it.stockMStr}</div>`
        : '';
      const stockCell = (tnPart || mPart)
        ? (tnPart + mPart)
        : '<span style="color: #16a34a; font-weight: 600; font-size: 10.5px;">В наличии</span>';

      tablesHtml += `
        <tr>
          <td style="padding: 12px 8px; border-bottom: 1px solid #f1f5f9; vertical-align: middle; font-size: 10.5px; font-weight: 600; color: #475569;">
            ${it.category}
          </td>
          <td style="padding: 12px 8px; border-bottom: 1px solid #f1f5f9; font-size: 12px; font-weight: 600; color: #0f172a; vertical-align: middle;">
            ${it.name}
          </td>
          <td style="padding: 12px 8px; border-bottom: 1px solid #f1f5f9; font-size: 12px; text-align: center; color: #334155; vertical-align: middle;">
            ${it.diam}
          </td>
          <td style="padding: 12px 8px; border-bottom: 1px solid #f1f5f9; font-size: 12px; text-align: center; color: #334155; vertical-align: middle;">
            ${it.wall}
          </td>
          <td style="padding: 12px 8px; border-bottom: 1px solid #f1f5f9; font-size: 11px; text-align: center; color: #64748b; vertical-align: middle; line-height: 1.3;">
            ${gostDisplay}
          </td>
          <td style="padding: 12px 8px; border-bottom: 1px solid #f1f5f9; text-align: center; vertical-align: middle;">
            ${stockCell}
          </td>
          <td style="padding: 12px 8px; border-bottom: 1px solid #f1f5f9; font-size: 11.5px; font-weight: 500; color: #64748b; text-align: right; vertical-align: middle; white-space: nowrap;">
            ${it.priceM}
          </td>
          <td style="padding: 12px 8px; border-bottom: 1px solid #f1f5f9; font-size: 12.5px; font-weight: 800; color: #0f172a; text-align: right; vertical-align: middle; white-space: nowrap;">
            ${it.priceTn}
          </td>
        </tr>
      `;
    });
  });

  const pdfHtml = `
    <!DOCTYPE html>
    <html lang="ru">
    <head>
      <meta charset="utf-8">
      <style>
        @page { size: A4 portrait; margin: 12mm 12mm; }
        body {
          font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Helvetica Neue', Helvetica, Arial, sans-serif;
          color: #1e293b;
          margin: 0;
          padding: 0;
          font-size: 12px;
          line-height: 1.45;
          background-color: #ffffff;
        }
        table { width: 100%; border-collapse: collapse; background-color: #ffffff; }
        tr { page-break-inside: avoid; }
      </style>
    </head>
    <body>
      <table style="margin-bottom: 18px; border-bottom: 1px solid #e2e8f0; padding-bottom: 16px;">
        <tr>
          <td style="vertical-align: top; padding-bottom: 12px;">
            <div style="font-size: 11px; font-weight: 700; color: #2563eb; text-transform: uppercase; letter-spacing: 0.9px; margin-bottom: 4px;">
              СКЛАДСКАЯ СПРАВКА: ТРУБНЫЙ МЕТАЛЛОПРОКАТ
            </div>
            <div style="font-size: 28px; font-weight: 900; color: #0f172a; letter-spacing: -0.7px; line-height: 1.05; margin-bottom: 8px;">
              КАПИТАЛ-ПЛЮС
            </div>
            <div style="color: #64748b; font-size: 11px; font-weight: 500;">
              Актуально на: <strong style="color:#0f172a;">${generatedDate}</strong> &nbsp;<span style="color:#cbd5e1;">|</span>&nbsp; База: <strong style="color:#0f172a;">г. Воронеж</strong>
            </div>
          </td>
          <td style="text-align: right; vertical-align: top; width: 230px; padding-bottom: 12px;">
            <div style="font-size: 10px; color: #64748b; text-transform: uppercase; font-weight: 700; letter-spacing: 0.5px;">ОТДЕЛ ПРОДАЖ</div>
            <div style="font-size: 13.5px; font-weight: 700; color: #1e293b; margin-top: 3px;">Кирилл</div>
            <div style="font-size: 18px; font-weight: 900; color: #2563eb; margin-top: 2px; letter-spacing: -0.3px;">8 (906) 588-37-77</div>
          </td>
        </tr>
      </table>

      <table>
        <thead>
          <tr style="color: #64748b; font-size: 9.5px; text-transform: uppercase; letter-spacing: 0.5px;">
            <th style="padding: 10px 8px; text-align: left; font-weight: 700; border-top: 1px solid #e2e8f0; border-bottom: 1px solid #e2e8f0; width: 12%;">ГРУППА</th>
            <th style="padding: 10px 8px; text-align: left; font-weight: 700; border-top: 1px solid #e2e8f0; border-bottom: 1px solid #e2e8f0; width: 24%;">НАИМЕНОВАНИЕ ПРОФИЛЯ</th>
            <th style="padding: 10px 8px; text-align: center; font-weight: 700; border-top: 1px solid #e2e8f0; border-bottom: 1px solid #e2e8f0; width: 9%;">Ø (ММ)</th>
            <th style="padding: 10px 8px; text-align: center; font-weight: 700; border-top: 1px solid #e2e8f0; border-bottom: 1px solid #e2e8f0; width: 9%;">СТЕНКА</th>
            <th style="padding: 10px 8px; text-align: center; font-weight: 700; border-top: 1px solid #e2e8f0; border-bottom: 1px solid #e2e8f0; width: 14%;">ГОСТ / СТАЛЬ</th>
            <th style="padding: 10px 8px; text-align: center; font-weight: 700; border-top: 1px solid #e2e8f0; border-bottom: 1px solid #e2e8f0; width: 12%;">ОСТАТОК</th>
            <th style="padding: 10px 8px; text-align: right; font-weight: 700; border-top: 1px solid #e2e8f0; border-bottom: 1px solid #e2e8f0; width: 10%;">ЦЕНА / М</th>
            <th style="padding: 10px 8px; text-align: right; font-weight: 700; border-top: 1px solid #e2e8f0; border-bottom: 1px solid #e2e8f0; width: 10%;">ЦЕНА / ТН</th>
          </tr>
        </thead>
        <tbody>
          ${tablesHtml}
        </tbody>
      </table>

      <table style="margin-top: 26px; padding-top: 14px; border-top: 1px solid #e2e8f0;">
        <tr>
          <td style="padding-top: 12px; font-size: 10.5px; color: #64748b; line-height: 1.55; vertical-align: top;">
            <strong style="color: #0f172a;">ООО «Капитал-Плюс»</strong><br>
            Складская база: г. Воронеж<br>
            Металл в наличии на складе. Помощь с автодоставкой по РФ.
          </td>
          <td style="padding-top: 12px; font-size: 10.5px; color: #64748b; line-height: 1.55; text-align: right; vertical-align: top;">
            Цены указаны с учетом НДС 22%.<br>
            По вопросам резервирования: <strong style="color: #0f172a;">8 (906) 588-37-77</strong>
          </td>
        </tr>
      </table>
    </body>
    </html>
  `;

  const cleanDate = Utilities.formatDate(new Date(), "Europe/Moscow", "dd.MM.yyyy");
  const blob = Utilities.newBlob(pdfHtml, MimeType.HTML, 'PriceList.html')
    .getAs(MimeType.PDF)
    .setName(`Прайс-лист_Капитал-Плюс_Воронеж_${cleanDate}.pdf`);
  return blob;
}

/**
 * Разрешает выбранные имена ЛПР для клиента
 */
function resolveChosenLprString_(client) {
  if (!client) return '';
  if (Array.isArray(client.selectedLprs)) {
    return client.selectedLprs.map(s => String(s || '').trim()).filter(Boolean).join(', ');
  }
  if (client.selectedLpr !== undefined && client.selectedLpr !== null) {
    return String(client.selectedLpr).trim();
  }
  return String(client.lpr || '').split(/[\n\/;,]+/)[0].trim();
}

/**
 * Записывает лог в N (История) и ставит дедлайн +2 дня при копировании КП в мессенджер или скачивании PDF
 */
function logClientPriceAction(payload) {
  try {
    if (!payload || !payload.clients || payload.clients.length === 0) {
      return { success: false, message: "Клиент не выбран" };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timeZone = ss.getSpreadsheetTimeZone() || "Europe/Moscow";
    const nowStamp = Utilities.formatDate(new Date(), timeZone, "dd.MM.yyyy HH:mm");

    const mode = payload.mode || 'whatsapp';
    const isFull = !!payload.isFullPrice;
    const diams = payload.diams || [];
    const diamsBrief = (!isFull && diams.length > 0) ? ` (Ø ${diams.slice(0, 5).join(', ')})` : (isFull ? ': Полный' : '');

    let loggedCount = 0;
    const loggedNames = [];

    payload.clients.forEach(client => {
      if (!client || !client.sheet || !client.row || client.row < 2) return;
      const targetSheet = typeof getSheetByNameSafe === 'function' ? getSheetByNameSafe(ss, client.sheet) : ss.getSheetByName(client.sheet);
      if (!targetSheet) return;

      const hijVals = targetSheet.getRange(client.row, 8, 1, 3).getValues()[0];
      const curNeed = String(hijVals[0] || '').trim();
      const curType = String(hijVals[1] || '').trim();
      const curApp  = String(hijVals[2] || '').trim();
      const snap    = typeof buildNeedSnapshot_ === 'function' ? buildNeedSnapshot_(curNeed, curType, curApp) : '';

      const chosenLpr = resolveChosenLprString_(client);
      const lprTag    = chosenLpr ? ` (для: ${chosenLpr})` : '';

      let actionLabel = mode === 'pdf'
        ? `📄 Выгружен PDF прайс${diamsBrief}${lprTag}`
        : `💬 КП в Мессенджер${diamsBrief}${lprTag}`;

      const tag = `[${actionLabel}${snap} - ${nowStamp}]`;

      const curHistory = String(targetSheet.getRange(client.row, 14).getValue() || '').trim();
      targetSheet.getRange(client.row, 14).setValue(curHistory ? `${tag}\n${curHistory}` : tag);

      const nextFollowUp = new Date();
      nextFollowUp.setDate(nextFollowUp.getDate() + 2);
      if (nextFollowUp.getDay() === 6) nextFollowUp.setDate(nextFollowUp.getDate() + 2);
      if (nextFollowUp.getDay() === 0) nextFollowUp.setDate(nextFollowUp.getDate() + 1);
      const nextFollowUpStr = Utilities.formatDate(nextFollowUp, timeZone, "yyyy-MM-dd");

      const curDeadline = targetSheet.getRange(client.row, 12).getValue();
      if (!curDeadline) {
        targetSheet.getRange(client.row, 12).setValue(nextFollowUpStr);
      }

      if (client.sheet.trim() === "3. Лист ожидания" && typeof refreshSingleRowTactic_ === 'function') {
        refreshSingleRowTactic_(targetSheet, client.row);
      }

      loggedCount++;
      loggedNames.push(client.company || 'Клиент');
    });

    return {
      success: loggedCount > 0,
      loggedCount: loggedCount,
      loggedNames: loggedNames
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * ПУНКТ 1: Динамическое получение подписи из настроек Gmail (Gmail.Users.Settings.SendAs)
 * - На ПК сохраняет ваш оригинальный крупный логотип (не режет его!) и ставит ширину 100% от карточки 660px,
 *   поэтому реквизиты «ООО «Капитал-Плюс» | ИНН ... | ОГРН ... | г. Воронеж...» на ПК идут в ОДНУ красивую строку (как на скрине image_802cff.jpg).
 * - На телефоне:
 *   1) Заменяет неразрывные пробелы (&nbsp;) внутри текста на обычные пробелы, чтобы «ОГРН 1083668004914» и слова в «Конфиденциальность» свободно переносились на новую строку, а не резались справа.
 *   2) Добавляет table-layout: fixed на внешнюю таблицу подписи и word-break: break-all на ссылку сайта capital-two-pi.vercel.app, чтобы сайт и ОГРН никогда не вылезали за правый край экрана телефона.
 */
function getSenderSignature() {
  try {
    if (typeof Gmail !== 'undefined' && Gmail.Users && Gmail.Users.Settings) {
      const sendAs = Gmail.Users.Settings.SendAs.list("me").sendAs || [];
      const primary = sendAs.find(a => a.isDefault) || sendAs[0];
      if (primary && primary.signature) {
        let cleanedSig = String(primary.signature)
          .replace(/<!DOCTYPE[^>]*>/gi, '')
          .replace(/<head[\s\S]*?<\/head>/gi, '')
          .replace(/<\/?html[^>]*>/gi, '')
          .replace(/<\/?body[^>]*>/gi, '')
          // Убираем пустые верхние параграфы-распорки
          .replace(/^(?:\s*<p[^>]*>(?:&nbsp;|\s)*<\/p>\s*)+/i, '')
          .replace(/<div[^>]*padding-top:\s*(?:14|18|28)px[^>]*>(?:&nbsp;|\s)*<\/div>/gi, '')
          // Заменяем неразрывные пробелы (&nbsp; и \u00A0) на обычные пробелы, чтобы ОГРН и Конфиденциальность не склеивались в длинную неразрывную строку на телефоне
          .replace(/&nbsp;|\u00A0/gi, ' ')
          // На внешней таблице подписи ставим width: 100%; max-width: 100%; table-layout: fixed; — на ПК это дает полные 612px (реквизиты в 1 строку!), а на телефоне жестко держит рамку экрана
          .replace(/max-width:\s*(?:\d+px|100%)\s*;?/gi, 'max-width: 100%; width: 100%; table-layout: fixed; box-sizing: border-box;')
          // В таблице контактов возвращаем table-layout: auto; width: auto; чтобы телефоны стояли рядом с подписями
          .replace(/(<table[^>]*font-size:\s*1[23](?:\.5)?px[^>]*style="[^"]*)/gi, '$1; table-layout: auto; width: auto; max-width: 100%;')
          // Убираем жесткий min-width у ячейки логотипа, но САМУ КАРТИНКУ ЛОГОТИПА НЕ УМЕНЬШАЕМ (чтобы на ПК она была крупной и четкой, как было)
          .replace(/min-width:\s*85px\s*;?/gi, '')
          // Для длинных ссылок (сайт capital-two-pi.vercel.app и email) обязательно ставим word-break: break-all, чтобы на телефоне сайт не обрезался
          .replace(/word-break:\s*break-word/gi, 'word-break: break-all')
          .replace(/<td colspan="2"([^>]*)>/gi, '<td colspan="2"$1 style="padding-top: 14px; width: 100%; max-width: 100%; box-sizing: border-box; word-wrap: break-word; overflow-wrap: break-word; white-space: normal;">')
          .replace(/НДС\s*20\s*%/gi, 'НДС 22%')
          .trim();

        return {
          html: `
            <div class="sig-wrap" style="margin-top: 20px; padding-top: 18px; border-top: 1px dashed #e2e8f0; width: 100%; max-width: 100%; box-sizing: border-box; word-wrap: break-word; overflow-wrap: break-word;">
              ${cleanedSig}
            </div>
          `,
          isLive: true
        };
      }
    }
  } catch (e) {
    // Если Gmail API недоступен — возвращаем пустую строку без дубля
  }

  return { html: '', isLive: false };
}

/**
 * УТИЛИТА ДЛЯ ОБНОВЛЕНИЯ ПОДПИСИ В САМОМ GMAIL (сохраняет крупный логотип на ПК и защищает от обрезки на телефоне):
 * Запустите makeMyGmailSignatureFluid() 1 раз в редакторе Apps Script (кнопка «Выполнить»).
 */
function makeMyGmailSignatureFluid() {
  if (typeof Gmail === 'undefined' || !Gmail.Users || !Gmail.Users.Settings) {
    throw new Error('Включите службу Gmail API в разделе «Службы» (+) слева в редакторе Apps Script.');
  }

  const sendAsList = Gmail.Users.Settings.SendAs.list("me").sendAs || [];
  const primary = sendAsList.find(a => a.isDefault) || sendAsList[0];
  if (!primary) {
    throw new Error('Не найден профиль отправителя в Gmail.');
  }

  // Сохраняем ваш оригинальный тег <img> логотипа в полном размере без урезания до 64px
  const currentSig = String(primary.signature || '');
  const imgMatch = currentSig.match(/<img\b[^>]*>/i);
  const logoHtml = imgMatch
    ? imgMatch[0].replace(/max-width:\s*64px\s*;?/gi, 'max-width: 86px; width: 86px; display: block;')
    : '<span style="color:#999999; font-size:12px;">[ЛОГОТИП]</span>';

  const fluidSignatureHtml = `
<table cellpadding="0" cellspacing="0" border="0" width="100%" style="font-family: Arial, sans-serif; font-size: 13px; color: #333333; width: 100%; max-width: 100%; table-layout: fixed; line-height: 1.4; border-collapse: collapse; box-sizing: border-box;">
  <tr>
    <td width="88" style="padding-right: 14px; vertical-align: middle; width: 88px; max-width: 88px; text-align: center;">
      ${logoHtml}
    </td>
    <td style="border-left: 2px solid #0056b3; padding-left: 14px; vertical-align: middle; word-wrap: break-word; overflow-wrap: break-word;">
      <div style="font-size: 17px; font-weight: bold; color: #000000; margin-bottom: 2px; line-height: 1.2;">ООО Капитал-Плюс</div>
      <div style="font-size: 12px; color: #666666; margin-bottom: 10px; line-height: 1.3;">Отдел продаж и развития</div>
      <table cellpadding="0" cellspacing="0" border="0" style="font-size: 12.5px; font-family: Arial, sans-serif; width: auto; max-width: 100%; table-layout: auto; border-collapse: collapse;">
        <tr>
          <td style="font-weight: bold; color: #0056b3; padding-right: 8px; padding-bottom: 4px; white-space: nowrap; vertical-align: top;">Моб / TG / WA:</td>
          <td style="padding-bottom: 4px; white-space: nowrap;"><a href="tel:+79065883777" style="color: #333333; text-decoration: none; font-weight: 600;">+7 (906) 588-37-77</a></td>
        </tr>
        <tr>
          <td style="font-weight: bold; color: #0056b3; padding-right: 8px; padding-bottom: 4px; white-space: nowrap; vertical-align: top;">Тел. офис:</td>
          <td style="padding-bottom: 4px; white-space: nowrap;"><a href="tel:+74732797685" style="color: #333333; text-decoration: none;">+7 (473) 279-76-85</a></td>
        </tr>
        <tr>
          <td style="font-weight: bold; color: #0056b3; padding-right: 8px; padding-bottom: 4px; white-space: nowrap; vertical-align: top;">Email:</td>
          <td style="padding-bottom: 4px; word-break: break-all;"><a href="mailto:mail@kapital-p.ru" style="color: #0056b3; text-decoration: none; word-break: break-all;">mail@kapital-p.ru</a></td>
        </tr>
        <tr>
          <td style="font-weight: bold; color: #0056b3; padding-right: 8px; white-space: nowrap; vertical-align: top;">Сайт:</td>
          <td style="word-break: break-all;"><a href="https://capital-two-pi.vercel.app" target="_blank" style="color: #0056b3; text-decoration: none; word-break: break-all;">capital-two-pi.vercel.app</a></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td colspan="2" style="padding-top: 14px; width: 100%; max-width: 100%; box-sizing: border-box; word-wrap: break-word; overflow-wrap: break-word; white-space: normal;">
      <hr style="border: none; border-top: 1px solid #e0e0e0; margin: 0 0 12px 0;">
      <div style="font-size: 10.5px; color: #888888; line-height: 1.45; word-wrap: break-word; overflow-wrap: break-word; white-space: normal;">
        <strong>ООО «Капитал-Плюс»</strong> | ИНН 3663070544 | ОГРН 1083668004914 | НДС 22% | г. Воронеж, ул. Волгоградская, д. 48, оф. 76<br><br>
        <strong>Конфиденциальность:</strong> Данное электронное сообщение и любые приложения к нему являются конфиденциальными и предназначены исключительно для использования лицом или организацией, которым они адресованы. Если Вы не являетесь надлежащим адресатом, пожалуйста, немедленно сообщите об этом отправителю и удалите данное письмо.
      </div>
    </td>
  </tr>
</table>`.trim();

  Gmail.Users.Settings.SendAs.patch(
    { signature: fluidSignatureHtml },
    "me",
    primary.sendAsEmail
  );

  try {
    SpreadsheetApp.getUi().alert('Готово! Подпись в Gmail обновлена: на ПК крупный логотип и реквизиты в 1 строку, на телефоне ничего не обрезается.');
  } catch (e) {
    Logger.log('Подпись в Gmail успешно обновлена.');
  }
}

/**
 * Грамотное перечисление через запятую и союз «и» (например: «159, 219 и 530 мм»)
 */
function joinRussianList_(arr) {
  if (!arr || arr.length === 0) return '';
  if (arr.length === 1) return arr[0];
  if (arr.length === 2) return arr[0] + ' и ' + arr[1];
  return arr.slice(0, -1).join(', ') + ' и ' + arr[arr.length - 1];
}

/**
 * ПУНКТ 3: Вариативный генератор тем писем с полноценной логикой рандомизации:
 * - Для 1 позиции — тема с размером и состоянием.
 * - Для 2–4 позиций — перечисление диаметров через союз «и».
 * - Для полного прайса — диапазон (мин–макс диаметры) со складом в Воронеже и ценой «от ... ₽/тн».
 * - Обращение по имени-отчеству ЛПР и название компании.
 */
function buildDynamicSubject(items, clientCompany, chosenLpr, variantSeed) {
  const companyTag = clientCompany ? ` — ${clientCompany}` : '';
  const pickRandom = (templates) => {
    const offset = typeof variantSeed === 'number' ? variantSeed : 0;
    const rnd = Math.floor(Math.random() * templates.length);
    return templates[(rnd + offset) % templates.length];
  };

  let lprGreeting = "";
  if (chosenLpr) {
    const lprClean = chosenLpr.split(/[\n\/;,]+/)[0].trim();
    const lprLower = lprClean.toLowerCase();
    const isService = lprLower.includes('приемная') ||
                      lprLower.includes('снабжен') ||
                      lprLower.includes('бухгалтер') ||
                      lprLower.includes('склад') ||
                      lprLower.includes('отдел');

    if (!isService && lprClean) {
      const parts = lprClean.split(/\s+/);
      if (parts.length === 3) {
        lprGreeting = `${parts[1]} ${parts[2]}, `;
      } else if (parts.length === 2) {
        lprGreeting = `${parts[0]} ${parts[1]}, `;
      } else if (parts.length === 1 && parts[0].length > 1) {
        lprGreeting = `${parts[0]}, `;
      }
    }
  }

  if (!items || items.length === 0) {
    return `${lprGreeting}актуальное наличие стальных труб на складе в Воронеже${companyTag}`;
  }

  const diamList = [];
  const prices = [];
  const condSet = new Set();

  items.forEach(it => {
    const d = safeParseFloat(it.diam);
    if (d > 0 && !diamList.includes(d)) diamList.push(d);

    const p = safeParseFloat(it.priceTn);
    if (p > 0) prices.push(p);

    if (it.condition) condSet.add(it.condition.trim());
  });

  diamList.sort((a, b) => a - b);
  const minPrice = prices.length > 0 ? Math.min(...prices) : 0;
  const minPriceStr = minPrice > 0
    ? ` от ${Math.round(minPrice).toString().replace(/\B(?=(\d{3})+(?!\d))/g, " ")} ₽/тн`
    : '';

  const condArr = Array.from(condSet);
  let condSummary = "";
  if (condArr.length === 1) {
    condSummary = condArr[0].toLowerCase();
  } else if (condArr.length > 1) {
    const hasNew  = condArr.some(c => c.toLowerCase().includes('нов'));
    const hasLay  = condArr.some(c => c.toLowerCase().includes('лежал'));
    const hasRec  = condArr.some(c => c.toLowerCase().includes('восст'));
    const hasUsed = condArr.some(c => c.toLowerCase().includes('б/у'));

    const parts = [];
    if (hasNew) parts.push('новые');
    if (hasLay) parts.push('лежалые');
    if (hasRec) parts.push('восст.');
    if (hasUsed) parts.push('б/у');
    condSummary = joinRussianList_(parts);
  }

  // 1 позиция — тема с размером и состоянием
  if (items.length === 1) {
    const it = items[0];
    const diamWall = (it.diam && it.wall) ? `${it.diam}х${it.wall}` : it.name;
    const cond = it.condition ? ` (${it.condition.toLowerCase()})` : '';
    const stockBrief = (it.stock && it.stock !== 'В наличии') ? `, объем ${it.stock}` : '';
    const priceBrief = (it.priceTn && it.priceTn !== 'По запросу') ? ` — ${it.priceTn}/тн` : '';

    const singleTemplates = [
      `${lprGreeting}труба ${diamWall}${cond} в наличии на складе в Воронеже${priceBrief}${companyTag}`,
      `${lprGreeting}коммерческое предложение: труба ${diamWall}${cond}${stockBrief}${priceBrief}${companyTag}`,
      `${lprGreeting}расчет и наличие по трубе ${diamWall}${cond} (склад г. Воронеж)${companyTag}`,
      `${lprGreeting}спецификация на трубу ${diamWall}${cond} из наличия${priceBrief}${companyTag}`
    ];
    return pickRandom(singleTemplates);
  }

  // 2–4 позиции / диаметра — перечисление диаметров через союз «и»
  if (items.length >= 2 && items.length <= 4 && diamList.length <= 4) {
    const listForJoin = diamList.length >= 2
      ? diamList.map(d => `${d} мм`)
      : items.map(it => (it.diam && it.wall) ? `${it.diam}х${it.wall}` : it.name);
    const diamStr = joinRussianList_(listForJoin);
    const desc = condSummary ? ` (${condSummary})` : '';

    const multiTemplates = [
      `${lprGreeting}коммерческое предложение на трубы ${diamStr}${desc} со склада в Воронеже${companyTag}`,
      `${lprGreeting}подборка стальных труб ${diamStr}${desc} из наличия${minPriceStr}${companyTag}`,
      `${lprGreeting}наличие и цены по трубам ${diamStr}${desc} • Капитал-Плюс${companyTag}`,
      `${lprGreeting}спецификация по трубам ${diamStr}${desc} (склад г. Воронеж)${minPriceStr}${companyTag}`
    ];
    return pickRandom(multiTemplates);
  }

  if (diamList.length >= 2 && diamList.length <= 4) {
    const diamStr = joinRussianList_(diamList.map(d => `${d} мм`));
    const desc = condSummary ? ` (${condSummary})` : '';

    const multiTemplates = [
      `${lprGreeting}коммерческое предложение на трубы ${diamStr}${desc} со склада в Воронеже${companyTag}`,
      `${lprGreeting}подборка стальных труб ${diamStr}${desc} из наличия${minPriceStr}${companyTag}`,
      `${lprGreeting}наличие и цены по трубам ${diamStr}${desc} • Капитал-Плюс${companyTag}`,
      `${lprGreeting}спецификация по трубам ${diamStr}${desc} (склад г. Воронеж)${minPriceStr}${companyTag}`
    ];
    return pickRandom(multiTemplates);
  }

  // Полный прайс — диапазон (мин–макс диаметры) со складом в Воронеже и ценой «от ... ₽/тн»
  if (diamList.length >= 2) {
    const minD = diamList[0];
    const maxD = diamList[diamList.length - 1];
    const desc = condSummary ? ` (${condSummary})` : '';

    const fullTemplates = [
      `${lprGreeting}наличие труб ${minD}–${maxD} мм на складе в Воронеже${minPriceStr}${companyTag}`,
      `${lprGreeting}складская справка по стальным трубам ${minD}–${maxD} мм${desc} • г. Воронеж${minPriceStr}${companyTag}`,
      `${lprGreeting}актуальный прайс-лист на трубы ${minD}–${maxD} мм со склада в Воронеже${minPriceStr}${companyTag}`,
      `${lprGreeting}коммерческое предложение: трубный прокат ${minD}–${maxD} мм${desc} (г. Воронеж)${minPriceStr}${companyTag}`
    ];
    return pickRandom(fullTemplates);
  }

  return `${lprGreeting}актуальное наличие труб со склада в Воронеже${minPriceStr}${companyTag}`;
}

/**
 * Пакетная отправка КП:
 * - На ПК (Gmail и Яндекс.Почта): широкая представительная карточка max-width: 660px (1-в-1 как на скриншотах «как было» image_8028fc.png и image_802cff.jpg).
 * - На телефоне: адаптивное уменьшение масштаба подписи и отступов через @media (max-width: 600px) + table-layout: fixed, чтобы ОГРН, сайт и Конфиденциальность входили целиком без обрезки.
 */
function sendBatchProposals(data) {
  try {
    let clients = data.clients || [];
    const manualEmail = String(data.manualEmail || '').trim();

    if (clients.length === 0 && manualEmail) {
      const manualList = manualEmail.split(/[,;\s]+/).map(e => e.trim()).filter(e => e && e.includes('@'));
      if (manualList.length > 0) {
        clients = [{
          company: 'Партнёр', lpr: '', selectedLpr: '', emailStr: manualList.join(', '), sheet: '', row: 0
        }];
      }
    }

    if (clients.length === 0) {
      return { success: false, message: "Выберите хотя бы одну компанию с Email или введите адрес вручную!" };
    }

    const items = data.items || [];
    if (items.length === 0) {
      return { success: false, message: "Не выбраны позиции труб для отправки!" };
    }

    const customComment = String(data.customComment || '').trim();
    const isDraftOnly = (data.createDraftOnly === true);
    const attachFullPdf = (data.attachFullPdf !== false);
    const isFullPrice = (items.length > 15);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timeZone = ss.getSpreadsheetTimeZone() || "Europe/Moscow";
    const warehouseState = getWarehouseGroupedData_(ss);

    let fullPricePdfBlob = null;
    if (attachFullPdf && !warehouseState.error) {
      try {
        fullPricePdfBlob = buildFullPricePdfBlob_(ss, warehouseState);
      } catch (pdfErr) {}
    }

    const aliases = GmailApp.getAliases();
    const senderAlias = aliases.length > 0 ? aliases[0] : null;
    const userSignature = getSenderSignature();

    // Группируем позиции строго в порядке: Новые -> Лежалые -> Восстановленные -> Б/у
    const groupedBlocks = groupItemsByConditionOrder_(items);

    const sentDiams = Array.from(new Set(items.map(it => safeParseFloat(it.diam)).filter(d => d > 0)));
    const sentDiamsBrief = (!isFullPrice && sentDiams.length > 0) ? ` (Ø ${sentDiams.slice(0, 5).join(', ')})` : '';

    let sentCount = 0;
    const errors = [];

    for (let cIdx = 0; cIdx < clients.length; cIdx++) {
      const client = clients[cIdx];
      const activeEmailsArr = Array.isArray(client.selectedEmails) && client.selectedEmails.length > 0
        ? client.selectedEmails
        : String(client.emailStr || client.email || '').split(/[,;\s\n]+/).map(x => x.trim()).filter(x => x && x.includes('@'));
      const recipientEmails = activeEmailsArr.join(', ');

      if (!recipientEmails || !recipientEmails.includes('@')) {
        continue;
      }

      const isManual = !client.sheet;
      const clientCompany = client.company && client.company !== 'Партнёр' ? client.company : '';
      const chosenPerson = resolveChosenLprString_(client);
      const clientLprGreeting = chosenPerson ? `, ${chosenPerson}` : '';

      const subject = buildDynamicSubject(items, clientCompany, chosenPerson, cIdx);

      // Текстовая версия письма с группировкой по категориям
      let plain = [];
      plain.push(`Здравствуйте${clientLprGreeting}!\r\n`);
      plain.push(clientCompany
        ? `Направляем актуальное коммерческое предложение по наличию трубного металлопроката со склада для компании «${clientCompany}»:\r\n`
        : `Направляем актуальное коммерческое предложение по наличию трубного металлопроката с нашей базы:\r\n`
      );

      if (customComment) {
        plain.push(`Условия поставки / Примечание:\r\n${customComment}\r\n`);
      }

      groupedBlocks.forEach(block => {
        plain.push(`\r\n${block.groupTitle.toUpperCase()}:`);
        block.items.forEach((it, i) => {
          const gostFormatted = formatGostClean_(it.gost || it.std);
          const techParts = [it.category, gostFormatted].filter(Boolean).join(', ');
          const techSuffix = techParts ? ` (${techParts})` : '';
          plain.push(`${i + 1}. ${it.name}${techSuffix} | Ø ${it.diam} × ${it.wall} | Остаток: ${it.stock || it.stockText || 'В наличии'} | Цена: ${it.priceTn}/тн (${it.priceM}/м)`);
        });
      });

      plain.push("\r\n----------------------------------------");
      plain.push("• База отгрузки: г. Воронеж.");
      plain.push("• Логистика: Помощь с организацией попутной и прямой доставки по РФ.");
      plain.push("• Расчёт: Все цены указаны с учётом НДС 22%. Выставим счёт за 10 минут.");

      const plainBody = plain.join("\r\n");

      // СТРОКИ ТАБЛИЦЫ С АККУРАТНОЙ ГРУППИРОВКОЙ
      let rowsHtml = "";
      groupedBlocks.forEach((block, bIdx) => {
        const topPad = bIdx === 0 ? "14px" : "20px";
        rowsHtml += `
          <tr>
            <td colspan="4" style="padding: ${topPad} 10px 7px 10px; background-color: #ffffff; border-bottom: 1.5px solid #cbd5e1; text-align: left;">
              <span style="font-size: 13px; font-weight: 800; color: #0f172a; letter-spacing: -0.1px;">${block.groupTitle}</span>
              <span style="font-size: 11px; font-weight: 500; color: #64748b; margin-left: 6px;">(${block.items.length} поз.)</span>
            </td>
          </tr>
        `;

        block.items.forEach((it, idx) => {
          const rowBg = idx % 2 === 0 ? "#ffffff" : "#f8fafc";
          const formattedGost = formatGostClean_(it.gost || it.std);
          const techInfo = [it.category, it.condition, formattedGost].filter(Boolean).join(' &bull; ');
          const techLine = techInfo
            ? '<div style="font-size: 11px; color: #64748b; margin-top: 2px; line-height: 1.25;">' + techInfo + '</div>'
            : '';

          const rawStock = String(it.stock || it.stockText || '').trim();
          const stockParts = rawStock.split('/').map(s => s.trim()).filter(Boolean);
          let stockCellHtml = '';
          if (stockParts.length === 2) {
            stockCellHtml = `<div style="font-size: 12.5px; font-weight: 700; color: #0f172a; line-height: 1.2;">${stockParts[0]}</div><div style="font-size: 10.5px; color: #64748b; font-weight: 500; margin-top: 2px; line-height: 1.2;">${stockParts[1]}</div>`;
          } else {
            stockCellHtml = `<div style="font-size: 12px; font-weight: 600; color: #334155; line-height: 1.2;">${rawStock}</div>`;
          }

          rowsHtml += `
            <tr style="background-color: ${rowBg}; border-bottom: 1px solid #f1f5f9;">
              <td class="cell-pad" style="padding: 8px 4px; color: #334155; font-size: 12.5px; text-align: left; vertical-align: middle; word-wrap: break-word; overflow-wrap: break-word;">
                <div style="font-weight: 700; color: #0f172a; font-size: 13px; line-height: 1.25;">${it.name}</div>
                ${techLine}
              </td>
              <td class="cell-pad" style="padding: 8px 2px; text-align: center; color: #1e293b; font-size: 12px; font-weight: 600; white-space: normal; word-break: break-word; vertical-align: middle;">
                &empty;&nbsp;${it.diam}&times;${it.wall}
              </td>
              <td class="cell-pad" style="padding: 8px 2px; text-align: center; vertical-align: middle;">
                ${stockCellHtml}
              </td>
              <td class="cell-pad" style="padding: 8px 4px; text-align: right; white-space: nowrap; vertical-align: middle;">
                <div style="font-size: 13px; font-weight: 800; color: #0f172a; line-height: 1.2;">
                  ${it.priceTn}
                </div>
                <div style="font-size: 10px; color: #64748b; font-weight: 500; margin-top: 2px;">
                  ${it.priceM}/м
                </div>
              </td>
            </tr>
          `;
        });
      });

      const commentHtml = customComment ? `
        <div style="margin: 14px 0 16px 0; padding: 14px 16px; background-color: #f8fafc; border-left: 3px solid #475569; border-radius: 6px; font-size: 13px; color: #334155; line-height: 1.5; word-wrap: break-word;">
          <strong style="color: #0f172a; text-transform: uppercase; font-size: 10.5px; letter-spacing: 0.5px; display: block; margin-bottom: 4px;">Условия поставки / Примечание:</strong>
          ${customComment.replace(/\n/g, '<br>')}
        </div>
      ` : "";

      const introText = clientCompany
        ? `Направляем актуальное коммерческое предложение по наличию трубного металлопроката со склада для компании «${clientCompany}»:`
        : `Направляем актуальное коммерческое предложение по наличию трубного металлопроката с нашей базы:`;

      // ШИРОКАЯ КАРТОЧКА 660px КАК НА СКРИНШОТАХ «КАК БЫЛО» (image_8028fc.png и image_802cff.jpg)
      // + МОБИЛЬНАЯ АДАПТАЦИЯ ПОДПИСИ НА ТЕЛЕФОНЕ (чуть меньше масштаб внизу, чтобы ОГРН, сайт и Конфиденциальность не резались)
      const htmlBody = `<!DOCTYPE html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
    @media only screen and (max-width: 620px) {
      .outer-body { padding: 4px 0 !important; }
      .card-header-cell { padding: 14px 10px !important; }
      .card-body-cell { padding: 12px 4px !important; }
      .cell-pad { padding: 6px 2px !important; }
      .cell-pad div { font-size: 11.5px !important; }
      .sig-wrap { font-size: 12.5px !important; }
      .sig-wrap table { font-size: 12.5px !important; }
      .sig-wrap img { max-width: 60px !important; height: auto !important; }
      .sig-wrap td { padding-top: 3px !important; padding-bottom: 3px !important; }
    }
  </style>
</head>
<body class="outer-body" style="margin: 0; padding: 20px 10px; background-color: #f1f5f9; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; color: #0f172a; -webkit-text-size-adjust: 100%;">
  <div style="width: 100%; max-width: 760px; margin: 0 auto; box-sizing: border-box;">
    <table border="0" cellpadding="0" cellspacing="0" width="100%" style="width: 100%; max-width: 760px; table-layout: fixed; background-color: #ffffff; border: 1px solid #e2e8f0; border-radius: 12px; overflow: hidden; border-collapse: separate; box-sizing: border-box;">
      <tr>
        <td class="card-header-cell" style="background-color: #0f172a; padding: 24px 26px; color: #ffffff; border-radius: 11px 11px 0 0;">
          <table width="100%" border="0" cellpadding="0" cellspacing="0" style="width: 100%; border-collapse: collapse;">
            <tr>
              <td style="vertical-align: middle; text-align: left; padding-right: 10px;">
                <div style="font-size: 21px; font-weight: 900; letter-spacing: -0.4px; color: #ffffff; line-height: 1.15;">
                  ООО «КАПИТАЛ-ПЛЮС»
                </div>
                <div style="font-size: 12.5px; color: #94a3b8; margin-top: 4px; line-height: 1.3;">
                  Поставки трубного металлопроката со склада &bull; г. Воронеж
                </div>
              </td>
              <td style="vertical-align: middle; text-align: right; width: 130px; white-space: nowrap;">
                <span style="display: inline-block; padding: 6px 13px; background-color: #1e293b; border: 1px solid #334155; border-radius: 8px; color: #e2e8f0; font-size: 11.5px; font-weight: 600; white-space: nowrap;">
                  База: г. Воронеж
                </span>
              </td>
            </tr>
          </table>
        </td>
      </tr>
      <tr>
        <td class="card-body-cell" style="padding: 24px 20px; width: 100%; box-sizing: border-box;">
          <div style="font-size: 16px; font-weight: 700; color: #0f172a; margin-bottom: 6px;">
            Здравствуйте${clientLprGreeting}!
          </div>
          <div style="font-size: 13.5px; color: #475569; line-height: 1.5; margin-bottom: 16px; word-wrap: break-word;">
            ${introText}
          </div>
          ${commentHtml}
          <table width="100%" border="0" cellpadding="0" cellspacing="0" style="width: 100%; max-width: 100%; table-layout: fixed; border-collapse: collapse; margin-top: 8px;">
            <thead>
              <tr style="background-color: #f8fafc; border-bottom: 2px solid #e2e8f0; text-transform: uppercase; font-size: 10px; color: #475569; letter-spacing: 0.4px;">
                <th align="left" style="padding: 8px 4px; font-weight: 700; width: 34%;">Наименование</th>
                <th align="center" style="padding: 8px 2px; font-weight: 700; width: 22%;">Размер</th>
                <th align="center" style="padding: 8px 2px; font-weight: 700; width: 22%;">Остаток</th>
                <th align="right" style="padding: 8px 4px; font-weight: 700; width: 22%;">Цена</th>
              </tr>
            </thead>
            <tbody>
              ${rowsHtml}
            </tbody>
          </table>

          <div style="margin-top: 22px; padding: 14px 16px; background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; font-size: 12.5px; color: #475569; line-height: 1.65; word-wrap: break-word;">
            &bull; <strong style="color: #0f172a;">База отгрузки:</strong> г. Воронеж (металл в наличии на складе).<br>
            &bull; <strong style="color: #0f172a;">Логистика:</strong> Помощь с организацией попутной и прямой доставки по РФ.<br>
            &bull; <strong style="color: #0f172a;">Расчёт:</strong> Все цены указаны с учётом НДС 22%. Выставим счёт за 10 минут.
          </div>

          ${userSignature.html}
        </td>
      </tr>
    </table>
  </div>
</body>
</html>`;

      const emailOptions = {
        htmlBody: htmlBody,
        name: "Кирилл Манюхин"
      };

      if (fullPricePdfBlob) {
        emailOptions.attachments = [fullPricePdfBlob];
      }

      if (senderAlias) {
        emailOptions.from = senderAlias;
        emailOptions.replyTo = senderAlias;
      }

      try {
        if (isDraftOnly) {
          GmailApp.createDraft(recipientEmails, subject, plainBody, emailOptions);
        } else {
          GmailApp.sendEmail(recipientEmails, subject, plainBody, emailOptions);
        }
        sentCount++;

        // ФИКСАЦИЯ В ТАБЛИЦЕ (КОЛОНКА N / 14 - ИСТОРИЯ)
        if (!isManual && client.sheet && client.row && client.row >= 2) {
          const targetSheet = typeof getSheetByNameSafe === 'function' ? getSheetByNameSafe(ss, client.sheet) : ss.getSheetByName(client.sheet);
          if (targetSheet) {
            const nowStamp = Utilities.formatDate(new Date(), timeZone, "dd.MM.yyyy HH:mm");
            const hijVals = targetSheet.getRange(client.row, 8, 1, 3).getValues()[0];
            const curNeed = String(hijVals[0] || '').trim();
            const curType = String(hijVals[1] || '').trim();
            const curApp  = String(hijVals[2] || '').trim();
            const snap    = typeof buildNeedSnapshot_ === 'function' ? buildNeedSnapshot_(curNeed, curType, curApp) : '';
            const lprTag  = chosenPerson ? ` (для: ${chosenPerson})` : '';
            const pdfTag  = fullPricePdfBlob ? ' + PDF' : '';

            let actionType = "📧 Отправлено";
            if (isDraftOnly) actionType = `📝 Черновик КП${sentDiamsBrief}${pdfTag}${lprTag}`;
            else if (isFullPrice) actionType = `📧 Прайс: Полный${pdfTag}${lprTag}`;
            else actionType = `📧 КП${sentDiamsBrief}${pdfTag}${lprTag}`;

            const tag = `[${actionType}${snap} - ${nowStamp}]`;

            const curHistory = String(targetSheet.getRange(client.row, 14).getValue() || '').trim();
            targetSheet.getRange(client.row, 14).setValue(curHistory ? `${tag}\n${curHistory}` : tag);

            const nextFollowUp = new Date();
            nextFollowUp.setDate(nextFollowUp.getDate() + 2);
            if (nextFollowUp.getDay() === 6) nextFollowUp.setDate(nextFollowUp.getDate() + 2);
            if (nextFollowUp.getDay() === 0) nextFollowUp.setDate(nextFollowUp.getDate() + 1);
            const nextFollowUpStr = Utilities.formatDate(nextFollowUp, timeZone, "yyyy-MM-dd");

            const curDeadline = targetSheet.getRange(client.row, 12).getValue();
            if (!curDeadline) {
              targetSheet.getRange(client.row, 12).setValue(nextFollowUpStr);
            }

            if (client.sheet.trim() === "3. Лист ожидания" && typeof refreshSingleRowTactic_ === 'function') {
              refreshSingleRowTactic_(targetSheet, client.row);
            }
          }
        }
      } catch (errOne) {
        errors.push(`${recipientEmails}: ${errOne.message}`);
      }

      // ПУНКТ 4: Защита от спам-блокировок при рассылке разным клиентам
      if (clients.length > 1 && cIdx < clients.length - 1 && !isDraftOnly) {
        Utilities.sleep(Math.floor(Math.random() * 1000) + 2500);
      }
    }

    return {
      success: true,
      sentCount: sentCount,
      totalCount: clients.length,
      isDraft: isDraftOnly,
      errors: errors
    };
  } catch (err) {
    return { success: false, message: err.message };
  }
}