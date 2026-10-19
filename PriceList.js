/**
 * ЕДИНЫЙ ЦЕНТР УПРАВЛЕНИЯ ПРОДАЖАМИ И СКЛАДОМ (CTRM «Капитал-Плюс» v6.4)
 * Файл: PriceList.gs — Генератор прайс-листа и КП:
 * 1. Динамический подхват активной строки клиента из таблицы в реальном времени (даже при открытом окне).
 * 2. Выбор конкретного ЛПР (на чье имя отправлять), если в ячейке D записано несколько человек.
 * 3. Разделение труб в теле письма по группам (Новые / Лежалые / Восстановленные / Б/у) — 1-в-1 как в PDF.
 * 4. Автоматическое формирование и прикрепление PDF-файла ПОЛНОГО прайса к каждому письму.
 * 5. Автоматическая запись лога в N (История) при копировании в WhatsApp/TG и при выгрузке PDF.
 */

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
  template.generatedDate = warehouseState.generatedDate;
  template.prefillClientKey = clientsState.prefillClientKey;
  template.clientsJson = JSON.stringify(clientsState.allClientsList);

  const html = template.evaluate()
    .setWidth(1180)
    .setHeight(860)
    .setTitle('Прайс-лист и КП • КАПИТАЛ-ПЛЮС');

  SpreadsheetApp.getUi().showModelessDialog(html, 'Прайс-лист и КП • КАПИТАЛ-ПЛЮС');
}

/**
 * Позволяет открытому окну прайса в любой момент подтянуть клиента,
 * на строку которого менеджер только что кликнул в таблице!
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
 * Собирает всех клиентов с 3 листов и разбивает колонку D (ЛПР) на массив отдельных имен,
 * чтобы в интерфейсе можно было в 1 клик выбрать, на чье имя отправлять КП!
 */
function getClientsAndActiveSelection_(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const activeSheetName = activeSheet ? activeSheet.getName().trim() : "";
  const activeRange = activeSheet ? activeSheet.getActiveRange() : null;

  const clientSheets = ["1. Обзвон (База)", "2. Золотой список (Киты)", "3. Лист ожидания"];
  const allClientsList = [];
  let prefillClientKey = "";

  clientSheets.forEach(sName => {
    const sh = ss.getSheetByName(sName);
    if (!sh) return;
    const lRow = sh.getLastRow();
    const lCol = sh.getLastColumn();
    if (lRow < 2 || lCol < 2) return;

    const data = sh.getRange(2, 1, lRow - 1, Math.min(14, lCol)).getValues();
    data.forEach((r, idx) => {
      const rowNum = idx + 2;
      const inn    = String(r[0] || '').trim();
      const comp   = String(r[1] || '').trim();
      const lprRaw = String(r[3] || '').trim();
      const phRaw  = String(r[4] || '').trim();
      const rawEm  = String(r[5] || '').trim();
      const need   = String(r[7] || '').trim();

      const emList = rawEm.split(/[,;\s\n]+/).map(x => x.trim()).filter(x => x && x.includes('@'));
      // Разбиваем ЛПР по переносам строк, слэшам и точкам с запятой, чтобы можно было выбрать конкретного человека
      const lprList = lprRaw
        .split(/[\n\/;]+/)
        .map(s => s.replace(/\(.*?\)/g, '').trim())
        .filter(Boolean);

      if (comp || emList.length > 0 || phRaw) {
        const clientKey = sName + '_' + rowNum;
        allClientsList.push({
          key: clientKey,
          sheet: sName,
          row: rowNum,
          inn: inn,
          company: comp || 'Без названия',
          lpr: lprRaw,
          lprList: lprList,
          phone: phRaw.split('\n')[0] || '',
          emailStr: emList.join(', '),
          emailList: emList,
          need: need
        });

        if (activeRange && activeSheetName === sName && activeRange.getRow() === rowNum) {
          prefillClientKey = clientKey;
        }
      }
    });
  });

  allClientsList.sort((a, b) => (a.company || '').localeCompare(b.company || ''));
  return { allClientsList: allClientsList, prefillClientKey: prefillClientKey };
}

/**
 * Считывает все позиции Склада в наличии и группирует их по категориям:
 * Новые -> Лежалые -> Восстановленные -> Б/у
 */
function getWarehouseGroupedData_(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Склад');
  if (!sheet) return { error: '⚠️ Ошибка: лист «Склад» не найден!' };

  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(14, sheet.getLastColumn());
  if (lastRow < 3) return { error: '⚠️ На листе «Склад» нет данных!' };

  const headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0].map(h => String(h).trim().toLowerCase());
  const colCategory  = headers.indexOf('категория') !== -1 ? headers.indexOf('категория') : 0;
  const colName      = headers.indexOf('позиция') !== -1 ? headers.indexOf('позиция') : 1;
  const colCondition = headers.indexOf('состояние') !== -1 ? headers.indexOf('состояние') : 2;
  const colDiam      = headers.findIndex(h => h.includes('ø') || h.includes('диаметр'));
  const colWall      = headers.findIndex(h => h.includes('стенка'));
  const colGost      = headers.findIndex(h => h.includes('гост') || h.includes('сырье'));

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

  const groupOrder = [
    "Трубы новые",
    "Трубы лежалые",
    "Трубы восстановленные",
    "Трубы б/у"
  ];

  const groupedData = {};
  groupOrder.forEach(grp => {
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
  const groupOrder = [
    "Трубы новые",
    "Трубы лежалые",
    "Трубы восстановленные",
    "Трубы б/у"
  ];
  const map = {};
  groupOrder.forEach(g => { map[g] = []; });

  (items || []).forEach(it => {
    const grp = it.group || resolvePipeGroupTitle_(it.condition);
    if (!map[grp]) map[grp] = [];
    map[grp].push(it);
  });

  return groupOrder.filter(g => map[g] && map[g].length > 0).map(g => ({
    groupTitle: g,
    items: map[g]
  }));
}

/**
 * Автоматически генерирует официальный PDF-файл ПОЛНОГО прайса склада («КАПИТАЛ-ПЛЮС»)
 * для вложения в письмо или прямого скачивания!
 */
function buildFullPricePdfBlob_(ss, warehouseState) {
  if (!warehouseState) warehouseState = getWarehouseGroupedData_(ss);
  const groupedData = warehouseState.groupedData || {};
  const generatedDate = warehouseState.generatedDate || '';

  let tablesHtml = '';
  for (const grp in groupedData) {
    tablesHtml += `
      <tr style="background-color:#e2e8f0;">
        <td colspan="6" style="padding:8px 10px; font-weight:800; font-size:12px; color:#0f172a; text-transform:uppercase; letter-spacing:0.5px;">
          ▪ ${grp}
        </td>
      </tr>
    `;
    groupedData[grp].forEach((it, idx) => {
      const bg = idx % 2 === 0 ? '#ffffff' : '#f8fafc';
      tablesHtml += `
        <tr style="background-color:${bg}; border-bottom:1px solid #e2e8f0;">
          <td style="padding:6px 8px; font-size:11px; color:#334155;">${it.category}</td>
          <td style="padding:6px 8px; font-size:11px; font-weight:700; color:#0f172a;">${it.name}</td>
          <td style="padding:6px 8px; font-size:11px; text-align:center; font-weight:600;">Ø ${it.diam}×${it.wall}</td>
          <td style="padding:6px 8px; font-size:10px; text-align:center; color:#475569;">${it.std || '—'}</td>
          <td style="padding:6px 8px; font-size:11px; text-align:center; font-weight:600; color:#0f172a;">${it.stockText}</td>
          <td style="padding:6px 8px; font-size:11px; text-align:right; font-weight:700; color:#0f172a;">
            ${it.priceTn} <span style="font-weight:400; font-size:10px; color:#64748b;">(${it.priceM}/м)</span>
          </td>
        </tr>
      `;
    });
  }

  const pdfHtml = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="utf-8">
      <style>
        @page { size: A4 portrait; margin: 12mm 10mm; }
        body { font-family: Arial, Helvetica, sans-serif; color: #0f172a; margin: 0; padding: 0; font-size: 11px; }
        table { width: 100%; border-collapse: collapse; }
        tr { page-break-inside: avoid; }
      </style>
    </head>
    <body>
      <table style="border-bottom: 2px solid #0f172a; padding-bottom: 10px; margin-bottom: 14px;">
        <tr>
          <td>
            <div style="font-size:10px; font-weight:700; color:#2563eb; text-transform:uppercase; letter-spacing:1px;">Полный прайс-лист складского наличия</div>
            <div style="font-size:22px; font-weight:900; color:#0f172a; margin:2px 0;">ООО «КАПИТАЛ-ПЛЮС»</div>
            <div style="font-size:11px; color:#475569;">Актуально на: <strong>${generatedDate}</strong> &nbsp;|&nbsp; База отгрузки: <strong>г. Воронеж</strong></div>
          </td>
          <td style="text-align:right; vertical-align:top;">
            <div style="font-size:11px; font-weight:700; color:#0f172a;">Кирилл Манюхин</div>
            <div style="font-size:15px; font-weight:900; color:#2563eb; margin-top:2px;">+7 (906) 588-37-77</div>
            <div style="font-size:10px; color:#475569;">Тел. офис: +7 (473) 279-76-85 | mail@kapital-p.ru</div>
          </td>
        </tr>
      </table>

      <table style="border: 1px solid #cbd5e1;">
        <thead>
          <tr style="background-color:#0f172a; color:#ffffff; font-size:10px; text-transform:uppercase;">
            <th style="padding:7px 8px; text-align:left; width:10%;">Тип</th>
            <th style="padding:7px 8px; text-align:left; width:26%;">Наименование</th>
            <th style="padding:7px 8px; text-align:center; width:13%;">Размер</th>
            <th style="padding:7px 8px; text-align:center; width:15%;">ГОСТ / Сырье</th>
            <th style="padding:7px 8px; text-align:center; width:16%;">Остаток</th>
            <th style="padding:7px 8px; text-align:right; width:20%;">Цена с НДС</th>
          </tr>
        </thead>
        <tbody>
          ${tablesHtml}
        </tbody>
      </table>

      <div style="margin-top:12px; font-size:10px; color:#475569; line-height:1.5;">
        <strong>Условия поставки:</strong> Склад г. Воронеж (погрузка в открытый автотранспорт). Помощь с организацией прямой и попутной доставки (CPT) по РФ. Все цены указаны с учетом НДС 20%.
      </div>
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
 * Записывает лог в N (История) и ставит дедлайн +2 дня при копировании КП в WhatsApp/TG или скачивании PDF!
 */
function logClientPriceAction(payload) {
  try {
    if (!payload || !payload.clients || payload.clients.length === 0) {
      return { success: false, message: "Клиент не выбран" };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timeZone = ss.getSpreadsheetTimeZone() || "Europe/Moscow";
    const nowStamp = Utilities.formatDate(new Date(), timeZone, "dd.MM.yyyy HH:mm");

    const mode = payload.mode || 'whatsapp'; // 'whatsapp' | 'pdf'
    const isFull = !!payload.isFullPrice;
    const diams = payload.diams || [];
    const diamsBrief = (!isFull && diams.length > 0) ? ` (Ø ${diams.slice(0, 5).join(', ')})` : (isFull ? ': Полный' : '');

    let loggedCount = 0;
    const loggedNames = [];

    payload.clients.forEach(client => {
      if (!client || !client.sheet || !client.row || client.row < 2) return;
      const targetSheet = ss.getSheetByName(client.sheet);
      if (!targetSheet) return;

      const hijVals = targetSheet.getRange(client.row, 8, 1, 3).getValues()[0];
      const curNeed = String(hijVals[0] || '').trim();
      const curType = String(hijVals[1] || '').trim();
      const curApp  = String(hijVals[2] || '').trim();
      const snap    = typeof buildNeedSnapshot_ === 'function' ? buildNeedSnapshot_(curNeed, curType, curApp) : '';

      const chosenLpr = String(client.selectedLpr || '').trim();
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

function getSenderSignature() {
  try {
    if (typeof Gmail !== 'undefined' && Gmail.Users && Gmail.Users.Settings) {
      const sendAs = Gmail.Users.Settings.SendAs.list("me").sendAs || [];
      const primary = sendAs.find(a => a.isDefault) || sendAs[0];
      if (primary && primary.signature) {
        return {
          html: '<div style="margin-top: 24px; padding-top: 14px; border-top: 1px solid #cbd5e1;">' + primary.signature + '</div>',
          isLive: true
        };
      }
    }
  } catch (e) {}

  const fallbackHtml = `
    <div style="margin-top: 24px; padding-top: 16px; border-top: 2px solid #e2e8f0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; font-size: 13px; color: #334155; line-height: 1.5;">
      <div style="font-weight: 800; font-size: 16px; color: #0f172a; letter-spacing: -0.3px;">Кирилл Манюхин</div>
      <div style="color: #64748b; font-size: 12px; margin-bottom: 8px;">Менеджер по продажам и развитию в маркетинге</div>
      
      <div style="font-size: 12px; color: #1e293b; margin-bottom: 4px;">
        <strong>Моб / TG / WA:</strong> <a href="tel:+79065883777" style="color: #2563eb; text-decoration: none; font-weight: 600;">+7 (906) 588-37-77</a> &nbsp;|&nbsp; 
        <strong>Тел. офис:</strong> +7 (473) 279-76-85
      </div>
      <div style="font-size: 12px; color: #1e293b; margin-bottom: 8px;">
        <strong>Email:</strong> <a href="mailto:mail@kapital-p.ru" style="color: #2563eb; text-decoration: none;">mail@kapital-p.ru</a> &nbsp;|&nbsp; 
        <strong>Сайт:</strong> <a href="https://capital-two-pi.vercel.app" style="color: #2563eb; text-decoration: none;" target="_blank">capital-two-pi.vercel.app</a>
      </div>
      
      <div style="font-size: 11px; color: #64748b; line-height: 1.4; margin-bottom: 8px;">
        ООО «Капитал-Плюс» | ИНН 3663070544 | ОГРН 1083668004914 | НДС 20% | г. Воронеж, ул. Волгоградская, д. 48, оф. 76
      </div>

      <div style="font-size: 10px; color: #94a3b8; line-height: 1.3; font-style: italic;">
        Конфиденциальность: Данное электронное сообщение предназначено исключительно для лица или организации, которым оно адресовано.
      </div>
    </div>
  `;

  return { html: fallbackHtml, isLive: false };
}

function buildDynamicSubject(items, clientCompany, chosenLpr) {
  const companyTag = clientCompany ? ` • ${clientCompany}` : '';

  let lprGreeting = "";
  if (chosenLpr) {
    const lprClean = chosenLpr.split(/[\n\/;]+/)[0].trim();
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
    return `${lprGreeting}наличие трубного металлопроката со склада Воронеж${companyTag}`;
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
    ? ` от ${Math.round(minPrice).toString().replace(/\B(?=(\d{3})+(?!\d))/g, " ")} ₽` 
    : '';

  const condArr = Array.from(condSet);
  let condSummary = "";
  if (condArr.length === 1) {
    condSummary = condArr[0].toLowerCase();
  } else if (condArr.length > 1) {
    const hasNew = condArr.some(c => c.toLowerCase().includes('нов'));
    const hasRec = condArr.some(c => c.toLowerCase().includes('восст'));
    const hasUsed = condArr.some(c => c.toLowerCase().includes('б/у'));
    const hasLay = condArr.some(c => c.toLowerCase().includes('лежал'));

    const parts = [];
    if (hasNew) parts.push('новые');
    if (hasLay) parts.push('лежалые');
    if (hasRec) parts.push('восст.');
    if (hasUsed) parts.push('б/у');
    condSummary = parts.join(' и ');
  }

  if (items.length === 1) {
    const it = items[0];
    const diamWall = (it.diam && it.wall) ? `${it.diam}×${it.wall}` : it.name;
    const cond = it.condition ? ` (${it.condition.toLowerCase()})` : '';
    const stockBrief = (it.stock && it.stock !== 'В наличии') ? ` [${it.stock}]` : '';
    const priceBrief = (it.priceTn && it.priceTn !== 'По запросу') ? ` по ${it.priceTn}` : '';
    return `${lprGreeting}наличие: Труба ${diamWall}${cond}${stockBrief}${priceBrief} со склада Воронеж${companyTag}`;
  }

  if (diamList.length >= 2 && diamList.length <= 4) {
    const diamStr = diamList.map(d => `Ø${d}`).join(', ');
    const desc = condSummary ? ` (${condSummary})` : '';
    return `${lprGreeting}наличие: трубы ${diamStr}${desc}${minPriceStr} со склада Воронеж${companyTag}`;
  }

  if (diamList.length > 4) {
    const minD = diamList[0];
    const maxD = diamList[diamList.length - 1];
    const desc = condSummary ? ` (${condSummary})` : '';
    return `${lprGreeting}наличие: трубы Ø ${minD}–${maxD} мм${desc}${minPriceStr} со склада Воронеж${companyTag}`;
  }

  return `${lprGreeting}актуальное наличие труб со склада Воронеж${minPriceStr}${companyTag}`;
}

/**
 * Пакетная отправка КП:
 * - В теле письма: выбранные позиции (или весь прайс), РАЗДЕЛЕННЫЕ ПО ГРУППАМ (Новые / Лежалые / Восстановленные / Б/у).
 * - Во вложении (attachments): ВСЕГДА прикрепляется официальный PDF-файл ПОЛНОГО прайса склада!
 * - В обращении: используется конкретное выбранное имя ЛПР (selectedLpr).
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
      return { success: false, message: "Укажите адрес Email или выберите компанию из списка!" };
    }

    const items = data.items || [];
    if (items.length === 0) {
      return { success: false, message: "Не выбраны позиции труб для отправки!" };
    }

    const customComment = String(data.customComment || '').trim();
    const isDraftOnly = (data.createDraftOnly === true);
    const attachFullPdf = (data.attachFullPdf !== false); // По умолчанию ВСЕГДА прикрепляем PDF полного прайса
    const isFullPrice = (items.length > 15);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timeZone = ss.getSpreadsheetTimeZone() || "Europe/Moscow";
    const warehouseState = getWarehouseGroupedData_(ss);

    // Генерируем PDF полного прайса 1 раз перед отправкой
    let fullPricePdfBlob = null;
    if (attachFullPdf && !warehouseState.error) {
      try {
        fullPricePdfBlob = buildFullPricePdfBlob_(ss, warehouseState);
      } catch (pdfErr) {}
    }

    const aliases = GmailApp.getAliases();
    const senderAlias = aliases.length > 0 ? aliases[0] : null;
    const userSignature = getSenderSignature();

    const groupedBodySections = groupItemsByConditionOrder_(items);

    const sentDiams = Array.from(new Set(items.map(it => safeParseFloat(it.diam)).filter(d => d > 0)));
    const sentDiamsBrief = (!isFullPrice && sentDiams.length > 0) ? ` (Ø ${sentDiams.slice(0, 5).join(', ')})` : '';

    let sentCount = 0;
    const errors = [];

    for (let cIdx = 0; cIdx < clients.length; cIdx++) {
      const client = clients[cIdx];
      const recipientEmails = String(client.emailStr || client.email || '').trim();

      if (!recipientEmails || !recipientEmails.includes('@')) {
        continue;
      }

      const isManual = !client.sheet;
      const clientCompany = client.company && client.company !== 'Партнёр' ? client.company : '';
      // Берем выбранное менеджером имя ЛПР (или первое имя из списка, чтобы никогда не склеивались 3 человека в одну строку!)
      const chosenPerson = (client.selectedLpr !== undefined && client.selectedLpr !== null)
        ? String(client.selectedLpr).trim()
        : String(client.lpr || '').split(/[\n\/;]+/)[0].trim();
      const clientLprGreeting = chosenPerson ? `, ${chosenPerson}` : '';

      const subject = buildDynamicSubject(items, clientCompany, chosenPerson);

      let plain = [];
      plain.push(`Здравствуйте${clientLprGreeting}!\r\n`);
      plain.push(clientCompany
        ? `Направляем актуальное предложение по наличию трубного металлопроката со склада для компании «${clientCompany}»:\r\n`
        : `Направляем актуальное предложение по наличию трубного металлопроката со склада:\r\n`
      );

      if (customComment) {
        plain.push(`Примечание / условия поставки:\r\n${customComment}\r\n`);
      }

      groupedBodySections.forEach(sec => {
        plain.push(`\r\n🔹 ${sec.groupTitle.toUpperCase()}:`);
        sec.items.forEach((it, i) => {
          const gostVal = it.gost || it.std || '—';
          const details = [it.category, it.condition, (gostVal !== '—' && gostVal !== '-') ? 'ГОСТ: ' + gostVal : ''].filter(Boolean).join(', ');
          plain.push(`${i + 1}. ${it.name} (${details}) | Ø ${it.diam}x${it.wall} | Наличие: ${it.stock || it.stockText || '—'} | Цена: ${it.priceTn} (${it.priceM}/м)`);
        });
      });

      if (fullPricePdfBlob) {
        plain.push("\r\n📎 Полный прайс-лист всех позиций склада прикреплен к письму в формате PDF.");
      }

      plain.push("\r\n────────────────────────────────────────");
      plain.push("• Склад отгрузки: г. Воронеж (открытый транспорт).");
      plain.push("• Доставка: прямым и попутным транспортом по РФ.");
      plain.push("• Цены указаны с учетом НДС.");
      plain.push("────────────────────────────────────────\r\n");
      plain.push("Кирилл Манюхин");
      plain.push("Менеджер по продажам и развитию в маркетинге");
      plain.push("Моб / TG / WA: +7 (906) 588-37-77 | Тел. офис: +7 (473) 279-76-85");
      plain.push("Email: mail@kapital-p.ru | Сайт: capital-two-pi.vercel.app");
      plain.push("ООО «Капитал-Плюс» | ИНН 3663070544 | ОГРН 1083668004914 | НДС 20% | г. Воронеж, ул. Волгоградская, д. 48, оф. 76");

      const plainBody = plain.join("\r\n");

      // Формируем HTML-таблицу с четким разделением на группы (Трубы новые / Трубы лежалые / Трубы восстановленные / Трубы б/у)
      let rowsHtml = "";
      groupedBodySections.forEach(sec => {
        rowsHtml += `
          <tr style="background-color: #e2e8f0; border-bottom: 2px solid #cbd5e1; border-top: 1px solid #cbd5e1;">
            <td colspan="4" style="padding: 10px 14px; font-size: 12px; font-weight: 800; color: #0f172a; text-transform: uppercase; letter-spacing: 0.6px;">
              ▪ ${sec.groupTitle}
            </td>
          </tr>
        `;

        sec.items.forEach((it, idx) => {
          const rowBg = idx % 2 === 0 ? "#ffffff" : "#f8fafc";
          let rawGost = (it.gost || it.std || '').toString().trim();
          let formattedGost = '';
          if (rawGost && rawGost !== '—' && rawGost !== '-') {
            const lower = rawGost.toLowerCase();
            if (lower.startsWith('гост') || lower.startsWith('ту')) formattedGost = rawGost;
            else if (lower.includes('нефт')) formattedGost = 'из-под нефти';
            else if (lower.includes('газ')) formattedGost = 'из-под газа';
            else if (lower.includes('вод')) formattedGost = 'из-под воды';
            else if (lower.includes('пар')) formattedGost = 'из-под пара';
            else formattedGost = rawGost;
          }

          const techInfo = [it.category, it.condition, formattedGost].filter(Boolean).join(' • ');
          const techLine = techInfo ? '<div style="font-size: 11px; color: #64748b; margin-top: 2px;">' + techInfo + '</div>' : '';

          rowsHtml += `
            <tr style="background-color: ${rowBg}; border-bottom: 1px solid #e2e8f0;">
              <td style="padding: 10px 14px; color: #0f172a; font-size: 13px; text-align: left; vertical-align: middle;">
                <div style="font-weight: 700; color: #0f172a;">${it.name}</div>
                ${techLine}
              </td>
              <td style="padding: 10px 10px; text-align: center; color: #1e293b; font-size: 13px; font-weight: 600; white-space: nowrap; vertical-align: middle;">
                &empty;&nbsp;${it.diam}&nbsp;&times;&nbsp;${it.wall}
              </td>
              <td style="padding: 10px 10px; text-align: center; color: #334155; font-size: 12px; font-weight: 600; white-space: nowrap; vertical-align: middle;">
                ${it.stock || it.stockText || ''}
              </td>
              <td style="padding: 10px 14px; text-align: right; white-space: nowrap; vertical-align: middle;">
                <div style="font-size: 15px; font-weight: 800; color: #0f172a; line-height: 1.2;">
                  ${it.priceTn}
                </div>
                <div style="font-size: 11px; color: #64748b; font-weight: 500; margin-top: 2px;">
                  ${it.priceM}/м
                </div>
              </td>
            </tr>
          `;
        });
      });

      const commentHtml = customComment ? `
        <div style="margin: 16px 0; padding: 12px 16px; background-color: #f8fafc; border-left: 3px solid #2563eb; border-radius: 4px; font-size: 13px; color: #334155; line-height: 1.5;">
          <strong style="color: #0f172a; text-transform: uppercase; font-size: 11px; letter-spacing: 0.5px; display: block; margin-bottom: 4px;">Условия поставки / Примечание:</strong>
          ${customComment.replace(/\n/g, '<br>')}
        </div>
      ` : "";

      const pdfNoticeHtml = fullPricePdfBlob ? `
        <div style="margin-top: 14px; padding: 10px 14px; background-color: #eff6ff; border: 1px solid #bfdbfe; border-radius: 6px; font-size: 12px; color: #1e40af; font-weight: 600;">
          📎 К письму прикреплен полный официальный прайс-лист всех позиций склада в формате PDF.
        </div>
      ` : "";

      const introText = clientCompany
        ? `Направляем актуальное коммерческое предложение по наличию трубного металлопроката со склада для компании «${clientCompany}»:`
        : `Направляем актуальное коммерческое предложение по наличию трубного металлопроката с нашей базы:`;

      const htmlBody = `
        <!DOCTYPE html>
        <html>
        <head>
          <meta charset="utf-8">
          <meta name="viewport" content="width=device-width, initial-scale=1.0">
        </head>
        <body style="margin: 0; padding: 24px 12px; background-color: #f1f5f9; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;">
          <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 720px; background-color: #ffffff; border: 1px solid #cbd5e1; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);">
            <tr>
              <td style="background-color: #0f172a; padding: 22px 28px; color: #ffffff;">
                <table width="100%" border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td>
                      <div style="font-size: 22px; font-weight: 900; letter-spacing: -0.5px; color: #ffffff; line-height: 1;">ООО «КАПИТАЛ-ПЛЮС»</div>
                      <div style="font-size: 12px; color: #94a3b8; margin-top: 5px; letter-spacing: 0.3px;">Поставки трубного металлопроката со склада • г. Воронеж</div>
                    </td>
                    <td align="right" style="vertical-align: middle;">
                      <span style="background-color: #1e293b; border: 1px solid #334155; padding: 5px 10px; border-radius: 4px; font-size: 11px; color: #cbd5e1; white-space: nowrap; font-weight: 600;">
                        База: г. Воронеж
                      </span>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
            <tr>
              <td style="padding: 24px 28px;">
                <div style="font-size: 15px; font-weight: 700; color: #0f172a; margin-bottom: 8px;">
                  Здравствуйте${clientLprGreeting}!
                </div>
                <div style="font-size: 13px; color: #475569; line-height: 1.5; margin-bottom: 14px;">
                  ${introText}
                </div>
                ${commentHtml}
                <table width="100%" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; margin-top: 14px; border: 1px solid #e2e8f0; border-radius: 6px; overflow: hidden;">
                  <thead>
                    <tr style="background-color: #0f172a; color: #ffffff; text-transform: uppercase; font-size: 11px; letter-spacing: 0.5px;">
                      <th align="left" style="padding: 10px 14px; font-weight: 700;">Наименование</th>
                      <th align="center" style="padding: 10px 10px; font-weight: 700; width: 17%;">Размер</th>
                      <th align="center" style="padding: 10px 10px; font-weight: 700; width: 16%;">Остаток</th>
                      <th align="right" style="padding: 10px 14px; font-weight: 700; width: 22%;">Цена с НДС</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${rowsHtml}
                  </tbody>
                </table>
                ${pdfNoticeHtml}
                <table width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-top: 18px; padding: 12px 16px; background-color: #f8fafc; border-radius: 6px; border: 1px solid #e2e8f0; font-size: 12px; color: #475569; line-height: 1.6;">
                  <tr>
                    <td>
                      &bull; <strong>База отгрузки:</strong> г. Воронеж (погрузка в открытый автотранспорт).<br>
                      &bull; <strong>Логистика:</strong> Помощь с организацией попутной и прямой доставки по РФ.<br>
                      &bull; <strong>Расчёт:</strong> Все цены указаны с учётом НДС. Выставим счёт за 10 минут.
                    </td>
                  </tr>
                </table>

                ${userSignature.html}
              </td>
            </tr>
          </table>
        </body>
        </html>
      `;

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

        // ФИКСАЦИЯ В ТАБЛИЦЕ (КОЛОНКА N / 14 - ИСТОРИЯ) СО СЛЕПКОМ H+I+J И ВЫБРАННЫМ ИМЕНЕМ ЛПР!
        if (!isManual && client.sheet && client.row && client.row >= 2) {
          const targetSheet = ss.getSheetByName(client.sheet);
          if (targetSheet) {
            const nowStamp = Utilities.formatDate(new Date(), timeZone, "dd.MM.yyyy HH:mm");
            const hijVals = targetSheet.getRange(client.row, 8, 1, 3).getValues()[0];
            const curNeed = String(hijVals[0] || '').trim();
            const curType = String(hijVals[1] || '').trim();
            const curApp  = String(hijVals[2] || '').trim();
            const snap    = typeof buildNeedSnapshot_ === 'function' ? buildNeedSnapshot_(curNeed, curType, curApp) : '';
            const lprTag  = chosenPerson ? ` (для: ${chosenPerson})` : '';
            const pdfTag  = fullPricePdfBlob ? ' + 📎PDF' : '';

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