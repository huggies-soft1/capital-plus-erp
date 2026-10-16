/**
 * ЕДИНЫЙ ЦЕНТР УПРАВЛЕНИЯ ПРОДАЖАМИ И СКЛАДОМ
 * Файл: PriceList.gs — Генератор прайс-листа реального наличия труб и отправка КП
 */

function openPriceListModal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const activeSheetName = activeSheet.getName().trim();
  const activeRange = activeSheet.getActiveRange();

  const clientSheets = ["1. Обзвон (База)", "2. Золотой список (Киты)", "3. Лист ожидания"];
  const allClientsList = [];
  let prefillClientKey = "";

  // 1. Сбор контактов со всех листов базы
  clientSheets.forEach(sName => {
    const sh = ss.getSheetByName(sName);
    if (!sh) return;
    const lRow = sh.getLastRow();
    const lCol = sh.getLastColumn();
    if (lRow < 2 || lCol < 2) return;

    const headers = sh.getRange(1, 1, 1, lCol).getValues()[0].map(h => String(h || '').trim().toLowerCase());
    const cInn   = headers.findIndex(h => h.includes('инн'));
    const cComp  = headers.findIndex(h => h.includes('компан') || h.includes('клиент'));
    const cLpr   = headers.findIndex(h => h.includes('лпр') || h.includes('контакт'));
    const cEmail = headers.findIndex(h => h.includes('mail') || h.includes('почт') || h.includes('мейл'));

    const data = sh.getRange(2, 1, lRow - 1, lCol).getValues();
    data.forEach((r, idx) => {
      const rowNum = idx + 2;
      const comp = String((cComp !== -1 ? r[cComp] : r[1]) || '').trim();
      const inn = String((cInn !== -1 ? r[cInn] : r[0]) || '').trim();
      const lpr = String((cLpr !== -1 ? r[cLpr] : r[3]) || '').trim();
      const rawEm = String((cEmail !== -1 ? r[cEmail] : r[5]) || '').trim();

      const emList = rawEm.split(/[,;\s\n]+/).map(x => x.trim()).filter(x => x && x.includes('@'));

      if (comp || emList.length > 0) {
        const clientKey = sName + '_' + rowNum;
        allClientsList.push({
          key: clientKey,
          sheet: sName,
          row: rowNum,
          inn: inn,
          company: comp || 'Без названия',
          lpr: lpr,
          emailStr: emList.join(', ')
        });

        if (activeRange && activeSheetName === sName && activeRange.getRow() === rowNum) {
          prefillClientKey = clientKey;
        }
      }
    });
  });

  allClientsList.sort((a, b) => (a.company || '').localeCompare(b.company || ''));

  // 2. Считывание остатков склада (14 или 15 колонок, Закупка скрыта от клиентов!)
  const sheet = ss.getSheetByName('Склад');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('⚠️ Ошибка: лист «Склад» не найден!');
    return;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(14, sheet.getLastColumn());
  if (lastRow < 3) {
    SpreadsheetApp.getUi().alert('⚠️ На листе «Склад» нет данных!');
    return;
  }

  const headers = sheet.getRange(2, 1, 1, lastCol).getValues()[0].map(h => String(h).trim().toLowerCase());
  const colCategory  = headers.indexOf('категория') !== -1 ? headers.indexOf('категория') : 0;
  const colName      = headers.indexOf('позиция') !== -1 ? headers.indexOf('позиция') : 1;
  const colCondition = headers.indexOf('состояние') !== -1 ? headers.indexOf('состояние') : 2;
  const colDiam      = headers.findIndex(h => h.includes('ø') || h.includes('диаметр'));
  const colWall      = headers.findIndex(h => h.includes('стенка'));
  const colGost      = headers.findIndex(h => h.includes('гост') || h.includes('сырье'));

  // Ищем строго ПРОДАЖУ (чтобы колонка Закупки никогда не попала в прайс)
  let colPriceTn = headers.findIndex(h => h.includes('продаж') && h.includes('тн'));
  if (colPriceTn === -1) colPriceTn = headers.findIndex(h => h.includes('цена') && h.includes('тн'));

  let colPriceM = headers.findIndex(h => h.includes('продаж') && h.includes('м'));
  if (colPriceM === -1) colPriceM = headers.findIndex(h => h.includes('цена') && h.includes('м'));

  const colStatus  = headers.findIndex(h => h.includes('наличи')) !== -1 ? headers.findIndex(h => h.includes('наличи')) : 8;
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

    // Отсекаем позиции со статусом «Нет» И позиции-призраки без фактического остатка
    if (rawStatus === 'нет') return;
    if (stockTn <= 0 && stockM <= 0) return;

    const rawCategory = (row[colCategory] || 'ЭСВ').toString().trim();
    const condition = (row[colCondition] || 'Б/у').toString().trim();
    const condLower = condition.toLowerCase();

    let groupTitle = "Трубы б/у";
    if (condLower.includes('восстанов') || condLower.includes('восст')) {
      groupTitle = "Трубы восстановленные";
    } else if (condLower.includes('лежал')) {
      groupTitle = "Трубы лежалые";
    } else if (condLower.includes('нов')) {
      groupTitle = "Трубы новые";
    }

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
    if (stockTnStr && stockMStr) {
      stockText = stockTnStr + " / " + stockMStr;
    } else if (stockTnStr) {
      stockText = stockTnStr;
    } else if (stockMStr) {
      stockText = stockMStr;
    }

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
      stockTnStr: stockTnStr,
      stockMStr: stockMStr,
      stockText: stockText,
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
    let items = rawItems.filter(x => x.group === grp);
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

  const template = HtmlService.createTemplateFromFile('PriceListModal');
  template.groupedData = groupedData;
  template.generatedDate = dateStr;
  template.prefillClientKey = prefillClientKey;
  template.clientsJson = JSON.stringify(allClientsList);

  const html = template.evaluate()
    .setWidth(1180)
    .setHeight(860)
    .setTitle('Прайс-лист • КАПИТАЛ-ПЛЮС');

  SpreadsheetApp.getUi().showModelessDialog(html, 'Прайс-лист и КП • КАПИТАЛ-ПЛЮС');
}

/**
 * Получение подписи пользователя: из настроек Gmail API или резервный шаблон
 */
function getSenderSignature() {
  // 1. Попытка получить живую подпись из настроек Gmail
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
  } catch (e) {
    // Gmail API не включен в службах, переходим к фирменному шаблону
  }

  // 2. Ваша точная фирменная подпись со всеми контактами
  const fallbackHtml = `
    <div style="margin-top: 24px; padding-top: 16px; border-top: 2px solid #e2e8f0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; font-size: 13px; color: #334155; line-height: 1.5;">
      <div style="font-weight: 800; font-size: 16px; color: #0f172a; letter-spacing: -0.3px;">ООО Капитал-Плюс</div>
      <div style="color: #64748b; font-size: 12px; margin-bottom: 8px;">Отдел продаж и развития</div>
      
      <div style="font-size: 12px; color: #1e293b; margin-bottom: 4px;">
        <strong>Моб / TG / WA:</strong> <a href="tel:+79065883777" style="color: #2563eb; text-decoration: none; font-weight: 600;">+7 (906) 588-37-77</a> &nbsp;|&nbsp; 
        <strong>Тел. офис:</strong> +7 (473) 279-76-85
      </div>
      <div style="font-size: 12px; color: #1e293b; margin-bottom: 8px;">
        <strong>Email:</strong> <a href="mailto:mail@kapital-p.ru" style="color: #2563eb; text-decoration: none;">mail@kapital-p.ru</a> &nbsp;|&nbsp; 
        <strong>Сайт:</strong> <a href="https://capital-two-pi.vercel.app" style="color: #2563eb; text-decoration: none;" target="_blank">capital-two-pi.vercel.app</a>
      </div>
      
      <div style="font-size: 11px; color: #64748b; line-height: 1.4; margin-bottom: 8px;">
        ООО «Капитал-Плюс» | ИНН 3663070544 | ОГРН 1083668004914 | НДС 22% | г. Воронеж, ул. Волгоградская, д. 48, оф. 76
      </div>

      <div style="font-size: 10px; color: #94a3b8; line-height: 1.3; font-style: italic;">
        Конфиденциальность: Данное электронное сообщение предназначено исключительно для лица или организации, которым оно адресовано.
      </div>
    </div>
  `;

  return { html: fallbackHtml, isLive: false };
}

/**
 * Вариант 4: Углубленная динамическая генерация тем по выбранным позициям
 */
function buildDynamicSubject(items, clientCompany, clientLpr) {
  const companyTag = clientCompany ? ` • ${clientCompany}` : '';

  let lprGreeting = "";
  if (clientLpr) {
    const lprClean = clientLpr.trim();
    const lprLower = lprClean.toLowerCase();
    const isService = lprLower.includes('приемная') || 
                      lprLower.includes('снабжен') || 
                      lprLower.includes('бухгалтер') || 
                      lprLower.includes('склад') ||
                      lprLower.includes('отдел');

    if (!isService) {
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
 * Пакетная отправка коммерческих предложений с защитой от спама
 */
function sendBatchProposals(data) {
  try {
    let clients = data.clients || [];
    const manualEmail = String(data.manualEmail || '').trim();

    if (clients.length === 0 && manualEmail) {
      const manualList = manualEmail.split(/[,;\s]+/).map(e => e.trim()).filter(e => e && e.includes('@'));
      if (manualList.length > 0) {
        clients = [{
          company: 'Партнёр', lpr: '', emailStr: manualList.join(', '), sheet: '', row: 0
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
    const isFullPrice = (items.length > 15);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timeZone = ss.getSpreadsheetTimeZone() || "Europe/Moscow";

    const aliases = GmailApp.getAliases();
    const senderAlias = aliases.length > 0 ? aliases[0] : null;

    // Подтягиваем подпись (из Gmail API или красивый встроенный шаблон)
    const userSignature = getSenderSignature();

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
      const clientLpr = client.lpr ? `, ${client.lpr}` : '';

      const subject = buildDynamicSubject(items, clientCompany, client.lpr);

      let plain = [];
      plain.push(`Здравствуйте${clientLpr}!\r\n`);
      plain.push(clientCompany
        ? `Направляем актуальное предложение по наличию трубного металлопроката со склада для компании «${clientCompany}»:\r\n`
        : `Направляем актуальное предложение по наличию трубного металлопроката со склада:\r\n`
      );

      if (customComment) {
        plain.push(`Примечание / условия поставки:\r\n${customComment}\r\n`);
      }

      plain.push("СПИСОК ПОЗИЦИЙ В НАЛИЧИИ:");
      items.forEach((it, i) => {
        const gostVal = it.gost || it.std || '—';
        const details = [it.category, it.condition, (gostVal !== '—' && gostVal !== '-') ? 'ГОСТ: ' + gostVal : ''].filter(Boolean).join(', ');

        plain.push(`${i + 1}. ${it.name} (${details}) | Ø ${it.diam}x${it.wall} | Наличие: ${it.stock || it.stockText || '—'} | Цена: ${it.priceTn} (${it.priceM}/м)`);
      });

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

      let rowsHtml = "";
      items.forEach((it, idx) => {
        const rowBg = idx % 2 === 0 ? "#ffffff" : "#f8fafc";
        let rawGost = (it.gost || it.std || '').toString().trim();
        let formattedGost = '';
        if (rawGost && rawGost !== '—' && rawGost !== '-') {
          const lower = rawGost.toLowerCase();
          if (lower.startsWith('гост') || lower.startsWith('ту')) {
            formattedGost = rawGost;
          } else if (lower.includes('нефт')) {
            formattedGost = 'из-под нефти';
          } else if (lower.includes('газ')) {
            formattedGost = 'из-под газа';
          } else if (lower.includes('вод')) {
            formattedGost = 'из-под воды';
          } else if (lower.includes('пар')) {
            formattedGost = 'из-под пара';
          } else {
            formattedGost = rawGost;
          }
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

      const commentHtml = customComment ? `
        <div style="margin: 16px 0; padding: 12px 16px; background-color: #f8fafc; border-left: 3px solid #475569; border-radius: 4px; font-size: 13px; color: #334155; line-height: 1.5;">
          <strong style="color: #0f172a; text-transform: uppercase; font-size: 11px; letter-spacing: 0.5px; display: block; margin-bottom: 4px;">Условия поставки / Примечание:</strong>
          ${customComment.replace(/\n/g, '<br>')}
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
                  Здравствуйте${clientLpr}!
                </div>
                <div style="font-size: 13px; color: #475569; line-height: 1.5; margin-bottom: 14px;">
                  ${introText}
                </div>
                ${commentHtml}
                <table width="100%" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse; margin-top: 14px; border: 1px solid #e2e8f0; border-radius: 6px; overflow: hidden;">
                  <thead>
                    <tr style="background-color: #f8fafc; border-bottom: 2px solid #cbd5e1; text-transform: uppercase; font-size: 11px; color: #475569; letter-spacing: 0.5px;">
                      <th align="left" style="padding: 10px 14px; font-weight: 700;">Наименование</th>
                      <th align="center" style="padding: 10px 10px; font-weight: 700; width: 17%;">Размер</th>
                      <th align="center" style="padding: 10px 10px; font-weight: 700; width: 16%;">Остаток</th>
                      <th align="right" style="padding: 10px 14px; font-weight: 700; width: 22%;">Цена</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${rowsHtml}
                  </tbody>
                </table>
                <table width="100%" border="0" cellpadding="0" cellspacing="0" style="margin-top: 18px; padding: 12px 16px; background-color: #f8fafc; border-radius: 6px; border: 1px solid #e2e8f0; font-size: 12px; color: #475569; line-height: 1.6;">
                  <tr>
                    <td>
                      &bull; <strong>База отгрузки:</strong> г. Воронеж (погрузка в открытый автотранспорт).<br>
                      &bull; <strong>Логистика:</strong> Помощь с организацией попутной и прямой доставки по РФ.<br>
                      &bull; <strong>Расчёт:</strong> Все цены указаны с учётом НДС. Выставим счёт за 10 минут.
                    </td>
                  </tr>
                </table>

                <!-- ВСТАВКА ВАШЕЙ ПОДПИСИ -->
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
          const targetSheet = ss.getSheetByName(client.sheet);
          if (targetSheet) {
            const nowStamp = Utilities.formatDate(new Date(), timeZone, "dd.MM.yyyy HH:mm");

            let actionType = "📧 Отправлено";
            if (isDraftOnly) actionType = "📝 Черновик КП";
            else if (isFullPrice) actionType = "📧 Прайс: Полный";
            else actionType = "💰 Выставлен СЧЕТ/КП";

            const tag = `[${actionType} - ${nowStamp}]`;

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