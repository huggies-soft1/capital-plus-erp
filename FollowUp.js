/**
 * ЕДИНЫЙ ЦЕНТР УПРАВЛЕНИЯ ПРОДАЖАМИ И СКЛАДОМ (CTRM «Капитал-Плюс» v7.1)
 * Файл: FollowUp.js — Самостоятельный модуль «Аналитика отношений, Торг 1С и Рассылка Напоминалок»
 * Строгий B2B-стиль без эмодзи (100% защита от битых символов при отправке через алиас).
 */

function stripAllEmojisClean_(str) {
  if (!str) return '';
  return String(str)
    .replace(/[\uD800-\uDBFF][\uDC00-\uDFFF]/g, '')
    .replace(/[\u2600-\u27BF\uFE00-\uFE0F\u200D\u2300-\u23FF\u2B50\u2B55\u2934\u2935\u25AA-\u25FE]/g, '')
    .replace(/[ \t]{2,}/g, ' ')
    .trim();
}

function openFollowUpModal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const state = getFollowUpRelationshipData_(ss);

  const template = HtmlService.createTemplateFromFile('FollowUpModal');
  template.clientsJson = JSON.stringify(state.clients);
  template.summaryJson = JSON.stringify(state.summary);
  template.prefillClientKey = state.prefillClientKey;
  template.generatedDate = state.generatedDate;

  const html = template.evaluate()
    .setWidth(1180)
    .setHeight(860)
    .setTitle('Аналитика отношений и Напоминалка • КАПИТАЛ-ПЛЮС');

  SpreadsheetApp.getUi().showModelessDialog(html, 'Аналитика отношений и Напоминалка • КАПИТАЛ-ПЛЮС');
}

function getFollowUpRelationshipLive() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return getFollowUpRelationshipData_(ss);
}

function getFollowUpRelationshipData_(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const now = new Date();
  const activeSheet = ss.getActiveSheet();
  const activeSheetName = activeSheet ? activeSheet.getName().trim() : '';
  const activeRange = activeSheet ? activeSheet.getActiveRange() : null;

  const months = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];
  const generatedDate = `${now.getDate()} ${months[now.getMonth()]} ${now.getFullYear()} г.`;

  const clientSheets = [
    '2. Золотой список (Киты)',
    '3. Лист ожидания',
    '1. Обзвон (База)'
  ];

  const clients = [];
  let prefillClientKey = '';

  let total1CTons = 0;
  let total1CRevenue = 0;
  let needReminderCount = 0;
  let whalesCount = 0;

  clientSheets.forEach(sName => {
    const sh = typeof getSheetByNameSafe === 'function' ? getSheetByNameSafe(ss, sName) : ss.getSheetByName(sName);
    if (!sh) return;
    const lRow = sh.getLastRow();
    const lCol = sh.getLastColumn();
    if (lRow < 2 || lCol < 2) return;

    const numCols = Math.min(19, lCol);
    const data = sh.getRange(2, 1, lRow - 1, numCols).getValues();
    const notesB = sh.getRange(2, 2, lRow - 1, 1).getNotes();

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
      const notesK  = String(r[10] || '').trim();
      const dlVal   = r[11];
      const history = String(r[13] || '').trim();
      const tactic  = numCols >= 15 ? String(r[14] || '').trim() : '';

      const segP    = numCols >= 16 ? String(r[15] || '').trim() : '';
      const tonsQ   = numCols >= 17 ? safeParseFloat(r[16]) : 0;
      const revR    = numCols >= 18 ? safeParseFloat(r[17]) : 0;
      const s1cInfo = numCols >= 19 ? String(r[18] || '').trim() : '';

      if (!comp && !inn && !phRaw && !rawEm) return;

      const dossierB = String((notesB[idx] && notesB[idx][0]) || '').trim();
      const emList = Array.from(new Set(
        rawEm.split(/[,;\s\n]+/).map(x => x.trim()).filter(x => x && x.includes('@'))
      ));
      const lprList = typeof parseLprList_ === 'function' ? parseLprList_(lprRaw) : (lprRaw ? [lprRaw] : []);
      const cityClean = typeof extractShortCity_ === 'function' ? extractShortCity_(cityRaw) : cityRaw;

      const stats1C = parseFollowUp1CStats_(segP, tonsQ, revR, s1cInfo, now);
      const statsKp = parseFollowUpHistoryStats_(history, now);

      if (stats1C.has1C) {
        total1CTons += stats1C.tons;
        total1CRevenue += stats1C.revenue;
        whalesCount++;
      }

      let relationStatus = 'Новый контакт';
      let needsReminder = false;
      if (statsKp.kpCount > 0 && statsKp.daysSinceKp !== null) {
        if (statsKp.daysSinceKp >= 14) {
          relationStatus = `Ждёт дожима (${statsKp.daysSinceKp} дн. после КП)`;
          needsReminder = true;
        } else if (statsKp.daysSinceKp >= 4) {
          relationStatus = `Пора напомнить (${statsKp.daysSinceKp} дн. после КП)`;
          needsReminder = true;
        } else {
          relationStatus = `Свежее КП (${statsKp.daysSinceKp} дн. назад)`;
        }
      } else if (stats1C.has1C) {
        if (stats1C.daysSinceOrder !== null && stats1C.daysSinceOrder > 180) {
          relationStatus = `Спящий Кит (посл. УПД ${stats1C.lastOrderDate})`;
          needsReminder = true;
        } else {
          relationStatus = `Активный партнёр (${stats1C.tonsFormatted})`;
        }
      }
      if (needsReminder) needReminderCount++;

      let deadlineStr = '';
      if (dlVal instanceof Date && !isNaN(dlVal.getTime())) {
        deadlineStr = Utilities.formatDate(dlVal, ss.getSpreadsheetTimeZone() || 'Europe/Moscow', 'dd.MM.yyyy');
      } else if (dlVal) {
        deadlineStr = String(dlVal).trim();
      }

      const cleanNeed = (need && need.toLowerCase() !== 'нет' && need !== '—' && need !== '-') ? need : '';
      const cleanDossier = [notesK, dossierB !== notesK ? dossierB : ''].filter(Boolean).join(' | ');

      const clientKey = `${sName}_${rowNum}`;
      clients.push({
        key: clientKey,
        sheet: sName,
        row: rowNum,
        inn: inn,
        company: comp || 'Без названия',
        city: cityClean,
        lpr: lprRaw,
        lprList: lprList,
        phone: phRaw.split('\n')[0] || '',
        emailStr: emList.join(', '),
        emailList: emList,
        sphere: [sphere, pApp, pType].filter(Boolean).join(' • '),
        need: cleanNeed,
        dossier: cleanDossier,
        deadline: deadlineStr,
        tactic: stripAllEmojisClean_(tactic),
        relationStatus: relationStatus,
        needsReminder: needsReminder,
        stats1C: stats1C,
        statsKp: statsKp
      });

      const actualSheetName = sh.getName().trim();
      if (activeRange && (activeSheetName === sName || activeSheetName === actualSheetName) && activeRange.getRow() === rowNum) {
        prefillClientKey = clientKey;
      }
    });
  });

  return {
    generatedDate: generatedDate,
    prefillClientKey: prefillClientKey,
    clients: clients,
    summary: {
      totalClients: clients.length,
      whalesCount: whalesCount,
      needReminderCount: needReminderCount,
      total1CTons: Math.round(total1CTons),
      total1CRevenueMln: (total1CRevenue / 1000000).toFixed(1)
    }
  };
}

function parseFollowUp1CStats_(segP, tonsQ, revR, s1cInfo, now) {
  const tons = Number((tonsQ || 0).toFixed(2));
  const rev = Math.round(revR || 0);
  const avgPrice = (tons > 0 && rev > 0) ? Math.round(rev / tons) : 0;

  let updCount = 0;
  let yearsStr = '';
  let lastOrderDate = '';
  let daysSinceOrder = null;
  let pipesPurchased = '';

  const s = String(s1cInfo || '').trim();
  if (s) {
    const updMatch = s.match(/(?:УПД:\s*(\d+)|(\d+)\s*УПД)/i);
    if (updMatch) updCount = parseInt(updMatch[1] || updMatch[2], 10) || 0;

    const lastMatch = s.match(/посл\.\s*(\d{2})\.(\d{2})\.(\d{4})/i);
    if (lastMatch) {
      lastOrderDate = `${lastMatch[1]}.${lastMatch[2]}.${lastMatch[3]}`;
      const dt = new Date(parseInt(lastMatch[3], 10), parseInt(lastMatch[2], 10) - 1, parseInt(lastMatch[1], 10));
      if (!isNaN(dt.getTime())) {
        daysSinceOrder = Math.max(0, Math.floor((now.getTime() - dt.getTime()) / 86400000));
      }
    }

    const yearsParenMatch = s.match(/\(\s*(\d{4}(?:\s*[–\-]\s*\d{4})?)\s*,/);
    const yearsBracketMatch = s.match(/\[[^\]]+\]\s*([^|]+?)\s*\|\s*УПД:/i);
    if (yearsParenMatch) {
      const rawY = yearsParenMatch[1].trim();
      yearsStr = /[–\-]/.test(rawY) ? `${rawY} гг.` : `${rawY} г.`;
    } else if (yearsBracketMatch) {
      yearsStr = yearsBracketMatch[1].trim();
    }

    const colonSplit = s.match(/^[^:]+?\([^)]*УПД[^)]*\)\s*:\s*(.+)$/i);
    if (colonSplit && colonSplit[1]) {
      pipesPurchased = colonSplit[1].trim();
    } else {
      const parts = s.split('|').map(p => p.trim());
      if (parts.length >= 3) {
        pipesPurchased = parts.slice(2).join(' | ').trim();
      } else {
        pipesPurchased = s.replace(/^\[[^\]]+\]\s*/, '').trim();
      }
    }
    pipesPurchased = String(pipesPurchased || '')
      .replace(/^(?:КП\+РЛ|КП|РЛ|1С)\s*(?:\([^)]*\))?\s*:?\s*/i, '')
      .trim();
  }

  let partnerStatus = 'Действующий контрагент';
  if (tons >= 50 || segP.includes('VIP') || segP.includes('A') || segP.includes('Кит') || segP.includes('Крупный')) {
    partnerStatus = 'Приоритетный партнёр';
  } else if (tons >= 15 || segP.includes('B') || segP.includes('Постоянный') || segP.includes('Средний')) {
    partnerStatus = 'Постоянный партнёр';
  } else if (tons > 0) {
    partnerStatus = 'Партнёр по поставкам';
  }

  return {
    has1C: (tons > 0 || rev > 0 || Boolean(s)),
    partnerStatus: partnerStatus,
    tons: tons,
    tonsFormatted: tons > 0 ? `${tons.toString().replace('.', ',')} тн` : '',
    revenue: rev,
    revenueFormatted: rev > 0 ? rev.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ' ') + ' ₽' : '',
    avgPrice: avgPrice,
    avgPriceFormatted: avgPrice > 0 ? avgPrice.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ' ') + ' ₽/тн' : '',
    updCount: updCount,
    years: yearsStr || (lastOrderDate ? lastOrderDate.slice(-4) + ' г.' : ''),
    lastOrderDate: lastOrderDate,
    daysSinceOrder: daysSinceOrder,
    pipesPurchased: pipesPurchased
  };
}

function parseFollowUpHistoryStats_(historyRaw, now) {
  const raw = String(historyRaw || '').trim();
  if (!raw) {
    return {
      kpCount: 0,
      totalTouches: 0,
      lastKpDate: '',
      daysSinceKp: null,
      quotedPrices: [],
      quotedSummary: '',
      lastHistoryLine: ''
    };
  }

  const lines = raw.split('\n').map(l => l.trim()).filter(Boolean);
  let kpCount = 0;
  let lastKpDate = '';
  let daysSinceKp = null;
  const quotedPrices = [];

  lines.forEach(line => {
    if (!/Черновик/i.test(line) && /^\[(?:📧|💬|📄|📋)?\s*(?:Прайс|КП\b|Выгружен PDF|Напоминалка|Напоминание|Дожим)/i.test(line)) {
      kpCount++;
      const dm = line.match(/-\s*(\d{2})\.(\d{2})\.(\d{4})(?:\s+\d{2}:\d{2})?\]$/) || line.match(/(\d{2})\.(\d{2})\.(\d{4})/);
      if (!lastKpDate && dm) {
        lastKpDate = `${dm[1]}.${dm[2]}.${dm[3]}`;
        const dt = new Date(parseInt(dm[3], 10), parseInt(dm[2], 10) - 1, parseInt(dm[1], 10));
        if (!isNaN(dt.getTime())) {
          daysSinceKp = Math.max(0, Math.floor((now.getTime() - dt.getTime()) / 86400000));
        }
      }
      const pm = line.match(/\d{2,4}[хx\*]\d+(?:[.,]\d+)?[^,;\]\)]*?\d[\d\s]*₽(?:\/тн|\/т)?/gi);
      if (pm) {
        pm.forEach(p => {
          const clean = p.trim();
          if (!quotedPrices.includes(clean)) quotedPrices.push(clean);
        });
      }
      if (quotedPrices.length === 0) {
        const dMatch = line.match(/Ø\s*([\d,\s]+)/i);
        if (dMatch) quotedPrices.push('Ø ' + dMatch[1].trim() + ' мм');
      }
    }
  });

  return {
    kpCount: kpCount,
    totalTouches: lines.length,
    lastKpDate: lastKpDate,
    daysSinceKp: daysSinceKp,
    quotedPrices: quotedPrices,
    quotedSummary: quotedPrices.slice(0, 4).join(', '),
    lastHistoryLine: stripAllEmojisClean_(lines[0] || '')
  };
}

function sendBatchFollowUpEmails(data) {
  try {
    let clients = data.clients || [];
    if (clients.length === 0) {
      return { success: false, message: 'Выберите хотя бы одного клиента из списка или нажмите «Тестовая фирма (на себя)»!' };
    }

    const scenarioKey = String(data.scenarioKey || 'remind_kp').trim();
    const customMessageText = stripAllEmojisClean_(data.customMessageText || '');
    const customOfferLine = stripAllEmojisClean_(data.customOfferLine || '');
    const isDraftOnly = (data.createDraftOnly === true);
    const includeRubles = (data.includeRubles !== false);

    if (!isDraftOnly) {
      try {
        const remainingQuota = MailApp.getRemainingDailyQuota();
        if (remainingQuota < clients.length) {
          return {
            success: false,
            message: `Недостаточно дневной квоты Gmail: осталось ${remainingQuota} писем, выбрано ${clients.length}. Уменьшите выборку или сохраните в черновики.`
          };
        }
      } catch (qErr) {}
    }
    const nextDays = Math.max(1, parseInt(data.nextDays, 10) || 3);
    const customStatsOverride = data.customStatsOverride || null;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timeZone = ss.getSpreadsheetTimeZone() || 'Europe/Moscow';

    const aliases = GmailApp.getAliases();
    const senderAlias = aliases.find(a => String(a).toLowerCase().includes('kapital-p.ru')) || (aliases.length > 0 ? aliases[0] : null);
    const userSignature = typeof getSenderSignature === 'function' ? getSenderSignature() : { html: '', isLive: false };

    let sentCount = 0;
    const errors = [];

    for (let cIdx = 0; cIdx < clients.length; cIdx++) {
      const client = clients[cIdx];
      const activeEmailsArr = Array.isArray(client.selectedEmails) && client.selectedEmails.length > 0
        ? client.selectedEmails
        : String(client.emailStr || client.email || '').split(/[,;\s\n]+/).map(x => x.trim()).filter(x => x && x.includes('@'));
      const recipientEmails = activeEmailsArr.join(', ');

      if (!recipientEmails || !recipientEmails.includes('@')) continue;

      const isManualOrTest = !client.sheet || client.key === '__TEST_COMPANY__' || Boolean(client.isTestSelf);
      const rawCompName = stripAllEmojisClean_(client.company || '');
      const clientCompany = rawCompName && rawCompName !== 'Партнёр' ? rawCompName.replace(/\s*\(Тест на себя\)$/i, '').trim() : '';
      const chosenPerson = stripAllEmojisClean_(typeof resolveChosenLprString_ === 'function' ? resolveChosenLprString_(client) : (client.selectedLpr || client.lpr || ''));
      const clientLprGreeting = chosenPerson ? `, ${chosenPerson}` : '';

      const s1c = (clients.length === 1 && customStatsOverride) ? customStatsOverride.stats1C : (client.stats1C || { has1C: false });
      const sKp = (clients.length === 1 && customStatsOverride) ? customStatsOverride.statsKp : (client.statsKp || { kpCount: 0 });
      const clientNeed = stripAllEmojisClean_((clients.length === 1 && customStatsOverride && customStatsOverride.need !== undefined) ? customStatsOverride.need : (client.need || ''));

      const subject = stripAllEmojisClean_(buildRelationshipSubject_(scenarioKey, clientCompany, chosenPerson, s1c, sKp, cIdx));

      const cleanCompanyDisplay = clientCompany
        ? clientCompany
            .replace(/^[«"”„']+|[»"”„']+$/g, '')
            .replace(/[«"”„']+/g, '')
            .replace(/^(ООО|АО|ЗАО|ПАО|ИП)\s+(.+)$/i, '$1 «$2»')
            .trim()
        : '';
      const companyFormattedPlain = cleanCompanyDisplay
        ? (cleanCompanyDisplay.includes('«') ? cleanCompanyDisplay : `«${cleanCompanyDisplay}»`)
        : '';

      const paragraphs = (clients.length === 1 && customMessageText)
        ? customMessageText.split(/\n+/).map(p => p.trim()).filter(Boolean)
        : buildDefaultScenarioParagraphs_(scenarioKey, companyFormattedPlain, s1c, sKp, clientNeed);

      const plain = [];
      plain.push(`Здравствуйте${clientLprGreeting}!\r\n`);
      paragraphs.forEach(p => plain.push(p + '\r\n'));

      if (s1c && s1c.has1C) {
        plain.push(`НАША ИСТОРИЯ СОТРУДНИЧЕСТВА (ООО «Капитал-Плюс»):`);
        plain.push(`• Статус: ${s1c.partnerStatus || 'Постоянный партнёр'}${s1c.years ? ' (' + s1c.years + ')' : ''}`);
        if (s1c.tonsFormatted) {
          plain.push(`• Суммарный объём отгрузок: ${s1c.tonsFormatted}${s1c.updCount ? ' (' + s1c.updCount + ' партий / УПД)' : ''}${s1c.lastOrderDate ? ', крайняя поставка: ' + s1c.lastOrderDate : ''}`);
        }
        if (includeRubles && s1c.revenueFormatted) {
          plain.push(`• Объём сделок с НДС: ${s1c.revenueFormatted}${s1c.avgPriceFormatted ? ' (средняя цена: ' + s1c.avgPriceFormatted + ')' : ''}`);
        }
        if (s1c.pipesPurchased) {
          plain.push(`• Поставлявшиеся позиции: ${s1c.pipesPurchased}`);
        }
      }
      const hasQuotedSummary = Boolean(
        sKp && sKp.quotedSummary && sKp.quotedSummary !== 'Сверка потребности на ближайшие объекты'
      );
      if ((sKp && sKp.lastKpDate) || hasQuotedSummary || clientNeed) {
        plain.push(`• Текущее КП / Потребность в работе: ${[
          sKp && sKp.lastKpDate ? 'КП от ' + sKp.lastKpDate : '',
          hasQuotedSummary ? sKp.quotedSummary : '',
          clientNeed ? 'Заявка: ' + clientNeed : ''
        ].filter(Boolean).join(' | ')}`);
      }
      if (customOfferLine) {
        plain.push(`\r\nПерсональное предложение / Условия торга:\r\n${customOfferLine}`);
      }

      plain.push('\r\n----------------------------------------');
      plain.push('• Для наших партнёров: приоритетный резерв на базе в г. Воронеже и индивидуальное согласование цены от объёма.');
      plain.push('• Отдел продаж: +7 (906) 588-37-77 (Кирилл).');
      const plainBody = stripAllEmojisClean_(plain.join('\r\n'));

      const partnerStatsHtml = buildPartnerStatsCardHtml_(s1c, sKp, clientNeed, includeRubles);

      const paragraphsHtml = paragraphs
        .map(p => `<div style="font-size:14px;color:#1e293b;line-height:1.55;margin-bottom:12px;">${p}</div>`)
        .join('');

      const offerBoxHtml = customOfferLine
        ? `<div style="margin:14px 0;padding:14px 16px;background:#fffbeb;border:1px solid #fde68a;border-left:4px solid #d97706;border-radius:8px;font-size:13.5px;color:#0f172a;line-height:1.5;"><b style="color:#b45309;text-transform:uppercase;font-size:11px;letter-spacing:0.5px;display:block;margin-bottom:4px;">Согласование цены и спец-условия поставки:</b>${customOfferLine.replace(/\n/g, '<br>')}</div>`
        : '';

      const rawHtmlBody = `<!DOCTYPE html><html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"><style>
body,table,td,a{-webkit-text-size-adjust:100%;-ms-text-size-adjust:100%}
@media only screen and (max-width:640px){
.ow{padding:0!important}
.ec{width:100%!important;max-width:100%!important;border-radius:0!important;border-left:none!important;border-right:none!important}
.ch{padding:16px 12px!important}
.ht{font-size:16.5px!important}
.cb{padding:16px 12px!important}
.sc{display:block!important;width:100%!important;padding:6px 0!important;border-left:none!important;border-top:1px solid #bae6fd!important}
}
</style></head>
<body style="margin:0;padding:0;width:100%!important;background:#f1f5f9;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;color:#0f172a;">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#f1f5f9" style="width:100%;background:#f1f5f9;border-collapse:collapse;">
<tr><td align="center" class="ow" style="padding:16px 8px;">
<table class="ec" align="center" width="680" border="0" cellpadding="0" cellspacing="0" bgcolor="#ffffff" style="width:100%;max-width:680px;background:#ffffff;border:1px solid #cbd5e1;border-radius:12px;overflow:hidden;border-collapse:separate;">
<tr><td class="ch" bgcolor="#0f172a" style="background:#0f172a;padding:20px 24px;color:#ffffff;border-bottom:3px solid #0284c7;">
<table width="100%" border="0" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;"><tr>
<td align="left" valign="middle" style="padding-right:8px;">
<div class="ht" style="font-size:19px;font-weight:900;letter-spacing:-0.3px;color:#ffffff;line-height:1.15;white-space:nowrap;">ООО «КАПИТАЛ‑ПЛЮС»</div>
<div style="font-size:12px;color:#94a3b8;margin-top:4px;line-height:1.3;">Поставки трубного металлопроката со склада • г. Воронеж</div>
</td>
<td align="right" valign="middle" style="white-space:nowrap;">
<span style="display:inline-block;padding:5px 10px;background:#1e293b;border:1px solid #334155;border-radius:6px;color:#38bdf8;font-size:11px;font-weight:700;white-space:nowrap;">Карта партнёра • г. Воронеж</span>
</td>
</tr></table>
</td></tr>
<tr><td class="cb" style="padding:22px 24px;width:100%;box-sizing:border-box;">
<div style="font-size:16px;font-weight:700;color:#0f172a;margin-bottom:10px;">Здравствуйте${clientLprGreeting}!</div>
${paragraphsHtml}
${partnerStatsHtml}
${offerBoxHtml}
<div style="margin-top:16px;padding:12px 14px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;font-size:12.5px;color:#475569;line-height:1.6;">
• <b style="color:#0f172a;">Индивидуальный подход:</b> Готовы обсудить встречное предложение по проходной цене и зафиксировать объём под Ваш объект.<br>
• <b style="color:#0f172a;">Прямая связь:</b> Направьте спецификацию ответным письмом или в мессенджер <b>+7 (906) 588‑37‑77</b> — дадим расчёт с доставкой за 10 минут.
</div>
${userSignature.html}
</td></tr>
</table>
</td></tr></table></body></html>`;

      const htmlBody = stripAllEmojisClean_(rawHtmlBody);

      const emailOptions = {
        htmlBody: htmlBody,
        name: 'Кирилл Манюхин'
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

        if (!isManualOrTest && client.sheet && client.row && client.row >= 2) {
          const targetSheet = typeof getSheetByNameSafe === 'function' ? getSheetByNameSafe(ss, client.sheet) : ss.getSheetByName(client.sheet);
          const actualRow = typeof resolveClientTargetRowSafe_ === 'function' ? resolveClientTargetRowSafe_(targetSheet, client) : client.row;
          if (targetSheet && actualRow >= 2) {
            const nowStamp = Utilities.formatDate(new Date(), timeZone, 'dd.MM.yyyy HH:mm');
            const lprTag = chosenPerson ? ` (для: ${chosenPerson})` : '';
            const offerShort = customOfferLine ? ` [${customOfferLine.slice(0, 45)}]` : '';
            const actLabel = isDraftOnly
              ? `Черновик Напоминалки${offerShort}${lprTag}`
              : `Напоминалка / Дожим${offerShort}${lprTag}`;
            const tag = stripAllEmojisClean_(`[${actLabel} - ${nowStamp}]`);

            const curHistory = String(targetSheet.getRange(actualRow, 14).getValue() || '').trim();
            targetSheet.getRange(actualRow, 14).setValue(curHistory ? `${tag}\n${curHistory}` : tag);

            const rawNextFollowUp = typeof getNextWorkingDateObj_ === 'function'
              ? getNextWorkingDateObj_(new Date(), nextDays)
              : new Date(new Date().getFullYear(), new Date().getMonth(), new Date().getDate() + nextDays);
            const nextFollowUp = typeof getSmartBalancedDeadlineObj_ === 'function'
              ? getSmartBalancedDeadlineObj_(targetSheet, rawNextFollowUp, timeZone, actualRow)
              : rawNextFollowUp;
            const nextFollowUpIso = Utilities.formatDate(nextFollowUp, timeZone, 'yyyy-MM-dd');

            const todayIso = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
            const curDeadline = targetSheet.getRange(actualRow, 12).getValue();
            const curDlObj = typeof parseDeadlineObj_ === 'function' ? parseDeadlineObj_(curDeadline, timeZone) : { isoDate: '', hasTime: false };
            const keepExistingDeadline =
              (curDlObj.isoDate === todayIso && curDlObj.hasTime) ||
              (curDlObj.isoDate && curDlObj.isoDate > todayIso);

            if (!keepExistingDeadline && (!curDlObj.isoDate || curDlObj.isoDate < nextFollowUpIso)) {
              if (typeof setDeadlineCellDate_ === 'function') {
                setDeadlineCellDate_(targetSheet, actualRow, nextFollowUp, false);
              } else {
                targetSheet.getRange(actualRow, 12).setNumberFormat('dd.MM.yyyy').setValue(nextFollowUp);
              }
            }

            if (typeof refreshSingleRowTactic_ === 'function') {
              refreshSingleRowTactic_(targetSheet, actualRow);
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

function buildRelationshipSubject_(scenarioKey, clientCompany, chosenLpr, s1c, sKp, seed) {
  const companyTag = clientCompany ? ` — ${clientCompany}` : '';
  let lprGreeting = '';
  if (chosenLpr) {
    const lprClean = chosenLpr.split(/[\n\/;,]+/)[0].trim();
    const lprLower = lprClean.toLowerCase();
    const isService = lprLower.includes('приемная') || lprLower.includes('снабжен') || lprLower.includes('бухгалтер') || lprLower.includes('склад') || lprLower.includes('отдел');
    if (!isService && lprClean) {
      const parts = lprClean.split(/\s+/);
      if (parts.length === 3) lprGreeting = `${parts[1]} ${parts[2]}, `;
      else if (parts.length === 2) lprGreeting = `${parts[0]} ${parts[1]}, `;
      else if (parts.length === 1 && parts[0].length > 1) lprGreeting = `${parts[0]}, `;
    }
  }

  if (scenarioKey === 'partner_1c' || (s1c && s1c.has1C && scenarioKey !== 'remind_kp')) {
    return `${lprGreeting}итоги нашего сотрудничества (${s1c && s1c.tonsFormatted ? s1c.tonsFormatted : 'поставки труб'}) и планы снабжения • Капитал-Плюс${companyTag}`;
  }
  if (scenarioKey === 'bargain_price') {
    return `${lprGreeting}согласование индивидуальной цены и условий поставки труб • Капитал-Плюс${companyTag}`;
  }
  const kpDateStr = (sKp && sKp.lastKpDate) ? ` от ${sKp.lastKpDate}` : '';
  return `${lprGreeting}уточнение по направленному коммерческому предложению${kpDateStr} • Капитал-Плюс${companyTag}`;
}

function buildDefaultScenarioParagraphs_(scenarioKey, companyFormattedPlain, s1c, sKp, clientNeed) {
  const compText = companyFormattedPlain ? `компанией ${companyFormattedPlain}` : 'Вашей компанией';
  const kpDateStr = (sKp && sKp.lastKpDate) ? ` от ${sKp.lastKpDate}` : '';
  const quotedStr = (sKp && sKp.quotedSummary) ? ` (${sKp.quotedSummary})` : '';

  if (scenarioKey === 'partner_1c') {
    return [
      `Благодарим за плодотворное сотрудничество с ООО «Капитал-Плюс»! Ниже в карточке партнёра мы собрали краткую сводку наших совместных поставок стальной трубы.`,
      `Сейчас мы формируем график отгрузок и резервов на складе в г. Воронеже. Подскажите, планируются ли у Вас в ближайшее время новые объекты или потребность по профильным диаметрам? Для постоянных партнёров готовы заранее придержать нужный объём и дать приоритетную цену.`
    ];
  }
  if (scenarioKey === 'bargain_price') {
    return [
      `В продолжение наших переговоров с ${compText} высылаем сводку по нашему взаимодействию и предлагаем обсудить индивидуальные условия поставки трубного проката.`,
      `Если у Вас есть целевая (проходная) цена по текущей спецификации или предложение конкурентов — напишите нам, мы оперативно согласуем дополнительную скидку от объёма и зафиксируем металл в резерве.`
    ];
  }
  return [
    `Ранее мы направляли для ${compText} коммерческое предложение${kpDateStr}${quotedStr} со склада в г. Воронеже.`,
    `Подскажите, пожалуйста, удалось ли рассмотреть наше предложение и сохраняется ли актуальность закупки под текущие задачи? Если нужно скорректировать объём, стенку или обсудить цену для прохождения в смету — будем рады оперативно дать обратную связь.`
  ];
}

function buildPartnerStatsCardHtml_(s1c, sKp, clientNeed, includeRubles) {
  const has1C = Boolean(s1c && s1c.has1C);
  const cleanQuoted = (sKp && sKp.quotedSummary && sKp.quotedSummary !== 'Сверка потребности на ближайшие объекты')
    ? sKp.quotedSummary
    : '';
  const hasKp = Boolean((sKp && sKp.lastKpDate) || cleanQuoted);
  if (!has1C && !hasKp && !clientNeed) return '';

  if (has1C) {
    const col1Val = s1c.tonsFormatted || 'В работе';
    const col1Sub = s1c.updCount ? `${s1c.updCount} партий (УПД)` : 'По истории поставок';

    const col2Val = (includeRubles && s1c.revenueFormatted) ? s1c.revenueFormatted : (s1c.partnerStatus || 'Постоянный партнёр');
    const col2Sub = (includeRubles && s1c.avgPriceFormatted) ? `Ср. цена: ${s1c.avgPriceFormatted}` : (s1c.years || 'Прямые поставки');

    const col3Val = s1c.lastOrderDate || (s1c.years || 'Действующий');
    const col3Sub = s1c.years ? `Период: ${s1c.years}` : 'Крайняя отгрузка';

    let detailsLines = [];
    if (s1c.pipesPurchased) {
      detailsLines.push(`<div style="margin-top:8px;padding-top:8px;border-top:1px solid #cbd5e1;font-size:12.5px;color:#1e293b;line-height:1.45;"><b style="color:#0f172a;">Ранее поставлявшиеся позиции:</b> ${s1c.pipesPurchased}</div>`);
    }
    if (sKp && sKp.lastKpDate) {
      detailsLines.push(`<div style="margin-top:4px;font-size:12.5px;color:#334155;line-height:1.45;"><b style="color:#0284c7;">Крайнее направленное КП:</b> от ${sKp.lastKpDate}${cleanQuoted ? ' (' + cleanQuoted + ')' : ''}</div>`);
    } else if (cleanQuoted) {
      detailsLines.push(`<div style="margin-top:4px;font-size:12.5px;color:#15803d;line-height:1.45;"><b>Текущий запрос / спецификация:</b> ${cleanQuoted}</div>`);
    }
    if (clientNeed) {
      detailsLines.push(`<div style="margin-top:4px;font-size:12.5px;color:#15803d;line-height:1.45;"><b>Текущая потребность в работе:</b> ${clientNeed}</div>`);
    }

    return `<div style="margin:14px 0 16px 0;padding:15px 18px;background:#f0f9ff;border:1px solid #bae6fd;border-left:4px solid #0284c7;border-radius:8px;">` +
      `<div style="font-size:11px;font-weight:800;color:#0369a1;text-transform:uppercase;letter-spacing:0.6px;margin-bottom:10px;">Сводка нашего сотрудничества и история отгрузок</div>` +
      `<table width="100%" border="0" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;"><tr>` +
      `<td width="33%" valign="top" style="padding-right:8px;"><div style="font-size:10.5px;color:#64748b;font-weight:600;text-transform:uppercase;">Поставленный объём</div><div style="font-size:16.5px;font-weight:900;color:#0f172a;margin-top:2px;">${col1Val}</div><div style="font-size:11.5px;color:#475569;margin-top:1px;">${col1Sub}</div></td>` +
      `<td class="sc" width="34%" valign="top" style="padding:0 8px;border-left:1px solid #bae6fd;"><div style="font-size:10.5px;color:#64748b;font-weight:600;text-transform:uppercase;">${includeRubles && s1c.revenueFormatted ? 'Объём сделок (с НДС)' : 'Статус партнёра'}</div><div style="font-size:16.5px;font-weight:900;color:#0f172a;margin-top:2px;">${col2Val}</div><div style="font-size:11.5px;color:#475569;margin-top:1px;">${col2Sub}</div></td>` +
      `<td class="sc" width="33%" valign="top" style="padding-left:8px;border-left:1px solid #bae6fd;"><div style="font-size:10.5px;color:#64748b;font-weight:600;text-transform:uppercase;">Крайняя отгрузка</div><div style="font-size:16.5px;font-weight:900;color:#0f172a;margin-top:2px;">${col3Val}</div><div style="font-size:11.5px;color:#475569;margin-top:1px;">${col3Sub}</div></td>` +
      `</tr></table>` +
      `${detailsLines.join('')}` +
      `</div>`;
  }

  let kpRows = [];
  if (sKp && sKp.lastKpDate) {
    kpRows.push(`<div style="font-size:13px;color:#1e293b;line-height:1.5;"><b style="color:#0f172a;">Ранее направленное предложение:</b> от ${sKp.lastKpDate}${cleanQuoted ? ' — ' + cleanQuoted : ''}</div>`);
  } else if (cleanQuoted) {
    kpRows.push(`<div style="font-size:13px;color:#1e293b;line-height:1.5;"><b style="color:#0f172a;">Текущий запрос / спецификация:</b> ${cleanQuoted}</div>`);
  }
  if (clientNeed) {
    kpRows.push(`<div style="font-size:13px;color:#1e293b;line-height:1.5;margin-top:4px;"><b style="color:#0f172a;">Зафиксированная потребность:</b> ${clientNeed}</div>`);
  }
  kpRows.push(`<div style="font-size:12px;color:#475569;margin-top:6px;">Готовы оперативно согласовать индивидуальную цену от объёма и зафиксировать резерв на складе в г. Воронеже.</div>`);

  return `<div style="margin:14px 0 16px 0;padding:14px 16px;background:#f0f9ff;border:1px solid #bae6fd;border-left:4px solid #0284c7;border-radius:8px;">` +
    `<div style="font-size:11px;font-weight:800;color:#0369a1;text-transform:uppercase;letter-spacing:0.6px;margin-bottom:6px;">Статус взаимодействия и согласования КП</div>` +
    `${kpRows.join('')}` +
    `</div>`;
}

function logFollowUpClientAction(payload) {
  try {
    if (!payload || !payload.clients || payload.clients.length === 0) {
      return { success: false, message: 'Клиент не выбран' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timeZone = ss.getSpreadsheetTimeZone() || 'Europe/Moscow';
    const nowStamp = Utilities.formatDate(new Date(), timeZone, 'dd.MM.yyyy HH:mm');
    const nextDays = Math.max(1, parseInt(payload.nextDays, 10) || 3);
    const cleanOffer = stripAllEmojisClean_(payload.offerLine || '');
    const offerShort = cleanOffer ? ` [${cleanOffer.slice(0, 45)}]` : '';

    let loggedCount = 0;
    payload.clients.forEach(client => {
      if (!client || !client.sheet || !client.row || client.row < 2 || client.key === '__TEST_COMPANY__' || client.isTestSelf) return;
      const targetSheet = typeof getSheetByNameSafe === 'function' ? getSheetByNameSafe(ss, client.sheet) : ss.getSheetByName(client.sheet);
      if (!targetSheet) return;

      const actualRow = typeof resolveClientTargetRowSafe_ === 'function' ? resolveClientTargetRowSafe_(targetSheet, client) : client.row;
      if (actualRow < 2) return;

      const chosenLpr = stripAllEmojisClean_(typeof resolveChosenLprString_ === 'function' ? resolveChosenLprString_(client) : (client.selectedLpr || client.lpr || ''));
      const lprTag = chosenLpr ? ` (для: ${chosenLpr})` : '';
      const tag = `[Напоминалка в Мессенджер${offerShort}${lprTag} - ${nowStamp}]`;

      const curHistory = String(targetSheet.getRange(actualRow, 14).getValue() || '').trim();
      targetSheet.getRange(actualRow, 14).setValue(curHistory ? `${tag}\n${curHistory}` : tag);

      const rawNextFollowUp = typeof getNextWorkingDateObj_ === 'function'
        ? getNextWorkingDateObj_(new Date(), nextDays)
        : new Date(new Date().getFullYear(), new Date().getMonth(), new Date().getDate() + nextDays);
      const nextFollowUp = typeof getSmartBalancedDeadlineObj_ === 'function'
        ? getSmartBalancedDeadlineObj_(targetSheet, rawNextFollowUp, timeZone, actualRow)
        : rawNextFollowUp;
      const nextFollowUpIso = Utilities.formatDate(nextFollowUp, timeZone, 'yyyy-MM-dd');
      const todayIso = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
      const curDeadline = targetSheet.getRange(actualRow, 12).getValue();
      const curDlObj = typeof parseDeadlineObj_ === 'function' ? parseDeadlineObj_(curDeadline, timeZone) : { isoDate: '', hasTime: false };
      const keepExistingDeadline =
        (curDlObj.isoDate === todayIso && curDlObj.hasTime) ||
        (curDlObj.isoDate && curDlObj.isoDate > todayIso);

      if (!keepExistingDeadline && (!curDlObj.isoDate || curDlObj.isoDate < nextFollowUpIso)) {
        if (typeof setDeadlineCellDate_ === 'function') {
          setDeadlineCellDate_(targetSheet, actualRow, nextFollowUp, false);
        } else {
          targetSheet.getRange(actualRow, 12).setNumberFormat('dd.MM.yyyy').setValue(nextFollowUp);
        }
      }

      if (typeof refreshSingleRowTactic_ === 'function') {
        refreshSingleRowTactic_(targetSheet, actualRow);
      }
      loggedCount++;
    });

    return { success: loggedCount > 0, loggedCount: loggedCount };
  } catch (e) {
    return { success: false, message: e.message };
  }
}