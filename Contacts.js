/**
 * ЕДИНЫЙ ЦЕНТР УПРАВЛЕНИЯ ПРОДАЖАМИ И СКЛАДОМ
 * Файл: Contacts.gs — Синхронизация с Контактами iPhone (People API)
 * Без дублей компаний через «/» и без дублирования полей внутри карточки телефона.
 */

function addSelectedRowToContacts() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName().trim();
  const activeRange = sheet.getActiveRange();

  if (!activeRange) {
    ui.alert('Предупреждение', 'Выделите строку с клиентом.', ui.ButtonSet.OK);
    return;
  }

  if (sheetName === 'Склад') {
    ui.alert('Внимание', 'На листе «Склад» находятся трубы! Перейдите на клиентский лист.', ui.ButtonSet.OK);
    return;
  }

  const row = activeRange.getRow();
  if (row <= 1) {
    ui.alert('Внимание', 'Выделена строка с заголовками. Выберите строку с клиентом.', ui.ButtonSet.OK);
    return;
  }

  // 1. Считываем данные строки (A–N) + полное Досье из заметки ячейки B
  const rowVals = sheet.getRange(row, 1, 1, 14).getValues()[0];
  const bNoteRaw = String(sheet.getRange(row, 2).getNote() || '').trim();

  const inn        = String(rowVals[0] || '').trim(); // A: ИНН
  let company      = String(rowVals[1] || '').trim(); // B: Компания
  const city       = String(rowVals[2] || '').trim(); // C: Город (адрес)
  const contactRaw = String(rowVals[3] || '').trim(); // D: ЛПР
  const phoneRaw   = String(rowVals[4] || '').trim(); // E: Телефон
  const emailRaw   = String(rowVals[5] || '').trim(); // F: E-mail
  const sphere     = String(rowVals[6] || '').trim(); // G: Сфера
  const kNoteRaw   = String(rowVals[10] || '').trim(); // K: Примечание

  company = typeof cleanCompanyName === 'function' ? cleanCompanyName(company) : company;

  if (!company && !contactRaw && !phoneRaw && !emailRaw) {
    ui.alert('Ошибка', 'В выбранной строке нет контактных данных!', ui.ButtonSet.OK);
    return;
  }

  const cleanNotes = extractCleanNoteForContact_(bNoteRaw, kNoteRaw);

  const meta = {
    company: company,
    city: city,
    inn: inn,
    sphere: sphere,
    notes: cleanNotes
  };

  // 2. Разбираем людей и телефоны
  const personProfiles = buildPersonProfilesFromRow_(contactRaw, phoneRaw, emailRaw);

  if (personProfiles.length === 0) {
    ui.alert('Ошибка', 'Не удалось распознать телефон, email или имя контакта!', ui.ButtonSet.OK);
    return;
  }

  try {
    const allContacts = fetchAllGoogleContactsLive_();
    const resultsLog = [];

    const updatesByResource = {};
    const newProfilesToCreate = [];

    personProfiles.forEach(profile => {
      const existing = findMatchingContactInList_(allContacts, profile.phones, profile.emails);
      if (existing && existing.resourceName) {
        const rName = existing.resourceName;
        if (!updatesByResource[rName]) {
          updatesByResource[rName] = {
            existing: existing,
            combinedProfile: {
              personName: profile.personName,
              phones: profile.phones.slice(),
              emails: profile.emails.slice()
            }
          };
        } else {
          if (profile.personName && !updatesByResource[rName].combinedProfile.personName.includes(profile.personName)) {
            updatesByResource[rName].combinedProfile.personName =
              [updatesByResource[rName].combinedProfile.personName, profile.personName].filter(Boolean).join(' / ');
          }
          updatesByResource[rName].combinedProfile.phones.push.apply(updatesByResource[rName].combinedProfile.phones, profile.phones);
          updatesByResource[rName].combinedProfile.emails.push.apply(updatesByResource[rName].combinedProfile.emails, profile.emails);
        }
      } else {
        newProfilesToCreate.push(profile);
      }
    });

    let apiWriteCount = 0;

    // 3. Обновление существующих контактов (с очисткой возможных старых дублей компаний через "/")
    Object.keys(updatesByResource).forEach(rName => {
      const entry = updatesByResource[rName];
      const mergeResult = mergeContactData_(entry.existing, entry.combinedProfile, meta);

      if (!mergeResult.hasChanges) {
        resultsLog.push(`✅ Уже актуален: ${mergeResult.resource.names[0].givenName}`);
        return;
      }

      if (apiWriteCount > 0) Utilities.sleep(1200);

      callPeopleApiWithRetry_(() => {
        return People.People.updateContact(mergeResult.resource, rName, {
          updatePersonFields: 'names,organizations,addresses,emailAddresses,phoneNumbers,biographies'
        });
      });
      apiWriteCount++;
      resultsLog.push(`🔄 Обновлен: ${mergeResult.resource.names[0].givenName}`);
    });

    // 4. Создание новых контактов
    newProfilesToCreate.forEach(profile => {
      if (apiWriteCount > 0) Utilities.sleep(1200);

      const newResource = createFreshContactResource_(profile, meta);
      const created = callPeopleApiWithRetry_(() => {
        return People.People.createContact(newResource);
      });
      if (created) allContacts.push(created);
      apiWriteCount++;
      resultsLog.push(`➕ Создан: ${newResource.names[0].givenName}`);
    });

    ss.toast(resultsLog.join('\n'), '📱 Контакты iPhone', 5);
  } catch (err) {
    ui.alert(
      'Ошибка People API',
      'Не удалось сохранить/обновить контакт.\nДетали: ' + err.message,
      ui.ButtonSet.OK
    );
  }
}

/**
 * Достает чистый текст заметок о клиенте (без системной шапки Досье)
 */
function extractCleanNoteForContact_(bNoteRaw, kNoteRaw) {
  if (bNoteRaw) {
    return bNoteRaw
      .replace(/^🗂️\s*ДОСЬЕ КЛИЕНТА[^\n]*\n+/i, '')
      .trim();
  }
  if (kNoteRaw) {
    return kNoteRaw.replace(/^📌\s*/, '').trim();
  }
  return '';
}

function callPeopleApiWithRetry_(apiFunc) {
  const delays = [1500, 3000, 5000];
  for (let attempt = 0; attempt <= delays.length; attempt++) {
    try {
      return apiFunc();
    } catch (e) {
      const msg = (e && e.message) ? e.message.toLowerCase() : '';
      const isQuotaError = msg.includes('exhausted') || msg.includes('quota') || msg.includes('rate') || msg.includes('429');
      if (isQuotaError && attempt < delays.length) {
        Utilities.sleep(delays[attempt]);
        continue;
      }
      throw e;
    }
  }
}

function buildPersonProfilesFromRow_(contactRaw, phoneRaw, emailRaw) {
  const contactLines = contactRaw.split(/\n+/).map(s => s.trim()).filter(Boolean);
  const rawPhoneLines = phoneRaw.split(/[\n;]+/).map(s => s.trim()).filter(Boolean);
  const emails = Array.from(new Set(
    emailRaw.split(/[,;\n\s]+/).map(e => e.trim()).filter(e => e && e.includes('@'))
  ));

  const cleanPhones = [];
  rawPhoneLines.forEach(line => {
    const matches = line.match(/(?:\+7|8|7)?[\s\-\.\(\)]*\d{3,5}[\s\-\.\(\)]*\d{2,3}[\s\-\.]*\d{2}[\s\-\.]*\d{2}|\d{10,11}/g) || [line];
    matches.forEach(m => {
      const digits = m.replace(/\D/g, '');
      let formatted = '';
      if (digits.length === 11 && (digits.startsWith('7') || digits.startsWith('8'))) {
        formatted = '+7' + digits.substring(1);
      } else if (digits.length === 10) {
        formatted = '+7' + digits;
      }
      if (formatted && !cleanPhones.includes(formatted)) {
        cleanPhones.push(formatted);
      }
    });
  });

  if (contactLines.length > 1 && contactLines.length === cleanPhones.length) {
    return contactLines.map((personName, idx) => ({
      personName: personName,
      phones: [{ value: cleanPhones[idx], type: 'work' }],
      emails: idx === 0 ? emails.map((em, i) => ({ value: em, type: i === 0 ? 'work' : 'other' })) : []
    }));
  }

  const singleName = contactLines.join(' / ');
  return [{
    personName: singleName,
    phones: cleanPhones.map((ph, idx) => ({
      value: ph,
      type: idx === 0 ? 'work' : 'mobile'
    })),
    emails: emails.map((em, idx) => ({
      value: em,
      type: idx === 0 ? 'work' : 'other'
    }))
  }];
}

function createFreshContactResource_(profile, meta) {
  const givenName = formatIphoneGivenName_(profile.personName, meta.company);
  const bioText = buildCompanyNotesBlock_(meta);

  return {
    names: [{ givenName: givenName, familyName: '' }],
    organizations: meta.company ? [{ name: meta.company, title: profile.personName ? 'ЛПР' : 'Контакт' }] : [],
    addresses: meta.city ? [{ streetAddress: meta.city, formattedValue: meta.city, type: 'work' }] : [],
    emailAddresses: profile.emails,
    phoneNumbers: profile.phones,
    biographies: bioText ? [{ value: bioText, contentType: 'TEXT_PLAIN' }] : []
  };
}

/**
 * Обновление контакта БЕЗ дублирования компаний через "/" и БЕЗ дублирования заметок
 */
function mergeContactData_(existingContact, profile, meta) {
  const existingGiven = (existingContact.names && existingContact.names[0] && existingContact.names[0].givenName)
    ? existingContact.names[0].givenName.trim()
    : '';

  // Достаем старую компанию только на случай, если в таблице колонка компании пустая (берем строго первую до "/")
  let existingSingleCompany = '';
  const bracketMatch = existingGiven.match(/\[([^\]]+)\]/);
  if (bracketMatch && bracketMatch[1]) {
    existingSingleCompany = bracketMatch[1].split(/\s*\/\s*/)[0].trim();
  } else if (existingContact.organizations && existingContact.organizations.length > 0 && existingContact.organizations[0].name) {
    existingSingleCompany = existingContact.organizations[0].name.split(/\s*\/\s*/)[0].trim();
  }

  // Строго ОДНА компания (никаких склеек через "/")
  const finalCompany = meta.company || existingSingleCompany;

  // Очищаем имя ЛПР от любых старых скобок [Компания]
  const existingPersonNameOnly = existingGiven.replace(/\s*\[[^\]]*\]/g, '').trim();
  const finalPersonName = profile.personName || existingPersonNameOnly;
  const finalGivenName = formatIphoneGivenName_(finalPersonName, finalCompany);

  // Телефоны без дублей
  const mergedPhones = [];
  const seenPhones10 = new Set();
  const allCandidatePhones = (existingContact.phoneNumbers || []).concat(profile.phones || []);
  allCandidatePhones.forEach(p => {
    const val = String(p.value || '').trim();
    const d10 = val.replace(/\D/g, '').slice(-10);
    if (d10.length >= 7 && !seenPhones10.has(d10)) {
      seenPhones10.add(d10);
      mergedPhones.push({ value: val, type: p.type || (mergedPhones.length === 0 ? 'work' : 'mobile') });
    }
  });

  // E-mail без дублей
  const mergedEmails = [];
  const seenEmails = new Set();
  const allCandidateEmails = (existingContact.emailAddresses || []).concat(profile.emails || []);
  allCandidateEmails.forEach(e => {
    const val = String(e.value || '').trim();
    const norm = val.toLowerCase();
    if (norm && norm.includes('@') && !seenEmails.has(norm)) {
      seenEmails.add(norm);
      mergedEmails.push({ value: val, type: e.type || (mergedEmails.length === 0 ? 'work' : 'other') });
    }
  });

  // Заметки (ИНН, Сфера, Примечание) — чистая перезапись актуальными данными без наслоения через "════"
  const existingBio = (existingContact.biographies && existingContact.biographies[0] && existingContact.biographies[0].value)
    ? existingContact.biographies[0].value.trim()
    : '';
  const newBlock = buildCompanyNotesBlock_(meta);
  const finalBio = newBlock || existingBio;

  const existingOrgName = (existingContact.organizations && existingContact.organizations[0] && existingContact.organizations[0].name)
    ? existingContact.organizations[0].name.trim()
    : '';
  const existingAddrObj = (existingContact.addresses && existingContact.addresses[0]) ? existingContact.addresses[0] : null;
  const existingCity = existingAddrObj
    ? String(existingAddrObj.streetAddress || existingAddrObj.formattedValue || existingAddrObj.city || '').trim()
    : '';

  const phonesChanged = mergedPhones.length !== (existingContact.phoneNumbers || []).length;
  const emailsChanged = mergedEmails.length !== (existingContact.emailAddresses || []).length;
  const orgChanged    = finalCompany !== existingOrgName || (existingContact.organizations || []).length > 1;
  const cityChanged   = Boolean(meta.city && (meta.city !== existingCity || (existingAddrObj && !existingAddrObj.streetAddress)));

  const hasChanges = (finalGivenName !== existingGiven) ||
                     orgChanged ||
                     cityChanged ||
                     phonesChanged ||
                     emailsChanged ||
                     (finalBio !== existingBio);

  return {
    hasChanges: hasChanges,
    resource: {
      etag: existingContact.etag,
      names: [{ givenName: finalGivenName, familyName: '' }],
      organizations: finalCompany ? [{ name: finalCompany, title: finalPersonName ? 'ЛПР' : 'Контакт' }] : [],
      addresses: meta.city ? [{ streetAddress: meta.city, formattedValue: meta.city, type: 'work' }] : (existingContact.addresses || []),
      emailAddresses: mergedEmails,
      phoneNumbers: mergedPhones,
      biographies: finalBio ? [{ value: finalBio, contentType: 'TEXT_PLAIN' }] : []
    }
  };
}

/**
 * Формат имени: строго "ЛПР [Компания]" (одна компания без "/")
 */
function formatIphoneGivenName_(personName, singleCompany) {
  const cleanComp = String(singleCompany || '').trim();
  const cleanPerson = String(personName || '').trim();
  if (cleanPerson && cleanComp) {
    return `${cleanPerson} [${cleanComp}]`;
  }
  return cleanPerson || cleanComp || 'Контакт';
}

/**
 * В Заметки карточки идут ТОЛЬКО те поля, которых нет в отдельных графах iPhone:
 * Сфера, ИНН и Примечание (Компания и Город уже стоят в своих штатных полях карточки и не дублируются).
 */
function buildCompanyNotesBlock_(meta) {
  const lines = [];
  if (meta.sphere) lines.push(`Сфера: ${meta.sphere}`);
  if (meta.inn)    lines.push(`ИНН: ${meta.inn}`);
  if (meta.notes)  lines.push(`Примечание:\n${meta.notes}`);
  return lines.join('\n');
}

function fetchAllGoogleContactsLive_() {
  const all = [];
  let pageToken = null;

  do {
    const resp = callPeopleApiWithRetry_(() => {
      return People.People.Connections.list('people/me', {
        personFields: 'names,emailAddresses,phoneNumbers,organizations,addresses,biographies,metadata',
        pageSize: 1000,
        pageToken: pageToken
      });
    });
    if (resp && resp.connections) {
      all.push.apply(all, resp.connections);
    }
    pageToken = resp ? resp.nextPageToken : null;
  } while (pageToken);

  return all;
}

function findMatchingContactInList_(allContacts, phones, emails) {
  const targetPhones10 = phones
    .map(p => (p.value || '').replace(/\D/g, ''))
    .map(d => d.length >= 10 ? d.slice(-10) : d)
    .filter(d => d.length >= 7);

  const targetEmails = emails
    .map(e => (e.value || '').toLowerCase().trim())
    .filter(Boolean);

  for (let i = 0; i < allContacts.length; i++) {
    const person = allContacts[i];

    if (targetPhones10.length > 0 && person.phoneNumbers) {
      const phoneMatch = person.phoneNumbers.some(p => {
        const d = (p.value || '').replace(/\D/g, '');
        const last10 = d.length >= 10 ? d.slice(-10) : d;
        return targetPhones10.includes(last10);
      });
      if (phoneMatch) return person;
    }

    if (targetEmails.length > 0 && person.emailAddresses) {
      const emailMatch = person.emailAddresses.some(e => {
        return targetEmails.includes((e.value || '').toLowerCase().trim());
      });
      if (emailMatch) return person;
    }
  }

  return null;
}