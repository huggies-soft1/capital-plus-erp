/**
 * ЕДИНЫЙ ЦЕНТР УПРАВЛЕНИЯ ПРОДАЖАМИ И СКЛАДОМ
 * Файл: Contacts.gs — Синхронизация с Контактами iPhone (People API) с защитой от лимитов квот
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

  // 1. Считываем все 14 колонок (Единый стандарт A–N)
  const rowVals = sheet.getRange(row, 1, 1, 14).getValues()[0];
  const inn        = String(rowVals[0] || '').trim(); // A: ИНН
  let company      = String(rowVals[1] || '').trim(); // B: Компания
  const city       = String(rowVals[2] || '').trim(); // C: Город
  const contactRaw = String(rowVals[3] || '').trim(); // D: ЛПР
  const phoneRaw   = String(rowVals[4] || '').trim(); // E: Телефон
  const emailRaw   = String(rowVals[5] || '').trim(); // F: Email
  const sphere     = String(rowVals[6] || '').trim(); // G: Сфера
  const need       = String(rowVals[7] || '').trim(); // H: Потребность
  const notesRaw   = String(rowVals[10] || '').trim(); // K: Примечание

  company = typeof cleanCompanyName === 'function' ? cleanCompanyName(company) : company;

  if (!company && !contactRaw && !phoneRaw && !emailRaw) {
    ui.alert('Ошибка', 'В выбранной строке нет контактных данных!', ui.ButtonSet.OK);
    return;
  }

  const meta = {
    company: company,
    city: city,
    inn: inn,
    sphere: sphere,
    need: need,
    notes: notesRaw
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

    // Группируем профили, если они ведут в один и тот же существующий контакт (чтобы не бить по API дважды!)
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
          // Сливаем второй телефон/имя в тот же запрос
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

    // 3. Выполняем обновления существующих контактов (по 1 запросу на уникальный контакт)
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

    // 4. Создаем новые контакты (если совпадений не нашлось)
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
 * Авто-повтор запроса при секундном лимите квоты Google (Resource has been exhausted)
 */
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

/**
 * Разбор ячейки ЛПР и Телефонов:
 * - Если несколько строк в ЛПР совпадают по числу со строками телефонов -> создает отдельные контакты на каждого ЛПР.
 * - Иначе -> привязывает все номера одному человеку / компании.
 */
function buildPersonProfilesFromRow_(contactRaw, phoneRaw, emailRaw) {
  const contactLines = contactRaw.split(/\n+/).map(s => s.trim()).filter(Boolean);
  const rawPhoneLines = phoneRaw.split(/[\n;]+/).map(s => s.trim()).filter(Boolean);
  const emails = emailRaw.split(/[,;\n\s]+/).map(e => e.trim()).filter(e => e && e.includes('@'));

  const cleanPhones = [];
  rawPhoneLines.forEach(line => {
    const matches = line.match(/(?:\+7|8|7)?[\s\-\.\(\)]*\d{3}[\s\-\.\(\)]*\d{3}[\s\-\.]*\d{2}[\s\-\.]*\d{2}|\d{10,11}/g) || [line];
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

  // Если в ячейке ЛПР явно записано несколько человек с новой строки (по числу телефонов)
  if (contactLines.length > 1 && contactLines.length === cleanPhones.length) {
    return contactLines.map((personName, idx) => ({
      personName: personName,
      phones: [{ value: cleanPhones[idx], type: 'work' }],
      emails: idx === 0 ? emails.map((em, i) => ({ value: em, type: i === 0 ? 'work' : 'other' })) : []
    }));
  }

  // Один человек (или карточка компании) со всеми телефонами
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
  const givenName = formatIphoneGivenName_(profile.personName, [meta.company]);
  const bioText = buildCompanyNotesBlock_(meta);

  return {
    names: [{ givenName: givenName, familyName: '' }],
    organizations: meta.company ? [{ name: meta.company, title: profile.personName ? 'ЛПР' : 'Контакт' }] : [],
    addresses: meta.city ? [{ city: meta.city, type: 'work' }] : [],
    emailAddresses: profile.emails,
    phoneNumbers: profile.phones,
    biographies: bioText ? [{ value: bioText, contentType: 'TEXT_PLAIN' }] : []
  };
}

/**
 * Умное слияние данных + проверка, изменилось ли хоть что-то (чтобы не тратить квоту зря)
 */
function mergeContactData_(existingContact, profile, meta) {
  const existingGiven = (existingContact.names && existingContact.names[0] && existingContact.names[0].givenName)
    ? existingContact.names[0].givenName.trim()
    : '';

  let companiesList = [];
  const bracketMatch = existingGiven.match(/\[([^\]]+)\]/);
  if (bracketMatch && bracketMatch[1]) {
    companiesList = bracketMatch[1].split(/\s*\/\s*/).map(s => s.trim()).filter(Boolean);
  } else if (existingContact.organizations && existingContact.organizations.length > 0) {
    existingContact.organizations.forEach(org => {
      if (org.name && !companiesList.includes(org.name.trim())) {
        companiesList.push(org.name.trim());
      }
    });
  }

  if (meta.company) {
    const normNew = meta.company.toLowerCase();
    const alreadyHas = companiesList.some(c => c.toLowerCase() === normNew);
    if (!alreadyHas) {
      companiesList.push(meta.company);
    }
  }

  const existingPersonNameOnly = existingGiven.replace(/\s*\[[^\]]*\]/g, '').trim();
  const finalPersonName = profile.personName || existingPersonNameOnly;
  const finalGivenName = formatIphoneGivenName_(finalPersonName, companiesList);

  const mergedPhones = (existingContact.phoneNumbers || []).map(p => ({ value: p.value, type: p.type || 'work' }));
  let phonesAdded = false;
  profile.phones.forEach(newPh => {
    const new10 = newPh.value.replace(/\D/g, '').slice(-10);
    const exists = mergedPhones.some(oldPh => (oldPh.value || '').replace(/\D/g, '').slice(-10) === new10);
    if (!exists && new10.length >= 7) {
      mergedPhones.push(newPh);
      phonesAdded = true;
    }
  });

  const mergedEmails = (existingContact.emailAddresses || []).map(e => ({ value: e.value, type: e.type || 'work' }));
  let emailsAdded = false;
  profile.emails.forEach(newEm => {
    const normEm = newEm.value.toLowerCase().trim();
    const exists = mergedEmails.some(oldEm => (oldEm.value || '').toLowerCase().trim() === normEm);
    if (!exists) {
      mergedEmails.push(newEm);
      emailsAdded = true;
    }
  });

  const existingBio = (existingContact.biographies && existingContact.biographies[0] && existingContact.biographies[0].value)
    ? existingContact.biographies[0].value.trim()
    : '';
  const newBlock = buildCompanyNotesBlock_(meta);
  let finalBio = existingBio;

  if (newBlock) {
    if (!existingBio) {
      finalBio = newBlock;
    } else if (meta.company && existingBio.includes(`Компания: ${meta.company}`)) {
      const escapedComp = meta.company.replace(/[.*+?^\${}()|[\]\\]/g, '\\$&');
      const blockRegex = new RegExp(`(?:🏢 )?Компания: ${escapedComp}[\\s\\S]*?(?=\\n═|$)`, 'i');
      finalBio = blockRegex.test(existingBio)
        ? existingBio.replace(blockRegex, newBlock)
        : existingBio + '\n════════════════════\n' + newBlock;
    } else {
      finalBio = existingBio + '\n════════════════════\n' + newBlock;
    }
  }

  const hasChanges = (finalGivenName !== existingGiven) || phonesAdded || emailsAdded || (finalBio.trim() !== existingBio.trim());

  return {
    hasChanges: hasChanges,
    resource: {
      etag: existingContact.etag,
      names: [{ givenName: finalGivenName, familyName: '' }],
      organizations: companiesList.map(cName => ({ name: cName, title: finalPersonName ? 'ЛПР' : 'Контакт' })),
      addresses: meta.city ? [{ city: meta.city, type: 'work' }] : (existingContact.addresses || []),
      emailAddresses: mergedEmails,
      phoneNumbers: mergedPhones,
      biographies: finalBio ? [{ value: finalBio, contentType: 'TEXT_PLAIN' }] : []
    }
  };
}

function formatIphoneGivenName_(personName, companiesList) {
  const cleanCompanies = (companiesList || []).filter(Boolean).join(' / ');
  if (personName && cleanCompanies) {
    return `${personName} [${cleanCompanies}]`;
  }
  return personName || cleanCompanies || 'Контакт';
}

function buildCompanyNotesBlock_(meta) {
  const lines = [];
  if (meta.company) lines.push(`🏢 Компания: ${meta.company}`);
  if (meta.sphere)  lines.push(`Сфера деятельности: ${meta.sphere}`);
  if (meta.need)    lines.push(`Потребность: ${meta.need}`);
  if (meta.city)    lines.push(`Город: ${meta.city}`);
  if (meta.inn)     lines.push(`ИНН: ${meta.inn}`);
  if (meta.notes)   lines.push(`Примечание: ${meta.notes}`);
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