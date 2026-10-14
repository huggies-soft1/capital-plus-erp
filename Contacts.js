/**
 * ЕДИНЫЙ ЦЕНТР УПРАВЛЕНИЯ ПРОДАЖАМИ И СКЛАДОМ
 * Файл: Contacts.gs — Добавление / Обновление клиента в Google Контактах (People API / iPhone)
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

  // Защита от вызова на складе
  if (sheetName === 'Склад') {
    ui.alert('Внимание', 'На листе «Склад» находятся трубы! Перейдите на клиентский лист.', ui.ButtonSet.OK);
    return;
  }

  const row = activeRange.getRow();
  if (row <= 1) {
    ui.alert('Внимание', 'Выделена строка с заголовками. Выберите строку с клиентом.', ui.ButtonSet.OK);
    return;
  }

  // 1. Считываем данные строки
  const rowVals = sheet.getRange(row, 1, 1, 13).getValues()[0];
  const inn           = String(rowVals[0] || '').trim(); // A
  let company         = String(rowVals[1] || '').trim(); // B
  const city          = String(rowVals[2] || '').trim(); // C
  const contactPerson = String(rowVals[3] || '').trim(); // D
  const phoneRaw      = String(rowVals[4] || '').trim(); // E
  const emailRaw      = String(rowVals[5] || '').trim(); // F

  let sphere = '';
  let notesRaw = '';

  // Считываем сферу и примечания (Инфоповод не трогаем вообще)
  if (sheetName === "1. Обзвон (База)") {
    sphere   = String(rowVals[6] || '').trim();  // G: Сфера деятельности
    notesRaw = String(rowVals[10] || '').trim(); // K: Примечание о клиенте
  } else if (sheetName === "2. Золотой список (Киты)" || sheetName === "3. Лист ожидания") {
    sphere   = String(rowVals[6] || '').trim();  // G: Сфера деятельности
    notesRaw = String(rowVals[10] || '').trim(); // K: Примечания
  } else if (sheetName === "4.Архив лидов") {
    notesRaw = `Причина отказа: ${rowVals[7]}\n${rowVals[8]}`.trim();
  } else {
    sphere   = String(rowVals[6] || '').trim();
    notesRaw = String(rowVals[10] || '').trim();
  }

  company = typeof cleanCompanyName === 'function' ? cleanCompanyName(company) : company;

  if (!company && !contactPerson && !phoneRaw && !emailRaw) {
    ui.alert('Ошибка', 'В выбранной строке нет контактных данных!', ui.ButtonSet.OK);
    return;
  }

   // 1. Формирование имени для iPhone (ВСЁ в Имя, Фамилия пустая — чтобы сортировка шла строго по ИМЕНИ человека на «А», «С», «И»!)
  let givenName = '';
  if (contactPerson && company) {
    givenName = `${contactPerson} [${company}]`;
  } else {
    givenName = contactPerson || company;
  }
  let familyName = ''; // Оставляем пустой, чтобы iOS не кидала контакт на букву «О» (из-за ООО)

  // 3. E-mail
  const emailList = emailRaw.split(/[,;\n\s]+/)
    .map(e => e.trim())
    .filter(e => e && e.includes('@'));
  const emailPayload = emailList.map((email, idx) => ({
    value: email,
    type: idx === 0 ? 'work' : 'other'
  }));

  // 4. Приведение телефонов к международному формату +7...
  const phoneLines = phoneRaw.split(/[\n,;]+/).map(p => p.trim()).filter(Boolean);
  const phonePayload = [];
  phoneLines.forEach((p, idx) => {
    let cleanDigits = p.replace(/[^\d+]/g, '');
    if (cleanDigits.startsWith('+7')) {
      cleanDigits = '+7' + cleanDigits.substring(2).replace(/\D/g, '');
    } else if (cleanDigits.length === 11 && (cleanDigits.startsWith('8') || cleanDigits.startsWith('7'))) {
      cleanDigits = '+7' + cleanDigits.substring(1);
    } else if (cleanDigits.length === 10) {
      cleanDigits = '+7' + cleanDigits;
    }
    if (cleanDigits) {
      phonePayload.push({
        value: cleanDigits,
        type: idx === 0 ? 'work' : 'other'
      });
    }
  });

  // 5. В Заметки контакта (СФЕРА ДЕЯТЕЛЬНОСТИ на первом месте)
  let notesLines = [];
  if (sphere)   notesLines.push(`Сфера деятельности: ${sphere}`);
  if (company)  notesLines.push(`Компания: ${company}`);
  if (city)     notesLines.push(`Город: ${city}`);
  if (inn)      notesLines.push(`ИНН: ${inn}`);
  if (notesRaw) notesLines.push(`Примечание:\n${notesRaw}`);

  // 6. Формирование объекта контакта для People API
  const contactResource = {
    names: [{ givenName: givenName, familyName: familyName }],
    organizations: company ? [{ name: company, title: contactPerson ? 'ЛПР' : 'Контактное лицо' }] : [],
    addresses: city ? [{ city: city, type: 'work' }] : [],
    emailAddresses: emailPayload,
    phoneNumbers: phonePayload,
    biographies: notesLines.length > 0 ? [{ value: notesLines.join('\n\n'), contentType: 'TEXT_PLAIN' }] : []
  };

  try {
    // 7. Поиск: ищем ТОЛЬКО по номеру телефона или точному Email
    const existingContact = findExistingContactStrict_(phonePayload, emailPayload);

    if (existingContact && existingContact.resourceName) {
      // Обновляем существующий контакт (найден по совпадению номера или почты)
      contactResource.etag = existingContact.etag;
      People.People.updateContact(contactResource, existingContact.resourceName, {
        updatePersonFields: 'names,organizations,addresses,emailAddresses,phoneNumbers,biographies'
      });
      const labelText = [givenName, familyName].filter(Boolean).join(' ');
      ss.toast(`«${labelText}» обновлен в контактах!`, '🔄 Контакт обновлен', 4);
    } else {
      // Создаем новый контакт
      People.People.createContact(contactResource);
      const labelText = [givenName, familyName].filter(Boolean).join(' ');
      ss.toast(`«${labelText}» добавлен в контакты!`, '✅ Новый контакт', 4);
    }
  } catch (err) {
    ui.alert('Ошибка People API', 
      'Не удалось сохранить/обновить контакт.\nУбедитесь, что служба People API включена в редакторе Apps Script (раздел «Службы» слева).\n\nДетали: ' + err.message, 
      ui.ButtonSet.OK);
  }
}

/**
 * СТРОГИЙ поиск контакта: ТОЛЬКО по телефону (10 цифр) или Email.
 * По именам («Алексей», «Иван») поиск ЗАПРЕЩЕН, чтобы не склеивать тезок!
 */
function findExistingContactStrict_(phones, emails) {
  // 1. Поиск по телефону с обязательной сверкой цифр
  for (let i = 0; i < phones.length; i++) {
    const digits = phones[i].value.replace(/\D/g, '');
    const last10 = digits.length >= 10 ? digits.slice(-10) : digits;
    if (last10.length >= 7) {
      const candidates = searchPeopleCandidates_(last10);
      for (let c = 0; c < candidates.length; c++) {
        const person = candidates[c];
        if (personPhoneMatches_(person, last10)) {
          return person; // Найдено точное совпадение по номеру
        }
      }
    }
  }

  // 2. Поиск по Email с точной сверкой
  for (let i = 0; i < emails.length; i++) {
    const targetEmail = (emails[i].value || '').toLowerCase().trim();
    if (targetEmail) {
      const candidates = searchPeopleCandidates_(targetEmail);
      for (let c = 0; c < candidates.length; c++) {
        const person = candidates[c];
        if (personEmailMatches_(person, targetEmail)) {
          return person; // Найдено точное совпадение по email
        }
      }
    }
  }

  return null; // Совпадений нет — это новый человек
}

function searchPeopleCandidates_(query) {
  try {
    const res = People.People.searchContacts({
      query: query,
      readMask: 'names,emailAddresses,phoneNumbers,metadata'
    });
    if (res && res.results && res.results.length > 0) {
      return res.results.map(r => r.person).filter(Boolean);
    }
  } catch (e) {
    Logger.log("Ошибка поиска в People API: " + e.message);
  }
  return [];
}

function personPhoneMatches_(person, targetLast10) {
  if (!person || !person.phoneNumbers) return false;
  return person.phoneNumbers.some(p => {
    const pDigits = (p.value || '').replace(/\D/g, '');
    const pLast10 = pDigits.length >= 10 ? pDigits.slice(-10) : pDigits;
    return pLast10 === targetLast10;
  });
}

function personEmailMatches_(person, targetEmail) {
  if (!person || !person.emailAddresses) return false;
  return person.emailAddresses.some(e => {
    return (e.value || '').toLowerCase().trim() === targetEmail;
  });
}