function doGet(e) {
  // GET-запит. Повертаємо помилку про необхідність використання POST
  let response = {};
  response.state = false;
  response.error = { message: ["please, use POST method"] };
  return ContentService.createTextOutput(JSON.stringify(response));
}
function doPost(e) {
  // POST-запит
  try {
    let param, sheet,
      response = { state: true };
    try {
      // Перетворення рядка тіла запиту на масив
      param = JSON.parse(e.postData.contents);
    } catch {
      // Повертаємо помилку отримання тіла запиту (якщо тіло пусте, або не являється валідним JSON)
      response.state = false;
      response.error = { message: ["failed parse input data"] };
      return ContentService.createTextOutput(JSON.stringify(response));
    }
    if (!("action" in param)) {
      // Помилка про відсутність параметру action
      response.state = false;
      response.error = { message: ["'action' is missing"] };
      return ContentService.createTextOutput(JSON.stringify(response));
    }
    if ("fileId" in param && param.fileId != null && param.fileId != "") {
      try {
        if ("sheetId" in param && param.sheetId != null && param.sheetId != "") {
          let sheets = SpreadsheetApp.openById(param.fileId).getSheets();
          for (let s=0; s<sheets.length; s++) {
            if (sheets[s].getSheetId() == param.sheetId) {
              sheet = sheets[s];
              break;
            }
          }
          if (sheet == null) {
            // Помилка отримання листа з файлу
            response.state = false;
            response.error = { message: ["'sheet' is not found"] };
            return ContentService.createTextOutput(JSON.stringify(response));
          }
        } else if ("sheet" in param && param.sheet != null && param.sheet != "") {
          sheet = SpreadsheetApp.openById(param.fileId).getSheetByName(param.sheet);
          if (sheet == null) {
            // Помилка отримання листа з файлу
            response.state = false;
            response.error = { message: ["'sheet' is not found"] };
            return ContentService.createTextOutput(JSON.stringify(response));
          }
        } else {
          // Помилка про відсутність назви/ідентифікатора листа
          response.state = false;
          response.error = { message: ["'sheet' or 'sheetId' is missing"] };
          return ContentService.createTextOutput(JSON.stringify(response));
        }
      } catch {
        response.state = false;
        response.error = { message: ["failed load sheet from fileId"] };
        return ContentService.createTextOutput(JSON.stringify(response));
      }
    } else if ("file" in param && param.file != null && param.file != "") {
      // Пошук файлів за іменем
      try {
        let files = DriveApp.getFilesByName(param.file);
        if (!files.hasNext()) {
          // Файлів з таким іменем немає на диску
          response.state = false;
          response.error = { message: ["file is not found"] };
          return ContentService.createTextOutput(JSON.stringify(response));
        }
        // Завантаження першого файлу з пошуку
        let file = files.next();
        if (file == null) {
          // Помилка завантаження файлу
          response.state = false;
          response.error = { message: ["failed loading file"] };
          return ContentService.createTextOutput(JSON.stringify(response));
        }
        if ("sheetId" in param && param.sheetId != null && param.sheetId != "") {
          let sheets = SpreadsheetApp.open(file).getSheets();
          for (let s=0; s<sheets.length; s++) {
            if (sheets[s].getSheetId() == param.sheetId) {
              sheet = sheets[s];
              break;
            }
          }
          if (sheet == null) {
            // Помилка отримання листа з файлу
            response.state = false;
            response.error = { message: ["'sheet' is not found"] };
            return ContentService.createTextOutput(JSON.stringify(response));
          }
        } else if ("sheet" in param && param.sheet != null && param.sheet != "") {
          sheet = SpreadsheetApp.open(file).getSheetByName(param.sheet);
          if (sheet == null) {
            // Помилка отримання листа з файлу
            response.state = false;
            response.error = { message: ["'sheet' is not found"] };
            return ContentService.createTextOutput(JSON.stringify(response));
          }
        } else {
          // Помилка про відсутність назви/ідентифікатора листа
          response.state = false;
          response.error = { message: ["'sheet' or 'sheetId' is missing"] };
          return ContentService.createTextOutput(JSON.stringify(response));
        }
      } catch {
        response.state = false;
        response.error = { message: ["failed load sheet from fileName"] };
        return ContentService.createTextOutput(JSON.stringify(response));
      }
    } else {
      // Помилка про відсутність назви файлу
      response.state = false;
      response.error = { message: ["'file' or 'fileId' is missing"] };
      return ContentService.createTextOutput(JSON.stringify(response));
    }
    // return ContentService.createTextOutput(JSON.stringify(response));
    if (param.action == "read") {
      // Отримання всього діапазону з даними
      let values = sheet.getDataRange().getValues();
      let columns = 0;
      let rows = [];
      for (let r = 0; r < values.length; r++) {
        if (r == 0) {
          // Визначення кількості стовпців в таблиці (до першого відсутнього значення)
          for (let v = 0; v < values[r].length; v++) {
            if (values[r][v] == "") {
              columns = v;
              break;
            }
            columns = v + 1;
          }
        } else {
          // Створення асоціативного масиву з інших рядків таблиці (ключами являються значення першого рядка)
          let oneRow = {};
          for (let v = 0; v < values[r].length; v++) {
            if (v < columns) {
              oneRow[values[0][v]] = values[r][v];
            }
          }
          rows.push(oneRow);
        }
      }
      let offset = 0,
        limit = 20;
      if ("limit" in param && typeof param.limit == "number") {
        limit = param.limit;
      }
      if ("offset" in param && typeof param.offset == "number") {
        offset = param.offset;
      }
      if ("search" in param) {
        // Пошук рядків, що відповідають заданим параметрам із search
        let approvedRows = [];
        response.mode = "search";
        for (let r = 0; r < rows.length; r++) {
          let approved = true;
          for (let key in rows[r]) {
            if (key in param.search) {
              if (rows[r][key] != param.search[key]) {
                approved = false;
                break;
              }
            }
          }
          if (approved) {
            approvedRows.push(rows[r]);
          }
        }
        response.count = approvedRows.length; // Загальна кількість рядків, що відповідають фільтру
        response.rows = approvedRows.splice(offset, limit); // Вибірка з рядків відповідно до параметрів offset та limit
      } else {
        response.mode = "not search";
        response.count = rows.length; // Загальна кількість рядків
        response.rows = rows.splice(offset, limit); // Вибірка з рядків відповідно до параметрів offset та limit
      }
    } else if (param.action == "insert") {
      // Додавання рядків внизу таблиці
      if ("row" in param) {
        // Додавання одного рядка
        if (typeof param.row != "object") {
          // Помилка, що параметр row не являється масивом
          response.state = false;
          response.error = { message: ["'row' must by an object"] };
          response.rowType = typeof param.row;
          return ContentService.createTextOutput(JSON.stringify(response));
        }
        // Підготовка списку заголовків
        let firstRow = sheet.getDataRange().getValues()[0];
        let appendRow = [];
        for (let v = 0; v < firstRow.length; v++) {
          if (firstRow[v] == "") {
            break;
          }
          if (firstRow[v] in param.row) {
            // Є значення для відповідного стовпця
            appendRow.push(param.row[firstRow[v]]);
          } else {
            // відсутнє значення для відповідного стовпця
            appendRow.push("");
          }
        }
        // Записуваний рядок у відповідь
        response.appendData = appendRow;
        // Додавання рядка в таблицю
        sheet.appendRow(appendRow);
      } else if ("rows" in param) {
        // Додавання декількох рядків
        if (typeof param.rows != "object") {
          // Помилка, що параметр rows не являється масивом
          response.state = false;
          response.error = { message: ["'rows' must by an array"] };
          response.rowType = typeof param.rows;
          return ContentService.createTextOutput(JSON.stringify(response));
        }
        // Підготовка списку заголовків таблиці
        let firstRow = sheet.getDataRange().getValues()[0];
        response.appendData = [];
        for (let r = 0; r < param.rows.length; r++) {
          // Перебір записуваних рядків
          if (typeof param.rows[r] != "object") {
            // Помилка, що окремий рядок не являється масивом
            response.state = "warning";
            response.error = { message: [`'rows[${r}]' must by an object`] };
            response.rowType = typeof param.rows[r];
            continue;
          }
          let appendRow = [];
          for (let v = 0; v < firstRow.length; v++) {
            if (firstRow[v] == "") {
              break;
            }
            if (firstRow[v] in param.rows[r]) {
              // Є значення для відповідного стовпця
              appendRow.push(param.rows[r][firstRow[v]]);
            } else {
              // відсутнє значення для відповідного стовпця
              appendRow.push("");
            }
          }
          // Записуваний рядок у відповідь
          response.appendData.push(appendRow);
          // Додавання рядка в таблицю
          sheet.appendRow(appendRow);
        }
      } else {
        response.state = false;
        response.error = { message: ["'row' or 'rows' is missing"] };
        return ContentService.createTextOutput(JSON.stringify(response));
      }
    } else if (param.action == "update") {
      // Оновлення рядка
      // Отримання всього діапазону з даними
      let values = sheet.getDataRange().getValues();
      let columns = 0;
      let rows = [];
      for (let r = 0; r < values.length; r++) {
        if (r == 0) {
          // Визначення кількості стовпців в таблиці (до першого відсутнього значення)
          for (let v = 0; v < values[r].length; v++) {
            if (values[r][v] == "") {
              columns = v;
              break;
            }
            columns = v + 1;
          }
        } else {
          // Створення асоціативного масиву з інших рядків таблиці (ключами являються значення першого рядка)
          let oneRow = {};
          for (let v = 0; v < values[r].length; v++) {
            if (v < columns) {
              oneRow[values[0][v]] = values[r][v];
            }
          }
          rows.push(oneRow);
        }
      }
      if ("search" in param) {
        // Пошук рядків, що відповідають заданим параметрам із search
        response.editedCells = [];
        for (let r = 0; r < rows.length; r++) {
          let approved = true;
          for (let key in rows[r]) {
            if (key in param.search) {
              if (rows[r][key] != param.search[key]) {
                approved = false;
                break;
              }
            }
          }
          if (approved) {
            // Рядок підходить за фільтрами, починаємо оновлення
            if ("row" in param) {
              if (typeof param.row != "object") {
                // Помилка, що параметр row не являється масивом
                response.state = false;
                response.error = { message: ["'row' must by an object"] };
                response.rowType = typeof param.row;
                return ContentService.createTextOutput(JSON.stringify(response));
              }
              // Пошук в масиві row значень з ключами, що відповідають заголовкам стовпців
              let firstRow = sheet.getDataRange().getValues()[0];
              for (let v = 0; v < firstRow.length; v++) {
                if (firstRow[v] == "") {
                  break;
                }
                if (firstRow[v] in param.row) {
                  // Є потрібне значення
                  // Повертаємо дані про оновлення клітинки у відповідь
                  response.editedCells.push({
                    row: r + 2,
                    column: v + 1,
                    oldValue: sheet.getRange(r + 2, v + 1).getValue(),
                    newValue: param.row[firstRow[v]],
                  });
                  // Оновлюємо клітинку
                  sheet.getRange(r + 2, v + 1).setRichTextValue(SpreadsheetApp.newRichTextValue().setText(param.row[firstRow[v]]).build());
                }
              }
            } else {
              // Помилка, що масив row відсутній
              response.state = false;
              response.error = { message: ["'row' is missing"] };
              return ContentService.createTextOutput(JSON.stringify(response));
            }
            // Завершуємо роботу після оновлення одного рядка
            break;
          }
        }
      } else {
        // Помилка, що параметр search відсутній
        response.state = false;
        response.error = { message: ["'search' is missing"] };
        return ContentService.createTextOutput(JSON.stringify(response));
      }
    } else if (param.action == "remove") {
      // Видалення рядків
      let values = sheet.getDataRange().getValues();
      let columns = 0;
      let rows = [];
      for (let r = 0; r < values.length; r++) {
        if (r == 0) {
          // Визначення кількості стовпців в таблиці (до першого відсутнього значення)
          for (let v = 0; v < values[r].length; v++) {
            if (values[r][v] == "") {
              columns = v;
              break;
            }
            columns = v + 1;
          }
        } else {
          // Створення асоціативного масиву з інших рядків таблиці (ключами являються значення першого рядка)
          let oneRow = {};
          for (let v = 0; v < values[r].length; v++) {
            if (v < columns) {
              oneRow[values[0][v]] = values[r][v];
            }
          }
          rows.push(oneRow);
        }
      }
      let offset = 0,
        limit = 20;
      if ("limit" in param && typeof param.limit == "number") {
        limit = param.limit;
      }
      if ("offset" in param && typeof param.offset == "number") {
        offset = param.offset;
      }
      if ("search" in param) {
        // Пошук рядків, що відповідають фільтру
        let missed = 0,
          used = 0;
        response.deletedRows = [];
        for (let r = 0; r < rows.length; r++) {
          let approved = true;
          for (let key in rows[r]) {
            if (key in param.search) {
              if (rows[r][key] != param.search[key]) {
                approved = false;
                break;
              }
            }
          }
          if (approved) {
            // Рядок повністю відповідає фільтрам
            if (missed < offset) {
              // Рядок пропускається через offset
              missed++;
              continue;
            }
            // Рядок видаляється
            sheet.deleteRow(r - used + 2);
            // Вміст видаленого рядка повертається у відповідь
            response.deletedRows.push({
              row: rows[r],
              number: r + 2,
            });
            used++;
            if (used >= limit) {
              // Досягнуто обмеження видялених рядків. Завершуємо цикл
              break;
            }
          }
        }
        // Повертаємо у відповідь кількість пропущених та видалених рядків
        response.deleted = `missed ${missed} rows, deleted ${used} rows`;
      } else {
        // Видаляємо діапазон рядків згідно параметрів offset та limit
        sheet.deleteRows(offset + 2, limit);
        response.deleted = `missed ${offset} rows, deleted ${limit} rows`;
      }
    } else if (param.action == "clear") {
      // Очищення таблиці
      if (!"approvedClear" in param || param.approvedClear !== true) {
        // Помилка про відсутність параметру для підтвердження очищення таблиці
        response.state = false;
        response.error = { message: ["Confirmation for clearing the table sheet was not canceled"] };
        return ContentService.createTextOutput(JSON.stringify(response));
      }
      // Отримання рядку заголовків таблиці
      let firstRow = sheet.getDataRange().getValues()[0];
      // Видалення всього вмісту таблиці
      response.cleaning = sheet.clear();
      // Запис в таблицю попередньо отриманого рядку заголовків
      response.columnsName = sheet.appendRow(firstRow);
    } else {
      // Помилка про невідоме значення параметру action
      response.state = false;
      response.error = { message: ["'action' is not supported"] };
      return ContentService.createTextOutput(JSON.stringify(response));
    }
    // Повернення відповіді
    return ContentService.createTextOutput(JSON.stringify(response));
  } catch {
    // Помилка виконання функції
    let response = {
      state: false,
      error: {
        message: ["failed execute function"],
      },
    };
    return ContentService.createTextOutput(JSON.stringify(response));
  }
}
