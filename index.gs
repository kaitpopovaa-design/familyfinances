function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  try {
    var params;
    if (e.postData && e.postData.contents) {
      try { params = JSON.parse(e.postData.contents); } catch(err) { params = e.parameter; }
    } else {
      params = e.parameter;
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var action = params.action || "read";

    if (action === "read") {
      var sheetProjects = getSheetFuzzy(ss, "проекты");
      var projects = sheetProjects ? getSheetDataAsJson(sheetProjects) : [];
      var sheetExpenses = getSheetFuzzy(ss, "расходы");
      var expenses = sheetExpenses ? getSheetDataAsJson(sheetExpenses) : [];
      var sheetCategories = getSheetFuzzy(ss, "категории");
      var categories = sheetCategories ? getSheetDataAsJson(sheetCategories) : [];
      return ContentService.createTextOutput(JSON.stringify({ projects: projects, expenses: expenses, categories: categories })).setMimeType(ContentService.MimeType.JSON);
    }
    else if (action === "addExpense") {
      // Строительные расходы: POST + идемпотентный RequestKey.
      // Даже если один и тот же запрос физически придет дважды, второй вызов
      // вернет уже существующую строку и НЕ создаст дубль.
      var writeLock = LockService.getScriptLock();
      writeLock.waitLock(30000);
      try {
        var sheetExpenses = getSheetFuzzy(ss, "расходы");
        if (!sheetExpenses) throw new Error('Лист "Расходы" не найден');

        var lastColumn = sheetExpenses.getLastColumn();
        if (lastColumn < 1) throw new Error('На листе "Расходы" отсутствуют заголовки');

        var lastRow = sheetExpenses.getLastRow();
        var rawHeaders = sheetExpenses.getRange(1, 1, 1, lastColumn).getValues()[0];

        // Техническая колонка не участвует в аналитике и нужна только для защиты записи от повторов.
        var requestKeyColIndex = -1;
        for (var rk = 0; rk < rawHeaders.length; rk++) {
          var rawKey = String(rawHeaders[rk] || '').toLowerCase().replace(/[\s_\-]/g, '');
          if (rawKey === 'requestkey' || rawKey === 'ключзапроса') {
            requestKeyColIndex = rk;
            break;
          }
        }
        if (requestKeyColIndex === -1) {
          requestKeyColIndex = rawHeaders.length;
          sheetExpenses.getRange(1, requestKeyColIndex + 1).setValue('RequestKey');
          rawHeaders.push('RequestKey');
          lastColumn = rawHeaders.length;
          SpreadsheetApp.flush();
        }

        var headers = rawHeaders.map(simplifyHeader);
        var requestKey = String(params.requestKey || '').trim();

        // Если такой пользовательский submit уже был сохранен, возвращаем его ID.
        // Это защищает и от повторного клика, и от повторной доставки HTTP-запроса.
        if (requestKey && lastRow > 1) {
          var requestRange = sheetExpenses.getRange(2, requestKeyColIndex + 1, lastRow - 1, 1);
          var foundCell = requestRange.createTextFinder(requestKey).matchEntireCell(true).findNext();
          if (foundCell) {
            var existingRow = foundCell.getRow();
            var existingIdColIndex = -1;
            for (var ei = 0; ei < headers.length; ei++) {
              if (headers[ei].indexOf('id') === 0 && headers[ei].length <= 3) {
                existingIdColIndex = ei;
                break;
              }
            }
            var existingId = existingIdColIndex !== -1
              ? String(sheetExpenses.getRange(existingRow, existingIdColIndex + 1).getValue()).trim()
              : String(existingRow);
            if (!existingId) existingId = String(existingRow);
            return ContentService.createTextOutput(JSON.stringify({
              success: true,
              id: existingId,
              row: existingRow,
              deduplicated: true
            })).setMimeType(ContentService.MimeType.JSON);
          }
        }

        var nextId = 1;
        var idColIndex = -1;
        for (var j = 0; j < headers.length; j++) {
          if (headers[j].indexOf('id') === 0 && headers[j].length <= 3) {
            idColIndex = j;
            break;
          }
        }

        if (lastRow > 1 && idColIndex !== -1) {
          var ids = sheetExpenses.getRange(2, idColIndex + 1, lastRow - 1, 1).getValues();
          var maxId = 0;
          for (var i = 0; i < ids.length; i++) {
            var currentId = parseInt(ids[i][0], 10);
            if (!isNaN(currentId) && currentId > maxId) maxId = currentId;
          }
          nextId = maxId + 1;
        } else if (lastRow > 1) {
          nextId = lastRow + 1;
        }

        var projectId = String(params.projectId || '').trim();
        var expenseDate = String(params.date || '').trim();
        var mainCategory = String(params.mainCategory || '').trim();
        var detailCategory = String(params.category || '').trim();
        var amountValue = parseFloat(params.amount);
        var commentValue = String(params.comment || '').trim();

        if (!projectId) throw new Error('Не передан объект строительства');
        if (!expenseDate) throw new Error('Не передана дата расхода');
        if (!mainCategory) throw new Error('Не передана общая категория');
        if (!isFinite(amountValue) || amountValue <= 0) throw new Error('Некорректная сумма расхода');

        var newRow = headers.map(function(header, idx) {
          var rawKey = String(rawHeaders[idx] || '').toLowerCase().replace(/[\s_\-]/g, '');
          if (rawKey === 'requestkey' || rawKey === 'ключзапроса') return requestKey;
          if (header.indexOf('id') === 0 && header.length <= 3) return nextId;
          if (header.indexOf('idпроекта') !== -1 || header.indexOf('проект') !== -1 || header.indexOf('объект') !== -1) return projectId;
          if (header.indexOf('дата') !== -1 || header.indexOf('date') !== -1) return expenseDate;
          if (header.indexOf('общаякатегория') !== -1 || header.indexOf('группа') !== -1 || header.indexOf('тип') !== -1) return mainCategory || 'Прочее';
          if (header.indexOf('категория') !== -1 || header.indexOf('детали') !== -1 || header.indexOf('детализация') !== -1 || header.indexOf('описание') !== -1) return detailCategory;
          if (header.indexOf('сумма') !== -1 || header.indexOf('amount') !== -1 || header.indexOf('цена') !== -1 || header.indexOf('стоимость') !== -1) return amountValue;
          if (header.indexOf('комментарий') !== -1 || header.indexOf('заметка') !== -1 || header.indexOf('примечание') !== -1) return commentValue;
          return '';
        });

        var targetRow = sheetExpenses.getLastRow() + 1;
        sheetExpenses.getRange(targetRow, 1, 1, newRow.length).setValues([newRow]);
        SpreadsheetApp.flush();

        // success возвращаем только после проверки физической записи строки.
        if (idColIndex !== -1) {
          var writtenId = String(sheetExpenses.getRange(targetRow, idColIndex + 1).getValue()).trim();
          if (writtenId !== String(nextId)) {
            throw new Error('Google Таблица не подтвердила записанный ID расхода');
          }
        }
        if (requestKey) {
          var writtenRequestKey = String(sheetExpenses.getRange(targetRow, requestKeyColIndex + 1).getValue()).trim();
          if (writtenRequestKey !== requestKey) {
            throw new Error('Google Таблица не подтвердила технический ключ записи');
          }
        }

        return ContentService.createTextOutput(JSON.stringify({
          success: true,
          id: nextId,
          row: targetRow,
          deduplicated: false
        })).setMimeType(ContentService.MimeType.JSON);
      } finally {
        writeLock.releaseLock();
      }
    }
    else if (action === "editExpense") {
      var sheetExpenses = getSheetFuzzy(ss, "расходы");
      var lastRow = sheetExpenses.getLastRow();
      if (lastRow > 1) {
        var headers = sheetExpenses.getRange(1, 1, 1, sheetExpenses.getLastColumn() || 1).getValues()[0].map(simplifyHeader);
        var idColIndex = -1;
        for (var j = 0; j < headers.length; j++) {
          if (headers[j].indexOf('id') === 0 && headers[j].length <= 3) {
            idColIndex = j;
            break;
          }
        }
        
        var rowIdx = findRowIndexById(sheetExpenses, params.id, idColIndex);
        if (rowIdx !== -1) {
          var updateRow = headers.map(function(h, idx) {
            if (h.indexOf('id') === 0 && h.length <= 3) return params.id;
            if (h.indexOf('idпроекта') !== -1 || h.indexOf('проект') !== -1 || h.indexOf('объект') !== -1) {
              return String(params.projectId).trim();
            }
            if (h.indexOf('дата') !== -1 || h.indexOf('date') !== -1) return params.date;
            
            // КРИТИЧЕСКИЙ ФИКС: Сначала проверяем длинный заголовок "общаякатегория"
            if (h.indexOf('общаякатегория') !== -1 || h.indexOf('группа') !== -1 || h.indexOf('тип') !== -1) {
              return String(params.mainCategory || "Прочее").trim();
            }
            // Только если это не общая категория, проверяем заголовок деталей
            if (h.indexOf('категория') !== -1 || h.indexOf('детали') !== -1 || h.indexOf('детализация') !== -1 || h.indexOf('описание') !== -1) {
              return String(params.category || "").trim();
            }
            
            if (h.indexOf('сумма') !== -1 || h.indexOf('amount') !== -1 || h.indexOf('цена') !== -1 || h.indexOf('стоимость') !== -1) {
              return parseFloat(params.amount) || 0;
            }
            if (h.indexOf('комментарий') !== -1 || h.indexOf('заметка') !== -1 || h.indexOf('примечание') !== -1) {
              return String(params.comment || "").trim();
            }
            return sheetExpenses.getRange(rowIdx, idx + 1).getValue(); 
          });
          sheetExpenses.getRange(rowIdx, 1, 1, headers.length).setValues([updateRow]);
          return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: "Record not found" })).setMimeType(ContentService.MimeType.JSON);
    }
    else if (action === "deleteExpense") {
      var sheetExpenses = getSheetFuzzy(ss, "расходы");
      var lastRow = sheetExpenses.getLastRow();
      if (lastRow > 1) {
        var headers = sheetExpenses.getRange(1, 1, 1, sheetExpenses.getLastColumn() || 1).getValues()[0].map(simplifyHeader);
        var idColIndex = -1;
        for (var j = 0; j < headers.length; j++) {
          if (headers[j].indexOf('id') === 0 && headers[j].length <= 3) {
            idColIndex = j;
            break;
          }
        }
        
        var rowToDel = findRowIndexById(sheetExpenses, params.id, idColIndex);
        if (rowToDel !== -1) {
          sheetExpenses.deleteRow(rowToDel);
          return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
    }
    else if (action === "deleteProject") {
      var sheetProjects = getSheetFuzzy(ss, "проекты");
      var sheetExpenses = getSheetFuzzy(ss, "расходы");
      
      var lastRowProj = sheetProjects.getLastRow();
      if (lastRowProj > 1) {
        var headersProj = sheetProjects.getRange(1, 1, 1, sheetProjects.getLastColumn() || 1).getValues()[0].map(simplifyHeader);
        var idProjColIndex = -1;
        for (var j = 0; j < headersProj.length; j++) {
          if (headersProj[j].indexOf('id') === 0 && headersProj[j].length <= 3) {
            idProjColIndex = j;
            break;
          }
        }
        
        var rowToDelProj = findRowIndexById(sheetProjects, params.id, idProjColIndex);
        if (rowToDelProj !== -1) {
          sheetProjects.deleteRow(rowToDelProj);
        }
      }
      
      var lastRowExp = sheetExpenses.getLastRow();
      if (lastRowExp > 1) {
        var headersExp = sheetExpenses.getRange(1, 1, 1, sheetExpenses.getLastColumn() || 1).getValues()[0].map(simplifyHeader);
        var projIdColIndex = -1;
        for (var j = 0; j < headersExp.length; j++) {
          var h = headersExp[j];
          if (h.indexOf('idпроекта') !== -1 || h.indexOf('проект') !== -1 || h.indexOf('объект') !== -1) {
            projIdColIndex = j;
            break;
          }
        }
        if (projIdColIndex !== -1) {
          var dataExp = sheetExpenses.getRange(2, projIdColIndex + 1, lastRowExp - 1, 1).getValues();
          for (var i = dataExp.length - 1; i >= 0; i--) {
            if (String(dataExp[i][0]).trim() === String(params.id).trim()) {
              sheetExpenses.deleteRow(i + 2);
            }
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ success: true })).setMimeType(ContentService.MimeType.JSON);
    }
    else if (action === "addProject") {
      var sheetProjects = getSheetFuzzy(ss, "проекты");
      var lastRow = sheetProjects.getLastRow();
      var numColsProj = sheetProjects.getLastColumn() || 1;
      var headersProj = sheetProjects.getRange(1, 1, 1, numColsProj).getValues()[0].map(simplifyHeader);
      var nextId = 1;
      if (lastRow > 1) {
        var idColIndexProj = -1;
        for (var j = 0; j < headersProj.length; j++) {
          if (headersProj[j].indexOf('id') === 0 && headersProj[j].length <= 3) {
            idColIndexProj = j;
            break;
          }
        }
        if (idColIndexProj !== -1) {
          var ids = sheetProjects.getRange(2, idColIndexProj + 1, lastRow - 1, 1).getValues();
          var maxId = 0;
          for (var i = 0; i < ids.length; i++) {
            var currentId = parseInt(ids[i][0], 10);
            if (!isNaN(currentId) && currentId > maxId) maxId = currentId;
          }
          nextId = maxId + 1;
        } else { nextId = lastRow + 1; }
      }
      var newRow = headersProj.map(function(header) {
        if (header.indexOf('id') === 0 && header.length <= 3) return nextId;
        if (header.indexOf('название') !== -1 || header.indexOf('name') !== -1 || header.indexOf('объект') !== -1) {
          return String(params.name || "").trim();
        }
        if (header.indexOf('бюджет') !== -1 || header.indexOf('budget') !== -1 || header.indexOf('лимит') !== -1) {
          return parseFloat(params.budget) || "";
        }
        if (header.indexOf('статус') !== -1 || header.indexOf('status') !== -1) {
          return String(params.status || "В процессе").trim();
        }
        return "";
      });
      sheetProjects.appendRow(newRow);
      return ContentService.createTextOutput(JSON.stringify({ success: true, id: nextId })).setMimeType(ContentService.MimeType.JSON);
    }
    else if (String(action).indexOf("expense") === 0) {
      return expenseHandleAction_(ss, String(action), params || {});
    }

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

function findRowIndexById(sheet, id, idColIndex) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  
  if (idColIndex !== -1) {
    var data = sheet.getRange(2, idColIndex + 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < data.length; i++) {
      var rowIdx = i + 2;
      var cellVal = String(data[i][0]).trim();
      
      // Если ячейка пустая, её ID равен номеру строки (совпадает с getSheetDataAsJson)
      var effectiveId = (cellVal === "") ? String(rowIdx) : cellVal;
      if (effectiveId === String(id).trim()) {
        return rowIdx;
      }
    }
  } else {
    // Если колонки ID нет вообще, то ID равен физическому номеру строки
    var parsedRow = parseInt(id, 10);
    if (!isNaN(parsedRow) && parsedRow >= 2 && parsedRow <= lastRow) {
      return parsedRow;
    }
  }
  return -1;
}

function getSheetFuzzy(ss, nameTarget) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().toLowerCase().replace(/[\s_]/g, '') === nameTarget.toLowerCase().replace(/[\s_]/g, '')) return sheets[i];
  }
  return null;
}

function simplifyHeader(str) {
  var s = String(str).toLowerCase().replace(/[\s_\-\(\)\"\'№]/g, '');
  var map = { 'a': 'а', 'c': 'с', 'e': 'е', 'o': 'о', 'p': 'р', 'x': 'х', 'y': 'у', 'b': 'в' };
  return s.split('').map(function(c) { return map[c] || c; }).join('');
}

function getSheetDataAsJson(sheet) {
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return []; 
  var data = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  var headers = data[0].map(function(h) { return String(h).trim(); });
  var jsonArray = [];
  
  var idColIndex = -1;
  for (var j = 0; j < headers.length; j++) {
    if (simplifyHeader(headers[j]).indexOf('id') === 0 && simplifyHeader(headers[j]).length <= 3) {
      idColIndex = j;
      break;
    }
  }
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var obj = {};
    var isEmpty = true;
    for (var j = 0; j < headers.length; j++) {
      var val = row[j];
      if (val instanceof Date) {
        val = val.getFullYear() + "-" + ("0" + (val.getMonth() + 1)).slice(-2) + "-" + ("0" + val.getDate()).slice(-2); 
      } else if (typeof val === 'string' && val.trim() !== '') {
        val = val.trim();
      }
      if (val !== "") isEmpty = false;
      obj[headers[j]] = val;
    }
    if (!isEmpty) {
      if (idColIndex === -1 || !obj[headers[idColIndex]] || String(obj[headers[idColIndex]]).trim() === "") {
        obj["ID"] = String(i + 1); 
      } else {
        obj["ID"] = String(obj[headers[idColIndex]]).trim();
      }
      jsonArray.push(obj);
    }
  }
  return jsonArray;
}

// ============================================================================
// СЕМЕЙНЫЕ РАСХОДЫ — ИЗОЛИРОВАННЫЙ МОДУЛЬ
// DETAIL = FamilyExpenses, SUMMARY = FamilyExpenseMonthlySummary.
// Для каждого месяца активен только один режим: DETAIL или SUMMARY.
// Все действия имеют префикс expense и не затрагивают недвижимость.
// ============================================================================

var EXPENSE_SHEET_OPERATIONS = "FamilyExpenses";
var EXPENSE_SHEET_CATEGORIES = "ExpenseCategories";
var EXPENSE_SHEET_IMPORT_LOG = "ExpenseImportLog";
var EXPENSE_SHEET_MONTHLY_SUMMARY = "FamilyExpenseMonthlySummary";
var EXPENSE_SHEET_MONTH_STATE = "ExpenseMonthState";

var EXPENSE_OPERATION_HEADERS = [
  "ID", "OperationKey", "Date", "Amount", "Description", "Category", "Subcategory",
  "OperationType", "Source", "Person", "Comment", "Month", "ImportId", "ImportedAt",
  "CreatedAt", "UpdatedAt", "DeletedAt", "SourceOperationId", "Occurrence"
];

var EXPENSE_CATEGORY_HEADERS = [
  "ID", "Category", "Subcategory", "Active", "Sort", "CreatedAt", "UpdatedAt"
];

// Первые 9 колонок сохранены в прежнем порядке для обратной совместимости.
var EXPENSE_IMPORT_HEADERS = [
  "ImportId", "FileHash", "Month", "FileName", "TotalRows", "ImportedRows",
  "Duplicates", "Errors", "ImportedAt", "ImportType", "StartMonth", "EndMonth", "ReplacedMonths"
];

var EXPENSE_SUMMARY_HEADERS = [
  "ID", "SummaryKey", "Month", "Category", "Amount", "Source", "Comment",
  "ImportId", "ImportedAt", "Active", "ReplacedAt", "ReplacedByImportId"
];

var EXPENSE_MONTH_STATE_HEADERS = [
  "Month", "Mode", "ImportId", "UpdatedAt"
];

var EXPENSE_ALLOWED_TYPES = [
  "Расход",
  "Доход",
  "Перевод между своими счетами",
  "Перевод внутри семьи",
  "Пополнение накоплений",
  "Снятие с накоплений",
  "Возврат покупки",
  "Инвестиции",
  "Корректировка"
];

var EXPENSE_ALLOWED_PERSONS = ["Общее", "Катя", "Коля"];

var EXPENSE_DEFAULT_CATEGORIES = [
  "Продукты",
  "Кафе и рестораны",
  "Дом и быт",
  "Одежда",
  "Красота",
  "Здоровье",
  "Автомобиль",
  "Транспорт",
  "Развлечения",
  "Путешествия",
  "Подарки",
  "Семья",
  "Животные",
  "Подписки и сервисы",
  "Связь",
  "Образование",
  "Налоги и обязательные платежи",
  "Крупные покупки",
  "Прочее"
];

var EXPENSE_HISTORICAL_CONTROL = {
  "2026-01": 231126,
  "2026-02": 137775,
  "2026-03": 183471,
  "2026-04": 123434,
  "2026-05": 265433,
  "2026-06": 166037
};
var EXPENSE_HISTORICAL_CONTROL_TOTAL = 1107276;

function expenseHandleAction_(ss, action, params) {
  expenseEnsureSheets_(ss);

  if (action === "expenseSetup") {
    return expenseJson_({
      success: true,
      sheets: [
        EXPENSE_SHEET_OPERATIONS,
        EXPENSE_SHEET_CATEGORIES,
        EXPENSE_SHEET_IMPORT_LOG,
        EXPENSE_SHEET_MONTHLY_SUMMARY,
        EXPENSE_SHEET_MONTH_STATE
      ]
    });
  }

  if (action === "expenseGetDashboard" || action === "expenseGetDashboardPeriod") {
    var currentRange;
    var previousRange;

    if (action === "expenseGetDashboard") {
      var month = expenseNormalizeMonth_(params.month) || expenseCurrentMonth_(ss);
      currentRange = expenseMonthRange_(month);
      previousRange = expenseMonthRange_(expenseShiftMonth_(month, -1));
    } else {
      var startDate = expenseNormalizeDate_(params.startDate);
      var endDate = expenseNormalizeDate_(params.endDate);
      if (!startDate || !endDate || startDate > endDate) {
        return expenseJson_({ success: false, error: "Некорректный интервал расходов." });
      }
      currentRange = { startDate: startDate, endDate: endDate };

      var prevStart = expenseNormalizeDate_(params.previousStartDate);
      var prevEnd = expenseNormalizeDate_(params.previousEndDate);
      if (!prevStart || !prevEnd || prevStart > prevEnd) {
        previousRange = expensePreviousEquivalentRange_(currentRange);
      } else {
        previousRange = { startDate: prevStart, endDate: prevEnd };
      }
    }

    var personFilter = expenseNormalizePersonFilter_(params.person);

    // PERFORMANCE: читаем FamilyExpenses / SUMMARY / MonthState / Categories ОДИН раз
    // на весь запрос дашборда. Раньше каждый current/previous/person проход заново
    // читал одни и те же листы несколько раз, из-за чего переключение месяцев
    // становилось заметно медленнее по мере роста истории.
    var dashboardSnapshot = expenseLoadDashboardSnapshot_(ss);
    var currentAll = expenseGetUnifiedRangeDataFromSnapshot_(dashboardSnapshot, currentRange.startDate, currentRange.endDate, "");
    var current = personFilter ? expenseGetUnifiedRangeDataFromSnapshot_(dashboardSnapshot, currentRange.startDate, currentRange.endDate, personFilter) : currentAll;
    var previous = expenseGetUnifiedRangeDataFromSnapshot_(dashboardSnapshot, previousRange.startDate, previousRange.endDate, personFilter);
    var categories = dashboardSnapshot.categories;

    return expenseJson_({
      success: true,
      startDate: currentRange.startDate,
      endDate: currentRange.endDate,
      previousStartDate: previousRange.startDate,
      previousEndDate: previousRange.endDate,
      summary: current.summary,
      previousSummary: previous.summary,
      categoryTotals: current.categoryTotals,
      operations: current.operations,
      summaryRows: current.summaryRows,
      monthlySeries: current.monthlySeries,
      monthStates: current.monthStates,
      partialSummaryMonths: current.partialSummaryMonths,
      categories: categories,
      allowedTypes: EXPENSE_ALLOWED_TYPES,
      allowedPersons: EXPENSE_ALLOWED_PERSONS,
      selectedPerson: personFilter,
      personTotals: expensePersonExpenseTotals_(currentAll.operations)
    });
  }

  if (action === "expenseGetOperations" || action === "expenseSearchOperations") {
    var filters = {
      month: expenseNormalizeMonth_(params.month),
      startDate: expenseNormalizeDate_(params.startDate),
      endDate: expenseNormalizeDate_(params.endDate),
      category: expenseClean_(params.category),
      subcategory: expenseClean_(params.subcategory),
      operationType: expenseClean_(params.operationType),
      source: expenseClean_(params.source),
      person: expenseClean_(params.person),
      q: expenseClean_(params.q)
    };
    return expenseJson_({
      success: true,
      operations: expenseGetOperationsData_(ss, filters)
    });
  }

  if (action === "expenseGetStats") {
    var statsMonth = expenseNormalizeMonth_(params.month) || expenseCurrentMonth_(ss);
    var statsRange = expenseMonthRange_(statsMonth);
    var statsData = expenseGetUnifiedRangeData_(ss, statsRange.startDate, statsRange.endDate);
    return expenseJson_({
      success: true,
      month: statsMonth,
      summary: statsData.summary,
      categoryTotals: statsData.categoryTotals,
      mode: expenseGetMonthMode_(ss, statsMonth)
    });
  }

  if (action === "expenseCompareMonths") {
    var firstMonth = expenseNormalizeMonth_(params.month) || expenseCurrentMonth_(ss);
    var secondMonth = expenseNormalizeMonth_(params.compareMonth) || expenseShiftMonth_(firstMonth, -1);
    var firstRange = expenseMonthRange_(firstMonth);
    var secondRange = expenseMonthRange_(secondMonth);
    return expenseJson_({
      success: true,
      month: firstMonth,
      compareMonth: secondMonth,
      summary: expenseGetUnifiedRangeData_(ss, firstRange.startDate, firstRange.endDate).summary,
      compareSummary: expenseGetUnifiedRangeData_(ss, secondRange.startDate, secondRange.endDate).summary
    });
  }

  if (action === "expenseGetMonthStates") {
    return expenseJson_({ success: true, states: expenseGetMonthStates_(ss) });
  }

  if (action === "expenseGetCategories") {
    return expenseJson_({
      success: true,
      categories: expenseGetCategoriesData_(ss, params.includeHidden === true || String(params.includeHidden) === "true")
    });
  }

  if (action === "expenseSaveCategory") {
    return expenseJson_(expenseSaveCategory_(ss, params));
  }

  if (action === "expenseUpdateCategory") {
    return expenseJson_(expenseUpdateCategory_(ss, params));
  }

  if (action === "expenseSaveOperation") {
    return expenseJson_(expenseSaveOperation_(ss, params));
  }

  if (action === "expenseUpdateOperation") {
    return expenseJson_(expenseUpdateOperation_(ss, params));
  }

  if (action === "expenseDeleteOperation") {
    return expenseJson_(expenseDeleteOperation_(ss, params.id));
  }

  if (action === "expenseImportPreview") {
    return expenseJson_(expenseImportPreview_(ss, params));
  }

  if (action === "expenseImportConfirm") {
    return expenseJson_(expenseImportConfirm_(ss, params));
  }

  if (action === "expenseHistoricalImportPreview") {
    return expenseJson_(expenseHistoricalImportPreview_(ss, params));
  }

  if (action === "expenseHistoricalImportConfirm") {
    return expenseJson_(expenseHistoricalImportConfirm_(ss, params));
  }

  return expenseJson_({ success: false, error: "Неизвестное действие модуля расходов: " + action });
}

function expenseEnsureSheets_(ss) {
  var operations = ss.getSheetByName(EXPENSE_SHEET_OPERATIONS);
  if (!operations) {
    operations = ss.insertSheet(EXPENSE_SHEET_OPERATIONS);
    operations.getRange(1, 1, 1, EXPENSE_OPERATION_HEADERS.length).setValues([EXPENSE_OPERATION_HEADERS]);
    operations.setFrozenRows(1);
  } else {
    expenseEnsureHeaders_(operations, EXPENSE_OPERATION_HEADERS);
  }

  var categories = ss.getSheetByName(EXPENSE_SHEET_CATEGORIES);
  if (!categories) {
    categories = ss.insertSheet(EXPENSE_SHEET_CATEGORIES);
    categories.getRange(1, 1, 1, EXPENSE_CATEGORY_HEADERS.length).setValues([EXPENSE_CATEGORY_HEADERS]);
    categories.setFrozenRows(1);
  } else {
    expenseEnsureHeaders_(categories, EXPENSE_CATEGORY_HEADERS);
  }

  var imports = ss.getSheetByName(EXPENSE_SHEET_IMPORT_LOG);
  if (!imports) {
    imports = ss.insertSheet(EXPENSE_SHEET_IMPORT_LOG);
    imports.getRange(1, 1, 1, EXPENSE_IMPORT_HEADERS.length).setValues([EXPENSE_IMPORT_HEADERS]);
    imports.setFrozenRows(1);
  } else {
    expenseEnsureHeaders_(imports, EXPENSE_IMPORT_HEADERS);
  }

  var monthlySummary = ss.getSheetByName(EXPENSE_SHEET_MONTHLY_SUMMARY);
  if (!monthlySummary) {
    monthlySummary = ss.insertSheet(EXPENSE_SHEET_MONTHLY_SUMMARY);
    monthlySummary.getRange(1, 1, 1, EXPENSE_SUMMARY_HEADERS.length).setValues([EXPENSE_SUMMARY_HEADERS]);
    monthlySummary.setFrozenRows(1);
  } else {
    expenseEnsureHeaders_(monthlySummary, EXPENSE_SUMMARY_HEADERS);
  }

  var monthState = ss.getSheetByName(EXPENSE_SHEET_MONTH_STATE);
  if (!monthState) {
    monthState = ss.insertSheet(EXPENSE_SHEET_MONTH_STATE);
    monthState.getRange(1, 1, 1, EXPENSE_MONTH_STATE_HEADERS.length).setValues([EXPENSE_MONTH_STATE_HEADERS]);
    monthState.setFrozenRows(1);
  } else {
    expenseEnsureHeaders_(monthState, EXPENSE_MONTH_STATE_HEADERS);
  }

  if (categories.getLastRow() < 2) {
    var now = expenseNowIso_();
    var rows = [];
    for (var i = 0; i < EXPENSE_DEFAULT_CATEGORIES.length; i++) {
      rows.push([
        Utilities.getUuid(),
        EXPENSE_DEFAULT_CATEGORIES[i],
        "",
        true,
        i + 1,
        now,
        now
      ]);
    }
    if (rows.length) categories.getRange(2, 1, rows.length, EXPENSE_CATEGORY_HEADERS.length).setValues(rows);
  }

  // Google Sheets иногда автоматически превращает YYYY-MM в дату.
  // Приводим технические колонки Month обратно к стабильному текстовому формату.
  expenseEnsureMonthTextColumn_(operations, "Month");
  expenseEnsureMonthTextColumn_(monthlySummary, "Month");
  expenseEnsureMonthTextColumn_(monthState, "Month");
}

function expenseEnsureHeaders_(sheet, requiredHeaders) {
  var lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    return;
  }

  var current = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function(v) { return String(v).trim(); });
  var missing = [];
  for (var i = 0; i < requiredHeaders.length; i++) {
    if (current.indexOf(requiredHeaders[i]) === -1) missing.push(requiredHeaders[i]);
  }
  if (missing.length) sheet.getRange(1, lastColumn + 1, 1, missing.length).setValues([missing]);
}

function expenseEnsureMonthTextColumn_(sheet, fieldName) {
  if (!sheet) return;
  var headers = expenseGetHeaders_(sheet);
  var col = headers.indexOf(fieldName || "Month");
  if (col === -1) return;

  var rowCount = Math.max(0, sheet.getLastRow() - 1);
  var range = rowCount > 0 ? sheet.getRange(2, col + 1, rowCount, 1) : null;

  // '@' = plain text. Ставим формат ДО записи, чтобы Google не превратил YYYY-MM обратно в дату.
  sheet.getRange(2, col + 1, Math.max(1, sheet.getMaxRows() - 1), 1).setNumberFormat('@');

  if (range) {
    var values = range.getValues();
    var changed = false;
    for (var i = 0; i < values.length; i++) {
      if (values[i][0] === "" || values[i][0] === null) continue;
      var normalized = expenseNormalizeMonth_(values[i][0]);
      if (normalized && String(values[i][0]).trim() !== normalized) {
        values[i][0] = normalized;
        changed = true;
      }
    }
    if (changed) range.setValues(values);
  }
}

function expenseGetOperationsData_(ss, filters) {
  filters = filters || {};
  var rows = expenseSheetToObjects_(ss.getSheetByName(EXPENSE_SHEET_OPERATIONS));
  var modeContext = expenseBuildModeContext_(ss);
  var month = expenseNormalizeMonth_(filters.month);
  var startDate = expenseNormalizeDate_(filters.startDate);
  var endDate = expenseNormalizeDate_(filters.endDate);
  var category = expenseNorm_(filters.category);
  var subcategory = expenseNorm_(filters.subcategory);
  var operationType = expenseNorm_(filters.operationType);
  var source = expenseNorm_(filters.source);
  var person = expenseNormalizePersonFilter_(filters.person);
  var q = expenseNorm_(filters.q);

  var result = rows.filter(function(row) {
    if (expenseClean_(row.DeletedAt)) return false;
    var rowMonth = expenseNormalizeMonth_(row.Month) || expenseNormalizeMonth_(expenseClean_(row.Date).slice(0, 7));
    if (expenseResolveMonthModeFromContext_(modeContext, rowMonth) !== "DETAIL") return false;
    if (month && rowMonth !== month) return false;
    var date = expenseClean_(row.Date);
    if (startDate && date < startDate) return false;
    if (endDate && date > endDate) return false;
    if (category && expenseNorm_(row.Category) !== category) return false;
    if (subcategory && expenseNorm_(row.Subcategory) !== subcategory) return false;
    if (operationType && expenseNorm_(row.OperationType) !== operationType) return false;
    if (source && expenseNorm_(row.Source) !== source) return false;
    if (person && expenseNormalizePerson_(row.Person) !== person) return false;
    if (q) {
      var haystack = expenseNorm_([
        row.Description, row.Category, row.Subcategory, row.Source, row.Person, row.Comment, row.OperationType
      ].join(" "));
      if (haystack.indexOf(q) === -1) return false;
    }
    return true;
  });

  result.sort(function(a, b) {
    var da = expenseClean_(a.Date);
    var db = expenseClean_(b.Date);
    if (da === db) return String(b.CreatedAt || "").localeCompare(String(a.CreatedAt || ""));
    return db.localeCompare(da);
  });

  return result;
}

function expenseGetCategoriesData_(ss, includeHidden) {
  var rows = expenseSheetToObjects_(ss.getSheetByName(EXPENSE_SHEET_CATEGORIES));
  var result = rows.filter(function(row) {
    if (includeHidden) return true;
    return expenseBool_(row.Active, true);
  });

  result.sort(function(a, b) {
    var sa = parseFloat(a.Sort) || 9999;
    var sb = parseFloat(b.Sort) || 9999;
    if (sa !== sb) return sa - sb;
    var ca = expenseClean_(a.Category);
    var cb = expenseClean_(b.Category);
    if (ca !== cb) return ca.localeCompare(cb, "ru");
    return expenseClean_(a.Subcategory).localeCompare(expenseClean_(b.Subcategory), "ru");
  });

  return result;
}

function expenseSaveCategory_(ss, params) {
  var category = expenseClean_(params.category);
  var subcategory = expenseClean_(params.subcategory);
  if (!category) return { success: false, error: "Укажите название категории." };

  var sheet = ss.getSheetByName(EXPENSE_SHEET_CATEGORIES);
  var rows = expenseSheetToObjects_(sheet);
  var targetCat = expenseNorm_(category);
  var targetSub = expenseNorm_(subcategory);

  for (var i = 0; i < rows.length; i++) {
    if (expenseNorm_(rows[i].Category) === targetCat && expenseNorm_(rows[i].Subcategory) === targetSub) {
      return { success: false, error: "Такая категория/подкатегория уже существует." };
    }
  }

  var now = expenseNowIso_();
  var maxSort = 0;
  for (var j = 0; j < rows.length; j++) maxSort = Math.max(maxSort, parseFloat(rows[j].Sort) || 0);

  var newObj = {
    ID: Utilities.getUuid(),
    Category: category,
    Subcategory: subcategory,
    Active: params.active === false ? false : true,
    Sort: parseFloat(params.sort) || (maxSort + 1),
    CreatedAt: now,
    UpdatedAt: now
  };

  expenseAppendObject_(sheet, newObj);
  return { success: true, category: newObj };
}

function expenseUpdateCategory_(ss, params) {
  var id = expenseClean_(params.id);
  if (!id) return { success: false, error: "Не указан ID категории." };

  var sheet = ss.getSheetByName(EXPENSE_SHEET_CATEGORIES);
  var found = expenseFindRowByField_(sheet, "ID", id);
  if (!found) return { success: false, error: "Категория не найдена." };

  var oldObj = found.object;
  var updated = {
    ID: oldObj.ID,
    Category: params.category !== undefined ? expenseClean_(params.category) : expenseClean_(oldObj.Category),
    Subcategory: params.subcategory !== undefined ? expenseClean_(params.subcategory) : expenseClean_(oldObj.Subcategory),
    Active: params.active !== undefined ? expenseBool_(params.active, true) : expenseBool_(oldObj.Active, true),
    Sort: params.sort !== undefined ? (parseFloat(params.sort) || oldObj.Sort || 9999) : oldObj.Sort,
    CreatedAt: oldObj.CreatedAt || expenseNowIso_(),
    UpdatedAt: expenseNowIso_()
  };

  if (!updated.Category) return { success: false, error: "Название категории не может быть пустым." };

  expenseWriteObjectToRow_(sheet, found.rowIndex, updated);
  return { success: true, category: updated };
}

function expenseSaveOperation_(ss, params) {
  var categories = expenseGetCategoriesData_(ss, false);
  var prepared = expensePrepareSingleOperation_(params, categories, true);
  if (prepared.errors.length) return { success: false, error: prepared.errors.join(" "), errors: prepared.errors };

  var month = prepared.operation.Month;
  var mode = expenseGetMonthMode_(ss, month);
  if (mode === "SUMMARY") {
    return {
      success: false,
      error: "За " + month + " активна историческая сводка SUMMARY. Нельзя добавить детальную операцию поверх неё. Сначала замените SUMMARY детальным импортом."
    };
  }
  if (mode === "CONFLICT") return { success: false, error: "В месяце обнаружен конфликт SUMMARY/DETAIL. Сначала исправьте режим месяца." };

  var sheet = ss.getSheetByName(EXPENSE_SHEET_OPERATIONS);
  var existingRows = expenseSheetToObjects_(sheet);

  if (!prepared.operation.Occurrence || parseInt(prepared.operation.Occurrence, 10) < 1) {
    prepared.operation.Occurrence = expenseNextOccurrence_(existingRows, prepared.operation);
    prepared.operation.OperationKey = expenseBuildOperationKey_(prepared.operation);
  }

  for (var i = 0; i < existingRows.length; i++) {
    if (expenseClean_(existingRows[i].OperationKey) === prepared.operation.OperationKey) {
      prepared.operation.Occurrence = expenseNextOccurrence_(existingRows, prepared.operation);
      prepared.operation.OperationKey = expenseBuildOperationKey_(prepared.operation);
      break;
    }
  }

  var now = expenseNowIso_();
  var obj = prepared.operation;
  obj.ID = Utilities.getUuid();
  obj.ImportId = "MANUAL";
  obj.ImportedAt = now;
  obj.CreatedAt = now;
  obj.UpdatedAt = now;
  obj.DeletedAt = "";

  expenseAppendObject_(sheet, obj);
  expenseSetMonthState_(ss, month, "DETAIL", "MANUAL");
  return { success: true, operation: obj };
}

function expenseUpdateOperation_(ss, params) {
  var id = expenseClean_(params.id);
  if (!id) return { success: false, error: "Не указан ID операции." };

  var sheet = ss.getSheetByName(EXPENSE_SHEET_OPERATIONS);
  var found = expenseFindRowByField_(sheet, "ID", id);
  if (!found) return { success: false, error: "Операция не найдена." };

  var categories = expenseGetCategoriesData_(ss, false);
  var merged = {};
  var oldObj = found.object;
  var keys = ["Date", "Amount", "Description", "Category", "Subcategory", "OperationType", "Source", "Person", "Comment", "SourceOperationId", "Occurrence"];
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    var paramKey = k.charAt(0).toLowerCase() + k.slice(1);
    merged[paramKey] = params[paramKey] !== undefined ? params[paramKey] : oldObj[k];
  }

  var prepared = expensePrepareSingleOperation_(merged, categories, false);
  if (prepared.errors.length) return { success: false, error: prepared.errors.join(" "), errors: prepared.errors };

  var newMonth = prepared.operation.Month;
  var newMode = expenseGetMonthMode_(ss, newMonth);
  if (newMode === "SUMMARY") {
    return { success: false, error: "За " + newMonth + " активна SUMMARY-сводка. Перенести туда детальную операцию нельзя." };
  }
  if (newMode === "CONFLICT") return { success: false, error: "В целевом месяце обнаружен конфликт SUMMARY/DETAIL." };

  var updated = prepared.operation;
  updated.ID = oldObj.ID;
  updated.ImportId = oldObj.ImportId || "MANUAL";
  updated.ImportedAt = oldObj.ImportedAt || oldObj.CreatedAt || expenseNowIso_();
  updated.CreatedAt = oldObj.CreatedAt || expenseNowIso_();
  updated.UpdatedAt = expenseNowIso_();
  updated.DeletedAt = oldObj.DeletedAt || "";

  var existingRows = expenseSheetToObjects_(sheet);
  for (var j = 0; j < existingRows.length; j++) {
    if (String(existingRows[j].ID) === String(id)) continue;
    if (expenseClean_(existingRows[j].OperationKey) === updated.OperationKey) {
      return { success: false, error: "После изменения операция совпадает с уже существующей записью." };
    }
  }

  expenseWriteObjectToRow_(sheet, found.rowIndex, updated);
  expenseSetMonthState_(ss, newMonth, "DETAIL", updated.ImportId || "MANUAL");
  return { success: true, operation: updated };
}

function expenseDeleteOperation_(ss, id) {
  id = expenseClean_(id);
  if (!id) return { success: false, error: "Не указан ID операции." };

  var sheet = ss.getSheetByName(EXPENSE_SHEET_OPERATIONS);
  var found = expenseFindRowByField_(sheet, "ID", id);
  if (!found) return { success: false, error: "Операция не найдена." };

  var headers = expenseGetHeaders_(sheet);
  var deletedCol = headers.indexOf("DeletedAt");
  var updatedCol = headers.indexOf("UpdatedAt");
  var now = expenseNowIso_();
  if (deletedCol !== -1) sheet.getRange(found.rowIndex, deletedCol + 1).setValue(now);
  if (updatedCol !== -1) sheet.getRange(found.rowIndex, updatedCol + 1).setValue(now);

  return { success: true };
}

function expenseImportPreview_(ss, params) {
  var rows = Array.isArray(params.rows) ? params.rows : [];
  var fileHash = expenseClean_(params.fileHash);
  var fileName = expenseClean_(params.fileName);
  var categories = expenseGetCategoriesData_(ss, false);
  var existingRows = expenseSheetToObjects_(ss.getSheetByName(EXPENSE_SHEET_OPERATIONS));
  var existingKeys = {};
  for (var i = 0; i < existingRows.length; i++) {
    var ek = expenseClean_(existingRows[i].OperationKey);
    if (ek) existingKeys[ek] = true;
  }

  var fileAlreadyImported = expenseWasFileImported_(ss, fileHash);
  var prepared = expensePrepareImportRows_(rows, categories);
  var knownCategoryMap = {};
  for (var kc = 0; kc < categories.length; kc++) {
    var knownName = expenseClean_(categories[kc].Category);
    if (knownName) knownCategoryMap[expenseNorm_(knownName)] = true;
  }
  var newCategoryMap = {};
  var seenKeys = {};
  var items = [];
  var newRows = 0;
  var duplicateRows = 0;
  var errorRows = 0;
  var conflictRows = 0;
  var validOperations = [];
  var monthsMap = {};
  var conflictMap = {};
  var modeContext = expenseBuildModeContext_(ss);

  for (var p = 0; p < prepared.length; p++) {
    var item = prepared[p];
    var status = "new";
    var duplicateReason = "";
    var op = item.operation;

    if (item.errors.length) {
      status = "error";
      errorRows++;
    } else {
      monthsMap[op.Month] = true;
      if (op.Category && !knownCategoryMap[expenseNorm_(op.Category)]) newCategoryMap[op.Category] = true;
      var mode = expenseResolveMonthModeFromContext_(modeContext, op.Month);
      if (mode === "SUMMARY") {
        conflictMap[op.Month] = "SUMMARY";
      } else if (mode === "CONFLICT") {
        item.errors.push("В месяце уже есть конфликт SUMMARY/DETAIL.");
        status = "error";
        errorRows++;
      }

      if (status !== "error") {
        var key = op.OperationKey;
        if (existingKeys[key]) {
          status = "duplicate";
          duplicateReason = "Операция уже есть в Google Таблице";
          duplicateRows++;
        } else if (seenKeys[key]) {
          status = "duplicate";
          duplicateReason = "Повтор внутри загружаемого файла";
          duplicateRows++;
        } else if (mode === "SUMMARY") {
          status = "conflict";
          conflictRows++;
          seenKeys[key] = true;
          newRows++;
        } else {
          seenKeys[key] = true;
          newRows++;
        }
        validOperations.push(op);
      }
    }

    items.push({
      rowNumber: item.rowNumber,
      status: status,
      duplicateReason: duplicateReason,
      errors: item.errors,
      operation: op
    });
  }

  var conflictMonths = Object.keys(conflictMap).sort().map(function(month) {
    return {
      month: month,
      mode: conflictMap[month],
      message: "За " + month + " уже есть историческая сводка. Детальный импорт приведет к задвоению расходов."
    };
  });

  return {
    success: true,
    importType: "DETAIL",
    fileAlreadyImported: fileAlreadyImported,
    fileHash: fileHash,
    fileName: fileName,
    totalRows: rows.length,
    newRows: newRows,
    duplicates: duplicateRows,
    errors: errorRows,
    conflicts: conflictRows,
    conflictMonths: conflictMonths,
    months: Object.keys(monthsMap).sort(),
    newCategories: Object.keys(newCategoryMap).sort(function(a,b){ return a.localeCompare(b, "ru"); }),
    summary: expenseBuildSummary_(validOperations),
    items: items
  };
}

function expenseImportConfirm_(ss, params) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var rows = Array.isArray(params.rows) ? params.rows : [];
    var fileHash = expenseClean_(params.fileHash);
    var fileName = expenseClean_(params.fileName);
    var importId = expenseClean_(params.importId) || ("DET-" + Utilities.getUuid());
    var replaceSummary = params.replaceSummary === true || String(params.replaceSummary) === "true";
    var categories = expenseGetCategoriesData_(ss, false);
    var prepared = expensePrepareImportRows_(rows, categories);

    var sheet = ss.getSheetByName(EXPENSE_SHEET_OPERATIONS);
    var existingRows = expenseSheetToObjects_(sheet);
    var existingKeys = {};
    for (var i = 0; i < existingRows.length; i++) {
      var oldKey = expenseClean_(existingRows[i].OperationKey);
      if (oldKey) existingKeys[oldKey] = true;
    }

    var modeContext = expenseBuildModeContext_(ss);
    var conflictMap = {};
    var batchMonths = {};
    for (var c = 0; c < prepared.length; c++) {
      if (prepared[c].errors.length) continue;
      var cm = prepared[c].operation.Month;
      batchMonths[cm] = true;
      var mode = expenseResolveMonthModeFromContext_(modeContext, cm);
      if (mode === "SUMMARY") conflictMap[cm] = true;
      if (mode === "CONFLICT") return { success: false, error: "В месяце " + cm + " уже есть конфликт SUMMARY/DETAIL." };
    }

    var conflictMonths = Object.keys(conflictMap).sort();
    if (conflictMonths.length && !replaceSummary) {
      return {
        success: false,
        requiresReplacement: true,
        conflictMonths: conflictMonths,
        error: "За частью месяцев уже есть SUMMARY. Подтвердите замену сводных данных детальными операциями."
      };
    }

    var seenBatch = {};
    var toWrite = [];
    var duplicates = 0;
    var errors = 0;
    var monthWriteCounts = {};
    var now = expenseNowIso_();

    for (var p = 0; p < prepared.length; p++) {
      var item = prepared[p];
      if (item.errors.length) {
        errors++;
        continue;
      }

      var op = item.operation;
      if (existingKeys[op.OperationKey] || seenBatch[op.OperationKey]) {
        duplicates++;
        continue;
      }

      seenBatch[op.OperationKey] = true;
      existingKeys[op.OperationKey] = true;
      op.ID = Utilities.getUuid();
      op.ImportId = importId;
      op.ImportedAt = now;
      op.CreatedAt = now;
      op.UpdatedAt = now;
      op.DeletedAt = "";

      toWrite.push(expenseObjectToRow_(op, EXPENSE_OPERATION_HEADERS));
      monthWriteCounts[op.Month] = (monthWriteCounts[op.Month] || 0) + 1;
    }

    // SUMMARY отключаем только для тех конфликтных месяцев, в которые реально пишется DETAIL.
    var replacedMonths = [];
    for (var r = 0; r < conflictMonths.length; r++) {
      var conflictMonth = conflictMonths[r];
      if ((monthWriteCounts[conflictMonth] || 0) > 0) {
        expenseDeactivateSummaryMonths_(ss, [conflictMonth], importId);
        replacedMonths.push(conflictMonth);
      }
    }

    if (toWrite.length) {
      expenseEnsureCategoriesFromPrepared_(ss, prepared);
      sheet.getRange(sheet.getLastRow() + 1, 1, toWrite.length, EXPENSE_OPERATION_HEADERS.length).setValues(toWrite);
    }

    var writtenMonths = Object.keys(monthWriteCounts);
    for (var w = 0; w < writtenMonths.length; w++) {
      expenseSetMonthState_(ss, writtenMonths[w], "DETAIL", importId);
    }

    var allMonths = Object.keys(batchMonths).sort();
    expenseLogImport_(ss, {
      ImportId: importId,
      FileHash: fileHash,
      Month: allMonths.length === 1 ? allMonths[0] : "",
      FileName: fileName,
      TotalRows: rows.length,
      ImportedRows: toWrite.length,
      Duplicates: duplicates,
      Errors: errors,
      ImportedAt: now,
      ImportType: "DETAIL",
      StartMonth: allMonths.length ? allMonths[0] : "",
      EndMonth: allMonths.length ? allMonths[allMonths.length - 1] : "",
      ReplacedMonths: replacedMonths.join(",")
    });

    return {
      success: true,
      importId: importId,
      importedRows: toWrite.length,
      duplicates: duplicates,
      errors: errors,
      replacedMonths: replacedMonths,
      months: allMonths
    };
  } finally {
    lock.releaseLock();
  }
}

function expenseEnsureCategoriesFromPrepared_(ss, prepared) {
  var sheet = ss.getSheetByName(EXPENSE_SHEET_CATEGORIES);
  var existing = expenseSheetToObjects_(sheet);
  var seen = {};
  for (var i = 0; i < existing.length; i++) {
    var ec = expenseClean_(existing[i].Category);
    var es = expenseClean_(existing[i].Subcategory);
    if (ec) seen[expenseNorm_(ec) + "|" + expenseNorm_(es)] = true;
  }
  var now = expenseNowIso_();
  var maxSort = existing.reduce(function(max, row) { return Math.max(max, parseFloat(row.Sort) || 0); }, 0);
  var rows = [];
  for (var p = 0; p < prepared.length; p++) {
    if (prepared[p].errors && prepared[p].errors.length) continue;
    var op = prepared[p].operation || {};
    var cat = expenseClean_(op.Category);
    var sub = expenseClean_(op.Subcategory);
    if (!cat) continue;
    var key = expenseNorm_(cat) + "|" + expenseNorm_(sub);
    if (seen[key]) continue;
    seen[key] = true;
    maxSort += 1;
    rows.push([Utilities.getUuid(), cat, sub, true, maxSort, now, now]);
  }
  if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, EXPENSE_CATEGORY_HEADERS.length).setValues(rows);
}

function expenseHistoricalImportPreview_(ss, params) {
  var rows = Array.isArray(params.rows) ? params.rows : [];
  var fileHash = expenseClean_(params.fileHash);
  var fileName = expenseClean_(params.fileName);
  var prepared = expensePrepareHistoricalRows_(rows);
  var existingSummary = expenseGetActiveSummaryRows_(ss);
  var existingKeys = {};
  var existingByMonth = {};
  for (var i = 0; i < existingSummary.length; i++) {
    var row = existingSummary[i];
    var key = expenseClean_(row.SummaryKey) || expenseBuildSummaryKey_(row);
    existingKeys[key] = true;
    var month = expenseNormalizeMonth_(row.Month);
    if (!month) continue;
    if (!existingByMonth[month]) existingByMonth[month] = {};
    existingByMonth[month][key] = true;
  }

  var incomingByMonth = {};
  var monthsMap = {};
  var monthTotals = {};
  var modeContext = expenseBuildModeContext_(ss);
  var errorRows = 0;
  var duplicateRows = 0;
  var conflictRows = 0;
  var newRows = 0;
  var items = [];

  for (var p = 0; p < prepared.length; p++) {
    var item = prepared[p];
    var status = "new";
    var duplicateReason = "";
    var op = item.summary;

    if (item.errors.length) {
      status = "error";
      errorRows++;
    } else {
      monthsMap[op.Month] = true;
      monthTotals[op.Month] = (monthTotals[op.Month] || 0) + (parseFloat(op.Amount) || 0);
      if (!incomingByMonth[op.Month]) incomingByMonth[op.Month] = {};
      incomingByMonth[op.Month][op.SummaryKey] = true;
    }

    items.push({
      rowNumber: item.rowNumber,
      status: status,
      duplicateReason: duplicateReason,
      errors: item.errors,
      summary: op
    });
  }

  var conflictMap = {};
  var months = Object.keys(monthsMap).sort();
  for (var m = 0; m < months.length; m++) {
    var monthName = months[m];
    var mode = expenseResolveMonthModeFromContext_(modeContext, monthName);
    if (mode === "DETAIL" || mode === "CONFLICT") {
      conflictMap[monthName] = mode;
    } else if (mode === "SUMMARY") {
      var incomingKeys = Object.keys(incomingByMonth[monthName] || {}).sort();
      var oldKeys = Object.keys(existingByMonth[monthName] || {}).sort();
      if (incomingKeys.join("|") !== oldKeys.join("|")) conflictMap[monthName] = "SUMMARY";
    }
  }

  for (var x = 0; x < items.length; x++) {
    var histItem = items[x];
    if (histItem.status === "error") continue;
    var summary = histItem.summary;
    if (conflictMap[summary.Month]) {
      histItem.status = "conflict";
      conflictRows++;
      newRows++;
    } else if (existingKeys[summary.SummaryKey]) {
      histItem.status = "duplicate";
      histItem.duplicateReason = "Такая сводная строка уже загружена";
      duplicateRows++;
    } else {
      newRows++;
    }
  }

  var control = expenseHistoricalControl_(monthTotals);
  var conflictMonths = Object.keys(conflictMap).sort().map(function(month) {
    var mode = conflictMap[month];
    return {
      month: month,
      mode: mode,
      message: mode === "DETAIL"
        ? "За " + month + " уже есть детальные операции DETAIL."
        : (mode === "SUMMARY" ? "За " + month + " уже есть другая SUMMARY-сводка." : "За " + month + " уже есть конфликт данных.")
    };
  });

  return {
    success: true,
    importType: "SUMMARY",
    fileAlreadyImported: expenseWasFileImported_(ss, fileHash),
    fileHash: fileHash,
    fileName: fileName,
    totalRows: rows.length,
    newRows: newRows,
    duplicates: duplicateRows,
    errors: errorRows,
    conflicts: conflictRows,
    conflictMonths: conflictMonths,
    months: months,
    monthTotals: expenseMonthTotalsArray_(monthTotals),
    grandTotal: expenseSumObjectValues_(monthTotals),
    control: control,
    items: items
  };
}

function expenseHistoricalImportConfirm_(ss, params) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var rows = Array.isArray(params.rows) ? params.rows : [];
    var fileHash = expenseClean_(params.fileHash);
    var fileName = expenseClean_(params.fileName);
    var replaceConflicts = params.replaceConflicts === true || String(params.replaceConflicts) === "true";
    var importId = expenseClean_(params.importId) || ("SUM-" + Utilities.getUuid());
    var prepared = expensePrepareHistoricalRows_(rows);
    var monthTotals = {};
    var validByMonth = {};
    var errorCount = 0;

    for (var i = 0; i < prepared.length; i++) {
      var item = prepared[i];
      if (item.errors.length) {
        errorCount++;
        continue;
      }
      var summary = item.summary;
      monthTotals[summary.Month] = (monthTotals[summary.Month] || 0) + (parseFloat(summary.Amount) || 0);
      if (!validByMonth[summary.Month]) validByMonth[summary.Month] = [];
      validByMonth[summary.Month].push(summary);
    }

    if (errorCount) return { success: false, error: "В историческом файле есть строки с ошибками. Исправьте файл перед импортом.", errors: errorCount };

    var control = expenseHistoricalControl_(monthTotals);
    if (control.applies && !control.ok) {
      return { success: false, error: "Контрольные суммы января–июня 2026 не совпадают. Импорт остановлен.", control: control };
    }

    var existingSummary = expenseGetActiveSummaryRows_(ss);
    var existingByMonth = {};
    var existingKeys = {};
    for (var e = 0; e < existingSummary.length; e++) {
      var existing = existingSummary[e];
      var existingKey = expenseClean_(existing.SummaryKey) || expenseBuildSummaryKey_(existing);
      existingKeys[existingKey] = true;
      var em = expenseNormalizeMonth_(existing.Month);
      if (!em) continue;
      if (!existingByMonth[em]) existingByMonth[em] = {};
      existingByMonth[em][existingKey] = true;
    }

    var modeContext = expenseBuildModeContext_(ss);
    var months = Object.keys(validByMonth).sort();
    var conflictMap = {};
    for (var m = 0; m < months.length; m++) {
      var month = months[m];
      var mode = expenseResolveMonthModeFromContext_(modeContext, month);
      if (mode === "DETAIL" || mode === "CONFLICT") {
        conflictMap[month] = mode;
      } else if (mode === "SUMMARY") {
        var incomingSet = {};
        for (var q = 0; q < validByMonth[month].length; q++) incomingSet[validByMonth[month][q].SummaryKey] = true;
        var incomingKeys = Object.keys(incomingSet).sort();
        var oldKeys = Object.keys(existingByMonth[month] || {}).sort();
        if (incomingKeys.join("|") !== oldKeys.join("|")) conflictMap[month] = "SUMMARY";
      }
    }

    var conflictMonths = Object.keys(conflictMap).sort();
    if (conflictMonths.length && !replaceConflicts) {
      return {
        success: false,
        requiresReplacement: true,
        conflictMonths: conflictMonths,
        error: "В части месяцев уже есть активные данные. Подтвердите явную замену."
      };
    }

    var now = expenseNowIso_();
    var replacedMonths = [];
    for (var r = 0; r < conflictMonths.length; r++) {
      var conflictMonth = conflictMonths[r];
      var conflictMode = conflictMap[conflictMonth];
      if (conflictMode === "DETAIL" || conflictMode === "CONFLICT") expenseSoftDeleteDetailMonths_(ss, [conflictMonth], importId);
      if (conflictMode === "SUMMARY" || conflictMode === "CONFLICT") expenseDeactivateSummaryMonths_(ss, [conflictMonth], importId);
      replacedMonths.push(conflictMonth);
    }

    // После замены перечитываем активные ключи.
    existingSummary = expenseGetActiveSummaryRows_(ss);
    existingKeys = {};
    for (var a = 0; a < existingSummary.length; a++) {
      existingKeys[expenseClean_(existingSummary[a].SummaryKey) || expenseBuildSummaryKey_(existingSummary[a])] = true;
    }

    var sheet = ss.getSheetByName(EXPENSE_SHEET_MONTHLY_SUMMARY);
    var toWrite = [];
    var duplicates = 0;
    var importedRows = 0;

    for (var mm = 0; mm < months.length; mm++) {
      var currentMonth = months[mm];
      var rowsForMonth = validByMonth[currentMonth];
      var isReplacementMonth = conflictMap[currentMonth] !== undefined;

      for (var z = 0; z < rowsForMonth.length; z++) {
        var rowObj = rowsForMonth[z];
        if (!isReplacementMonth && existingKeys[rowObj.SummaryKey]) {
          duplicates++;
          continue;
        }
        rowObj.ID = Utilities.getUuid();
        rowObj.ImportId = importId;
        rowObj.ImportedAt = now;
        rowObj.Active = true;
        rowObj.ReplacedAt = "";
        rowObj.ReplacedByImportId = "";
        toWrite.push(expenseObjectToRow_(rowObj, EXPENSE_SUMMARY_HEADERS));
        importedRows++;
      }
    }

    if (toWrite.length) sheet.getRange(sheet.getLastRow() + 1, 1, toWrite.length, EXPENSE_SUMMARY_HEADERS.length).setValues(toWrite);

    for (var s = 0; s < months.length; s++) {
      var stateMonth = months[s];
      if ((validByMonth[stateMonth] || []).length) expenseSetMonthState_(ss, stateMonth, "SUMMARY", importId);
    }

    expenseLogImport_(ss, {
      ImportId: importId,
      FileHash: fileHash,
      Month: months.length === 1 ? months[0] : "",
      FileName: fileName,
      TotalRows: rows.length,
      ImportedRows: importedRows,
      Duplicates: duplicates,
      Errors: 0,
      ImportedAt: now,
      ImportType: "SUMMARY",
      StartMonth: months.length ? months[0] : "",
      EndMonth: months.length ? months[months.length - 1] : "",
      ReplacedMonths: replacedMonths.join(",")
    });

    return {
      success: true,
      importId: importId,
      importedRows: importedRows,
      duplicates: duplicates,
      replacedMonths: replacedMonths,
      monthTotals: expenseMonthTotalsArray_(monthTotals),
      grandTotal: expenseSumObjectValues_(monthTotals),
      control: control
    };
  } finally {
    lock.releaseLock();
  }
}

function expensePrepareImportRows_(rows, categories) {
  var signatureCounters = {};
  var prepared = [];

  for (var i = 0; i < rows.length; i++) {
    var raw = rows[i] || {};
    var date = expenseNormalizeDate_(raw.Date !== undefined ? raw.Date : raw.date);
    var amount = expenseParseAmount_(raw.Amount !== undefined ? raw.Amount : raw.amount);
    var description = expenseClean_(raw.Description !== undefined ? raw.Description : raw.description);
    var source = expenseClean_(raw.Source !== undefined ? raw.Source : raw.source) || "Общая карта";
    var sourceOperationId = expenseClean_(raw.SourceOperationId !== undefined ? raw.SourceOperationId : raw.sourceOperationId);
    var occurrenceRaw = raw.Occurrence !== undefined ? raw.Occurrence : raw.occurrence;

    var signature = [
      date || "",
      isNaN(amount) ? "" : amount.toFixed(2),
      expenseNorm_(description),
      expenseNorm_(source)
    ].join("|");

    var occurrence = parseInt(occurrenceRaw, 10);
    if (isNaN(occurrence) || occurrence < 1) {
      signatureCounters[signature] = (signatureCounters[signature] || 0) + 1;
      occurrence = signatureCounters[signature];
    } else {
      signatureCounters[signature] = Math.max(signatureCounters[signature] || 0, occurrence);
    }

    var sourceObj = {
      date: date,
      amount: amount,
      description: description,
      category: raw.Category !== undefined ? raw.Category : raw.category,
      subcategory: raw.Subcategory !== undefined ? raw.Subcategory : raw.subcategory,
      operationType: raw.OperationType !== undefined ? raw.OperationType : raw.operationType,
      source: source,
      person: raw.Person !== undefined ? raw.Person : raw.person,
      comment: raw.Comment !== undefined ? raw.Comment : raw.comment,
      sourceOperationId: sourceOperationId,
      occurrence: occurrence
    };

    var item = expensePrepareSingleOperation_(sourceObj, categories, false, true);
    prepared.push({ rowNumber: i + 2, operation: item.operation, errors: item.errors });
  }

  return prepared;
}

function expensePrepareSingleOperation_(raw, categories, forManual, allowUnknownCategory) {
  raw = raw || {};
  var errors = [];

  var date = expenseNormalizeDate_(raw.date !== undefined ? raw.date : raw.Date);
  if (!date) errors.push("Некорректная дата.");

  var amount = expenseParseAmount_(raw.amount !== undefined ? raw.amount : raw.Amount);
  if (isNaN(amount) || amount <= 0) errors.push("Сумма должна быть больше нуля.");

  var description = expenseClean_(raw.description !== undefined ? raw.description : raw.Description);
  if (!description) errors.push("Не заполнено описание операции.");

  var categoryInput = expenseClean_(raw.category !== undefined ? raw.category : raw.Category);
  var category = "";
  var activeCategoryMap = {};
  for (var i = 0; i < categories.length; i++) {
    var c = expenseClean_(categories[i].Category);
    if (c && !activeCategoryMap[expenseNorm_(c)]) activeCategoryMap[expenseNorm_(c)] = c;
  }
  if (categoryInput && activeCategoryMap[expenseNorm_(categoryInput)]) {
    category = activeCategoryMap[expenseNorm_(categoryInput)];
  } else if (categoryInput) {
    category = categoryInput;
    if (!allowUnknownCategory) errors.push("Неизвестная категория: " + categoryInput + ".");
  } else {
    errors.push("Не заполнена категория.");
  }

  var operationTypeInput = expenseClean_(raw.operationType !== undefined ? raw.operationType : raw.OperationType) || "Расход";
  operationTypeInput = expenseNormalizeOperationType_(operationTypeInput);
  var operationType = "";
  for (var j = 0; j < EXPENSE_ALLOWED_TYPES.length; j++) {
    if (expenseNorm_(EXPENSE_ALLOWED_TYPES[j]) === expenseNorm_(operationTypeInput)) {
      operationType = EXPENSE_ALLOWED_TYPES[j];
      break;
    }
  }
  if (!operationType) {
    operationType = operationTypeInput;
    errors.push("Недопустимый тип операции: " + operationTypeInput + ".");
  }

  var source = expenseClean_(raw.source !== undefined ? raw.source : raw.Source) || "Общая карта";
  var personRaw = expenseClean_(raw.person !== undefined ? raw.person : raw.Person);
  var person = expenseNormalizePerson_(personRaw || "Общее");
  if (!person) {
    person = personRaw || "Общее";
    errors.push("Недопустимое значение Person: " + person + ". Используйте Общее, Катя или Коля.");
  }
  var occurrence = parseInt(raw.occurrence !== undefined ? raw.occurrence : raw.Occurrence, 10);
  if (isNaN(occurrence) || occurrence < 1) occurrence = forManual ? 0 : 1;

  var operation = {
    ID: "",
    OperationKey: "",
    Date: date || "",
    Amount: isNaN(amount) ? "" : amount,
    Description: description,
    Category: category,
    Subcategory: expenseClean_(raw.subcategory !== undefined ? raw.subcategory : raw.Subcategory),
    OperationType: operationType,
    Source: source,
    Person: person,
    Comment: expenseClean_(raw.comment !== undefined ? raw.comment : raw.Comment),
    Month: date ? date.slice(0, 7) : "",
    ImportId: "",
    ImportedAt: "",
    CreatedAt: "",
    UpdatedAt: "",
    DeletedAt: "",
    SourceOperationId: expenseClean_(raw.sourceOperationId !== undefined ? raw.sourceOperationId : raw.SourceOperationId),
    Occurrence: occurrence
  };

  if (date && !isNaN(amount) && amount > 0 && description && operationType) operation.OperationKey = expenseBuildOperationKey_(operation);
  return { operation: operation, errors: errors };
}

function expensePrepareHistoricalRows_(rows) {
  var prepared = [];
  for (var i = 0; i < rows.length; i++) {
    var raw = rows[i] || {};
    var errors = [];
    var recordType = expenseClean_(raw.RecordType);
    var month = expenseNormalizeMonth_(raw.Month);
    var category = expenseClean_(raw.Category);
    var amount = expenseParseAmount_(raw.Amount);
    var source = expenseClean_(raw.Source);
    var comment = expenseClean_(raw.Comment);

    if (expenseNorm_(recordType) !== "summary") errors.push("RecordType должен быть SUMMARY.");
    if (!month) errors.push("Month должен быть в формате YYYY-MM.");
    if (!category) errors.push("Категория не заполнена.");
    if (isNaN(amount) || amount < 0) errors.push("Amount должен быть числом не меньше нуля.");

    var summary = {
      ID: "",
      SummaryKey: "",
      Month: month,
      Category: category,
      Amount: isNaN(amount) ? "" : amount,
      Source: source,
      Comment: comment,
      ImportId: "",
      ImportedAt: "",
      Active: true,
      ReplacedAt: "",
      ReplacedByImportId: ""
    };
    if (!errors.length) summary.SummaryKey = expenseBuildSummaryKey_(summary);
    prepared.push({ rowNumber: i + 2, summary: summary, errors: errors });
  }
  return prepared;
}

function expenseNormalizePerson_(value) {
  var raw = expenseClean_(value);
  if (!raw) return "Общее";
  var n = expenseNorm_(raw);
  if (n === "общее" || n === "общие" || n === "семья" || n === "семейное") return "Общее";
  if (n === "катя" || n === "екатерина" || n === "kate" || n === "katya") return "Катя";
  if (n === "коля" || n === "николай" || n === "kolya" || n === "nikolay") return "Коля";
  return "";
}

function expenseNormalizePersonFilter_(value) {
  var raw = expenseClean_(value);
  if (!raw || expenseNorm_(raw) === "все" || expenseNorm_(raw) === "all") return "";
  return expenseNormalizePerson_(raw);
}

function expenseNormalizeOperationType_(value) {
  var raw = expenseClean_(value);
  var n = expenseNorm_(raw);
  var aliases = {
    "expense": "Расход",
    "income": "Доход",
    "refund": "Возврат покупки",
    "purchase refund": "Возврат покупки",
    "investment": "Инвестиции",
    "adjustment": "Корректировка",
    "internal transfer": "Перевод между своими счетами",
    "family transfer": "Перевод внутри семьи",
    "savings deposit": "Пополнение накоплений",
    "savings withdrawal": "Снятие с накоплений"
  };
  return aliases[n] || raw;
}

function expensePersonExpenseTotals_(operations) {
  var totals = { "Общее": 0, "Катя": 0, "Коля": 0 };
  for (var i = 0; i < operations.length; i++) {
    var op = operations[i];
    var type = expenseClean_(op.OperationType);
    if (type !== "Расход" && type !== "Возврат покупки") continue;
    var person = expenseNormalizePerson_(op.Person) || "Общее";
    var amount = parseFloat(op.Amount) || 0;
    totals[person] = (totals[person] || 0) + (type === "Возврат покупки" ? -amount : amount);
  }
  return totals;
}

function expenseBuildOperationKey_(operation) {
  var source = expenseNorm_(operation.Source || "Общая карта");
  var sourceOperationId = expenseClean_(operation.SourceOperationId);
  if (sourceOperationId) return expenseSha256_("bank|" + source + "|" + expenseNorm_(sourceOperationId));

  var canonical = [
    "row",
    expenseClean_(operation.Date),
    (parseFloat(operation.Amount) || 0).toFixed(2),
    expenseNorm_(operation.Description),
    source,
    String(parseInt(operation.Occurrence, 10) || 1)
  ].join("|");
  return expenseSha256_(canonical);
}

function expenseBuildSummaryKey_(summary) {
  var canonical = [
    "summary",
    expenseClean_(summary.Month),
    expenseNorm_(summary.Category),
    (parseFloat(summary.Amount) || 0).toFixed(2),
    expenseNorm_(summary.Source),
    expenseNorm_(summary.Comment)
  ].join("|");
  return expenseSha256_(canonical);
}

function expenseNextOccurrence_(existingRows, operation) {
  var maxOccurrence = 0;
  var date = expenseClean_(operation.Date);
  var amount = (parseFloat(operation.Amount) || 0).toFixed(2);
  var description = expenseNorm_(operation.Description);
  var source = expenseNorm_(operation.Source || "Общая карта");

  for (var i = 0; i < existingRows.length; i++) {
    var row = existingRows[i];
    if (
      expenseClean_(row.Date) === date &&
      (parseFloat(row.Amount) || 0).toFixed(2) === amount &&
      expenseNorm_(row.Description) === description &&
      expenseNorm_(row.Source || "Общая карта") === source &&
      !expenseClean_(row.SourceOperationId)
    ) {
      maxOccurrence = Math.max(maxOccurrence, parseInt(row.Occurrence, 10) || 1);
    }
  }
  return maxOccurrence + 1;
}

function expenseBuildSummary_(operations) {
  var expensesGross = 0;
  var refunds = 0;
  var income = 0;
  var transfers = 0;
  var savingsIn = 0;
  var savingsOut = 0;
  var investments = 0;

  for (var i = 0; i < operations.length; i++) {
    var op = operations[i];
    var amount = parseFloat(op.Amount) || 0;
    var type = expenseClean_(op.OperationType);
    if (type === "Расход") expensesGross += amount;
    else if (type === "Возврат покупки") refunds += amount;
    else if (type === "Доход") income += amount;
    else if (type === "Перевод между своими счетами" || type === "Перевод внутри семьи") transfers += amount;
    else if (type === "Пополнение накоплений") savingsIn += amount;
    else if (type === "Снятие с накоплений") savingsOut += amount;
    else if (type === "Инвестиции") investments += amount;
  }

  var expenses = expensesGross - refunds;
  return {
    expenses: expenses,
    expensesGross: expensesGross,
    refunds: refunds,
    income: income,
    transfers: transfers,
    savingsIn: savingsIn,
    savingsOut: savingsOut,
    investments: investments,
    cashFlow: income - expenses,
    operationCount: operations.length,
    summaryRowCount: 0,
    hasSummaryMonths: false,
    incomeComplete: true,
    transfersComplete: true,
    cashFlowComplete: true
  };
}

function expenseBuildUnifiedSummary_(operations, summaryRows) {
  var detail = expenseBuildSummary_(operations);
  var summaryExpenses = 0;
  for (var i = 0; i < summaryRows.length; i++) summaryExpenses += parseFloat(summaryRows[i].Amount) || 0;
  var hasSummary = summaryRows.length > 0;

  detail.summaryExpenses = summaryExpenses;
  detail.expenses += summaryExpenses;
  detail.expensesGross += summaryExpenses;
  detail.summaryRowCount = summaryRows.length;
  detail.hasSummaryMonths = hasSummary;
  detail.incomeComplete = !hasSummary;
  detail.transfersComplete = !hasSummary;
  detail.cashFlowComplete = !hasSummary;
  detail.cashFlow = hasSummary ? null : (detail.income - detail.expenses);
  return detail;
}

function expenseCategoryTotals_(operations) {
  var map = {};
  for (var i = 0; i < operations.length; i++) {
    var op = operations[i];
    var type = expenseClean_(op.OperationType);
    if (type !== "Расход" && type !== "Возврат покупки") continue;
    var category = expenseClean_(op.Category) || "Прочее";
    var amount = parseFloat(op.Amount) || 0;
    var effect = type === "Возврат покупки" ? -amount : amount;
    map[category] = (map[category] || 0) + effect;
  }
  return expenseCategoryMapToArray_(map);
}

function expenseUnifiedCategoryTotals_(operations, summaryRows) {
  var map = {};
  var detail = expenseCategoryTotals_(operations);
  for (var i = 0; i < detail.length; i++) map[detail[i].category] = (map[detail[i].category] || 0) + detail[i].amount;
  for (var j = 0; j < summaryRows.length; j++) {
    var category = expenseClean_(summaryRows[j].Category) || "Прочее";
    map[category] = (map[category] || 0) + (parseFloat(summaryRows[j].Amount) || 0);
  }
  return expenseCategoryMapToArray_(map);
}

function expenseCategoryMapToArray_(map) {
  var result = Object.keys(map).map(function(category) { return { category: category, amount: map[category] }; });
  result.sort(function(a, b) { return b.amount - a.amount; });
  return result;
}

function expenseLoadDashboardSnapshot_(ss) {
  var operations = expenseSheetToObjects_(ss.getSheetByName(EXPENSE_SHEET_OPERATIONS)).filter(function(row) {
    return !expenseClean_(row.DeletedAt);
  });
  var summaryRows = expenseGetActiveSummaryRows_(ss);
  var stateRows = expenseSheetToObjects_(ss.getSheetByName(EXPENSE_SHEET_MONTH_STATE));
  var categories = expenseGetCategoriesData_(ss, false);

  var snapshot = {
    operations: operations,
    summaryRows: summaryRows,
    stateRows: stateRows,
    categories: categories
  };
  snapshot.modeContext = expenseBuildModeContextFromSnapshot_(snapshot);
  return snapshot;
}

function expenseBuildModeContextFromSnapshot_(snapshot) {
  snapshot = snapshot || {};
  var stateRows = snapshot.stateRows || [];
  var detailRows = snapshot.operations || [];
  var summaryRows = snapshot.summaryRows || [];

  var stateMap = {};
  for (var i = 0; i < stateRows.length; i++) {
    var month = expenseNormalizeMonth_(stateRows[i].Month);
    var mode = expenseClean_(stateRows[i].Mode).toUpperCase();
    if (month && (mode === "SUMMARY" || mode === "DETAIL")) stateMap[month] = mode;
  }

  var detailMonths = {};
  for (var d = 0; d < detailRows.length; d++) {
    if (expenseClean_(detailRows[d].DeletedAt)) continue;
    var dm = expenseNormalizeMonth_(detailRows[d].Month) || expenseNormalizeMonth_(expenseClean_(detailRows[d].Date).slice(0, 7));
    if (dm) detailMonths[dm] = true;
  }

  var summaryMonths = {};
  for (var s = 0; s < summaryRows.length; s++) {
    var sm = expenseNormalizeMonth_(summaryRows[s].Month);
    if (sm) summaryMonths[sm] = true;
  }

  return { stateMap: stateMap, detailMonths: detailMonths, summaryMonths: summaryMonths };
}

function expenseGetUnifiedRangeDataFromSnapshot_(snapshot, startDate, endDate, personFilter) {
  personFilter = expenseNormalizePersonFilter_(personFilter);
  snapshot = snapshot || {};
  var modeContext = snapshot.modeContext || expenseBuildModeContextFromSnapshot_(snapshot);
  var allOperations = snapshot.operations || [];
  var allSummary = snapshot.summaryRows || [];
  var months = expenseMonthsBetweenDates_(startDate, endDate);
  var operations = [];
  var summaryRows = [];
  var monthlySeries = [];
  var partialSummaryMonths = [];
  var monthStates = [];

  // Быстрые индексы вместо filter() по всему листу для каждого месяца.
  var operationsByMonth = {};
  for (var oi = 0; oi < allOperations.length; oi++) {
    var op = allOperations[oi];
    var opMonth = expenseNormalizeMonth_(op.Month) || expenseNormalizeMonth_(expenseClean_(op.Date).slice(0, 7));
    if (!opMonth) continue;
    if (!operationsByMonth[opMonth]) operationsByMonth[opMonth] = [];
    operationsByMonth[opMonth].push(op);
  }

  var summaryByMonth = {};
  for (var si = 0; si < allSummary.length; si++) {
    var sr = allSummary[si];
    var sm = expenseNormalizeMonth_(sr.Month);
    if (!sm) continue;
    if (!summaryByMonth[sm]) summaryByMonth[sm] = [];
    summaryByMonth[sm].push(sr);
  }

  for (var i = 0; i < months.length; i++) {
    var month = months[i];
    var mode = expenseResolveMonthModeFromContext_(modeContext, month);
    var monthRange = expenseMonthRange_(month);
    monthStates.push({ month: month, mode: mode || "EMPTY" });

    if (mode === "SUMMARY") {
      if (personFilter) {
        monthlySeries.push({ month: month, mode: "SUMMARY", expenses: null, income: null, transfers: null, cashFlow: null, partial: false, personUnavailable: true });
        continue;
      }
      var fullyCovered = startDate <= monthRange.startDate && endDate >= monthRange.endDate;
      if (!fullyCovered) {
        partialSummaryMonths.push(month);
        monthlySeries.push({ month: month, mode: "SUMMARY", expenses: null, income: null, transfers: null, cashFlow: null, partial: true });
        continue;
      }
      var rowsForMonth = summaryByMonth[month] || [];
      var monthExpense = 0;
      for (var s = 0; s < rowsForMonth.length; s++) {
        summaryRows.push(rowsForMonth[s]);
        monthExpense += parseFloat(rowsForMonth[s].Amount) || 0;
      }
      monthlySeries.push({ month: month, mode: "SUMMARY", expenses: monthExpense, income: null, transfers: null, cashFlow: null, partial: false });
    } else if (mode === "DETAIL") {
      var sourceOps = operationsByMonth[month] || [];
      var monthOps = [];
      for (var mo = 0; mo < sourceOps.length; mo++) {
        var currentOp = sourceOps[mo];
        var date = expenseNormalizeDate_(currentOp.Date) || expenseClean_(currentOp.Date);
        if (!date || date < startDate || date > endDate) continue;
        if (personFilter && expenseNormalizePerson_(currentOp.Person) !== personFilter) continue;
        monthOps.push(currentOp);
      }
      for (var d = 0; d < monthOps.length; d++) operations.push(monthOps[d]);
      var detailSummary = expenseBuildSummary_(monthOps);
      monthlySeries.push({
        month: month,
        mode: "DETAIL",
        expenses: detailSummary.expenses,
        income: detailSummary.income,
        transfers: detailSummary.transfers,
        cashFlow: detailSummary.cashFlow,
        partial: false
      });
    } else if (mode === "CONFLICT") {
      monthlySeries.push({ month: month, mode: "CONFLICT", expenses: null, income: null, transfers: null, cashFlow: null, partial: false });
    } else {
      monthlySeries.push({ month: month, mode: "EMPTY", expenses: 0, income: 0, transfers: 0, cashFlow: 0, partial: false });
    }
  }

  operations.sort(function(a, b) { return expenseClean_(b.Date).localeCompare(expenseClean_(a.Date)); });
  return {
    operations: operations,
    summaryRows: summaryRows,
    summary: expenseBuildUnifiedSummary_(operations, summaryRows),
    categoryTotals: expenseUnifiedCategoryTotals_(operations, summaryRows),
    monthlySeries: monthlySeries,
    monthStates: monthStates,
    partialSummaryMonths: partialSummaryMonths
  };
}

function expenseGetUnifiedRangeData_(ss, startDate, endDate, personFilter) {
  return expenseGetUnifiedRangeDataFromSnapshot_(expenseLoadDashboardSnapshot_(ss), startDate, endDate, personFilter);
}

function expenseBuildModeContext_(ss) {
  var snapshot = expenseLoadDashboardSnapshot_(ss);
  return snapshot.modeContext;
}

function expenseResolveMonthModeFromContext_(context, month) {
  month = expenseNormalizeMonth_(month);
  if (!month) return "";
  var hasDetail = !!context.detailMonths[month];
  var hasSummary = !!context.summaryMonths[month];
  var state = context.stateMap[month] || "";

  // MonthState — приоритетный указатель, но он не должен скрывать реально
  // существующие данные, если когда-то осталась устаревшая служебная строка.
  if (state === "DETAIL" && hasDetail) return "DETAIL";
  if (state === "SUMMARY" && hasSummary) return "SUMMARY";
  if (state === "DETAIL" && !hasDetail && hasSummary) return "SUMMARY";
  if (state === "SUMMARY" && !hasSummary && hasDetail) return "DETAIL";

  if (hasDetail && hasSummary) return "CONFLICT";
  if (hasSummary) return "SUMMARY";
  if (hasDetail) return "DETAIL";
  return "";
}

function expenseGetMonthMode_(ss, month) {
  return expenseResolveMonthModeFromContext_(expenseBuildModeContext_(ss), month);
}

function expenseGetMonthStates_(ss) {
  var context = expenseBuildModeContext_(ss);
  var monthsMap = {};
  Object.keys(context.stateMap).forEach(function(m) { monthsMap[m] = true; });
  Object.keys(context.detailMonths).forEach(function(m) { monthsMap[m] = true; });
  Object.keys(context.summaryMonths).forEach(function(m) { monthsMap[m] = true; });
  return Object.keys(monthsMap).sort().map(function(month) {
    return { month: month, mode: expenseResolveMonthModeFromContext_(context, month) || "EMPTY" };
  });
}

function expenseSetMonthState_(ss, month, mode, importId) {
  month = expenseNormalizeMonth_(month);
  mode = expenseClean_(mode).toUpperCase();
  if (!month || (mode !== "SUMMARY" && mode !== "DETAIL")) return;

  var sheet = ss.getSheetByName(EXPENSE_SHEET_MONTH_STATE);
  var found = expenseFindRowByField_(sheet, "Month", month);
  var obj = { Month: month, Mode: mode, ImportId: expenseClean_(importId), UpdatedAt: expenseNowIso_() };
  if (found) expenseWriteObjectToRow_(sheet, found.rowIndex, obj);
  else expenseAppendObject_(sheet, obj);
}

function expenseGetActiveSummaryRows_(ss) {
  var rows = expenseSheetToObjects_(ss.getSheetByName(EXPENSE_SHEET_MONTHLY_SUMMARY));
  return rows.filter(function(row) {
    return expenseBool_(row.Active, true) && !expenseClean_(row.ReplacedAt);
  });
}

function expenseDeactivateSummaryMonths_(ss, months, importId) {
  if (!months || !months.length) return;
  var monthMap = {};
  for (var i = 0; i < months.length; i++) monthMap[months[i]] = true;
  var sheet = ss.getSheetByName(EXPENSE_SHEET_MONTHLY_SUMMARY);
  var headers = expenseGetHeaders_(sheet);
  var monthCol = headers.indexOf("Month");
  var activeCol = headers.indexOf("Active");
  var replacedAtCol = headers.indexOf("ReplacedAt");
  var replacedByCol = headers.indexOf("ReplacedByImportId");
  if (sheet.getLastRow() < 2 || monthCol === -1) return;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  var now = expenseNowIso_();
  var changed = false;
  for (var r = 0; r < data.length; r++) {
    var month = expenseNormalizeMonth_(data[r][monthCol]);
    var active = activeCol === -1 ? true : expenseBool_(data[r][activeCol], true);
    if (monthMap[month] && active) {
      if (activeCol !== -1) data[r][activeCol] = false;
      if (replacedAtCol !== -1) data[r][replacedAtCol] = now;
      if (replacedByCol !== -1) data[r][replacedByCol] = importId;
      changed = true;
    }
  }
  if (changed) sheet.getRange(2, 1, data.length, headers.length).setValues(data);
}

function expenseSoftDeleteDetailMonths_(ss, months, importId) {
  if (!months || !months.length) return;
  var monthMap = {};
  for (var i = 0; i < months.length; i++) monthMap[months[i]] = true;
  var sheet = ss.getSheetByName(EXPENSE_SHEET_OPERATIONS);
  var headers = expenseGetHeaders_(sheet);
  var monthCol = headers.indexOf("Month");
  var deletedCol = headers.indexOf("DeletedAt");
  var updatedCol = headers.indexOf("UpdatedAt");
  if (sheet.getLastRow() < 2 || monthCol === -1 || deletedCol === -1) return;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  var now = expenseNowIso_();
  var changed = false;
  for (var r = 0; r < data.length; r++) {
    var month = expenseNormalizeMonth_(data[r][monthCol]);
    var deleted = expenseClean_(data[r][deletedCol]);
    if (monthMap[month] && !deleted) {
      data[r][deletedCol] = now;
      if (updatedCol !== -1) data[r][updatedCol] = now;
      changed = true;
    }
  }
  if (changed) sheet.getRange(2, 1, data.length, headers.length).setValues(data);
}

function expenseHistoricalControl_(monthTotals) {
  var controlMonths = Object.keys(EXPENSE_HISTORICAL_CONTROL).sort();
  var presentControlMonths = controlMonths.filter(function(month) { return monthTotals[month] !== undefined; });
  var applies = presentControlMonths.length > 0;
  var details = [];
  var ok = true;

  for (var i = 0; i < presentControlMonths.length; i++) {
    var month = presentControlMonths[i];
    var actual = Math.round((parseFloat(monthTotals[month]) || 0) * 100) / 100;
    var expected = EXPENSE_HISTORICAL_CONTROL[month];
    var match = Math.abs(actual - expected) < 0.005;
    if (!match) ok = false;
    details.push({ month: month, expected: expected, actual: actual, match: match });
  }

  var isFullFirstSet = presentControlMonths.length === controlMonths.length && controlMonths.every(function(month) {
    return monthTotals[month] !== undefined;
  });
  var actualTotal = 0;
  for (var t = 0; t < controlMonths.length; t++) {
    if (monthTotals[controlMonths[t]] !== undefined) actualTotal += parseFloat(monthTotals[controlMonths[t]]) || 0;
  }
  var totalMatch = !isFullFirstSet || Math.abs(actualTotal - EXPENSE_HISTORICAL_CONTROL_TOTAL) < 0.005;
  if (!totalMatch) ok = false;

  return {
    applies: applies,
    ok: ok,
    isFullFirstSet: isFullFirstSet,
    expectedTotal: isFullFirstSet ? EXPENSE_HISTORICAL_CONTROL_TOTAL : null,
    actualTotal: isFullFirstSet ? actualTotal : null,
    totalMatch: totalMatch,
    details: details
  };
}

function expenseMonthTotalsArray_(monthTotals) {
  return Object.keys(monthTotals).sort().map(function(month) {
    return { month: month, amount: monthTotals[month] };
  });
}

function expenseSumObjectValues_(obj) {
  var total = 0;
  Object.keys(obj || {}).forEach(function(key) { total += parseFloat(obj[key]) || 0; });
  return total;
}

function expenseWasFileImported_(ss, fileHash) {
  if (!fileHash) return false;
  var rows = expenseSheetToObjects_(ss.getSheetByName(EXPENSE_SHEET_IMPORT_LOG));
  for (var i = 0; i < rows.length; i++) {
    if (expenseClean_(rows[i].FileHash) === fileHash) return true;
  }
  return false;
}

function expenseLogImport_(ss, obj) {
  expenseAppendObject_(ss.getSheetByName(EXPENSE_SHEET_IMPORT_LOG), obj);
}

function expenseSheetToObjects_(sheet) {
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return [];
  var data = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  var headers = data[0].map(function(v) { return String(v).trim(); });
  var result = [];

  for (var i = 1; i < data.length; i++) {
    var obj = {};
    var hasValue = false;
    for (var j = 0; j < headers.length; j++) {
      var value = data[i][j];
      if (value instanceof Date && !isNaN(value.getTime())) value = Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
      if (value !== "" && value !== null) hasValue = true;
      obj[headers[j]] = value;
    }
    if (hasValue) result.push(obj);
  }
  return result;
}

function expenseGetHeaders_(sheet) {
  if (!sheet || sheet.getLastColumn() < 1) return [];
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(v) { return String(v).trim(); });
}

function expenseFindRowByField_(sheet, field, value) {
  var headers = expenseGetHeaders_(sheet);
  var col = headers.indexOf(field);
  if (col === -1 || sheet.getLastRow() < 2) return null;
  var values = sheet.getRange(2, col + 1, sheet.getLastRow() - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    var matches = field === "Month"
      ? expenseNormalizeMonth_(values[i][0]) === expenseNormalizeMonth_(value)
      : String(values[i][0]).trim() === String(value).trim();
    if (matches) {
      var rowIndex = i + 2;
      var rowValues = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
      var obj = {};
      for (var j = 0; j < headers.length; j++) obj[headers[j]] = rowValues[j];
      return { rowIndex: rowIndex, object: obj };
    }
  }
  return null;
}

function expenseAppendObject_(sheet, obj) {
  var headers = expenseGetHeaders_(sheet);
  sheet.appendRow(expenseObjectToRow_(obj, headers));
}

function expenseObjectToRow_(obj, headers) {
  return headers.map(function(h) { return obj[h] !== undefined && obj[h] !== null ? obj[h] : ""; });
}

function expenseWriteObjectToRow_(sheet, rowIndex, obj) {
  var headers = expenseGetHeaders_(sheet);
  var row = headers.map(function(h) { return obj[h] !== undefined && obj[h] !== null ? obj[h] : ""; });
  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
}

function expenseNormalizeDate_(value) {
  if (!value) return "";
  if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");

  var s = String(value).trim();
  var m;
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/))) return expenseValidDate_(parseInt(m[1], 10), parseInt(m[2], 10), parseInt(m[3], 10));
  if ((m = s.match(/^(\d{1,2})[\.\/\-](\d{1,2})[\.\/\-](\d{4})$/))) return expenseValidDate_(parseInt(m[3], 10), parseInt(m[2], 10), parseInt(m[1], 10));
  return "";
}

function expenseValidDate_(year, month, day) {
  var d = new Date(year, month - 1, day);
  if (d.getFullYear() !== year || d.getMonth() !== month - 1 || d.getDate() !== day) return "";
  return String(year) + "-" + ("0" + month).slice(-2) + "-" + ("0" + day).slice(-2);
}

function expenseParseAmount_(value) {
  if (typeof value === "number") return value;
  if (value === undefined || value === null) return NaN;
  var s = String(value)
    .replace(/[\s\u00A0\u202F]/g, "")
    .replace(/[₽рРRUBrub]/g, "")
    .replace(",", ".")
    .replace(/[^0-9\.\-]/g, "");
  var n = parseFloat(s);
  return isNaN(n) ? NaN : n;
}

function expenseNormalizeMonth_(value) {
  if (value === undefined || value === null || value === "") return "";

  // Если Google Sheets уже превратил YYYY-MM в настоящую дату.
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM");
  }

  var s = expenseClean_(value);
  var m;

  // Нормальный технический формат.
  if ((m = s.match(/^(\d{4})-(\d{1,2})$/))) {
    var month1 = parseInt(m[2], 10);
    if (month1 >= 1 && month1 <= 12) return m[1] + "-" + ("0" + month1).slice(-2);
  }

  // Значение, которое пришло из ячейки-даты после чтения листа.
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/))) {
    var month2 = parseInt(m[2], 10);
    if (month2 >= 1 && month2 <= 12) return m[1] + "-" + ("0" + month2).slice(-2);
  }

  // На случай отображения/ручного значения в российском формате.
  if ((m = s.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})$/))) {
    var month3 = parseInt(m[2], 10);
    if (month3 >= 1 && month3 <= 12) return m[3] + "-" + ("0" + month3).slice(-2);
  }

  return "";
}

function expenseCurrentMonth_(ss) {
  var tz = ss && ss.getSpreadsheetTimeZone ? ss.getSpreadsheetTimeZone() : Session.getScriptTimeZone();
  return Utilities.formatDate(new Date(), tz, "yyyy-MM");
}

function expenseShiftMonth_(month, delta) {
  var m = String(month || "").match(/^(\d{4})-(\d{2})$/);
  if (!m) return "";
  var d = new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1 + delta, 1);
  return d.getFullYear() + "-" + ("0" + (d.getMonth() + 1)).slice(-2);
}

function expenseMonthRange_(month) {
  month = expenseNormalizeMonth_(month);
  if (!month) return { startDate: "", endDate: "" };
  var y = parseInt(month.slice(0, 4), 10);
  var m = parseInt(month.slice(5, 7), 10);
  var lastDay = new Date(y, m, 0).getDate();
  return { startDate: month + "-01", endDate: month + "-" + ("0" + lastDay).slice(-2) };
}

function expenseMonthsBetweenDates_(startDate, endDate) {
  var start = expenseNormalizeDate_(startDate);
  var end = expenseNormalizeDate_(endDate);
  if (!start || !end || start > end) return [];
  var sy = parseInt(start.slice(0, 4), 10);
  var sm = parseInt(start.slice(5, 7), 10) - 1;
  var ey = parseInt(end.slice(0, 4), 10);
  var em = parseInt(end.slice(5, 7), 10) - 1;
  var cursor = new Date(sy, sm, 1);
  var last = new Date(ey, em, 1);
  var result = [];
  while (cursor <= last && result.length < 600) {
    result.push(cursor.getFullYear() + "-" + ("0" + (cursor.getMonth() + 1)).slice(-2));
    cursor = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 1);
  }
  return result;
}

function expensePreviousEquivalentRange_(range) {
  var start = expenseNormalizeDate_(range.startDate);
  var end = expenseNormalizeDate_(range.endDate);
  var startDate = new Date(parseInt(start.slice(0,4),10), parseInt(start.slice(5,7),10)-1, parseInt(start.slice(8,10),10));
  var endDate = new Date(parseInt(end.slice(0,4),10), parseInt(end.slice(5,7),10)-1, parseInt(end.slice(8,10),10));
  var duration = Math.floor((endDate.getTime() - startDate.getTime()) / 86400000) + 1;
  var prevEnd = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate() - 1);
  var prevStart = new Date(prevEnd.getFullYear(), prevEnd.getMonth(), prevEnd.getDate() - duration + 1);
  return { startDate: expenseDateToIso_(prevStart), endDate: expenseDateToIso_(prevEnd) };
}

function expenseDateToIso_(d) {
  return d.getFullYear() + "-" + ("0" + (d.getMonth() + 1)).slice(-2) + "-" + ("0" + d.getDate()).slice(-2);
}

function expenseNowIso_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
}

function expenseClean_(value) {
  if (value === undefined || value === null) return "";
  return String(value).trim();
}

function expenseNorm_(value) {
  return expenseClean_(value).toLowerCase().replace(/\s+/g, " ");
}

function expenseBool_(value, defaultValue) {
  if (value === undefined || value === null || value === "") return defaultValue;
  if (value === false || String(value).toLowerCase() === "false" || String(value) === "0") return false;
  return true;
}

function expenseSha256_(text) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text), Utilities.Charset.UTF_8);
  return digest.map(function(byte) {
    var v = (byte < 0 ? byte + 256 : byte).toString(16);
    return v.length === 1 ? "0" + v : v;
  }).join("");
}

function expenseJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
