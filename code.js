function doGet(e) {
  try {
    let response = {
      state: true,
      params: e.parameter
    };
    if (!("sheet" in e.parameter)) {
      response = {state:false, error:{message:["'sheet' is missing"]}};
      return ContentService.createTextOutput(JSON.stringify(response))
    }
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(e.parameter.sheet);
    // let sheet = SpreadsheetApp.openById('1U7YG7EyTxvDxrplJ434OGelfSeaqMd9wtW13ylwjHoE').getSheetByName(e.parameter.sheet);
    if (sheet == null) {
      response = {state:false, error:{message:["'sheet' is not found"]}}
      return ContentService.createTextOutput(JSON.stringify(response))
    }
    let range = sheet.getDataRange();
    let values = range.getValues();
    let columns = 0;
    let rows = [];
    for (let r=0; r<values.length; r++) {
      if (r==0) {
        for (let v=0; v<values[r].length; v++) {
          if (values[r][v] == "") {
            columns = v;
            break;
          }
          columns = v+1;
        }
      } else {
        let oneRow = {};
        for (let v=0; v<values[r].length; v++) {
          if (v<columns) {
            oneRow[values[0][v]] = values[r][v];
          }
        }
        rows.push(oneRow);
      }
    }
    if ("search" in e.parameter) {
      try{
        let offset=0, missed=0, limit = 20;
        response.rows = [];
        if ("limit" in e.parameter && String(Number(e.parameter.limit)) == e.parameter.limit) {
          limit = Number(e.parameter.limit);
        }
        if ("offset" in e.parameter && String(Number(e.parameter.offset)) == e.parameter.offset) {
          offset = Number(e.parameter.offset);
        }
        let search = JSON.parse(e.parameter.search);
        response.mode = "search";
        for (let r=0; r<rows.length; r++) {
          let approved = true;
          for (let key in rows[r]) {
            if (key in search) {
              if (rows[r][key] != search[key]) {
                approved = false;
                break;
              }
            }
          }
          if (approved) {
            if (missed<offset) {
              missed++;
            } else {
              response.rows.push(rows[r]);
            }
          }
          if (limit <= response.rows.length) {
            break;
          }
        }
        if (response.rows.length < 1) {
          response.state = false;
          response.error = {message:["row is not found from search"]};
          return ContentService.createTextOutput(JSON.stringify(response));
        }
      } catch {
        let response = {
          state: false,
          error: {
            message: [
              "failed parse string 'search'"
            ]
          }
        };
        return ContentService.createTextOutput(JSON.stringify(response));
      }
    } else {
      let offset=0, limit=20;
      if ("offset" in e.parameter && String(Number(e.parameter.offset)) == e.parameter.offset) {
        offset = Number(e.parameter.offset);
      }
      if ("limit" in e.parameter && String(Number(e.parameter.limit)) == e.parameter.limit) {
        limit = Number(e.parameter.limit);
      }
      response.mode = "not search";
      response.rows = rows.splice(offset, limit);
    }
    return ContentService.createTextOutput(JSON.stringify(response));
  } catch {
    let response = {
      state: false,
      error: {
        message: [
          "failed execute function"
        ]
      }
    };
    return ContentService.createTextOutput(JSON.stringify(response));
  }
}
function doPost(e) {
  try {
    let param, response = {state: true};
    try {
      param = JSON.parse(e.postData.contents);
    } catch (error) {
      response.state = false;
      response.error = {message: ["failed parse input data", error]};
      return ContentService.createTextOutput(JSON.stringify(response));
    }
    response.input = param;
    if (!("action" in param)) {
      response.state = false;
      response.error = {message: ["'action' is missing"]};
      return ContentService.createTextOutput(JSON.stringify(response));
    }
    if (!("sheet" in param)) {
      response.state = false;
      response.error = {message:["'sheet' is missing"]};
      return ContentService.createTextOutput(JSON.stringify(response))
    }
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(param.sheet);
    // let sheet = SpreadsheetApp.openById('1U7YG7EyTxvDxrplJ434OGelfSeaqMd9wtW13ylwjHoE').getSheetByName(param.sheet);
    if (sheet == null) {
      response.state = false;
      response.error = {message:["'sheet' is not found"]}
      return ContentService.createTextOutput(JSON.stringify(response))
    }
    if (param.action == "insert") {
      if ("row" in param) {
        if (typeof param.row != "object") {
          response.state = false;
          response.error = {message: ["'row' must by an object"]};
          response.rowType = typeof param.row;
          return ContentService.createTextOutput(JSON.stringify(response));
        }
        let firstRow = sheet.getDataRange().getValues()[0];
        let appendRow = [];
        for (let v=0; v<firstRow.length; v++) {
          if (firstRow[v] == "") {
            break;
          }
          if (firstRow[v] in param.row) {
            appendRow.push(param.row[firstRow[v]]);
          } else {
            appendRow.push("");
          }
        }
        response.appendData = appendRow;
        sheet.appendRow(appendRow);
      } else if ("rows" in param) {
        if (typeof param.rows != "object") {
          response.state = false;
          response.error = {message: ["'rows' must by an array"]};
          response.rowType = typeof param.rows;
          return ContentService.createTextOutput(JSON.stringify(response));
        }
        let firstRow = sheet.getDataRange().getValues()[0];
        response.appendData = [];
        for (let r=0; r<param.rows.length; r++) {
          if (typeof param.rows[r] != "object") {
            response.state = "warning";
            response.error = {message: [`'rows[${r}]' must by an object`]};
            response.rowType = typeof param.rows[r];
            continue;
          }
          let appendRow = [];
          for (let v=0; v<firstRow.length; v++) {
            if (firstRow[v] == "") {
              break;
            }
            if (firstRow[v] in param.rows[r]) {
              appendRow.push(param.rows[r][firstRow[v]]);
            } else {
              appendRow.push("");
            }
          }
          response.appendData.push(appendRow);
          sheet.appendRow(appendRow);
        }
      } else {
        response.state = false;
        response.error = {message: ["'row' or 'rows' is missing"]};
        return ContentService.createTextOutput(JSON.stringify(response));
      }
    } else if (param.action == "update") {
      response.state = false;
      response.error = {message: ["'action' is development"]};
      return ContentService.createTextOutput(JSON.stringify(response));
    } else if (param.action == "remove") {
      response.state = false;
      response.error = {message: ["'action' is development"]};
      return ContentService.createTextOutput(JSON.stringify(response));
    } else if (param.action == "clear") {
      if (!("approvedClear") in param || param.approvedClear !== true) {
        response.state = false;
        response.error = {message: ["Confirmation for clearing the table sheet was not canceled"]};
        return ContentService.createTextOutput(JSON.stringify(response));
      }
      let firstRow = sheet.getDataRange().getValues()[0];
      response.cleaning = sheet.clear();
      response.columnsName = sheet.appendRow(firstRow);
    }

    return ContentService.createTextOutput(JSON.stringify(response));
  } catch {
    let response = {
      state: false,
      error: {
        message: [
          "failed execute function"
        ]
      }
    };
    return ContentService.createTextOutput(JSON.stringify(response));
  }
}








