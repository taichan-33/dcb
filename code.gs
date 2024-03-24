function doPost(e) {
  var json = JSON.parse(e.postData.contents);
  var userMessage = json.events[0].message ? json.events[0].message.text : null;
  var replyToken = json.events[0].replyToken;
  var userId = json.events[0].source.userId;

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var currentYear = new Date().getFullYear();
  var sheet = spreadsheet.getSheetByName(String(currentYear));

  // If the sheet for the current year does not exist, create it
  if (!sheet) {
    var templateSheetName = '原本';
    var newSheetName = String(currentYear);
    sheet = createNewSheet(spreadsheet, templateSheetName, newSheetName);
    spreadsheet.setActiveSheet(sheet);
  }

  var date = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');

  var LINE_ACCESS_TOKEN = 'Your access token here';
  var headers = {
    'Content-Type': 'application/json; charset=UTF-8',
    'Authorization': 'Bearer ' + LINE_ACCESS_TOKEN
  };

  var postData = {
    replyToken: replyToken,
    messages: []
  };

  var lastEntryRow = getLastEntryRow(sheet);

  // Get the user's message history from the Properties Service
  var userProperties = PropertiesService.getUserProperties();
  var messageHistory = JSON.parse(userProperties.getProperty(userId) || '[]');
  messageHistory.push(userMessage);
  userProperties.setProperty(userId, JSON.stringify(messageHistory));

  // Check the previous message
  var prevMessage = messageHistory.length >= 2 ? messageHistory[messageHistory.length - 2] : null;

  function createQuickReply() {
    return {
      items: [
        {type: 'action', action: {type: 'message', label: '項目', text: '項目'}},
        {type: 'action', action: {type: 'message', label: '収入', text: '収入'}},
        {type: 'action', action: {type: 'message', label: '支出', text: '支出'}},
        {type: 'action', action: {type: 'message', label: '備考', text: '備考'}},
        {type: 'action', action: {type: 'message', label: '月収支計算', text: '月収支計算'}}
      ]
    };
  }

  switch (userMessage) {
    case '項目':
      postData.messages.push({type: 'text', text: '項目を入力してください', quickReply: createQuickReply()});
      break;

    case '収入':
      postData.messages.push({type: 'text', text: '収入金額を入力してください', quickReply: createQuickReply()});
      break;

    case '支出':
      postData.messages.push({type: 'text', text: '支出金額を入力してください', quickReply: createQuickReply()});
      break;

    case '備考':
            postData.messages.push({type: 'text', text: '備考を入力してください', quickReply: createQuickReply()});
      break;

    case '月収支計算':
      postData.messages.push({type: 'text', text: '収支計算する月を入力してください', quickReply: createQuickReply()});
      break;

    default:
      switch (prevMessage) {
        case '項目':
  var newRow = lastEntryRow + 1;
  sheet.getRange('A' + newRow).setValue(date);
  sheet.getRange('B' + newRow).setValue(userMessage);
  postData.messages.push({type: 'text', text: '項目が記入されました', quickReply: createQuickReply()});
  break;

        case '収入':
          if (sheet.getRange('C' + lastEntryRow).getValue() === '') {
            sheet.getRange('C' + lastEntryRow).setValue(userMessage);
            postData.messages.push({type: 'text', text: '収入が記入されました', quickReply: createQuickReply()});
          } else {
            postData.messages.push({type: 'text', text: '既に収入が記入されています', quickReply: createQuickReply()});
          }
          break;

        case '支出':
          if (sheet.getRange('D' + lastEntryRow).getValue() === '') {
            sheet.getRange('D' + lastEntryRow).setValue(userMessage);
            postData.messages.push({type: 'text', text: '支出が記入されました', quickReply: createQuickReply()});
          } else {
            postData.messages.push({type: 'text', text: '既に支出が記入されています', quickReply: createQuickReply()});
          }
          break;

        case '備考':
          if (sheet.getRange('E' + lastEntryRow).getValue() === '') {
            sheet.getRange('E' + lastEntryRow).setValue(userMessage);
            postData.messages.push({type: 'text', text: '備考が記入されました', quickReply: createQuickReply()});
          } else {
            postData.messages.push({type: 'text', text: '既に備考が記入されています', quickReply: createQuickReply()});
          }
          break;

        case '月収支計算':
        var month = parseInt(userMessage, 10);
        if (isNaN(month) || month < 1 || month > 12) {
        postData.messages.push({type: 'text', text: '指定のキーワードを入力してください', quickReply: createQuickReply()});
       } else {
        var { incomeTotal, expenseTotal } = getMonthlyBalance(sheet, month);
       postData.messages.push({type: 'text', text: month + '月の売上合計は' + incomeTotal + '円です。\n出金合計は' + expenseTotal + '円です。', quickReply: createQuickReply()});
      }
       break;

        default:
          postData.messages.push({type: 'text', text: '指定のキーワードを入力してください', quickReply: createQuickReply()});
      }
  }

  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    headers: headers,
    method: 'post',
    payload: JSON.stringify(postData)
  });

    return ContentService.createTextOutput(JSON.stringify({content: 'success'})).setMimeType(ContentService.MimeType.JSON);
}

function createQuickReply() {
  return {
    items: [
      {type: 'action', action: {type: 'message', label: '項目', text: '項目'}},
      {type: 'action', action: {type: 'message', label: '収入', text: '収入'}},
      {type: 'action', action: {type: 'message', label: '支出', text: '支出'}},
      {type: 'action', action: {type: 'message', label: '備考', text: '備考'}},
      {type: 'action', action: {type: 'message', label: '月収支計算', text: '月収支計算'}}
    ]
  };
}

function getLastEntryRow(sheet) {
  var lastRow = sheet.getLastRow();
  var lastEntryRow = sheet.getRange('B1:B' + lastRow).getValues().filter(String).length;
  return lastEntryRow;
}

function createNewSheet(spreadsheet, templateSheetName, newSheetName) {
  var templateSheet = spreadsheet.getSheetByName(templateSheetName);
  if (!templateSheet) {
    throw new Error("Template sheet '" + templateSheetName + "' not found.");
  }
  var newSheet = templateSheet.copyTo(spreadsheet);
  newSheet.setName(newSheetName);
  return newSheet;
}


function getMonthlyBalance(sheet, month) {
  var lastRow = sheet.getLastRow();
  var dateRange = sheet.getRange('A1:A' + lastRow).getValues();
  var incomeRange = sheet.getRange('C1:C' + lastRow).getValues();
  var expenseRange = sheet.getRange('D1:D' + lastRow).getValues();
  var incomeTotal = 0;
  var expenseTotal = 0;

  for (var i = 1; i < dateRange.length; i++) {
    var currentDate = dateRange[i][0];
    if (currentDate instanceof Date && currentDate.getMonth() + 1 === month) {
      incomeTotal += parseInt(incomeRange[i][0], 10) || 0;
      expenseTotal += parseInt(expenseRange[i][0], 10) || 0;
    }
  }

  return { incomeTotal, expenseTotal };
}
