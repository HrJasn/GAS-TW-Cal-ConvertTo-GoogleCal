
//取得人事行政總處網頁辦公日曆表並寫入到Google日曆
function getTWCalToGCal(){

  const ActiveSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  ActiveSheet.clear();

  //定義人事行政總處網頁網域位址
  const dgpaPre = 'https://www.dgpa.gov.tw/';
  //定義從人事行政總處網頁辦公日曆表清單網頁頁面取得清單
  const dgpaTWCalLst = fetchToHTMLAndFiltedURL(dgpaPre + 'informationlist?uid=41', /(中華民國.*年.*日曆表)\s*$/);

  //迴圈處理辦公日曆表清單
  dgpaTWCalLst.forEach(dgpaResp1 => {
  
    console.log(`辦公日曆表頁面連結: ${dgpaResp1.href}`);
    //從辦公日曆表頁面取得辦公日曆表檔案連結
    const dgpaResp2 = fetchToHTMLAndFiltedURL(dgpaPre + dgpaResp1.href, /\.xls[\s\t\r\n ]*$/);
    console.log(`辦公日曆表檔案連結: ${JSON.stringify(dgpaResp2,null,2)}`);

    //辦公日曆表檔案下載到Google Drive
    const folder = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents().next();
    const filename = dgpaResp2[0].text;
    const ssFileName = filename.replace(/\.[^\.]+$/,'');
    let ssfiles = folder.getFilesByName(ssFileName);
    let SpreadsheetID = '';
    if (!SpreadsheetID && ssfiles.hasNext()) {
      SpreadsheetID = ssfiles.next().getId();
    }

    var files = folder.getFilesByName(filename);
    if(!files.hasNext()) {
      const excelUrl = dgpaPre + dgpaResp2[0].href;
      const response = UrlFetchApp.fetch(excelUrl, {muteHttpExceptions: true});
      const blob = response.getBlob().setContentType('application/vnd.ms-excel');
      const destFolders = [folder.getId()];
      updateFile(folder, filename, blob);
      console.log('destFolders: ', destFolders);
      // .xls檔案轉換為 Google 試算表
      if (!SpreadsheetID && !ssfiles.hasNext()) {
        SpreadsheetID = convertExcel2Sheets(blob, ssFileName, destFolders);
      }
    }

    console.log(`SpreadsheetID: ${SpreadsheetID}`);
    
    // 開啟轉換後的試算表
    const spreadsheet = SpreadsheetApp.openById(SpreadsheetID);
    const sheet = spreadsheet.getSheets()[0];
    const range = sheet.getDataRange();
    const rangeValues = range.getValues();

    // 透過試算表ID從API取得框線設定資料
    var fileId = spreadsheet.getId();
    var sheetName = sheet.getName();
    var fields = "sheets/data/rowData/values/userEnteredFormat/borders";
    var params = {
        method: "get",
        headers: {Authorization: "Bearer " +  ScriptApp.getOAuthToken()},
        muteHttpExceptions: true,
    };
    var rangeText = sheetName + "!" + range.getA1Notation();
    var url = "https://sheets.googleapis.com/v4/spreadsheets/" + fileId + "?ranges=" + encodeURIComponent(rangeText) + "&fields=" + encodeURIComponent(fields); 
    var res = UrlFetchApp.fetch(url, params);
    var RangesBordersFieldsData = JSON.parse(res.getContentText());
    const BorderFieldRows = RangesBordersFieldsData.sheets[0].data[0].rowData;

    // 辨識出年份資料後繼續轉換為方便行程表整合用的格式
    let year = null;
    const YearCell = matchesFromRangeValues(rangeValues, /西元[\s\t]*([0-9]+)[\s\t]*年/gm);
    if(YearCell && YearCell[0] && YearCell[0].value && YearCell[0].value[1]){
      year = YearCell[0].value[1];
    }
    if(year){
      
      const MonthCells = matchesFromRangeValues(rangeValues, /^[\s\t]*月[\s\t]*$/gm);
      const TWHldLst = updateMonthCellsToTWHldList(year, range, MonthCells, BorderFieldRows);
      // 排除週一到週五的上班日和週六、日的休假日
      let lastDay;
      const TWHldFiltedLst = TWHldLst.filter(day => {
        let cuDayChk = (
          ['一', '二', '三', '四', '五'].includes(day['星期']) && 
          day['上班或放假'] === '放假日'
        ) || (
          ['六', '日'].includes(day['星期']) && 
          day['上班或放假'] === '上班日'
        ) || (
          day['上班或放假'] === '放假日' && 
          day['節日'].match(/[節]/)
        ) || (lastDay && (
          (
            lastDay['上班或放假'] === '放假日' &&
            day['節日'].match(/[節]/)
          ) || (
            day['上班或放假'] === '放假日' &&
            lastDay['節日'].match(/[節]/)
          )
        ));
        lastDay = day;
        return cuDayChk;
      });
      // 控制台顯示過濾後的資料為表格格式
      printTable(TWHldFiltedLst);

      // 試算表若無資料則當作為第一行寫入標題資料
      if (ActiveSheet.getLastRow() == 0) {
        ActiveSheet.appendRow(Object.keys(TWHldFiltedLst[0]));
      }
      // 寫入過濾後行程資料到試算表
      if (TWHldFiltedLst.length > 0) {
        const dataRows = TWHldFiltedLst.map(day => Object.values(day));
        ActiveSheet.getRange(ActiveSheet.getLastRow()+1, 1, dataRows.length, dataRows[0].length).setValues(dataRows);
      }
      // 更新或新增剩餘不重複的行程資料到試算表名稱命名的日曆
      syncCalendarWithSpreadsheet(TWHldFiltedLst);

    }

  });

}

function updateMonthCellsToTWHldList(year, range, monthCells, BorderFieldRows) {
  const rangeValues = range.getValues();
  monthCells.forEach((cell) => {
    const row = cell.row;
    const col = cell.column;

    // Find the first cell with a left border
    for (let i = col; i >= 0; i--) {
      if (BorderFieldRows[row].values[i].userEnteredFormat && 
          BorderFieldRows[row].values[i].userEnteredFormat.borders.left &&
          BorderFieldRows[row].values[i].userEnteredFormat.borders.left.style == "SOLID_MEDIUM" &&
          BorderFieldRows[row].values[i].userEnteredFormat.borders.left.width == 2) {
        cell.left = i;
        break;
      } else if (i != col && 
                BorderFieldRows[row].values[i].userEnteredFormat && 
                BorderFieldRows[row].values[i].userEnteredFormat.borders.right &&
                BorderFieldRows[row].values[i].userEnteredFormat.borders.right.style == "SOLID_MEDIUM" &&
                BorderFieldRows[row].values[i].userEnteredFormat.borders.right.width == 2) {
        cell.left = i+1;
        break;
      }
    }

    // Find the first cell with a right border
    for (let i = col; i < BorderFieldRows[row].values.length; i++) {
      if (BorderFieldRows[row].values[i].userEnteredFormat && 
          BorderFieldRows[row].values[i].userEnteredFormat.borders.right &&
          BorderFieldRows[row].values[i].userEnteredFormat.borders.right.style == "SOLID_MEDIUM" &&
          BorderFieldRows[row].values[i].userEnteredFormat.borders.right.width == 2) {
        cell.right = i;
        break;
      } else if (i != col && 
                BorderFieldRows[row].values[i].userEnteredFormat && 
                BorderFieldRows[row].values[i].userEnteredFormat.borders.left &&
                BorderFieldRows[row].values[i].userEnteredFormat.borders.left.style == "SOLID_MEDIUM" &&
                BorderFieldRows[row].values[i].userEnteredFormat.borders.left.width == 2) {
        cell.right = i-1;
        break;
      }
    }

    // Find the first cell with a top border
    for (let i = row; i >= 0; i--) {
      if (BorderFieldRows[i].values[col].userEnteredFormat && 
          BorderFieldRows[i].values[col].userEnteredFormat.borders.top &&
          BorderFieldRows[i].values[col].userEnteredFormat.borders.top.style == "SOLID_MEDIUM" &&
          BorderFieldRows[i].values[col].userEnteredFormat.borders.top.width == 2) {
        cell.top = i;
        break;
      } else if (i != row && 
                BorderFieldRows[i].values[col].userEnteredFormat && 
                BorderFieldRows[i].values[col].userEnteredFormat.borders.bottom &&
                BorderFieldRows[i].values[col].userEnteredFormat.borders.bottom.style == "SOLID_MEDIUM" &&
                BorderFieldRows[i].values[col].userEnteredFormat.borders.bottom.width == 2) {
        cell.top = i+1;
        break;
      }
    }

    // Find the first cell with a bottom border
    for (let i = row; i < BorderFieldRows.length; i++) {
      if (BorderFieldRows[i].values[col].userEnteredFormat && 
          BorderFieldRows[i].values[col].userEnteredFormat.borders.bottom &&
          BorderFieldRows[i].values[col].userEnteredFormat.borders.bottom.style == "SOLID_MEDIUM" &&
          BorderFieldRows[i].values[col].userEnteredFormat.borders.bottom.width == 2) {
        cell.bottom = i;
        break;
      } else if (i != row && 
                BorderFieldRows[i].values[col].userEnteredFormat && 
                BorderFieldRows[i].values[col].userEnteredFormat.borders.top &&
                BorderFieldRows[i].values[col].userEnteredFormat.borders.top.style == "SOLID_MEDIUM" &&
                BorderFieldRows[i].values[col].userEnteredFormat.borders.top.width == 2) {
        cell.bottom = i-1;
        break;
      }
    }

    // 合併內容值為字串並取代 cell.value
    if (cell.left !== undefined && cell.right !== undefined) {
      let combinedValue = '';
      for (let i = cell.left; i <= cell.right; i++) {
        combinedValue += rangeValues[row][i];
      }
      const months = ['一月','二月','三月','四月','五月','六月','七月','八月','九月','十月','十一月','十二月'];
      // 若 combinedValue 匹配到月份陣列中的元素，則轉換為索引值加1的數字
      const monthIndex = months.indexOf(combinedValue);
      if (monthIndex !== -1) {
        combinedValue = (monthIndex + 1).toString();
      }
      cell.monthnumber = combinedValue.trim();
    }

    // 移除 cell.row 和 cell.column 屬性
    delete cell.row;
    delete cell.column;
    delete cell.value;

  });

  const matchedCells = [];
  monthCells.forEach((cell) => {
    const left = cell.left;
    const right = cell.right;
    const top = cell.top;
    const bottom = cell.bottom;
    const monthnumber = cell.monthnumber;

    for (let row = top + 2; row <= bottom; row += 2) {
      for (let col = left; col <= right; col++) {
        const celldate = rangeValues[row][col];
        let cellHld = '上班日';
        let cellcolor = '#ffffff';
        const targetRange = range.getCell(row+1, col+1);
        if(targetRange && targetRange.getBackground()){
          cellcolor = targetRange.getBackground();
        }
        if (cellcolor) {
          switch (cellcolor) {
            case '#ff99cc':
              cellHld = '放假日';
              break;
            case '#ffffff':
              cellHld = '上班日';
              break;
          }
        }
        if (celldate) {
          const cellpsco = {
            '年份': year,
            '月份': monthnumber,
            '日期': celldate,
            '星期': rangeValues[top + 1][col],
            '節日': rangeValues[row + 1][col].replace(/[\r\n\t ]+/gm,' '),
            '上班或放假': cellHld,
            'A1符號' : targetRange.getA1Notation()
          };
          matchedCells.push(cellpsco);
        }
      }
    }

  });

  return matchedCells;

}

function fetchToHTMLAndFiltedURL(url, pattern) {
  const response = UrlFetchApp.fetch(url);
  const htmlContent = response.getContentText();
  const $ = Cheerio.load(htmlContent);
  const links = $('a');
  const linkObjects = [];
  links.each((_, element) => {
    const href = $(element).attr('href');
    let text = $(element).text();
    if (text.match(pattern)) {
      text = text.match(pattern)[1] || text;
      linkObjects.push({ href, text });
    }
  });
  return linkObjects;
}

function updateFile(folder, filename, blob) {
  var files = folder.getFilesByName(filename);
  if (files.hasNext()) {
    var file = files.next();
    const fileId = file.getId();
    var url = 'https://www.googleapis.com/drive/v2/files/' + fileId;
    const base64Data = Utilities.base64Encode(blob.getBytes());
    var headers = {
      'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
      'Content-Type': blob.getContentType()
    };
    var options = {
      'method': 'patch',
      'headers': headers,
      'payload': JSON.stringify(blob)
    };
    UrlFetchApp.fetch(url, options);
  }else{
    var file = folder.createFile(blob).setName(filename);
  }
  return file;
}

function convertExcel2Sheets(excelFile, filename, parents = []) {
  // Combine upload parameters with content details
  const uploadParams = {
    method: 'post',
    contentType: 'application/vnd.ms-excel',
    payload: excelFile.getBytes(),
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
    },
  };

  // Upload with conversion and fetch response data in one step
  const uploadResponse = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v2/files/?uploadType=media&convert=true',
    uploadParams
  ).getContentText();
  const fileData = JSON.parse(uploadResponse);

  // Update metadata with filename and parent folders (if provided)
  const updateParams = {
    method: 'put',
    contentType: 'application/json',
    payload: JSON.stringify({
      title: filename,
      parents: parents.filter((id)=>DriveApp.getFolderById(id)).map((id) => ({ id })),
    }),
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken(),
    },
  };
  UrlFetchApp.fetch('https://www.googleapis.com/drive/v2/files/' + fileData.id, updateParams);

  // Return the converted spreadsheet
  return fileData.id;
}

function matchesFromRangeValues(rangeValues, pattern){
  let matchedCells = [];
  for (let row = 0; row < rangeValues.length-1; row++) {
    for (let col = 0; col < rangeValues[row].length-1; col++) {
      let cellValue = rangeValues[row][col];
      cellValue = cellValue.toString();
      if (cellValue && typeof cellValue === 'string') {
        const matchedValue = pattern.exec(cellValue);
        if (matchedValue) {
          matchedCells.push({
            row: row,
            column: col,
            value: matchedValue
          });
        }
      }
    }
  }
  return matchedCells;
}

function syncCalendarWithSpreadsheet(TWHldFiltedLst) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const calendarName = spreadsheet.getName();
  let calendar = findCalendarByName(calendarName);

  /*const year = TWHldFiltedLst[0]['年份'];
  const startDelDate = new Date(year, 0, 1); // 年初
  const endDelDate = new Date(year + 1, 0); // 下一年的年初午夜
  const events = calendar.getEvents(startDelDate, endDelDate); // 取得年份内的所有行程
  events.forEach(event => event.deleteEvent()); // 刪除每個行程*/

  if (!calendar) {
    calendar = CalendarApp.createCalendar(calendarName);
  }

  TWHldFiltedLst.forEach(day => {
    let title = `${day['上班或放假'] }`;
    if(day['節日'].match(/[節]/)){
      title = `${title}(${day['節日']})`
    }
    const startDate = new Date(`${day['年份']}-${day['月份']}-${day['日期']}`);
    const endDate = new Date(startDate);
    endDate.setDate(endDate.getDate() + 1); // 設置為結束日期的午夜

    const existingEvents = calendar.getEventsForDay(startDate);

    // 檢查是否有重複的事件
    const eventExists = existingEvents.some(event => event.getTitle() === title);

    if (!eventExists) {
      calendar.createAllDayEvent(title, startDate);
    }
  });
}

function findCalendarByName(name) {
  const calendars = CalendarApp.getAllOwnedCalendars();
  for (let i = 0; i < calendars.length; i++) {
    if (calendars[i].getName() === name) {
      return calendars[i];
    }
  }
  return null;
}

function printTable(data) {
  let header = Object.keys(data[0]);
  console.log(header.join('\t'));
  
  data.forEach(row => {
    let rowValues = [];
    header.forEach(col => {
      rowValues.push(row[col]);
    });
    console.log(rowValues.join('\t'));
  });
}
