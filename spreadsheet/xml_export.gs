/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen()
{
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Export Xml',  functionName: 'exportXml_'},
    {name: 'Export Json', functionName: 'exportJson_'}
  ];
  spreadsheet.addMenu('User Tools', menuItems);
}

function exportXml_()
{
    exportData('Xml');
}

function exportJson_()
{
    exportData('Json');
}

// ----------------------------------
// Get export data position.
// ----------------------------------
// specification:
//  - 'TABLE' : DATA TABLE keyword. from next row is DATA TABLE Begin.
//  - '*'     : is check the row comment out. write cell '//', same row is comment out.
// ----------------------------------
//
// example DATA TABLE).
// |-----|-----|-----|-----|-----|
// |TABLE|     |     |     |     |
// |-----|-----|-----|-----|-----|
// |*    |ID   |INFO1|INFO2|INFO3| 
// |-----|-----|-----|-----|-----|
// |     |1    |A    |B    |C    |
// |-----|-----|-----|-----|-----|
// |//   |2    |D    |E    |F    |
// |-----|-----|-----|-----|-----|
// 
// 1).  ID=1 is enable data.
// 2).  ID=2 is disable data. skip export.
// 
// ----------------------------------
function getDataRangeBeginPosition_(values)
{
  var nameRow = 0;
  var dataRow = 0;
  var dataColumn = 0;

  var isFound = false;
  for(var r = 0; r < values.length && !isFound; ++r)
  {
      for(var c = 0; c < values[r].length; ++c)
      {
        if(values[r][c] == 'TABLE')
        {
          isFound = true;
          nameRow = r + 1; // next row is Parameter Name.
          dataRow = nameRow + 1;
          dataColumn = c;

          // skip comment rows.
          for (; isSkip(values[dataRow][dataColumn]); ++dataRow) {}
          break;
        }
      }
  }

  return {isFound, nameRow, row: dataRow, column: dataColumn};
}

// ----------------------------------
// skip row.
// ----------------------------------
function isSkip(value)
{
  return value == '*' || value == '//';
}

// ----------------------------------
// export data.
// input: type           // type for function sheetPerser_()
// input: exportAllSheet // true / false
// return: dataList & fileNameList.
// ----------------------------------
function exportData(type = 'Xml', exportAllSheet = false)
{
  var spreadsheet = SpreadsheetApp.getActive();
  var dataList = [];
  var fileNameList = [];
  if (exportAllSheet)
  {
    spreadsheet.getSheets().forEach((sheet) =>
    {
      var {data, fileName} = sheetPerser_(sheet, type);
      dataList.push(data);
      fileNameList.push(fileName);
    });
  }
  else
  {
    // this getIndex start from 1.
    var index = spreadsheet.getActiveSheet().getIndex() - 1;
    var sheet = spreadsheet.getSheets()[index];
    var {data, fileName} = sheetPerser_(sheet, type);
    dataList.push(data);
    fileNameList.push(fileName);

    // // output (to cell)
    // var outCell = sheet.getRange(1, 1);
    // outCell.setValue(data);
    
    // output (to file)
    DriveApp.createFile(fileName, data);
  }
  return {dataList, fileNameList};
}

// ----------------------------------
// Perser.
// ----------------------------------
function sheetPerser_(sheet, type = 'Xml')
{
  if (type == 'Xml')
  {
    var xml = exportDataTableSheet2Xml_(sheet);
    return {data: xml, fileName: sheet.getName() + '.xml'};
  }
  else if (type == 'Json')
  {
    var json = exportDataTableSheet2Json_(sheet);
    return {data: json, fileName: sheet.getName() + '.json'};
  }

  return {data: "", fileName: sheet.getName()};
}

// ----------------------------------
// sheet convert to xml.
// return data(string)
// ----------------------------------
function exportDataTableSheet2Xml_(sheet)
{
  var dataRange = sheet.getDataRange();
  var displayValues = dataRange.getDisplayValues();
  var {isFound, nameRow, row, column} = getDataRangeBeginPosition_(displayValues);
  if (!isFound)
  {
    return "";
  }

  return xmlPerser_(sheet.getName(), displayValues, nameRow, row, column);
}

// ----------------------------------
// the exporter perse.
// return data(string)
// ----------------------------------
function xmlPerser_(tableName, values, nameRow, dataRangeRowBgn, dataRangeClmBgn)
{
  const filterByEnableRow = (row, index) =>
    !isSkip(row[dataRangeClmBgn]) &&
    index >= dataRangeRowBgn;

  const filterByEnableColumn = (column, index) =>
    index > dataRangeClmBgn;

  var enableValues = values.filter(filterByEnableRow);
  var columnNames  = values[nameRow].filter(filterByEnableColumn);

  // perse
  var xml = '<table name="'+tableName+'">';
  for(var r = 0; r < enableValues.length; ++r)
  {
    xml += '<info '
    enableValues[r]
      .filter(filterByEnableColumn)
      .forEach((column, index) =>
        xml += columnNames[index]+'="'+column+'"' + ' '
      );
    xml += '/>'
  }
  xml += '</table>'
  
  // text convert to xml file text.
  var document = XmlService.parse(xml);
  var persedXml = XmlService.getPrettyFormat().format(document);
  return persedXml;
}

// ----------------------------------
// sheet convert to json.
// return data(string)
// ----------------------------------
function exportDataTableSheet2Json_(sheet)
{
  var dataRange = sheet.getDataRange();
  var displayValues = dataRange.getDisplayValues();
  var {isFound, nameRow, row, column} = getDataRangeBeginPosition_(displayValues);
  if (!isFound)
  {
    return "";
  }

  return jsonPerser_(sheet.getName(), displayValues, nameRow, row, column);
}

// ----------------------------------
// sheet convert to json.
// return data(string)
// ----------------------------------
function jsonPerser_(tableName, values, nameRow, dataRangeRowBgn, dataRangeClmBgn)
{
  const filterByEnableRow = (row, index) =>
    !isSkip(row[dataRangeClmBgn]) &&
    index >= dataRangeRowBgn;

  const filterByEnableColumn = (column, index) =>
    index > dataRangeClmBgn;

  var enableValues = values.filter(filterByEnableRow);
  var columnNames  = values[nameRow].filter(filterByEnableColumn);

  // convert column: Array To Object
  for(var r = 0; r < enableValues.length; ++r)
  {
    enableValues[r] = enableValues[r]
      .filter(filterByEnableColumn)
      .reduce((pre, cur, i) =>
        (pre[columnNames[i]] = cur, pre), {}
      );
  }
  
  // perse
  var jsonReplacer = null;
  var jsonSpace = 4;
  var json = JSON.stringify(
    {
      name: tableName,
      list: enableValues,
    },
    jsonReplacer,
    jsonSpace
  );
  return json;
}

