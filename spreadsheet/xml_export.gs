/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Export Xml', functionName: 'exportXml_'}
  ];
  spreadsheet.addMenu('User Tools', menuItems);
  
  //exportXml_();
}

// ----------------------------------
// export XML.
// ----------------------------------
// specification:
//  - 'TABLE' : DATA TABLE keyword. from next row is DATA TABLE Begin.
//  - '*' : is check the row comment out. write cell '//', same row is comment out.
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
function exportXml_()
{
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getId();
  
  // this time is sheet[0] only.
  exportDataTableSheet2Xml_( spreadsheet.getSheets()[0] );
}

// ----------------------------------
// sheet convert to xml.
// ----------------------------------
function exportDataTableSheet2Xml_(sheet)
{
  var dataRange = sheet.getDataRange();
  var values = dataRange.getDisplayValues();

  // get Begin R/C
  var dataRangeRowBgn = 0;
  var dataRangeClmBgn = 0;
  var found_TABLE = false;
  for(var r = 0; r < values.length && !found_TABLE; ++r)
  {
      for(var c = 0; c < values[r].length; ++c)
      {
        if(values[r][c] == 'TABLE')
        {
          found_TABLE = true;
          // 'TABLE' next row.
          dataRangeRowBgn = r + 1;
          dataRangeClmBgn = c;
          break;
        }
      }
  }
  if(!found_TABLE)
  {
    return "";
  }
  
  // original perse.
  var xml = xmlPerser_(sheet.getName(), values, dataRangeRowBgn, dataRangeClmBgn);
  
  // output (to cell)
  //var outCell = sheet.getRange(1, 1);
  //outCell.setValue(xml);
  
  // output (to file)
  var fileName = sheet.getName()+'.xml';
  DriveApp.createFile(fileName, xml);
}

// ----------------------------------
// the exporter perse.
// ----------------------------------
function xmlPerser_(tableName, values, dataRangeRowBgn, dataRangeClmBgn)
{
  // xml text
  var xml = '';
  
  // perse
  xml = '<table name="'+tableName+'">';
  for(var r = dataRangeRowBgn + 1; r < values.length; ++r)
  {
    // disable?
    if(values[r][dataRangeClmBgn] == '//')
    {
      continue;
    }
    
    // enable
    xml += '<info '
    for(var c = dataRangeClmBgn + 1; c < values[r].length; ++c)
    {
      var tag = values[dataRangeRowBgn][c];
      var val = values[r][c];

      xml += tag+'="'+val+'"';
      // next attr.
      xml += ' ';
    }
    xml += '/>'
  }
  xml += '</table>'
  
  // text convert to xml file text.
  var document = XmlService.parse(xml);
  var persedXml = XmlService.getPrettyFormat().format(document);
  return persedXml;  
}