function onlyUnique(value, index, self) {
  return self.indexOf(value) === index;
}

function createFinancialReport() {
  var sheet = SpreadsheetApp.getActiveSheet()
  var lastRow = 100;
  var numColumns = 10
  var result = sheet.getRange("List!D2:D").getValues()
  var cards = sheet.getRange("List!B2:B").getValues()
  var numRows = 100
  Logger.log(numRows)
  var cardsUnique = sheet.getRange(2,2,lastRow,1).getValues().join().split(",").filter(onlyUnique)
  Logger.log(cardsUnique)
  var i = 0
  var dict = {
    'Яндекс.Драйв':'Каршеринг',
    'Яндекс Такси':'Такси',
    'Штрафы ГИБДД': 'Платные дороги, штрафы',
    'Супермаркеты': 'Питание (до расчётов)',
    'Фастфуд': 'Питание (до расчётов)',
    'Рестораны': 'Питание (до расчётов)',
    'Жкх': 'Коммунальные платежи',
    'Московский транспорт': 'Общественный транспорт'
  }
  var cardDict = {
    '*1422':'cred',
    '*6138': 'cred',
    '*3705': 'cred',
    '*6925': 'yandex',
    '*7074': 'debet',
    '*5986': 'debet',
    '*3443': 'schet',
    '*7000': 'debet',
    '*9891': 'debet',
    '*8711': 'mobile',
    '': 'empty'
  }
  while (result[i] != '  ') {
    for (key in dict) {
      if (result[i][0].indexOf(key) > -1){
        sheet.getRange(i+2,5).setValue(dict[key])
        Logger.log(result[i][0])
      }
    }
    i = i +1
  }
  i = 0
  Logger.log("Starting card nums processing")
  while (i < numRows) {
    if (cards[i][0] in cardDict) {
      Logger.log(cards[i][0])
      Logger.log(cardDict[cards[i][0]])
      sheet.getRange(i+2,2).setValue(cardDict[cards[i][0]])   
    }
    i = i +1
  }
    var cardsUnique = sheet.getRange(2,2,lastRow,1).getValues().join().split(",").filter(onlyUnique)
    var i = 0
    for (cardN in cardsUnique) {
      var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      var yourNewSheet = activeSpreadsheet.getSheetByName("Name of your new sheet");
      if (yourNewSheet != null) {
        activeSpreadsheet.deleteSheet(yourNewSheet);
        }
      yourNewSheet = activeSpreadsheet.insertSheet();
      if (cardsUnique[cardN] == '') {
        yourNewSheet.setName("empty");
        Logger.log(cardsUnique[cardN].toString())
        } else
          {
            yourNewSheet.setName(cardsUnique[cardN]);
            }
      }
  var cards = sheet.getRange("List!B2:B").getValues()
  i = 0
  Logger.log("Starting moving rows")

    
  while (i < numRows) {
    Logger.log(cards[i][0])
    var targetSheet = activeSpreadsheet.getSheetByName(cards[i][0]);
      var source = sheet.getRange(i+2,1,1,numColumns).getValues()
      Logger.log(source)
      var sourceLine = Array.from({ length: 100 })
      for (let i = 0; i <=numColumns; i++) {
        sourceLine[i] = source[0][i]
      }
    Logger.log(sourceLine)
    targetSheet.appendRow(sourceLine);
    i = i + 1
  }

}

