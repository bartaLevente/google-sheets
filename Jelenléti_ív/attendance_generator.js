//Adott dátum, adott hónap első és utolsó napja
const currDate = new Date(2024,0,12);
const lastDay = new Date(currDate.getFullYear(), currDate.getMonth() + 1, 0);
const firstDay = new Date(currDate.getFullYear(), currDate.getMonth(), 1);

//Munkaszüneti napok
//78b6ffb10241d0a3146b008db75262aacd23fbe625b3f37fe09d782b8a93596f
const key = '109b21dcd1528204f5d61f3d12dacdcc1964421211d507b40a73d94a362fd2da';
const year = currDate.getFullYear();
const url = 'https://szunetnapok.hu/api/' + key + '/' + year + '/'
const response = UrlFetchApp.fetch(url, { 'muteHttpExceptions': true });
const json = response.getContentText();
const jsonData = JSON.parse(json);
const holidays = jsonData['days'];

//nem munkanapok: munkaszüneti + (péntek-vasárnap)
const notWorkdays = getNotWorkDays();

const company = 'Cubicfox Kft.';

function main(){
  //dolgozók adatai
  const people = getPeopleInfo();

  //Új Spreadsheet dokumentum létrehozása 'év_hónap_jelenléti' néven 
  const newMonthName = currDate.getFullYear() + '_' + (currDate.getMonth() + 1) + '_jelenléti';
  const newMonthSpreadsheet = SpreadsheetApp.create(newMonthName);

  for(person of people){
    //Dolgozó nevével új sheet létrehozása, amennyiben még nem létezik ilyen
    const newPersonTitle = person.name;
    if (newMonthSpreadsheet.getSheetByName(newPersonTitle)) {
      Logger.log('Sheet with the name ' + newPersonTitle + ' already exists.');
    }
    newMonthSpreadsheet.insertSheet(newPersonTitle);

    //dolgozó szabdságinak meghatározása
    let dayOffs;
    try{
      dayOffs = getDayoffs(person);
    }catch(error){
      Logger.log(error.message);
      return;
    }
    //sheet előkészítése a dolgozó adataival
    setupSheet(person,company,newMonthSpreadsheet);

    //végleges tömb meghatározása egy hónapra
    const attendanceByMonth = getMonthAttandence(person,dayOffs);

    //adatok kirajzolása 2 oszlopban
    drawAttendence(true, person, newMonthSpreadsheet,attendanceByMonth);
    drawAttendence(false, person, newMonthSpreadsheet,attendanceByMonth);
  }

  //a létrehozáskor az alapértelmezetten létrhozott sheet törlése
  const defaultSheet = newMonthSpreadsheet.getSheetByName('Sheet1');
  if (defaultSheet) {
    newMonthSpreadsheet.deleteSheet(defaultSheet);
  }
}

//Dolgozók adatainak kigyűjtése az 'adatok' munkalapról
function getPeopleInfo(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('adatok');
  const data = sheet.getDataRange().getValues();
  data.shift(); //fejlécek
  const result = [];
  for(row of data){
    const formatedStart = row[2].toLocaleString('hu-HU', { hour: '2-digit', minute: '2-digit', hour12: false });
    const formatedEnd = row[3].toLocaleString('hu-HU', { hour: '2-digit', minute: '2-digit', hour12: false });
    //minden dolgozónak egy person objektum
    const person = {
      name:   row[0],
      email:  row[1],
      start:  formatedStart,
      end:    formatedEnd,
      hours:  row[4]
    };
    result.push(person);
  }
  return result;
}

function setupSheet(person,companyName, spreadsheet) {
  const sheet = spreadsheet.getSheetByName(person.name);
  sheet.clearContents();

  //Cím
  const title = sheet.getRange('A1:I1');
  title.merge();
  title.setValue('Jelenléti ív');
  title.setHorizontalAlignment('center').setFontWeight('bold').setFontSize(14);

  // Munkáltató adatai
  const employerrow = sheet.getRange('A4:C4');
  employerrow.setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange('A4').setValue('Munkáltató:');
  const employer = sheet.getRange('B4:C4');
  employer.merge();
  employer.setValue(companyName);
  employer.setHorizontalAlignment('center').setFontWeight('bold');

  //Munkavállaló adatai
  const employeerow1 = sheet.getRange('A6:C6');
  employeerow1.setHorizontalAlignment('center').setVerticalAlignment('middle').setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange('A6').setValue('Munkavállaló neve:');
  const employee = sheet.getRange('C6');
  employee.setValue(person.name);
  employee.setFontWeight('bold');
  sheet.autoResizeColumn(employee.getColumn()); // C and H is not equal

  const employeerow2 = sheet.getRange('D6:G6');
  employeerow2.setHorizontalAlignment('center').setVerticalAlignment('middle').setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange('D6').setValue(currDate.getFullYear());
  sheet.getRange('E6').setValue('év').setFontWeight('bold');
  sheet.getRange('F6').setValue(currDate.toLocaleString('hu-HU', { month: 'long' }));
  sheet.getRange('G6').setValue('hónap').setFontWeight('bold');

  
  const employerrow3 = sheet.getRange('H6');
  employerrow3.setValue('Kelt.:' + Utilities.formatDate(lastDay,Session.getScriptTimeZone(),'yyyy.MM.dd'));
  employerrow3.setHorizontalAlignment('center').setVerticalAlignment('middle').setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  //Aláírás
  const signatureText = sheet.getRange('F54:F55');
  signatureText.merge();
  signatureText.setValue('Aláírás:');
  signatureText.setHorizontalAlignment('center').setVerticalAlignment('middle').setFontWeight('bold');
  const signature = sheet.getRange('G54:I55');
  signature.merge();
  signature.setBorder(true, true, true, true, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

}

function getDayoffs(person) {
  //szabadságokat tartalmazó sheet adatainak beolvasása
  const dayOffSpreadSheetId = '1I7jCo_hC7VhZg8R6O6zOwAgAhWqLLb63s-9RFZrypWE';
  const dayOffSpreadSheetName = 'Sheet1';
  const dayOffSpreadSheet = SpreadsheetApp.openById(dayOffSpreadSheetId);
  const dayOffSheet = dayOffSpreadSheet.getSheetByName(dayOffSpreadSheetName);

  //hiba, ha nem sikerült megnyitni
  if (!dayOffSheet) {
    throw new Error('Source sheet not found.');
  }

  const dayOffData = dayOffSheet.getDataRange().getValues();

  const result = [];

  //kiszűri az adott dologzóhoz, adott hónaphoz tartozó 'Added'-del ellátott sorokat
  const filteredData = dayOffData.filter(row => {
    const emailMatches = row[1] === person.email;
    const added = row[6] === 'Added';
    const start = new Date(row[2]);
    const end = new Date(row[3]);
    const includesCurrMonth = (start <= lastDay) && (end >= firstDay);
    return emailMatches && added && includesCurrMonth;
  });

  //a result tömb feltöltése a szabadnapok dátumának napjával
  for (let row of filteredData) {

    const startDate = new Date(row[2]);
    const endDate = new Date(row[3]);

    let startDay = startDate.getDate();
    let endDay = endDate.getDate();

    //ha a szabadság a hónap első napja előtt kezdődött, akkor a startDay 1
    if(startDate <= firstDay){
      startDay = 1;
    }

    //ha a szabadság a hónap utolsó napja után ért véget, akkor az endDay a hónap utolsó napja
    if(endDate >= lastDay){
      endDay = lastDay.getDate();
    }
    //a result tömbbe startDaytől endDayig bekerülnek a dátumok napjai
    for (let i = startDay; i < (endDay + 1); i++) {
      result.push(i);
    }
  }
  return result;
}

function getSickLeaves(){
  return;
}

function getNotWorkDays(){
  
  const result = [];

  //az első péntek meghatározása
  var firstFriday = new Date(firstDay);
  while(firstFriday.getDay() !== 5){
    firstFriday.setDate(firstFriday.getDate() + 1);
  }
  
  //az összes péntek - szombat - vasárnap dátumát hozzáadja a result tömbhöz
  var currDay = new Date(firstFriday)
  while(currDay.getMonth() === currDate.getMonth()){
    
    var currWeekendDay = new Date(currDay);
    var weekendCount = 1;
    while(currDay.getMonth() === currWeekendDay.getMonth() && weekendCount <= 3){
      result.push(currWeekendDay.getDate());
      currWeekendDay.setDate(currWeekendDay.getDate() + 1);
      weekendCount ++;
    }
    currDay.setDate(currDay.getDate() + 7);
  }

  //az adott hónaphoz tartozó munkaszüneti napok is bekerülnek a resultba - amennyiben még nincsenek benne
  for(row of holidays){
    const holiDate = row.date;
    const parts = holiDate.split('-');
    const holiMonth = parseInt(parts[1]);
    const holiDay = parseInt(parts[2]);
    if(holiMonth === (currDate.getMonth() + 1) && !result.includes(holiDay)){
      result.push(holiDay);
    }
  }
  return result;
}

function getMonthAttandence(person,dayOffs){
  const length = lastDay.getDate();
  const defaultValue = person.hours;
  const result = new Array(length).fill(defaultValue);

  for(element of dayOffs){
    result[element-1] = 'FSZ';
  }
  for(element of notWorkdays){
    result[element-1] = '';
  }
  return result;
}

function drawAttendence(firstCol, person, spreadsheet, attendance){
  const sheet = spreadsheet.getSheetByName(person.name);
  var currentRow = 8;
  var startCol = 'A';
  var endCol = 'D';
  var startIndex = 0;
  var endIndex = 16;
  if(!firstCol){
    startCol = 'F';
    endCol = 'I';
    startIndex = 16;
    endIndex = attendance.length;
  }
  
  for(let i = startIndex; i <endIndex ; i++){
    //Egy naphoz tartozó blokk
    const section = sheet.getRange(startCol + currentRow + ':' + endCol + (currentRow+2));
    section.setHorizontalAlignment('center').setVerticalAlignment('middle').setBorder(true, true, true, true, true, true, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    //Fejlécek feliratai
    const dayText = sheet.getRange(startCol + currentRow + ':' + String.fromCharCode(startCol.charCodeAt(0) + 2) + currentRow);
    dayText.merge();
    dayText.setValue('Nap');
    sheet.getRange(endCol + currentRow).setValue('Ledolgozott óra');

    //Nap sorszáma
    const dayOfMonthText = sheet.getRange(startCol + (currentRow+1) + ':' + startCol + (currentRow+2));
    dayOfMonthText.merge();
    dayOfMonthText.setValue(i+1).setFontWeight('bold');

    //Érkezési - Távozási adatok
    sheet.getRange(String.fromCharCode(startCol.charCodeAt(0) + 1) + (currentRow+1)).setValue('Érkezett');
    sheet.getRange(String.fromCharCode(startCol.charCodeAt(0) + 1) + (currentRow+2)).setValue('Távozott');
    sheet.getRange(String.fromCharCode(startCol.charCodeAt(0) + 2) + (currentRow+1)).setValue(person.start);
    sheet.getRange(String.fromCharCode(startCol.charCodeAt(0) + 2) + (currentRow+2)).setValue(person.end);

    //Ledogozott óra - Szabadság - Táppénz
    const work = sheet.getRange(endCol + (currentRow+1) + ':' + endCol + (currentRow+2));
    work.merge();
    work.setValue(attendance[i]).setFontSize(14).setFontWeight('bold');
    currentRow += 3;
  }
}
