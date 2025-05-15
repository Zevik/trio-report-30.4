const SPREADSHEET_ID = '1UxQn7mAinamXXZ6WuK0Zp8aRdfYqXCQ6mf-n4fYVZ8c';
const SHEET_NAME = 'Shift_card';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('דוח שעות חודשי');
}

function getAvailableMonths() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1).getValues(); // עמודה A

  const uniqueMonths = new Set();

  for (let i = 0; i < data.length; i++) {
    const cell = data[i][0];
    let date;

    if (Object.prototype.toString.call(cell) === '[object Date]') {
      date = cell;
    } else {
      date = parseDate(cell);
    }

    if (!date || isNaN(date.getTime())) continue;

    const monthKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
    uniqueMonths.add(monthKey);
  }

  const sorted = Array.from(uniqueMonths).sort().reverse();
  Logger.log("📅 חודשים זמינים: " + JSON.stringify(sorted));
  return sorted;
}


function getMonthlyReport(monthKey) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  
  // סוגי משמרת
  const SHIFT_TYPES = ['רפואה שלמה', 'דמו', 'הכשרה', 'מיזם טריו'];
  const LOCATIONS = ['בית', 'מרפאה'];
  
  // { "שם רפואן": { 
  //   "רפואה שלמה": { "בית": { minutes: 0, shifts: 0 }, "מרפאה": { minutes: 0, shifts: 0 } },
  //   "דמו": { "בית": { minutes: 0, shifts: 0 }, "מרפאה": { minutes: 0, shifts: 0 } },
  //   ... וכן הלאה
  // }
  const results = {}; 

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowNum = i + 1;

    const rawDate = row[0];         // עמודה A - תאריך
    const name = row[1];            // עמודה B - רפואן/סטודנט
    const shiftType = row[2];       // עמודה C - סוג משמרת
    const durationI = row[8];       // עמודה I - משך משמרת ידני (עדיפות גבוהה)
    const durationH = row[7];       // עמודה H - משך משמרת מחושב
    const location = row[9] || '';  // עמודה J - מיקום המשמרת (בית/מרפאה)
    
    // אם יש ערך בעמודה I, נשתמש בו, אחרת נשתמש בעמודה H
    const duration = durationI || durationH;
    
    const date = parseDate(rawDate);

    if (!date || !name) {
      Logger.log(`⛔ שורה ${rowNum}: אין תאריך או שם — דילוג`);
      continue;
    }

    if (!shiftType) {
      Logger.log(`⛔ שורה ${rowNum}: אין סוג משמרת — דילוג`);
      continue;
    }
    
    // בדיקה שסוג המשמרת תקין
    if (!SHIFT_TYPES.includes(shiftType)) {
      Logger.log(`⚠️ שורה ${rowNum}: סוג משמרת לא מוכר: ${shiftType}`);
      // נמשיך בכל זאת כדי לא לאבד מידע
    }
    
    // בדיקה שהמיקום תקין
    const normalizedLocation = location.trim();
    if (normalizedLocation && !LOCATIONS.includes(normalizedLocation)) {
      Logger.log(`⚠️ שורה ${rowNum}: מיקום משמרת לא מוכר: ${normalizedLocation}`);
      // נמשיך בכל זאת
    }
    
    // מיקום ברירת מחדל אם לא צוין
    const effectiveLocation = normalizedLocation || 'לא צוין';

    const rowMonthKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
    if (rowMonthKey !== monthKey) continue;

    let minutes = 0;

    try {
      const type = Object.prototype.toString.call(duration);

      if (!duration) {
        Logger.log(`⛔ שורה ${rowNum}: ערך ריק בעמודות I ו-H`);
        continue;
      }

      if (type === '[object Date]') {
        const h = duration.getHours();
        const m = duration.getMinutes();
        minutes = h * 60 + m;
        Logger.log(`✅ שורה ${rowNum}: Date ${h}:${m} → ${minutes} דקות`);
      }

      else if (typeof duration === 'number') {
        minutes = Math.round(duration * 24 * 60);
        Logger.log(`✅ שורה ${rowNum}: Number ${duration} → ${minutes} דקות`);
      }

      else if (typeof duration === 'string') {
        const parts = duration.trim().split(':');
        const h = parseInt(parts[0]) || 0;
        const m = parseInt(parts[1]) || 0;
        minutes = h * 60 + m;
        Logger.log(`✅ שורה ${rowNum}: String "${duration}" → ${minutes} דקות`);
      }

      else {
        Logger.log(`⛔ שורה ${rowNum}: סוג לא מזוהה (${type})`);
        continue;
      }

    } catch (e) {
      Logger.log(`💥 שורה ${rowNum}: שגיאה בעיבוד משך (${duration}) → ${e.message}`);
      continue;
    }

    // ייצור מבנה נתונים עבור הרפואן אם לא קיים
    if (!results[name]) {
      results[name] = {};
      
      // אתחול המבנה עם כל סוגי המשמרות והמיקומים
      for (const type of SHIFT_TYPES) {
        results[name][type] = {};
        for (const loc of LOCATIONS) {
          results[name][type][loc] = { minutes: 0, shifts: 0 };
        }
      }
    }
    
    // הוספת זמן ומשמרת למבנה הנתונים
    if (!results[name][shiftType]) {
      results[name][shiftType] = {};
      for (const loc of LOCATIONS) {
        results[name][shiftType][loc] = { minutes: 0, shifts: 0 };
      }
    }
    
    if (!results[name][shiftType][effectiveLocation]) {
      results[name][shiftType][effectiveLocation] = { minutes: 0, shifts: 0 };
    }
    
    // הוספת הדקות והמשמרות
    results[name][shiftType][effectiveLocation].minutes += minutes;
    results[name][shiftType][effectiveLocation].shifts += 1;
  }

  // עיבוד התוצאות למבנה הדוח הסופי
  const report = {
    monthKey,
    monthName: getMonthName(monthKey),
    data: []
  };

  // מיפוי המידע לפורמט של הדוח
  for (const [name, shiftTypes] of Object.entries(results)) {
    const studentData = { name };
    
    // סיכום שעות ומשמרות לפי סוג
    for (const shiftType of SHIFT_TYPES) {
      if (!shiftTypes[shiftType]) continue;
      
      studentData[`${shiftType}_hours`] = {};
      studentData[`${shiftType}_shifts`] = {};
      
      let totalMinutes = 0;
      let totalShifts = 0;
      
      for (const location of LOCATIONS) {
        const data = shiftTypes[shiftType][location];
        if (data) {
          studentData[`${shiftType}_hours`][location] = formatMinutesToHHMM(data.minutes);
          studentData[`${shiftType}_shifts`][location] = data.shifts;
          totalMinutes += data.minutes;
          totalShifts += data.shifts;
        } else {
          studentData[`${shiftType}_hours`][location] = '00:00';
          studentData[`${shiftType}_shifts`][location] = 0;
        }
      }
      
      // סה"כ לכל סוג משמרת
      studentData[`${shiftType}_total_hours`] = formatMinutesToHHMM(totalMinutes);
      studentData[`${shiftType}_total_shifts`] = totalShifts;
    }
    
    // חישוב סה"כ כללי
    let grandTotalMinutes = 0;
    let grandTotalShifts = 0;
    
    for (const shiftType of SHIFT_TYPES) {
      for (const location of LOCATIONS) {
        if (shiftTypes[shiftType] && shiftTypes[shiftType][location]) {
          grandTotalMinutes += shiftTypes[shiftType][location].minutes;
          grandTotalShifts += shiftTypes[shiftType][location].shifts;
        }
      }
    }
    
    studentData.total_hours = formatMinutesToHHMM(grandTotalMinutes);
    studentData.total_shifts = grandTotalShifts;
    
    report.data.push(studentData);
  }
  
  // מיון לפי שם
  report.data.sort((a, b) => a.name.localeCompare(b.name));
  
  Logger.log('✅ דוח סופי:');
  report.data.forEach(student => Logger.log(`${student.name}: ${student.total_hours}, משמרות: ${student.total_shifts}`));
  
  return report;
}


function parseDate(value) {
  if (Object.prototype.toString.call(value) === '[object Date]') return value;
  try {
    const parts = value.toString().split(' ')[0].split('/');
    return new Date(`${parts[2]}-${parts[1]}-${parts[0]}`);
  } catch (e) {
    Logger.log(`שגיאה בפענוח תאריך: ${value}`);
    return null;
  }
}

function formatMinutesToHHMM(mins) {
  const hours = Math.floor(mins / 60);
  const minutes = mins % 60;
  return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
}

function getMonthName(monthKey) {
  const [year, month] = monthKey.split('-');
  
  const monthNames = {
    '01': 'ינואר',
    '02': 'פברואר',
    '03': 'מרץ',
    '04': 'אפריל',
    '05': 'מאי',
    '06': 'יוני',
    '07': 'יולי',
    '08': 'אוגוסט',
    '09': 'ספטמבר',
    '10': 'אוקטובר',
    '11': 'נובמבר',
    '12': 'דצמבר'
  };
  
  return `${monthNames[month]} ${year}`;
}

/**
 * יוצר קובץ אקסל מהדוח ומחזיר URL להורדה
 * משתמש בגיליון קבוע במקום ליצור חדש בכל פעם
 * @param {Object} report - דוח השעות המלא
 * @return {string} URL לקובץ האקסל
 */
function createExcelFile(report) {
  const FIXED_SPREADSHEET_ID = '1I_3XUG7FHR-SNUZ0MDOOpioCRl3TjmkzxO4QYICb-UM';
  const spreadsheet = SpreadsheetApp.openById(FIXED_SPREADSHEET_ID);
  const sheet = spreadsheet.getActiveSheet();
  
  // מחיקת כל התוכן הקודם בגיליון
  sheet.clear();
  
  // יצירת כותרת הדוח
  sheet.getRange("A1").setValue(`דוח שעות עבודה רפואנים לחודש ${report.monthName}`);
  sheet.getRange("A1:K1").merge();
  
  // עיצוב כותרת
  sheet.getRange("A1:K1").setFontWeight("bold");
  sheet.getRange("A1:K1").setHorizontalAlignment("center");
  sheet.getRange("A1:K1").setFontSize(14);
  
  // כותרות עמודות
  const headers = [
    ["שם", "רפואה שלמה", "", "מיזם טריו", "", "דמו", "", "הכשרה", "", "מיקום משמרת", ""],
    ["", "שעות", "משמרות", "שעות", "משמרות", "שעות", "משמרות", "שעות", "משמרות", "בית", "מרפאה"]
  ];
  
  // הכנסת כותרות
  sheet.getRange(3, 1, 2, 11).setValues(headers);
  sheet.getRange(3, 1, 2, 11).setFontWeight("bold");
  sheet.getRange(3, 1, 2, 11).setHorizontalAlignment("center");
  sheet.getRange(3, 1, 2, 11).setBackground("#f0f0f0");
  
  // מיזוג תאים עבור כותרות משנה
  sheet.getRange("B3:C3").merge();  // רפואה שלמה
  sheet.getRange("D3:E3").merge();  // מיזם טריו
  sheet.getRange("F3:G3").merge();  // דמו
  sheet.getRange("H3:I3").merge();  // הכשרה
  sheet.getRange("J3:K3").merge();  // מיקום משמרת
  sheet.getRange("A3:A4").merge();  // שם
  
  // הכנסת נתוני הדוח
  let rowIndex = 5;
  report.data.forEach(student => {
    const row = [student.name];
    
    // רפואה שלמה
    addShiftTypeToRow(row, student, 'רפואה שלמה');
    
    // מיזם טריו
    addShiftTypeToRow(row, student, 'מיזם טריו');
    
    // דמו
    addShiftTypeToRow(row, student, 'דמו');
    
    // הכשרה
    addShiftTypeToRow(row, student, 'הכשרה');
    
    // מיקום משמרת
    row.push(getLocationShifts(student, 'בית'));
    row.push(getLocationShifts(student, 'מרפאה'));
    
    sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    rowIndex++;
  });
  
  // התאמת רוחב עמודות
  for (let i = 1; i <= 15; i++) {
    sheet.autoResizeColumn(i);
  }
  
  // כיוון טקסט מימין לשמאל
  sheet.setRightToLeft(true);
  
  // יצירת קובץ לייצוא - שימוש ב-ID הקבוע
  const url = `https://docs.google.com/spreadsheets/d/${FIXED_SPREADSHEET_ID}/export?format=xlsx`;
  
  return url;
}

/**
 * מוסיף את נתוני סוג המשמרת לשורת אקסל
 */
function addShiftTypeToRow(row, student, shiftType) {
  const key = `${shiftType}_hours`;
  const shiftKey = `${shiftType}_shifts`;
  const totalKey = `${shiftType}_total_hours`;
  const totalShiftsKey = `${shiftType}_total_shifts`;
  
  if (student[key] && student[totalKey]) {
    // רק אם יש ערך שאינו אפס, נציג אותו
    row.push(student[totalKey] !== '00:00' ? student[totalKey] : ''); // שעות
    row.push(student[totalShiftsKey] > 0 ? student[totalShiftsKey] : ''); // משמרות
  } else {
    row.push('');
    row.push('');
  }
}

/**
 * מחזיר את מספר המשמרות לפי מיקום
 */
function getLocationShifts(student, location) {
  let total = 0;
  const shiftTypes = ['רפואה שלמה', 'דמו', 'הכשרה', 'מיזם טריו'];
  
  for (const type of shiftTypes) {
    const key = `${type}_shifts`;
    if (student[key] && student[key][location]) {
      total += student[key][location];
    }
  }
  
  // רק אם יש ערך שאינו אפס, נציג אותו
  return total > 0 ? total : '';
}

/**
 * יוצר קובץ CSV מהדוח ומחזיר URL להורדה
 * @param {Object} report - דוח השעות המלא
 * @return {string} URL לקובץ ה-CSV
 */
function createCSVFile(report) {
  let csv = `"דוח שעות עבודה רפואנים לחודש ${report.monthName}"\n\n`;
  
  // כותרות
  csv += '"שם","רפואה שלמה - שעות","רפואה שלמה - משמרות","מיזם טריו - שעות","מיזם טריו - משמרות","דמו - שעות","דמו - משמרות","הכשרה - שעות","הכשרה - משמרות","מיקום - בית","מיקום - מרפאה"\n';
  
  // נתוני הדוח
  report.data.forEach(student => {
    let row = [student.name];
    
    // רפואה שלמה
    addShiftTypeToCSVRow(row, student, 'רפואה שלמה');
    
    // מיזם טריו
    addShiftTypeToCSVRow(row, student, 'מיזם טריו');
    
    // דמו
    addShiftTypeToCSVRow(row, student, 'דמו');
    
    // הכשרה
    addShiftTypeToCSVRow(row, student, 'הכשרה');
    
    // מיקום משמרת
    row.push(getLocationShifts(student, 'בית'));
    row.push(getLocationShifts(student, 'מרפאה'));
    
    csv += row.map(item => `"${item}"`).join(',') + '\n';
  });
  
  // שמירת ה-CSV כקובץ בדרייב ויצירת קישור להורדה
  const fileName = `דוח שעות רפואנים - ${report.monthName}.csv`;
  const file = DriveApp.createFile(fileName, csv, MimeType.CSV);
  
  return file.getDownloadUrl();
}

/**
 * מוסיף את נתוני סוג המשמרת לשורת CSV
 */
function addShiftTypeToCSVRow(row, student, shiftType) {
  const key = `${shiftType}_hours`;
  const shiftKey = `${shiftType}_shifts`;
  const totalKey = `${shiftType}_total_hours`;
  const totalShiftsKey = `${shiftType}_total_shifts`;
  
  if (student[key] && student[totalKey]) {
    // רק אם יש ערך שאינו אפס, נציג אותו
    row.push(student[totalKey] !== '00:00' ? student[totalKey] : ''); // שעות
    row.push(student[totalShiftsKey] > 0 ? student[totalShiftsKey] : ''); // משמרות
  } else {
    row.push('');
    row.push('');
  }
}

/**
 * פותח את גיליון Google Sheets הקבוע עם הדוח ומחזיר URL לפתיחתו
 * @param {Object} report - דוח השעות המלא
 * @return {string} URL לגיליון
 */
function createGoogleSheet(report) {
  // משתמש באותו גיליון קבוע כמו בפונקציית createExcelFile
  // מחזיר URL ישיר אליו
  
  const FIXED_SPREADSHEET_ID = '1I_3XUG7FHR-SNUZ0MDOOpioCRl3TjmkzxO4QYICb-UM';
  const spreadsheet = SpreadsheetApp.openById(FIXED_SPREADSHEET_ID);
  const sheet = spreadsheet.getActiveSheet();
  
  // מחיקת כל התוכן הקודם בגיליון
  sheet.clear();
  
  // יצירת כותרת הדוח
  sheet.getRange("A1").setValue(`דוח שעות עבודה רפואנים לחודש ${report.monthName}`);
  sheet.getRange("A1:K1").merge();
  
  // עיצוב כותרת
  sheet.getRange("A1:K1").setFontWeight("bold");
  sheet.getRange("A1:K1").setHorizontalAlignment("center");
  sheet.getRange("A1:K1").setFontSize(14);
  
  // כותרות עמודות
  const headers = [
    ["שם", "רפואה שלמה", "", "מיזם טריו", "", "דמו", "", "הכשרה", "", "מיקום משמרת", ""],
    ["", "שעות", "משמרות", "שעות", "משמרות", "שעות", "משמרות", "שעות", "משמרות", "בית", "מרפאה"]
  ];
  
  // הכנסת כותרות
  sheet.getRange(3, 1, 2, 11).setValues(headers);
  sheet.getRange(3, 1, 2, 11).setFontWeight("bold");
  sheet.getRange(3, 1, 2, 11).setHorizontalAlignment("center");
  sheet.getRange(3, 1, 2, 11).setBackground("#f0f0f0");
  
  // מיזוג תאים עבור כותרות משנה
  sheet.getRange("B3:C3").merge();  // רפואה שלמה
  sheet.getRange("D3:E3").merge();  // מיזם טריו
  sheet.getRange("F3:G3").merge();  // דמו
  sheet.getRange("H3:I3").merge();  // הכשרה
  sheet.getRange("J3:K3").merge();  // מיקום משמרת
  sheet.getRange("A3:A4").merge();  // שם
  
  // הכנסת נתוני הדוח
  let rowIndex = 5;
  report.data.forEach(student => {
    const row = [student.name];
    
    // רפואה שלמה
    addShiftTypeToRow(row, student, 'רפואה שלמה');
    
    // מיזם טריו
    addShiftTypeToRow(row, student, 'מיזם טריו');
    
    // דמו
    addShiftTypeToRow(row, student, 'דמו');
    
    // הכשרה
    addShiftTypeToRow(row, student, 'הכשרה');
    
    // מיקום משמרת
    row.push(getLocationShifts(student, 'בית'));
    row.push(getLocationShifts(student, 'מרפאה'));
    
    sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    rowIndex++;
  });
  
  // התאמת רוחב עמודות
  for (let i = 1; i <= 15; i++) {
    sheet.autoResizeColumn(i);
  }
  
  // כיוון טקסט מימין לשמאל
  sheet.setRightToLeft(true);
  
  // מחזיר את הקישור לגיליון הקבוע
  return spreadsheet.getUrl();
}
