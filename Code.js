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

  const results = {}; // { "שם סטודנט": סך בדקות }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rawDate = row[0];    // עמודה A - תאריך
    const name = row[1];       // עמודה B - סטודנט
    const durationI = row[8];  // עמודה I - משך (עדיפות גבוהה)
    const durationH = row[7];  // עמודה H - משך
    // אם יש ערך בעמודה I, נשתמש בו, אחרת נשתמש בעמודה H
    const duration = durationI || durationH;

    const date = parseDate(rawDate);
    const rowNum = i + 1;

    if (!date || !name) {
      Logger.log(`⛔ שורה ${rowNum}: אין תאריך או שם — דילוג`);
      continue;
    }

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

    if (!results[name]) results[name] = 0;
    results[name] += minutes;
  }

  const formatted = Object.entries(results).map(([name, mins]) => ({
    name,
    total: formatMinutesToHHMM(mins),
  }));

  Logger.log('✅ דוח סופי:');
  formatted.forEach(r => Logger.log(`${r.name}: ${r.total}`));

  return formatted.sort((a, b) => a.name.localeCompare(b.name));
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
