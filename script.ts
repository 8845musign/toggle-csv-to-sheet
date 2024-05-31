// 型定義をインポートします。
/// <reference types="google-apps-script" />

// 日付をフォーマットする関数
function formatDate(date: Date): string {
  const year = date.getFullYear();
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const day = date.getDate().toString().padStart(2, '0');
  return `${year}-${month}-${day}`;
}

// 日付と時間を正規化する関数
const normalizeDateTime = (date: Date, time: Date): Date => {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), time.getHours(), time.getMinutes(), time.getSeconds());
}

// 時間をフォーマットする関数
function formatTime(date: Date): string {
  const hours = date.getHours().toString();
  const minutes = date.getMinutes().toString().padStart(2, '0');
  return `${hours}:${minutes}`;
}

// 指定された月の全ての日付を取得する関数
function getDatesOfMonth(year: number, month: number): string[] {
  const dates: string[] = [];
  let date = new Date(year, month - 1, 1);
  while (date.getMonth() === month - 1) {
    const yearStr = date.getFullYear().toString();
    const monthStr = ('0' + (date.getMonth() + 1)).slice(-2);
    const dayStr = ('0' + date.getDate()).slice(-2);
    dates.push(`${yearStr}-${monthStr}-${dayStr}`);
    date.setDate(date.getDate() + 1);
  }
  return dates;
}

// ミリ秒をHH:MM形式に変換する関数
function convertMillisecondsToHHMM(milliseconds: number): string {
  const hours = Math.floor(milliseconds / 3600000);
  const minutes = Math.floor((milliseconds % 3600000) / 60000);
  const formattedHours = hours.toString().padStart(2, '0');
  const formattedMinutes = minutes.toString().padStart(2, '0');
  return `${formattedHours}:${formattedMinutes}`;
}

// メイン関数
function main(): void {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("シート1");
  const outputSheet = ss.getSheetByName("シート2");

  if (!inputSheet || !outputSheet) {
    throw new Error("指定されたシートが見つかりません");
  }

  const inputData = inputSheet.getRange(`F2:K${inputSheet.getLastRow()}`).getValues();
  
  interface WorkTime {
    start: Date;
    end: Date;
    breaks: { start: Date, end: Date }[];
    content: string;
  }

  const workTimes: { [key: string]: WorkTime } = {};

  // 作業時間の集計
  inputData.forEach((row: any[]) => {
    const content = row[0] as string;
    const workDate = formatDate(row[2] as Date);
    const dateObj = row[2] as Date;
    const startTime = row[3] as Date;
    const endTime = row[5] as Date;

    const start = normalizeDateTime(dateObj, startTime);
    const end = normalizeDateTime(dateObj, endTime);

    if (!workTimes[workDate]) {
      workTimes[workDate] = { start, end, breaks: [], content };
    } else {
      workTimes[workDate].breaks.push({ start: workTimes[workDate].end, end: start });
      workTimes[workDate].end = end;
    }
  });

  const day1 = Object.keys(workTimes)[0];
  const thisYear = parseInt(day1.slice(0, 4), 10);
  const thisMonth = parseInt(day1.slice(5, 7), 10);

  const thisMonthDates = getDatesOfMonth(thisYear, thisMonth);

  // 休憩時間の計算と出力データの準備
  const outputData: (string | Date)[][] = thisMonthDates.map(date => {
    if (workTimes[date]) {
      const breakTimeTotal = workTimes[date].breaks.reduce((total, currentBreak) => {
        return total + (currentBreak.end.getTime() - currentBreak.start.getTime());
      }, 0);
      return [date, formatTime(workTimes[date].start), formatTime(workTimes[date].end), convertMillisecondsToHHMM(breakTimeTotal)];
    } else {
      return [date, '', '', ''];
    }
  });

  // 出力シートに書き込み
  if (outputData.length > 0) {
    outputSheet.getRange(2, 1, outputData.length, 4).setValues(outputData);
  }
}
