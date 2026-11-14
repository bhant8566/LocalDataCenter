'use strict';

// Taiwan DGPA office calendars, bundled for offline use. Verified 2026-11-14.
// Source: https://data.gov.tw/dataset/14718
// License: Government Data Open License, version 1 (https://data.gov.tw/license).
// 2026 CSV: https://www.dgpa.gov.tw/FileConversion?filename=dgpa%2Ffiles%2F202506%2Fa52331bd-a189-466b-b0f0-cae3062bbf74.csv&name=115.csv&nfix=
// 2027 CSV: https://www.dgpa.gov.tw/FileConversion?filename=dgpa%2Ffiles%2F202607%2Ff538b1ff-ba60-4c63-9477-10db8e6612d1.csv&name=116.csv&nfix=
// Only named days off are listed. Weekends are calculated independently.
const calendarHolidayYears = [2026, 2027];
const calendarHolidays = Object.freeze({
    '2026-01-01':'開國紀念日',
    '2026-02-15':'小年夜',
    '2026-02-16':'農曆除夕',
    '2026-02-17':'春節',
    '2026-02-18':'春節',
    '2026-02-19':'春節',
    '2026-02-20':'補假',
    '2026-02-27':'補假',
    '2026-02-28':'和平紀念日',
    '2026-04-03':'補假',
    '2026-04-04':'兒童節',
    '2026-04-05':'清明節',
    '2026-04-06':'補假',
    '2026-05-01':'勞動節',
    '2026-06-19':'端午節',
    '2026-09-25':'中秋節',
    '2026-09-28':'孔子誕辰紀念日/教師節',
    '2026-10-09':'補假',
    '2026-10-10':'國慶日',
    '2026-10-25':'臺灣光復暨金門古寧頭大捷紀念日',
    '2026-10-26':'補假',
    '2026-12-25':'行憲紀念日',
    '2027-01-01':'開國紀念日',
    '2027-02-04':'小年夜',
    '2027-02-05':'農曆除夕',
    '2027-02-06':'春節',
    '2027-02-07':'春節',
    '2027-02-08':'春節',
    '2027-02-09':'補假',
    '2027-02-10':'補假',
    '2027-02-28':'和平紀念日',
    '2027-03-01':'補假',
    '2027-04-04':'兒童節',
    '2027-04-05':'清明節',
    '2027-04-06':'補假',
    '2027-04-30':'補假',
    '2027-05-01':'勞動節',
    '2027-06-09':'端午節',
    '2027-09-15':'中秋節',
    '2027-09-28':'孔子誕辰紀念日/教師節',
    '2027-10-10':'國慶日',
    '2027-10-11':'補假',
    '2027-10-25':'臺灣光復暨金門古寧頭大捷紀念日',
    '2027-12-24':'補假',
    '2027-12-25':'行憲紀念日',
    '2027-12-31':'補假'
});

function calendarHolidayNote(year) {
    return calendarHolidayYears.includes(year) ? '' : '此年份尚未內建台灣假日資料，僅標示週末。';
}

function isCalendarWeekend(date) {
    const day = new Date(date + 'T00:00:00').getDay();
    return day === 0 || day === 6;
}
