// Multi-month Excel export extension kept outside app.js to minimize merge conflicts.
(function () {
  "use strict";

function heraExportGetMonthValueFromMeta(monthMeta) {
  return `${monthMeta.year}-${String(monthMeta.month).padStart(2, "0")}`;
}

function heraExportCompareMonthValues(a, b) {
  return String(a || "").localeCompare(String(b || ""));
}

function heraExportGetMonthRange(startValue, endValue) {
  const startMeta = getMonthMeta(startValue);
  const endMeta = getMonthMeta(endValue || startValue);
  if (!startMeta || !endMeta || heraExportCompareMonthValues(startValue, endValue || startValue) > 0) return null;
  const months = [];
  let year = startMeta.year;
  let month = startMeta.month;
  while (year < endMeta.year || (year === endMeta.year && month <= endMeta.month)) {
    const value = `${year}-${String(month).padStart(2, "0")}`;
    months.push({ value, ...getMonthMeta(value) });
    month += 1;
    if (month > 12) {
      month = 1;
      year += 1;
    }
  }
  return months;
}

function heraExportFormatMonthName(monthMeta) {
  return [
    "GENNAIO", "FEBBRAIO", "MARZO", "APRILE", "MAGGIO", "GIUGNO",
    "LUGLIO", "AGOSTO", "SETTEMBRE", "OTTOBRE", "NOVEMBRE", "DICEMBRE"
  ][Number(monthMeta?.month || 0) - 1] || heraExportGetMonthValueFromMeta(monthMeta);
}

function heraExportFormatMonthLabel(monthMeta) {
  return `${heraExportFormatMonthName(monthMeta)} ${monthMeta.year}`;
}

function heraExportFormatPeriodLabel(months) {
  if (!months.length) return "";
  const formatMonth = (meta) => `${String(meta.month).padStart(2, "0")}/${meta.year}`;
  return months.length === 1 ? formatMonth(months[0]) : `${formatMonth(months[0])} - ${formatMonth(months[months.length - 1])}`;
}

async function heraExportFetchReportsForPeriod(months, options = {}) {
  if (!Array.isArray(months) || !months.length) return [];
  const firstMonth = months[0];
  const lastMonth = months[months.length - 1];
  const fromDate = `${firstMonth.value}-01`;
  const toDate = `${lastMonth.value}-${String(lastMonth.daysInMonth).padStart(2, "0")}`;
  const includePendingApprovals = options?.includePendingApprovals === true;
  const reportsQuery = db.collection("oreReports")
    .where("date", ">=", fromDate)
    .where("date", "<=", toDate)
    .orderBy("date", "asc")
    .get();
  const approvalsQuery = includePendingApprovals
    ? db.collection("oreApprovalRequests")
      .where("date", ">=", fromDate)
      .where("date", "<=", toDate)
      .orderBy("date", "asc")
      .get()
    : Promise.resolve(null);
  const [reportsSnapshot, approvalsSnapshot] = await Promise.all([reportsQuery, approvalsQuery]);
  const reports = reportsSnapshot.docs.map((doc) => ({
    id: doc.id,
    sourceCollection: "oreReports",
    approvalStatus: "approved",
    ...doc.data()
  }));
  const pendingApprovals = approvalsSnapshot
    ? approvalsSnapshot.docs
      .map((doc) => ({
        id: doc.id,
        sourceCollection: "oreApprovalRequests",
        ...doc.data()
      }))
      .filter((request) => !["approved", "rejected"].includes(String(request.status || "").trim()))
    : [];
  return deduplicateHoursRecordsForDisplay([...reports, ...pendingApprovals])
    .sort((a, b) => String(a.date || "").localeCompare(String(b.date || "")));
}

async function heraExportHoursRangeWorkbook(options = {}) {
  const onlyCommessaId = String(options.onlyCommessaId || "").trim();
  const emptyMessage = String(options.emptyMessage || "Nessuna ora registrata nel periodo selezionato per l'export globale.");
  const monthValue = String(options.monthStartValue || ui.hoursTableMonth?.value || ui.hoursStatsMonth?.value || "").trim();
  const monthEndValue = String(options.monthEndValue || ui.hoursTableMonthEnd?.value || monthValue).trim();
  const monthRange = heraExportGetMonthRange(monthValue, monthEndValue);
  if (!monthRange) {
    alert("Seleziona un periodo mesi valido prima di esportare il file globale.");
    return;
  }
  const monthMeta = monthRange[0];
  const periodDaysInMonth = monthRange.reduce((max, item) => Math.max(max, item.daysInMonth), 0);
  if (!window.ExcelJS?.Workbook && window.HeraHeavyLibs?.ensure) {
    try { await window.HeraHeavyLibs.ensure("exceljs"); } catch (_) {}
  }
  if (!window.ExcelJS?.Workbook) {
    alert("Libreria Excel non disponibile. Controlla la connessione e riprova.");
    return;
  }

  try {
    logHoursDebug("mese selezionato", monthValue);
    logHoursDebug("periodo selezionato", { da: monthValue, a: monthEndValue, mesi: monthRange.map((item) => item.value) });
    const reports = await heraExportFetchReportsForPeriod(monthRange, { includePendingApprovals: true });
    logHoursDebug("record trovati", Array.isArray(reports) ? reports.length : 0);
    const monthMetaByValue = new Map(monthRange.map((item) => [item.value, item]));
    const commessaMap = new Map();
    const globalOperatorMonthMap = new Map();
    let totalValidGlobalRows = 0;
    reports.forEach((report) => {
      const reportDate = String(report.date || "");
      const reportMonthValue = reportDate.slice(0, 7);
      const reportMonthMeta = monthMetaByValue.get(reportMonthValue);
      const day = Number(reportDate.split("-")[2] || 0);
      if (!reportMonthMeta || !day || day < 1 || day > reportMonthMeta.daysInMonth) return;
      const entries = Array.isArray(report.entries) ? report.entries : [];
      entries.forEach((entry) => {
        const entryCommessaInfo = resolveHoursEntryCommessa(entry);
        const commessaId = String(entryCommessaInfo.id || entryCommessaInfo.key || "").trim();
        if (!commessaId) return;
        if (onlyCommessaId && commessaId !== onlyCommessaId) return;
        const commessaName = String(entryCommessaInfo.nome || entry.commessaName || commesseById.get(entryCommessaInfo.id)?.nome || "Commessa").trim() || "Commessa";
        const commessaCode = String(entryCommessaInfo.codice || commesseById.get(entryCommessaInfo.id)?.codice || "").trim();
        const commessaMonthKey = `${commessaId}::${reportMonthValue}`;
        if (!commessaMap.has(commessaMonthKey)) {
          commessaMap.set(commessaMonthKey, {
            commessaId,
            commessaName,
            commessaCode,
            monthValue: reportMonthValue,
            monthMeta: reportMonthMeta,
            monthLabelIt: heraExportFormatMonthLabel(reportMonthMeta),
            operatorsMap: new Map()
          });
        }
        const commessaBucket = commessaMap.get(commessaMonthKey);
        (Array.isArray(entry.rows) ? entry.rows : []).forEach((row) => {
          const operatore = String(row.operatore || "").trim();
          const ore = Number(row.ore || 0);
          if (!operatore || ore <= 0) return;
          totalValidGlobalRows += 1;
          const operatorNorm = operatore.toLocaleLowerCase("it-IT").replace(/\s+/g, " ").trim();
          if (!commessaBucket.operatorsMap.has(operatorNorm)) {
            commessaBucket.operatorsMap.set(operatorNorm, {
              displayName: operatore,
              days: Array.from({ length: periodDaysInMonth }, () => 0)
            });
          }
          const operatorMonthKey = `${operatorNorm}::${reportMonthValue}`;
          if (!globalOperatorMonthMap.has(operatorMonthKey)) {
            globalOperatorMonthMap.set(operatorMonthKey, {
              monthMeta: reportMonthMeta,
              days: Array.from({ length: periodDaysInMonth }, () => 0)
            });
          }
          commessaBucket.operatorsMap.get(operatorNorm).days[day - 1] += ore;
          globalOperatorMonthMap.get(operatorMonthKey).days[day - 1] += ore;
        });
      });
    });

  logHoursDebug("dati usati per export", { mode: "global", periodo: { da: monthValue, a: monthEndValue }, commesse: Array.from(commessaMap.values()).map((item) => ({
    commessaName: item.commessaName,
    commessaCode: item.commessaCode,
    monthValue: item.monthValue,
    operatori: Array.from(item.operatorsMap.values())
  })) });
  if (!commessaMap.size || totalValidGlobalRows <= 0) {
    alert(emptyMessage);
    if (ui.hoursTableFeedback) ui.hoursTableFeedback.textContent = emptyMessage;
    return;
  }

  const monthNameIt = monthRange.length === 1 ? heraExportFormatMonthName(monthRange[0]) : heraExportFormatPeriodLabel(monthRange);
  const monthLabelIt = monthNameIt;

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Hera App";
  workbook.created = new Date();
  const worksheet = workbook.addWorksheet("Export globale", {
    views: [{ state: "frozen", xSplit: 1, ySplit: 8 }]
  });

  const dayStartColumn = 2;
  const totalColumn = periodDaysInMonth + 2;
  const ordinaryHoursColumn = periodDaysInMonth + 3;
  const overtimeHoursColumn = periodDaysInMonth + 4;
  const lastColumn = overtimeHoursColumn;
  const dayHeaderFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
  const totalColumnFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9EAF7" } };
  const hoursFilledCell = { type: "pattern", pattern: "solid", fgColor: { argb: "FFEDE9DD" } };
  const weekendFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE5E5E5" } };
  const errorFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFC7CE" } };
  const whiteFill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
  const thinBorder = {
    top: { style: "thin", color: { argb: "FF000000" } },
    left: { style: "thin", color: { argb: "FF000000" } },
    bottom: { style: "thin", color: { argb: "FF000000" } },
    right: { style: "thin", color: { argb: "FF000000" } }
  };
  const mediumSide = { style: "medium", color: { argb: "FF000000" } };
  const thickSide = { style: "thick", color: { argb: "FF000000" } };
  const getEasterSunday = (year) => {
    const a = year % 19;
    const b = Math.floor(year / 100);
    const c = year % 100;
    const d = Math.floor(b / 4);
    const e = b % 4;
    const f = Math.floor((b + 8) / 25);
    const g = Math.floor((b - f + 1) / 3);
    const h = (19 * a + b - d - g + 15) % 30;
    const i = Math.floor(c / 4);
    const k = c % 4;
    const l = (32 + 2 * e + 2 * i - h - k) % 7;
    const m = Math.floor((a + 11 * h + 22 * l) / 451);
    const month = Math.floor((h + l - 7 * m + 114) / 31);
    const day = ((h + l - 7 * m + 114) % 31) + 1;
    return new Date(year, month - 1, day);
  };
  const getItalianHolidayKeys = (year) => {
    const easterMonday = getEasterSunday(year);
    easterMonday.setDate(easterMonday.getDate() + 1);
    const easterMondayKey = `${String(easterMonday.getMonth() + 1).padStart(2, "0")}-${String(easterMonday.getDate()).padStart(2, "0")}`;
    return new Set([
      "01-01",
      "01-06",
      easterMondayKey,
      "04-25",
      "05-01",
      "06-02",
      "08-15",
      "11-01",
      "12-08",
      "12-25",
      "12-26"
    ]);
  };
  const holidayKeysByYear = new Map();
  const getHolidayKeysForYear = (year) => {
    if (!holidayKeysByYear.has(year)) holidayKeysByYear.set(year, getItalianHolidayKeys(year));
    return holidayKeysByYear.get(year);
  };
  const isHolidayDay = (dayNumber, targetMonthMeta = monthMeta) => {
    const holidayKey = `${String(targetMonthMeta.month).padStart(2, "0")}-${String(dayNumber).padStart(2, "0")}`;
    return getHolidayKeysForYear(targetMonthMeta.year).has(holidayKey);
  };
  const isWeekendDay = (dayNumber, targetMonthMeta = monthMeta) => {
    const dayDate = new Date(targetMonthMeta.year, targetMonthMeta.month - 1, dayNumber);
    const weekday = dayDate.getDay();
    return weekday === 0 || weekday === 6;
  };
  const getOrdinaryHoursLimit = (dayNumber, targetMonthMeta = monthMeta) => {
    if (isWeekendDay(dayNumber, targetMonthMeta) || isHolidayDay(dayNumber, targetMonthMeta)) return 0;
    const weekday = new Date(targetMonthMeta.year, targetMonthMeta.month - 1, dayNumber).getDay();
    if (weekday >= 1 && weekday <= 4) return 8;
    if (weekday === 5) return 7;
    return 0;
  };
  const splitOrdinaryAndOvertimeHours = (hours, dayNumber, targetMonthMeta = monthMeta) => {
    const dailyHours = Number(hours || 0);
    if (!Number.isFinite(dailyHours) || dailyHours <= 0) return { ordinary: 0, overtime: 0 };
    const ordinaryLimit = getOrdinaryHoursLimit(dayNumber, targetMonthMeta);
    const ordinary = Math.min(dailyHours, ordinaryLimit);
    return {
      ordinary,
      overtime: Math.max(dailyHours - ordinaryLimit, 0)
    };
  };
  const getExcelNumberFormat = (value) => {
    const num = Number(value || 0);
    if (!Number.isFinite(num) || num <= 0) return null;
    return Number.isInteger(num) ? "0" : "0.##";
  };
  const setThinBorder = (cell) => {
    cell.border = thinBorder;
  };
  const setOuterBlockBorder = (startRow, endRow) => {
    for (let row = startRow; row <= endRow; row += 1) {
      for (let col = 1; col <= lastColumn; col += 1) {
        const cell = worksheet.getCell(row, col);
        const border = { ...(cell.border || {}) };
        if (row === startRow) border.top = mediumSide;
        if (row === endRow) border.bottom = mediumSide;
        if (col === 1) border.left = mediumSide;
        if (col === lastColumn) border.right = mediumSide;
        cell.border = border;
      }
    }
  };
  const addWeekSeparatorBorders = (rowIndex, targetMonthMeta = monthMeta) => {
    for (let day = 1; day <= targetMonthMeta.daysInMonth; day += 1) {
      const date = new Date(targetMonthMeta.year, targetMonthMeta.month - 1, day);
      const isSunday = date.getDay() === 0;
      if (!isSunday || day === targetMonthMeta.daysInMonth) continue;
      const dayCol = dayStartColumn + day - 1;
      const cell = worksheet.getCell(rowIndex, dayCol);
      const border = { ...(cell.border || {}) };
      border.right = thickSide;
      cell.border = border;
      const nextCol = dayCol + 1;
      if (nextCol <= dayStartColumn + targetMonthMeta.daysInMonth - 1) {
        const nextCell = worksheet.getCell(rowIndex, nextCol);
        const nextBorder = { ...(nextCell.border || {}) };
        nextBorder.left = thickSide;
        nextCell.border = nextBorder;
      }
    }
  };

  let rowPointer = 1;
  const commesseSorted = Array.from(commessaMap.values())
    .sort((a, b) => a.commessaName.localeCompare(b.commessaName, "it") || a.monthValue.localeCompare(b.monthValue));

  const totalCommesse = new Set(commesseSorted.map((item) => item.commessaId)).size;
  const totalOperatorsUnique = new Set(Array.from(globalOperatorMonthMap.keys()).map((key) => key.split("::")[0])).size;
  const totalOperatorsActive = commesseSorted.reduce((acc, commessaBlock) => (
    acc + Array.from(commessaBlock.operatorsMap.values()).filter((operator) => (
      operator.days.some((value) => Number(value || 0) > 0)
    )).length
  ), 0);
  const monthlyHourTotals = Array.from(globalOperatorMonthMap.values()).reduce((acc, operatorMonth) => {
    operatorMonth.days.forEach((value, dayIndex) => {
      const dailyHours = Number(value || 0);
      if (dailyHours <= 0 || dayIndex >= operatorMonth.monthMeta.daysInMonth) return;
      const dailyBreakdown = splitOrdinaryAndOvertimeHours(dailyHours, dayIndex + 1, operatorMonth.monthMeta);
      acc.ordinary += dailyBreakdown.ordinary;
      acc.overtime += dailyBreakdown.overtime;
      acc.total += dailyHours;
    });
    return acc;
  }, { ordinary: 0, overtime: 0, total: 0 });

  const summaryStartRow = rowPointer;
  worksheet.mergeCells(summaryStartRow, 1, summaryStartRow, lastColumn);
  const summaryTitleCell = worksheet.getCell(summaryStartRow, 1);
  summaryTitleCell.value = " VARGA CANTIERI   RIEPILOGO GESTIONE ORE GLOBAL";
  summaryTitleCell.font = { bold: true, size: 14, color: { argb: "FF000000" } };
  summaryTitleCell.alignment = { horizontal: "center", vertical: "middle" };
  summaryTitleCell.fill = whiteFill;
  rowPointer += 1;

  const formatSummaryValue = (value) => {
    if (typeof value !== "number") return value;
    return Number.isInteger(value) ? value : Number(value.toFixed(2));
  };
  const summaryCardRows = [
    [
      ["MESE DI RIFERIMENTO", monthNameIt],
      ["ANNO", monthRange.length === 1 ? String(monthMeta.year) : `${monthRange[0].year}-${monthRange[monthRange.length - 1].year}`],
      ["DATA ESPORTAZIONE", new Date().toLocaleDateString("it-IT")]
    ],
    [
      ["TOTALE ORE", formatSummaryValue(monthlyHourTotals.total)],
      ["ORE ORDINARIE", formatSummaryValue(monthlyHourTotals.ordinary)],
      ["ORE STRAORDINARIE", formatSummaryValue(monthlyHourTotals.overtime)]
    ],
    [
      ["TOTALE OPERATORI", totalOperatorsActive],
      ["TOTALE OPERATORI UNICI", totalOperatorsUnique],
      ["TOTALE COMMESSE", totalCommesse]
    ]
  ];
  const summaryColumnGroups = [
    [1, Math.floor(lastColumn / 3)],
    [Math.floor(lastColumn / 3) + 1, Math.floor((lastColumn * 2) / 3)],
    [Math.floor((lastColumn * 2) / 3) + 1, lastColumn]
  ];
  summaryCardRows.forEach((cards) => {
    const labelRowIndex = rowPointer;
    const valueRowIndex = rowPointer + 1;
    cards.forEach(([label, value], cardIndex) => {
      const [startCol, endCol] = summaryColumnGroups[cardIndex];
      worksheet.mergeCells(labelRowIndex, startCol, labelRowIndex, endCol);
      worksheet.mergeCells(valueRowIndex, startCol, valueRowIndex, endCol);
      const labelCell = worksheet.getCell(labelRowIndex, startCol);
      const valueCell = worksheet.getCell(valueRowIndex, startCol);
      labelCell.value = label;
      valueCell.value = value;
      labelCell.font = { bold: true, size: 8, color: { argb: "FF4B5563" } };
      valueCell.font = { bold: true, size: 13, color: { argb: "FF000000" } };
      labelCell.alignment = { horizontal: "center", vertical: "bottom" };
      valueCell.alignment = { horizontal: "center", vertical: "top" };
      if (typeof value === "number") valueCell.numFmt = getExcelNumberFormat(value) || "0";
      for (let row = labelRowIndex; row <= valueRowIndex; row += 1) {
        for (let col = startCol; col <= endCol; col += 1) {
          const cell = worksheet.getCell(row, col);
          cell.fill = whiteFill;
          setThinBorder(cell);
        }
      }
    });
    worksheet.getRow(labelRowIndex).height = 13;
    worksheet.getRow(valueRowIndex).height = 19;
    rowPointer += 2;
  });
  const summaryEndRow = rowPointer - 1;
  rowPointer += 1;
  const firstCommessaStartRow = rowPointer;

  commesseSorted.forEach((commessaBlock, idx) => {
    const blockMonthMeta = commessaBlock.monthMeta || monthMeta;
    const operatorRows = Array.from(commessaBlock.operatorsMap.values())
      .sort((a, b) => a.displayName.localeCompare(b.displayName, "it"));
    const operators = operatorRows.length ? operatorRows : [];

    const startRow = rowPointer;
    const commessaRow = worksheet.getRow(rowPointer);
    commessaRow.getCell(1).value = "COMMESSA";
    commessaRow.getCell(2).value = commessaBlock.commessaName;
    worksheet.mergeCells(rowPointer, 2, rowPointer, lastColumn);
    commessaRow.font = { bold: true, size: 14, color: { argb: "FF0B1F44" } };

    rowPointer += 1;
    const codeRow = worksheet.getRow(rowPointer);
    codeRow.getCell(1).value = "CODICE COMMESSA";
    codeRow.getCell(2).value = commessaBlock.commessaCode || "N/D";
    worksheet.mergeCells(rowPointer, 2, rowPointer, lastColumn);
    codeRow.font = { bold: true, size: 12, color: { argb: "FF0B1F44" } };

    rowPointer += 1;
    const meseRow = worksheet.getRow(rowPointer);
    meseRow.getCell(1).value = "MESE RIF.";
    meseRow.getCell(2).value = commessaBlock.monthLabelIt || monthLabelIt;
    worksheet.mergeCells(rowPointer, 2, rowPointer, lastColumn);

    rowPointer += 1;
    const headerRow = worksheet.getRow(rowPointer);
    headerRow.getCell(1).value = "OPERATORE";
    headerRow.getCell(1).fill = dayHeaderFill;
    headerRow.getCell(1).font = { bold: true, color: { argb: "FF000000" } };
    for (let day = 1; day <= periodDaysInMonth; day += 1) {
      headerRow.getCell(day + 1).value = day <= blockMonthMeta.daysInMonth ? day : "";
      headerRow.getCell(day + 1).fill = dayHeaderFill;
      headerRow.getCell(day + 1).font = { bold: true, color: { argb: "FF000000" } };
    }
    headerRow.getCell(totalColumn).value = "TOTALE";
    headerRow.getCell(totalColumn).fill = totalColumnFill;
    headerRow.getCell(totalColumn).font = { bold: true, color: { argb: "FF000000" } };
    headerRow.getCell(ordinaryHoursColumn).value = "ORE ORDINARIE";
    headerRow.getCell(ordinaryHoursColumn).fill = dayHeaderFill;
    headerRow.getCell(ordinaryHoursColumn).font = { bold: true, color: { argb: "FF000000" } };
    headerRow.getCell(overtimeHoursColumn).value = "ORE STRAORDINARIE";
    headerRow.getCell(overtimeHoursColumn).fill = dayHeaderFill;
    headerRow.getCell(overtimeHoursColumn).font = { bold: true, color: { argb: "FF000000" } };
    headerRow.height = 24;

    rowPointer += 1;
    let commessaTotal = 0;

    operators.forEach((operatorData, operatorIdx) => {
      const row = worksheet.getRow(rowPointer + operatorIdx);
      row.getCell(1).value = operatorData.displayName || "";
      let total = 0;
      let ordinaryHours = 0;
      let overtimeHours = 0;
      for (let dayIdx = 0; dayIdx < periodDaysInMonth; dayIdx += 1) {
        const value = Number(operatorData.days[dayIdx] || 0);
        const cell = row.getCell(dayIdx + 2);
        const isValidDay = dayIdx < blockMonthMeta.daysInMonth;
        if (!isValidDay) {
          cell.fill = whiteFill;
        } else if (isWeekendDay(dayIdx + 1, blockMonthMeta) || isHolidayDay(dayIdx + 1, blockMonthMeta)) {
          cell.fill = weekendFill;
        } else {
          cell.fill = whiteFill;
        }
        if (isValidDay && value > 0) {
          cell.value = value;
          cell.fill = value > 12 ? errorFill : hoursFilledCell;
          const numFmt = getExcelNumberFormat(value);
          if (numFmt) cell.numFmt = numFmt;
          const dailyBreakdown = splitOrdinaryAndOvertimeHours(value, dayIdx + 1, blockMonthMeta);
          ordinaryHours += dailyBreakdown.ordinary;
          overtimeHours += dailyBreakdown.overtime;
          total += value;
        } else {
          cell.value = null;
        }
      }
      row.getCell(totalColumn).value = total > 0 ? total : null;
      row.getCell(totalColumn).fill = totalColumnFill;
      if (total > 0) row.getCell(totalColumn).numFmt = getExcelNumberFormat(total);
      row.getCell(ordinaryHoursColumn).value = ordinaryHours > 0 ? ordinaryHours : null;
      if (ordinaryHours > 0) row.getCell(ordinaryHoursColumn).numFmt = getExcelNumberFormat(ordinaryHours);
      row.getCell(overtimeHoursColumn).value = overtimeHours > 0 ? overtimeHours : null;
      if (overtimeHours > 0) row.getCell(overtimeHoursColumn).numFmt = getExcelNumberFormat(overtimeHours);
      commessaTotal += total;
      row.height = 21;
    });

    const totalCommessaRowIndex = rowPointer + operators.length;
    const totalCommessaRow = worksheet.getRow(totalCommessaRowIndex);
    totalCommessaRow.getCell(1).value = "TOTALE COMMESSA";
    worksheet.mergeCells(totalCommessaRowIndex, 1, totalCommessaRowIndex, totalColumn - 1);
    totalCommessaRow.getCell(totalColumn).value = commessaTotal > 0 ? commessaTotal : null;
    totalCommessaRow.getCell(totalColumn).fill = totalColumnFill;
    if (commessaTotal > 0) totalCommessaRow.getCell(totalColumn).numFmt = getExcelNumberFormat(commessaTotal);
    totalCommessaRow.getCell(1).font = { bold: true, color: { argb: "FF000000" } };
    totalCommessaRow.getCell(totalColumn).font = { bold: true, color: { argb: "FF000000" } };
    totalCommessaRow.height = 22;

    const endRow = totalCommessaRowIndex;
    for (let row = startRow; row <= endRow; row += 1) {
      for (let col = 1; col <= lastColumn; col += 1) {
        const cell = worksheet.getCell(row, col);
        if (!cell.fill) cell.fill = whiteFill;
        setThinBorder(cell);
        if (row >= startRow + 2 && row < endRow) {
          const isOperatorName = col === 1;
          cell.alignment = {
            vertical: "middle",
            horizontal: isOperatorName ? "left" : "center"
          };
        } else if (row === endRow) {
          cell.alignment = {
            vertical: "middle",
            horizontal: col === 1 ? "left" : "center"
          };
        } else {
          cell.alignment = {
            vertical: "middle",
            horizontal: col === 1 ? "left" : "center"
          };
        }
        const isTitleLabel = col === 1 && row <= startRow + 2;
        const isHeaderRow = row === startRow + 2;
        if (isTitleLabel || isHeaderRow) {
          cell.font = { ...(cell.font || {}), bold: true, color: { argb: "FF000000" } };
        }
      }
      if (row <= startRow + 1) worksheet.getRow(row).height = 22;
      addWeekSeparatorBorders(row, blockMonthMeta);
    }

    setOuterBlockBorder(startRow, endRow);

    rowPointer = endRow + 2;
    if (idx < commesseSorted.length - 1) {
      worksheet.getRow(rowPointer - 1).height = 10;
    }
  });

  for (let col = 1; col <= lastColumn; col += 1) {
    const cell = worksheet.getCell(summaryStartRow, col);
    cell.fill = whiteFill;
    setThinBorder(cell);
  }
  worksheet.getRow(summaryStartRow).height = 24;
  setOuterBlockBorder(summaryStartRow, summaryEndRow);

  for (let col = 1; col <= lastColumn; col += 1) {
    if (col === 1) {
      worksheet.getColumn(col).width = 28;
      continue;
    }
    if (col >= dayStartColumn && col <= totalColumn - 1) {
      worksheet.getColumn(col).width = 4.2;
      continue;
    }
    if (col === totalColumn) {
      worksheet.getColumn(col).width = 11;
      continue;
    }
    if (col === ordinaryHoursColumn) {
      worksheet.getColumn(col).width = 16;
      continue;
    }
    worksheet.getColumn(col).width = 18;
  }

  worksheet.autoFilter = {
    from: { row: firstCommessaStartRow + 3, column: 1 },
    to: { row: firstCommessaStartRow + 3, column: 1 }
  };

  for (let row = 1; row <= worksheet.rowCount; row += 1) {
    const currentRow = worksheet.getRow(row);
    if (!currentRow.height) currentRow.height = 21;
  }

  const safeMonth = monthRange.length === 1 ? monthValue.replace("/", "-") : `${monthValue}_${monthEndValue}`;
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const safePrefix = String(options.fileNamePrefix || "ore_global").replace(/[\\/:*?"<>|]/g, "_").replace(/\s+/g, "_");
  const fileName = `${safePrefix}_${safeMonth}.xlsx`;
  if (window.navigator?.msSaveOrOpenBlob) {
    window.navigator.msSaveOrOpenBlob(blob, fileName);
    return;
  }
  const url = window.URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  window.URL.revokeObjectURL(url);
  } catch (error) {
    console.error("Errore export Excel Global ore:", error);
    if (ui.hoursTableFeedback) ui.hoursTableFeedback.textContent = "Errore export Excel Global. Controlla i dati o riprova.";
    alert("Errore export Excel Global. Controlla i dati o riprova.");
  }
}


  function getHoursExportStartMonth() {
    return String(document.getElementById("hours-table-month")?.value || ui?.hoursStatsMonth?.value || "").trim();
  }

  function getHoursExportEndMonth() {
    const start = getHoursExportStartMonth();
    return String(document.getElementById("hours-table-month-end")?.value || start).trim();
  }

  function syncHoursExportEndMonth() {
    const startInput = document.getElementById("hours-table-month");
    const endInput = document.getElementById("hours-table-month-end");
    if (!startInput || !endInput) return;
    if (!endInput.value) endInput.value = startInput.value || "";
    if (endInput.value && startInput.value && endInput.value < startInput.value) startInput.value = endInput.value;
  }

  async function exportSelectedCommessaRange(event) {
    event.preventDefault();
    event.stopImmediatePropagation();
    try {
      const commessaId = String(document.getElementById("hours-table-commessa-select")?.value || "").trim();
      const start = getHoursExportStartMonth();
      const end = getHoursExportEndMonth();
      if (!commessaId || !heraExportGetMonthRange(start, end)) {
        alert("Seleziona mese di inizio, mese di fine e commessa prima di esportare Excel.");
        return;
      }
      const commessaName = commesseById.get(commessaId)?.nome || "Commessa";
      await heraExportHoursRangeWorkbook({
        onlyCommessaId: commessaId,
        monthStartValue: start,
        monthEndValue: end,
        emptyMessage: "Nessuna ora registrata per questa commessa nel periodo selezionato.",
        fileNamePrefix: `ore_${String(commessaName || "commessa").replace(/[\\/:*?"<>|]/g, "_").replace(/\s+/g, "_")}`
      });
    } catch (error) {
      console.error("Errore export Excel ore periodo:", error);
      if (ui?.hoursTableFeedback) ui.hoursTableFeedback.textContent = "Errore export Excel ore. Controlla i dati o riprova.";
      alert("Errore export Excel ore. Controlla i dati o riprova.");
    }
  }

  async function exportGlobalRange(event) {
    event.preventDefault();
    event.stopImmediatePropagation();
    await heraExportHoursRangeWorkbook({
      monthStartValue: getHoursExportStartMonth(),
      monthEndValue: getHoursExportEndMonth()
    });
  }

  function initHoursRangeExportExtension() {
    const startInput = document.getElementById("hours-table-month");
    const endInput = document.getElementById("hours-table-month-end");
    if (startInput && endInput) {
      if (!endInput.value) endInput.value = startInput.value || new Date().toISOString().slice(0, 7);
      startInput.addEventListener("change", syncHoursExportEndMonth);
      endInput.addEventListener("change", syncHoursExportEndMonth);
    }
    ui?.hoursStatsMonth?.addEventListener("change", () => {
      const startInput = document.getElementById("hours-table-month");
      const endInput = document.getElementById("hours-table-month-end");
      if (startInput && ui.hoursStatsMonth?.value) startInput.value = ui.hoursStatsMonth.value;
      if (endInput && ui.hoursStatsMonth?.value) endInput.value = ui.hoursStatsMonth.value;
    });
    ui?.viewHoursBtn?.addEventListener("click", () => setTimeout(syncHoursExportEndMonth, 0), true);
    document.getElementById("hours-table-export-btn")?.addEventListener("click", exportSelectedCommessaRange, true);
    document.getElementById("hours-table-export-global-btn")?.addEventListener("click", exportGlobalRange, true);
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", initHoursRangeExportExtension);
  else initHoursRangeExportExtension();
})();

// Carica separatamente il controllo aggiornamenti per evitare modifiche ai file
// principali dell'app e ridurre i conflitti durante gli aggiornamenti del branch.
(function loadAppUpdateFeature() {
  if (document.querySelector('script[data-hera-app-update="true"]')) return;
  const script = document.createElement("script");
  script.src = "update-app-feature.js?v=20260824-deadlock1";
  script.dataset.heraAppUpdate = "true";
  document.head.appendChild(script);
})();
