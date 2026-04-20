function exportBPV_10AK() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfg = getConfig();
  const utils = createUtils();

  let ui = null;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {}

  try {
    const sheets = getRequiredSheets(ss, cfg);
    const sourceMeta = getSourceMeta(sheets.source, cfg);

    if (sourceMeta.numRows <= 0) {
      if (ui) ui.alert("Немає даних для експорту");
      return;
    }

    const idx = getSourceIndexes(utils.colNum);
    const data = readSourceData(sheets.source, cfg, sourceMeta.lastCol, sourceMeta.numRows);
    const commonValue = sheets.general.getRange("B2").getValue();

    writeTargetSheet(sheets.target, data, cfg, idx, utils);

    const rows = buildGroupedRows(data, idx, utils);
    const yearTotals = buildYearTotals(rows, utils);
    const visibleRows = rows.filter((row) => row.isNotWrittenOff);

    const summaryData = buildSummaryOutput(visibleRows, yearTotals, commonValue, cfg, utils);
    writeSummarySheet(sheets.summary, summaryData, cfg, utils);

    if (ui) ui.alert("✅ Дані оновлено");
  } catch (e) {
    if (ui) {
      ui.alert("❌ Помилка: " + e.message);
    } else {
      Logger.log("Помилка: " + e.message);
    }
    throw e;
  }
}

function getConfig() {
  return {
    source: "Списання_розширене_1_2_3_4_5",
    target: "БВП 10 АК",
    summary: "БПВ що вівторка",
    general: "Загальні дані",
    sourceStartRow: 2,
    targetStartRow: 10,
    summaryStartRow: 10
  };
}

function createUtils() {
  const colNum = (letters) => {
    let n = 0;
    for (const ch of String(letters).toUpperCase().trim()) {
      n = n * 26 + ch.charCodeAt(0) - 64;
    }
    return n;
  };

  const isBlank = (v) => v === null || v === "" || String(v).trim() === "";
  const trim = (v) => (v === null || v === undefined ? "" : String(v).trim());
  const num = (v) => {
    if (typeof v === "number") return v;
    if (isBlank(v)) return 0;
    const x = Number(String(v).replace(/\s/g, "").replace(",", "."));
    return isNaN(x) ? 0 : x;
  };

  return { colNum, isBlank, trim, num };
}

function getRequiredSheets(ss, cfg) {
  const source = ss.getSheetByName(cfg.source);
  const target = ss.getSheetByName(cfg.target);
  const summary = ss.getSheetByName(cfg.summary);
  const general = ss.getSheetByName(cfg.general);

  if (!source) throw new Error(`Не знайдено лист: ${cfg.source}`);
  if (!target) throw new Error(`Не знайдено лист: ${cfg.target}`);
  if (!summary) throw new Error(`Не знайдено лист: ${cfg.summary}`);
  if (!general) throw new Error(`Не знайдено лист: ${cfg.general}`);

  return { source, target, summary, general };
}

function getSourceMeta(sourceSheet, cfg) {
  const lastRow = sourceSheet.getLastRow();
  const lastCol = sourceSheet.getLastColumn();
  const numRows = lastRow - cfg.sourceStartRow + 1;
  return { lastRow, lastCol, numRows };
}

function getSourceIndexes(colNum) {
  return {
    B: colNum("B") - 1,
    C: colNum("C") - 1,
    E: colNum("E") - 1,
    F: colNum("F") - 1,
    G: colNum("G") - 1,
    H: colNum("H") - 1,
    I: colNum("I") - 1,
    M: colNum("M") - 1,
    P: colNum("P") - 1,
    Q: colNum("Q") - 1,
    T: colNum("T") - 1,
    Y: colNum("Y") - 1,
    Z: colNum("Z") - 1,
    AK: colNum("AK") - 1,
    AL: colNum("AL") - 1
  };
}

function readSourceData(sourceSheet, cfg, lastCol, numRows) {
  return sourceSheet.getRange(cfg.sourceStartRow, 1, numRows, lastCol).getValues();
}

function writeTargetSheet(targetSheet, data, cfg, idx, utils) {
  const targetMaxRows = targetSheet.getMaxRows();
  const targetColsCount = utils.colNum("Z");

  targetSheet
    .getRange(cfg.targetStartRow, 1, Math.max(1, targetMaxRows - cfg.targetStartRow + 1), targetColsCount)
    .clearContent();

  const targetOutput = buildTargetOutput(data, idx, targetColsCount, utils);

  targetSheet
    .getRange(cfg.targetStartRow, 1, targetOutput.length, targetColsCount)
    .setValues(targetOutput);
}

function buildTargetOutput(data, idx, targetColsCount, utils) {
  const targetCols = {
    A: utils.colNum("A") - 1,
    B: utils.colNum("B") - 1,
    C: utils.colNum("C") - 1,
    D: utils.colNum("D") - 1,
    E: utils.colNum("E") - 1,
    G: utils.colNum("G") - 1,
    I: utils.colNum("I") - 1,
    J: utils.colNum("J") - 1,
    K: utils.colNum("K") - 1,
    L: utils.colNum("L") - 1,
    M: utils.colNum("M") - 1,
    N: utils.colNum("N") - 1,
    O: utils.colNum("O") - 1,
    P: utils.colNum("P") - 1,
    Q: utils.colNum("Q") - 1,
    Z: utils.colNum("Z") - 1
  };

  const targetOutput = Array.from({ length: data.length }, () => new Array(targetColsCount).fill(""));

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const zVal = utils.trim(row[idx.AK]);

    targetOutput[i][targetCols.A] = row[idx.E];
    targetOutput[i][targetCols.B] = row[idx.F];
    targetOutput[i][targetCols.C] = row[idx.H];
    targetOutput[i][targetCols.D] = row[idx.B];
    targetOutput[i][targetCols.E] = row[idx.T];
    targetOutput[i][targetCols.G] = row[idx.G];
    targetOutput[i][targetCols.I] = row[idx.I];
    targetOutput[i][targetCols.J] = row[idx.M];
    targetOutput[i][targetCols.K] = row[idx.P];
    targetOutput[i][targetCols.P] = row[idx.Y];
    targetOutput[i][targetCols.Q] = row[idx.Z];
    targetOutput[i][targetCols.Z] = row[idx.AK];

    if (zVal !== "") {
      targetOutput[i][targetCols.L] = row[idx.I];
      targetOutput[i][targetCols.M] = row[idx.P];
    } else {
      targetOutput[i][targetCols.N] = row[idx.I];
      targetOutput[i][targetCols.O] = row[idx.P];
    }
  }

  return targetOutput;
}

function buildGroupedRows(data, idx, utils) {
  const grouped = new Map();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const order = utils.trim(row[idx.Y]);
    if (!order) continue;

    const year = row[idx.T];
    const key = `${order}||${utils.trim(year)}`;

    if (!grouped.has(key)) {
      grouped.set(key, {
        order,
        year: utils.isBlank(year) ? "" : year,
        qValue: utils.isBlank(row[idx.Q]) ? "" : row[idx.Q],
        fValue: utils.isBlank(row[idx.C]) ? "" : row[idx.C],
        sum: 0,
        unresolved: false,
        hasAK: false,
        hasQuestion: false
      });
    }

    const item = grouped.get(key);
    item.sum += utils.num(row[idx.P]);

    if (utils.isBlank(item.qValue) && !utils.isBlank(row[idx.Q])) item.qValue = row[idx.Q];
    if (utils.isBlank(item.fValue) && !utils.isBlank(row[idx.C])) item.fValue = row[idx.C];
    if (utils.isBlank(row[idx.Z])) item.unresolved = true;
    if (!utils.isBlank(row[idx.AK])) item.hasAK = true;
    if (!utils.isBlank(row[idx.AL]) && String(row[idx.AL]).includes("?")) item.hasQuestion = true;
  }

  const rows = Array.from(grouped.values()).map((x) => {
    const isWrittenOff = x.hasAK && !x.hasQuestion;
    return {
      ...x,
      isWrittenOff,
      isNotWrittenOff: !isWrittenOff
    };
  });

  rows.sort((a, b) => {
    const diff = utils.num(a.year) - utils.num(b.year);
    return diff !== 0 ? diff : a.order.localeCompare(b.order, "uk");
  });

  return rows;
}

function buildYearTotals(rows, utils) {
  const yearTotals = new Map();

  for (const row of rows) {
    const yearKey = utils.trim(row.year);

    if (!yearTotals.has(yearKey)) {
      yearTotals.set(yearKey, {
        writtenOffSum: 0,
        writtenOffCount: 0,
        notWrittenOffSum: 0
      });
    }

    const totals = yearTotals.get(yearKey);

    if (row.isWrittenOff) {
      totals.writtenOffSum += row.sum;
      totals.writtenOffCount += 1;
    } else {
      totals.notWrittenOffSum += row.sum;
    }
  }

  return yearTotals;
}

function buildSummaryOutput(visibleRows, yearTotals, commonValue, cfg, utils) {
  const summaryColsCount = utils.colNum("Q");
  const output = [];
  const mergeRows = [];

  const summaryCols = {
    A: utils.colNum("A") - 1,
    B: utils.colNum("B") - 1,
    C: utils.colNum("C") - 1,
    D: utils.colNum("D") - 1,
    E: utils.colNum("E") - 1,
    F: utils.colNum("F") - 1,
    G: utils.colNum("G") - 1,
    H: utils.colNum("H") - 1,
    I: utils.colNum("I") - 1,
    M: utils.colNum("M") - 1,
    N: utils.colNum("N") - 1,
    O: utils.colNum("O") - 1,
    Q: utils.colNum("Q") - 1
  };

  let currentYear = null;
  let sheetRow = cfg.summaryStartRow;

  for (const item of visibleRows) {
    const yearKey = utils.trim(item.year);

    if (yearKey !== currentYear) {
      currentYear = yearKey;
      const totals = yearTotals.get(yearKey) || {
        writtenOffSum: 0,
        writtenOffCount: 0,
        notWrittenOffSum: 0
      };

      const totalRow = new Array(summaryColsCount).fill("");
      totalRow[summaryCols.A] = commonValue;
      totalRow[summaryCols.B] = item.year;
      totalRow[summaryCols.C] = "Списані втрати";
      totalRow[summaryCols.G] = totals.writtenOffSum;
      totalRow[summaryCols.H] = totals.writtenOffCount;
      totalRow[summaryCols.N] = totals.notWrittenOffSum;

      output.push(totalRow);
      mergeRows.push(sheetRow);
      sheetRow++;
    }

    const detailRow = new Array(summaryColsCount).fill("");
    detailRow[summaryCols.A] = commonValue;
    detailRow[summaryCols.B] = item.year;
    detailRow[summaryCols.C] = item.qValue;
    detailRow[summaryCols.D] = item.order;
    detailRow[summaryCols.E] = item.sum < 1.7 ? "Командиром в/ч" : "Командувачем АК";
    detailRow[summaryCols.F] = item.fValue;
    detailRow[summaryCols.G] = item.sum;
    detailRow[summaryCols.H] = 1;
    detailRow[summaryCols.I] = "";
    detailRow[summaryCols.M] = 1;
    detailRow[summaryCols.N] = item.sum;
    detailRow[summaryCols.O] = 1;
    detailRow[summaryCols.Q] = item.unresolved ? "не завершене службове розслідування" : "";

    output.push(detailRow);
    sheetRow++;
  }

  return { output, mergeRows, summaryColsCount };
}

function writeSummarySheet(summarySheet, summaryData, cfg, utils) {
  const summaryMaxRows = summarySheet.getMaxRows();
  const workRange = summarySheet.getRange(
    cfg.summaryStartRow,
    1,
    Math.max(1, summaryMaxRows - cfg.summaryStartRow + 1),
    summaryData.summaryColsCount
  );

  const mergedRanges = workRange.getMergedRanges();
  if (mergedRanges.length) mergedRanges.forEach((range) => range.breakApart());

  workRange.clearContent();
  workRange.setBackground(null);

  if (!summaryData.output.length) return;

  summarySheet
    .getRange(cfg.summaryStartRow, 1, summaryData.output.length, summaryData.summaryColsCount)
    .setValues(summaryData.output);

  const formatCols = {
    B: utils.colNum("B"),
    C: utils.colNum("C"),
    G: utils.colNum("G"),
    H: utils.colNum("H"),
    M: utils.colNum("M"),
    N: utils.colNum("N"),
    O: utils.colNum("O"),
    Q: utils.colNum("Q")
  };

  summarySheet.getRange(cfg.summaryStartRow, formatCols.B, summaryData.output.length, 1).setNumberFormat("0");
  summarySheet.getRange(cfg.summaryStartRow, formatCols.G, summaryData.output.length, 1).setNumberFormat("#,##0.00");
  summarySheet.getRange(cfg.summaryStartRow, formatCols.H, summaryData.output.length, 1).setNumberFormat("0");
  summarySheet.getRange(cfg.summaryStartRow, formatCols.M, summaryData.output.length, 1).setNumberFormat("0");
  summarySheet.getRange(cfg.summaryStartRow, formatCols.N, summaryData.output.length, 1).setNumberFormat("#,##0.00");
  summarySheet.getRange(cfg.summaryStartRow, formatCols.O, summaryData.output.length, 1).setNumberFormat("0");

  for (const rowNum of summaryData.mergeRows) {
    summarySheet.getRange(rowNum, formatCols.C, 1, 3).merge();
  }

  for (const rowNum of summaryData.mergeRows) {
    summarySheet.getRange(rowNum, formatCols.C).setHorizontalAlignment("center").setFontWeight("bold");
    summarySheet.getRange(rowNum, 1, 1, summaryData.summaryColsCount).setBackground("#d9ead3");
  }

  const qRange = summarySheet.getRange(cfg.summaryStartRow, formatCols.Q, summaryData.output.length, 1);
  const qVals = qRange.getValues();
  qRange.setBackgrounds(qVals.map(([v]) => [utils.trim(v) ? "#f4cccc" : null]));
}
