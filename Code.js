function exportBPV_10AK() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const CFG = {
    source: "Списання_розширене_1_2_3_4_5",
    target: "БВП 10 АК",
    summary: "БПВ що вівторка",
    general: "Загальні дані",
    sourceStartRow: 2,
    targetStartRow: 10,
    summaryStartRow: 10
  };

  let ui = null;
  try {
    ui = SpreadsheetApp.getUi();
  } catch (e) {}

  try {
    const source = ss.getSheetByName(CFG.source);
    const target = ss.getSheetByName(CFG.target);
    const summary = ss.getSheetByName(CFG.summary);
    const general = ss.getSheetByName(CFG.general);

    if (!source) throw new Error(`Не знайдено лист: ${CFG.source}`);
    if (!target) throw new Error(`Не знайдено лист: ${CFG.target}`);
    if (!summary) throw new Error(`Не знайдено лист: ${CFG.summary}`);
    if (!general) throw new Error(`Не знайдено лист: ${CFG.general}`);

    const lastRow = source.getLastRow();
    const lastCol = source.getLastColumn();
    const numRows = lastRow - CFG.sourceStartRow + 1;

    if (numRows <= 0) {
      if (ui) ui.alert("Немає даних для експорту");
      return;
    }

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

    const IDX = {
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

    const data = source.getRange(CFG.sourceStartRow, 1, numRows, lastCol).getValues();
    const commonValue = general.getRange("B2").getValue();

    // =====================================================
    // 1. БВП 10 АК — ТУРБО, ОДИН БЛОК A:Z
    // =====================================================
    const targetMaxRows = target.getMaxRows();
    const targetColsCount = colNum("Z"); // 26

    // Очистити всю робочу зону A:Z
    target
      .getRange(
        CFG.targetStartRow,
        1,
        Math.max(1, targetMaxRows - CFG.targetStartRow + 1),
        targetColsCount
      )
      .clearContent();

    // Формуємо масив A:Z
    const targetOutput = Array.from({ length: numRows }, () => new Array(targetColsCount).fill(""));

    for (let i = 0; i < numRows; i++) {
      const row = data[i];
      const zVal = trim(row[IDX.AK]); // в Z листа БВП 10 АК іде AK джерела

      // Основні колонки
      targetOutput[i][colNum("A") - 1] = row[IDX.E];
      targetOutput[i][colNum("B") - 1] = row[IDX.F];
      targetOutput[i][colNum("C") - 1] = row[IDX.H];
      targetOutput[i][colNum("D") - 1] = row[IDX.B];
      targetOutput[i][colNum("E") - 1] = row[IDX.T];
      targetOutput[i][colNum("G") - 1] = row[IDX.G];
      targetOutput[i][colNum("I") - 1] = row[IDX.I];
      targetOutput[i][colNum("J") - 1] = row[IDX.M];
      targetOutput[i][colNum("K") - 1] = row[IDX.P];
      targetOutput[i][colNum("P") - 1] = row[IDX.Y];
      targetOutput[i][colNum("Q") - 1] = row[IDX.Z];
      targetOutput[i][colNum("Z") - 1] = row[IDX.AK];

      // Логіка по Z:
      // якщо AK заповнена -> I джерела в L, P джерела в M
      // якщо AK порожня -> I джерела в N, P джерела в O
      if (zVal !== "") {
        targetOutput[i][colNum("L") - 1] = row[IDX.I];
        targetOutput[i][colNum("M") - 1] = row[IDX.P];
      } else {
        targetOutput[i][colNum("N") - 1] = row[IDX.I];
        targetOutput[i][colNum("O") - 1] = row[IDX.P];
      }
    }

    target
      .getRange(CFG.targetStartRow, 1, targetOutput.length, targetColsCount)
      .setValues(targetOutput);

    // =====================================================
    // 2. ГРУПУВАННЯ ДЛЯ БПВ що вівторка
    // =====================================================
    const grouped = new Map();

    for (let i = 0; i < numRows; i++) {
      const row = data[i];
      const order = trim(row[IDX.Y]);
      if (!order) continue;

      const year = row[IDX.T];
      const key = `${order}||${trim(year)}`;

      if (!grouped.has(key)) {
        grouped.set(key, {
          order,
          year: isBlank(year) ? "" : year,
          qValue: isBlank(row[IDX.Q]) ? "" : row[IDX.Q],
          fValue: isBlank(row[IDX.C]) ? "" : row[IDX.C],
          sum: 0,
          unresolved: false,
          hasAK: false,
          hasQuestion: false
        });
      }

      const item = grouped.get(key);
      item.sum += num(row[IDX.P]);

      if (isBlank(item.qValue) && !isBlank(row[IDX.Q])) item.qValue = row[IDX.Q];
      if (isBlank(item.fValue) && !isBlank(row[IDX.C])) item.fValue = row[IDX.C];
      if (isBlank(row[IDX.Z])) item.unresolved = true;
      if (!isBlank(row[IDX.AK])) item.hasAK = true;
      if (!isBlank(row[IDX.AL]) && String(row[IDX.AL]).includes("?")) item.hasQuestion = true;
    }

    let rows = Array.from(grouped.values()).map(x => {
      const isWrittenOff = x.hasAK && !x.hasQuestion;
      return {
        ...x,
        isWrittenOff,
        isNotWrittenOff: !isWrittenOff
      };
    });

    rows.sort((a, b) => {
      const diff = num(a.year) - num(b.year);
      return diff !== 0 ? diff : a.order.localeCompare(b.order, "uk");
    });

    // =====================================================
    // 3. ПІДСУМКИ ПО РОКАХ
    // =====================================================
    const yearTotals = new Map();

    for (const r of rows) {
      const y = trim(r.year);

      if (!yearTotals.has(y)) {
        yearTotals.set(y, {
          writtenOffSum: 0,
          writtenOffCount: 0,
          notWrittenOffSum: 0
        });
      }

      const t = yearTotals.get(y);

      if (r.isWrittenOff) {
        t.writtenOffSum += r.sum;
        t.writtenOffCount += 1;
      } else {
        t.notWrittenOffSum += r.sum;
      }
    }

    const visible = rows.filter(r => r.isNotWrittenOff);

    // =====================================================
    // 4. БПВ що вівторка — ОДИН БЛОК A:Q
    // =====================================================
    const summaryColsCount = colNum("Q"); // 17
    const output = [];
    const mergeRows = [];
    let currentYear = null;
    let sheetRow = CFG.summaryStartRow;

    for (const item of visible) {
      const yearKey = trim(item.year);

      if (yearKey !== currentYear) {
        currentYear = yearKey;

        const t = yearTotals.get(yearKey) || {
          writtenOffSum: 0,
          writtenOffCount: 0,
          notWrittenOffSum: 0
        };

        const totalRow = new Array(summaryColsCount).fill("");
        totalRow[colNum("A") - 1] = commonValue;
        totalRow[colNum("B") - 1] = item.year;
        totalRow[colNum("C") - 1] = "Списані втрати";
        totalRow[colNum("G") - 1] = t.writtenOffSum;
        totalRow[colNum("H") - 1] = t.writtenOffCount;
        totalRow[colNum("N") - 1] = t.notWrittenOffSum;

        output.push(totalRow);
        mergeRows.push(sheetRow);
        sheetRow++;
      }

      const r = new Array(summaryColsCount).fill("");
      r[colNum("A") - 1] = commonValue;
      r[colNum("B") - 1] = item.year;
      r[colNum("C") - 1] = item.qValue;
      r[colNum("D") - 1] = item.order;
      r[colNum("E") - 1] = item.sum < 1.7 ? "Командиром в/ч" : "Командувачем АК";
      r[colNum("F") - 1] = item.fValue;
      r[colNum("G") - 1] = item.sum;
      r[colNum("H") - 1] = 1;
      r[colNum("I") - 1] = "";
      r[colNum("M") - 1] = 1;
      r[colNum("N") - 1] = item.sum;
      r[colNum("O") - 1] = 1;
      r[colNum("Q") - 1] = item.unresolved ? "не завершене службове розслідування" : "";

      output.push(r);
      sheetRow++;
    }

    const summaryMaxRows = summary.getMaxRows();
    const workRange = summary.getRange(
      CFG.summaryStartRow,
      1,
      Math.max(1, summaryMaxRows - CFG.summaryStartRow + 1),
      summaryColsCount
    );

    const mergedRanges = workRange.getMergedRanges();
    if (mergedRanges.length) mergedRanges.forEach(r => r.breakApart());

    workRange.clearContent();
    workRange.setBackground(null);

    if (output.length) {
      summary
        .getRange(CFG.summaryStartRow, 1, output.length, summaryColsCount)
        .setValues(output);

      summary.getRange(CFG.summaryStartRow, colNum("B"), output.length, 1).setNumberFormat("0");
      summary.getRange(CFG.summaryStartRow, colNum("G"), output.length, 1).setNumberFormat("#,##0.00");
      summary.getRange(CFG.summaryStartRow, colNum("H"), output.length, 1).setNumberFormat("0");
      summary.getRange(CFG.summaryStartRow, colNum("M"), output.length, 1).setNumberFormat("0");
      summary.getRange(CFG.summaryStartRow, colNum("N"), output.length, 1).setNumberFormat("#,##0.00");
      summary.getRange(CFG.summaryStartRow, colNum("O"), output.length, 1).setNumberFormat("0");

      for (const rowNum of mergeRows) {
        summary.getRange(rowNum, colNum("C"), 1, 3).merge();
      }

      for (const rowNum of mergeRows) {
        summary.getRange(rowNum, colNum("C")).setHorizontalAlignment("center").setFontWeight("bold");
        summary.getRange(rowNum, 1, 1, summaryColsCount).setBackground("#d9ead3");
      }

      const qRange = summary.getRange(CFG.summaryStartRow, colNum("Q"), output.length, 1);
      const qVals = qRange.getValues();
      qRange.setBackgrounds(qVals.map(([v]) => [trim(v) ? "#f4cccc" : null]));
    }

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
