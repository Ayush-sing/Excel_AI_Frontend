// src/taskpane/components/App.tsx
import React, { useEffect, useRef, useState } from "react";
import Loader from "./Loader";

/**
 * Excel AI Assistant - Taskpane App (TypeScript)
 *
 * Changes made:
 * - Removed window.confirm usage (not supported in some Office hosts)
 * - Added explicit overwrite workflow: when a cell is occupied we ask user to type "overwrite"
 * - Avoid reading properties without .load() + context.sync()
 * - Do not render "Place in specific cell" button for charts (charts: use Place in Excel / Append / New)
 * - Defensive calls for Excel API differences ((Excel as any).Placement)
 *
 * Backend assumed at http://127.0.0.1:8000/chat
 */

/* ---------------- Excel helpers ---------------- */

async function getExcelData() {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRangeOrNullObject();
    usedRange.load("values, address, rowCount, columnCount");
    await context.sync();
    return {
      values: usedRange.values || [],
      address: usedRange.address || "A1",
      rowCount: usedRange.rowCount || 0,
      columnCount: usedRange.columnCount || 0,
    };
  });
}

async function findPlacementCell(columnName: string) {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRangeOrNullObject();
    usedRange.load("values, rowCount");
    await context.sync();

    const values = usedRange.values || [];
    if (!values || values.length === 0) {
      return { found: false, cellAddress: "A1", occupied: false };
    }

    const headerRow = values[0] || [];
    const targetIndex = headerRow.findIndex((h: any) =>
      h && String(h).toLowerCase().includes(columnName.toLowerCase())
    );

    if (targetIndex === -1) {
      return { found: false, cellAddress: null, occupied: false };
    }

    // find last non-empty row
    let lastRow = 0;
    for (let i = 1; i < values.length; i++) {
      const v = values[i][targetIndex];
      if (v !== null && v !== "" && typeof v !== "undefined") lastRow = i;
    }

    const nextRow = lastRow + 2;
    const colLetter = String.fromCharCode(65 + targetIndex); // only A-Z safe; if >26 you'll need helper
    const cellAddress = `${colLetter}${nextRow}`;

    const cell = sheet.getRange(cellAddress);
    cell.load("values");
    await context.sync();

    const cellValue = cell.values?.[0]?.[0];
    return { found: true, cellAddress, occupied: cellValue !== null && cellValue !== "" };
  });
}

/**
 * writeResult(note, newSheet=false, cellAddress?, overwrite=false)
 * - overwrite: if true, replace existing value without error
 */
async function writeResult(
  note: string,
  newSheet = false,
  cellAddress?: string,
  overwrite = false
) {
  return await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    let sheet: Excel.Worksheet;

    if (newSheet) {
      // enumerate sheets to pick next AI_Results_n
      sheets.load("items/name");
      await context.sync();
      const names: string[] = (sheets.items || []).map((s: any) => s.name || "");
      const matches = names.filter((n) => n.startsWith("AI_Results"));
      let next = 1;
      if (matches.length > 0) {
        const nums = matches
          .map((m) => {
            const p = m.split("_").pop();
            const v = parseInt(p || "0");
            return isNaN(v) ? 0 : v;
          })
          .filter((v) => typeof v === "number");
        if (nums.length > 0) next = Math.max(...nums) + 1;
      }
      const newName = `AI_Results_${next}`;
      sheet = sheets.add(newName);
      await context.sync();
    } else {
      sheet = context.workbook.worksheets.getActiveWorksheet();
    }

    const addr = cellAddress ? cellAddress.toUpperCase() : "A1";
    if (!/^[A-Z]+[0-9]+$/i.test(addr)) {
      throw new Error(`Invalid cell address: ${addr}`);
    }

    const range = sheet.getRange(addr);
    range.load("values");
    await context.sync();

    const cur = range.values?.[0]?.[0];
    if (cur !== null && cur !== "" && !overwrite) {
      throw new Error(`Cell ${addr} already contains data.`);
    }

    const m = (note || "").toString().match(/[-+]?[0-9]*\.?[0-9]+/);
    const value: any = m ? parseFloat(m[0]) : note;
    range.values = [[value]];
    range.format.font.bold = true;
    range.format.horizontalAlignment = "Right";
    range.format.autofitColumns();
    await context.sync();
  });
}

/**
 * Insert chart image at a reasonable position on the *currently visible area* approximation.
 * This updated function tries to center on the user's active cell (so it appears in the user's viewport).
 * Fallback: if active cell bounds are unavailable, we fallback to previous used-range heuristic.
 */
async function insertChartImageOnSheet(chartBase64: string, sheetName?: string | null) {
  return await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    let sheet: Excel.Worksheet;
    if (sheetName) {
      const maybe = sheets.getItemOrNullObject(sheetName);
      await context.sync();
      sheet = (maybe && (maybe as any).isNullObject) ? sheets.add(sheetName) : sheets.getItem(sheetName);
    } else {
      sheet = context.workbook.worksheets.getActiveWorksheet();
    }

    const shapesAny = (sheet as any).shapes as any;

    // default size for image
    const desiredWidth = 340;
    const desiredHeight = 180;

    // Try to get the active cell (user viewport reference)
    try {
      const activeCell = context.workbook.getActiveCell();
      activeCell.load(["left", "top", "width", "height", "address"]);
      await context.sync();

      // create image first (base64 without data URL prefix if needed)
      let base64 = chartBase64;
      const prefixIndex = base64.indexOf("base64,");
      if (prefixIndex >= 0) base64 = base64.substring(prefixIndex + 7);

      const image = shapesAny.addImage(base64);

      // set the image desired size first (some hosts require sync to update width/height)
      try {
        image.width = desiredWidth;
        image.height = desiredHeight;
      } catch (e) {
        // ignore if host doesn't allow setting immediately
        console.warn("Could not set initial image size:", e);
      }

      // sync to ensure width/height are materialized
      await context.sync();

      // Load image width/height (defensive)
      try {
        image.load(["width", "height"]);
        await context.sync();
      } catch {
        // ignore if not supported
      }

      // Compute center position based on activeCell
      const acLeft = typeof activeCell.left === "number" ? activeCell.left : 20;
      const acTop = typeof activeCell.top === "number" ? activeCell.top : 20;
      const acWidth = typeof activeCell.width === "number" ? activeCell.width : 64;
      const acHeight = typeof activeCell.height === "number" ? activeCell.height : 20;

      // Determine final image size (use read value if available)
      const imgWidth = (image.width && typeof image.width === "number") ? image.width : desiredWidth;
      const imgHeight = (image.height && typeof image.height === "number") ? image.height : desiredHeight;

      // center image horizontally/vertically around active cell
      const finalLeft = Math.max(8, Math.round(acLeft + acWidth / 2 - imgWidth / 2));
      const finalTop = Math.max(8, Math.round(acTop + acHeight / 2 - imgHeight / 2));

      try {
        image.left = finalLeft;
        image.top = finalTop;
      } catch (e) {
        console.warn("Could not set image position directly:", e);
      }

      // placement (defensive)
      try {
        const placementAny = (Excel as any).Placement;
        if (placementAny && placementAny.twoCell) {
          image.placement = placementAny.twoCell;
        } else if (placementAny && placementAny.moveAndSize) {
          image.placement = placementAny.moveAndSize;
        } else {
          image.placement = Excel.Placement.absolute;
        }
      } catch {
        try {
          image.placement = Excel.Placement.absolute;
        } catch {
          // swallow
        }
      }

      // name the image
      try {
        image.name = `AIChart_${Date.now()}`;
      } catch { }

      await context.sync();
      return;
    } catch (activeCellErr) {
      // If active cell approach fails, fallback to used-range heuristic (previous behavior)
      console.warn("Active-cell placement failed, falling back to used-range heuristic:", activeCellErr);
      try {
        const used = sheet.getUsedRangeOrNullObject();
        used.load("rowCount");
        await context.sync();
        const rowCount = used.rowCount || 1;
        const top = Math.max(20, (rowCount + 2) * 18);
        const left = 20;

        let base64 = chartBase64;
        const prefixIndex = base64.indexOf("base64,");
        if (prefixIndex >= 0) base64 = base64.substring(prefixIndex + 7);

        const image = shapesAny.addImage(base64);
        try {
          image.left = left;
          image.top = top;
          image.width = desiredWidth;
          image.height = desiredHeight;
        } catch (e) {
          console.warn("Could not set some image properties in fallback:", e);
        }

        try {
          const placementAny = (Excel as any).Placement;
          if (placementAny && placementAny.twoCell) {
            image.placement = placementAny.twoCell;
          } else if (placementAny && placementAny.moveAndSize) {
            image.placement = placementAny.moveAndSize;
          } else {
            image.placement = Excel.Placement.absolute;
          }
        } catch {
          try {
            image.placement = Excel.Placement.absolute;
          } catch { }
        }

        try {
          image.name = `AIChart_${Date.now()}`;
        } catch { }

        await context.sync();
        return;
      } catch (fallbackErr) {
        console.error("Fallback placement also failed:", fallbackErr);
        // rethrow to let caller show an error message if needed
        throw fallbackErr;
      }
    }
  });
}

/* ---------------- Upload placement helpers ---------------- */

/**
 * Create a new worksheet and write the uploaded data (headers + rows)
 * dataRows: array of rows (each row array) WITHOUT headers (we pass headers separately)
 */
async function createSheetFromUpload(sheetName: string, headers: string[], dataRows: any[][]) {
  return await Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    const newSheet = sheets.add(sheetName);
    // write headers
    const headerRange = newSheet.getRange("A1").getResizedRange(0, headers.length - 1);
    headerRange.values = [headers];
    // write rows if any
    if (dataRows && dataRows.length > 0) {
      const rowsRange = newSheet.getRange(`A2`).getResizedRange(dataRows.length - 1, headers.length - 1);
      rowsRange.values = dataRows;
    }
    headerRange.format.font.bold = true;
    await context.sync();
    return { ok: true, sheetName: newSheet.name };
  });
}

/**
 * Append dataRows (array of arrays) to the active sheet, aligning by headers.
 * If headers match exactly -> append vertically at bottom aligned to columns
 * headers param is array of column names from uploaded file
 *
 * NOTE: Modified to always append horizontally (to the right) when invoked via the "Merge horizontally" button.
 */
// 🆕 Always place data horizontally (side-by-side), preserving existing headers and data
async function appendUploadToActiveSheet(headers: string[], dataRows: any[][]) {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // load the first row to check for existing headers
    const used = sheet.getUsedRangeOrNullObject();
    used.load("columnCount, values");
    await context.sync();

    let startColumn = 0; // default (A)
    let hasHeader = false;

    if (used && !(used as any).isNullObject && used.values && used.values.length > 0) {
      const firstRow = used.values[0] || [];
      // find last non-empty cell in the first row
      const lastFilled = firstRow.reduceRight(
        (acc, val, idx) => (acc === -1 && val !== "" && val !== null ? idx : acc),
        -1
      );
      if (lastFilled >= 0) {
        hasHeader = true;
        startColumn = lastFilled + 1; // next empty column
      }
    }

    // write headers at row 1, starting from startColumn
    const headerRange = sheet.getRangeByIndexes(0, startColumn, 1, headers.length);
    headerRange.values = [headers];
    headerRange.format.font.bold = true;

    // write data rows below headers
    if (dataRows.length > 0) {
      const dataRange = sheet.getRangeByIndexes(1, startColumn, dataRows.length, headers.length);
      dataRange.values = dataRows;
    }

    await context.sync();
    const startCell = `${indexToColumnName(startColumn + 1)}1`;
    return { ok: true, mode: "always_horizontal", startCell };
  });
}

// 🆕 Always append vertically (below existing data)
async function appendUploadToActiveSheetVertical(headers: string[], dataRows: any[][]) {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const used = sheet.getUsedRangeOrNullObject();
    used.load("rowCount, columnCount, values");
    await context.sync();

    let startRow = 0;
    let hasHeader = false;

    if (used && !(used as any).isNullObject && used.values && used.values.length > 0) {
      const firstRow = used.values[0] || [];
      const hasAnyHeader = firstRow.some((val) => val !== "" && val !== null);
      if (hasAnyHeader) {
        hasHeader = true;
        startRow = used.values.length; // append below all data
      }
    }

    // if no header, write header at top
    if (!hasHeader) {
      const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
      headerRange.values = [headers];
      headerRange.format.font.bold = true;
      startRow = 1;
    }

    // write data below existing rows
    const dataRange = sheet.getRangeByIndexes(startRow, 0, dataRows.length, headers.length);
    dataRange.values = dataRows;
    await context.sync();

    const startCell = `A${startRow + 1}`;
    return { ok: true, mode: "always_vertical", startCell };
  });
}


/** Convert column index (1-based) to Excel letter */
function indexToColumnName(index: number): string {
  let colName = "";
  while (index > 0) {
    const rem = (index - 1) % 26;
    colName = String.fromCharCode(65 + rem) + colName;
    index = Math.floor((index - 1) / 26);
  }
  return colName;
}


async function writeUploadToEmptySheet(headers: string[], dataRows: any[][]) {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const used = sheet.getUsedRangeOrNullObject();
    used.load("address, rowCount");
    await context.sync();
    if (used && !(used as any).isNullObject && used.rowCount > 0) {
      return { ok: false, reason: "sheet_not_empty" };
    }
    const headerRange = sheet.getRange("A1").getResizedRange(0, Math.max(0, headers.length - 1));
    headerRange.values = [headers];
    if (dataRows.length > 0) {
      const rowsRange = sheet.getRange("A2").getResizedRange(dataRows.length - 1, Math.max(0, headers.length - 1));
      rowsRange.values = dataRows;
    }
    headerRange.format.font.bold = true;
    await context.sync();
    return { ok: true, mode: "wrote_to_empty_sheet" };
  });
}

/* small util: case-insensitive array equality for first n */
function arraysEqualCI(a: string[], b: string[]) {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) {
    if (String(a[i]).toLowerCase() !== String(b[i]).toLowerCase()) return false;
  }
  return true;
}

/* ---------------- React component ---------------- */

type Message = { role: "user" | "ai"; text: string; chart?: string | null };

type PendingUpload = {
  fileId: string;
  fileName: string;
  headers: string[];
  rows: any[][];
};

export default function App(): JSX.Element {
  const [bootLoading, setBootLoading] = useState(true);
  // session-only persistence (per your preference)
  const [messages, setMessages] = useState<Message[]>(() => {
    try {
      const raw = sessionStorage.getItem("excel_ai_chat_v1");
      return raw ? (JSON.parse(raw) as Message[]) : [];
    } catch {
      return [];
    }
  });

  useEffect(() => {
    try {
      sessionStorage.setItem("excel_ai_chat_v1", JSON.stringify(messages));
    } catch { }
  }, [messages]);

  useEffect(() => {
    const t = setTimeout(() => setBootLoading(false), 2100);
    return () => clearTimeout(t);
  }, []);

  const [input, setInput] = useState("");
  const [pendingPlacement, setPendingPlacement] = useState<{
    note: string;
    chart?: string | null;
    suggestedColumn?: string | null;
    awaitingCell?: boolean; // when true, next user input is cell address
    awaitingOverwrite?: { sheet?: string | null; addr: string; chart?: boolean }; // when true, user must type 'overwrite'
  } | null>(null);

  const [pendingUpload, setPendingUpload] = useState<PendingUpload | null>(null);

  const [loading, setLoading] = useState(false);
  const [loadingMsg, setLoadingMsg] = useState("Thinking...");
  const loadingTimerRef = useRef<number | null>(null);
  const chatContainerRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    document.body.style.background =  "#0f1720";
    // small responsive hint for sidebar expansion
    document.body.style.maxWidth = "600px";
    document.body.style.transition = "width 0.2s ease";
  }, []);

  useEffect(() => {
    // autoscroll to bottom when messages update
    if (chatContainerRef.current) {
      chatContainerRef.current.scrollTop = chatContainerRef.current.scrollHeight;
    }
  }, [messages, pendingUpload]);

  const pushMessage = (m: Message) => setMessages((prev) => [...prev, m]);

  function clearPendingPlacement() {
    setPendingPlacement(null);
  }

  // ----- Placement actions -----

  /**
   * For charts we call this "Place in Excel" (keeps moveable behavior).
   * For non-charts this acts like Place below data (finds suggested column).
   */
  async function handlePlaceInExcelForChartOrBelow() {
    if (!pendingPlacement) return;
    setLoading(true);
    setLoadingMsg("Placing result...");
    try {
      if (pendingPlacement.chart) {
        // Chart placement flow (chart-only)
        await insertChartImageOnSheet(pendingPlacement.chart, null);
        pushMessage({ role: "ai", text: "✅ Chart placed at the top of the sheet (moveable chart)." });
      } else {
        // Non-chart: try suggested column or fallback to A1
        if (pendingPlacement.suggestedColumn) {
          const placement = await findPlacementCell(pendingPlacement.suggestedColumn);
          if (!placement.found) {
            pushMessage({ role: "ai", text: `❌ Could not find column '${pendingPlacement.suggestedColumn}'.` });
          } else if (placement.occupied) {
            // we cannot use confirm; instruct user to type 'overwrite' to confirm
            pushMessage({
              role: "ai",
              text: `⚠️ Suggested cell ${placement.cellAddress} is occupied. Type 'overwrite ${placement.cellAddress}' in the input and press Enter to force overwrite.`,
            });
            setPendingPlacement({
              ...pendingPlacement,
              awaitingOverwrite: { addr: placement.cellAddress, chart: false },
            });
            setLoading(false);
            setLoadingMsg("Thinking...");
            return;
          } else {
            await writeResult(pendingPlacement.note, false, placement.cellAddress);
            pushMessage({ role: "ai", text: `✅ Result placed at ${placement.cellAddress}` });
          }
        } else {
          await writeResult(pendingPlacement.note, false, "A1");
          pushMessage({ role: "ai", text: `✅ Result placed at A1 (fallback).` });
        }
      }
    } catch (e: any) {
      pushMessage({ role: "ai", text: `⚠️ Failed to place: ${String(e?.message || e)}` });
    } finally {
      clearPendingPlacement();
      setLoading(false);
      setLoadingMsg("Thinking...");
    }
  }

  // "Place in specific cell" now only shown for non-chart items;
  // when clicked we set awaitingCell=true and ask user to type address into the input.
  function handlePlaceInSpecificCell() {
    if (!pendingPlacement) {
      pushMessage({ role: "ai", text: "⚠️ No result waiting for placement." });
      return;
    }
    setPendingPlacement({ ...pendingPlacement, awaitingCell: true });
    pushMessage({
      role: "ai",
      text: "Please type the target cell (like A1 or C5) in the input box and press Enter to place the result.",
    });
  }

  async function handlePlaceOnResultsSheet(createNew = false) {
    if (!pendingPlacement) return;
    setLoading(true);
    setLoadingMsg("Placing on results sheet...");
    try {
      // find latest AI_Results or create new
      let targetSheetName: string | null = null;
      await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();
        const names = sheets.items.map((s: any) => s.name || "");
        const matches = names.filter((n: string) => n.startsWith("AI_Results"));
        if (matches.length === 0 || createNew) {
          targetSheetName = null;
        } else {
          matches.sort();
          targetSheetName = matches[matches.length - 1];
        }
      });

      if (!targetSheetName || createNew) {
        if (pendingPlacement.chart) {
          // create sheet and later insert
          await Excel.run(async (context) => {
            const sheets = context.workbook.worksheets;
            sheets.load("items/name");
            await context.sync();
            const names = sheets.items.map((s: any) => s.name || "");
            const matches = names.filter((n: string) => n.startsWith("AI_Results"));
            let next = 1;
            if (matches.length > 0) {
              const nums = matches
                .map((m: string) => {
                  const p = m.split("_").pop();
                  const v = parseInt(p || "0");
                  return isNaN(v) ? 0 : v;
                })
                .filter((v) => typeof v === "number");
              if (nums.length > 0) next = Math.max(...nums) + 1;
            }
            const newName = `AI_Results_${next}`;
            sheets.add(newName);
            await context.sync();
            targetSheetName = newName;
          });
        } else {
          // writeResult handles creating new sheet
          await writeResult(pendingPlacement.note, true, "A1");
          pushMessage({ role: "ai", text: `✅ Result created on new results sheet.` });
          clearPendingPlacement();
          setLoading(false);
          setLoadingMsg("Thinking...");
          return;
        }
      }

      // now insert into targetSheetName
      if (pendingPlacement.chart) {
        await insertChartImageOnSheet(pendingPlacement.chart, targetSheetName);
        pushMessage({ role: "ai", text: `✅ Chart placed on ${targetSheetName}` });
      } else {
        await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getItem(targetSheetName!);
          const used = sheet.getUsedRangeOrNullObject();
          used.load("rowCount");
          await context.sync();
          const row = used.rowCount ?? 0;
          const putCell = `A${row + 1}`;
          // Write full descriptive text instead of only numeric value
          // --- Clean descriptive text (no emojis, no gaps, no overlaps) ---
          const cleanText = pendingPlacement.note
            .replace(/^[^a-zA-Z0-9]+/, "") // remove starting emojis or symbols
            .replace(/\n+/g, " ")           // remove newlines
            .trim();

          const descriptiveText = `${cleanText}`;

          // reuse existing putCell, don’t redeclare it
          const range = sheet.getRange(putCell);
          range.values = [[descriptiveText]];
          range.format.font.bold = true;
          range.format.autofitColumns();
          range.format.autofitRows();


          await context.sync();
          pushMessage({ role: "ai", text: `✅ Result appended to ${targetSheetName} at ${putCell}` });
        });
      }

      clearPendingPlacement();
    } catch (e: any) {
      pushMessage({ role: "ai", text: `⚠️ Failed to place on results sheet: ${String(e?.message || e)}` });
    } finally {
      setLoading(false);
      setLoadingMsg("Thinking...");
    }
  }

  function handleDontPlace() {
    if (pendingPlacement) pushMessage({ role: "ai", text: "ℹ️ Placement skipped. Preview kept in chat." });
    clearPendingPlacement();
  }

  // ----- send to backend & handle input -----

  async function handleSend() {
    // 1) If awaiting overwrite confirmation: expect "overwrite <CELL>" or just "overwrite"
    if (pendingPlacement?.awaitingOverwrite) {
      const text = input.trim().toLowerCase();
      const addr = pendingPlacement.awaitingOverwrite.addr;
      if (text === `overwrite` || text === `overwrite ${addr.toLowerCase()}`) {
        // proceed to write with overwrite=true
        setLoading(true);
        setLoadingMsg(`Overwriting ${addr}...`);
        try {
          if (pendingPlacement.chart) {
            // for chart overwrite scenario, we still just insert chart (no per-cell overwrite)
            await insertChartImageOnSheet(pendingPlacement.chart, null);
            pushMessage({ role: "ai", text: `✅ Chart placed (overwrite flow).` });
          } else {
            await writeResult(pendingPlacement.note, false, addr, true);
            pushMessage({ role: "ai", text: `✅ Overwrote and placed result at ${addr}.` });
          }
          clearPendingPlacement();
        } catch (err: any) {
          pushMessage({ role: "ai", text: `⚠️ Failed to overwrite: ${String(err?.message || err)}` });
        } finally {
          setLoading(false);
          setLoadingMsg("Thinking...");
          setInput("");
        }
        return;
      } else {
        pushMessage({ role: "ai", text: "❌ To confirm overwrite type 'overwrite' (or 'overwrite <CELL>')." });
        setInput("");
        return;
      }
    }

    // 2) If awaiting a specific cell address (non-chart case)
    if (pendingPlacement?.awaitingCell) {
      const addr = input.trim().toUpperCase();
      if (!/^[A-Z]+[0-9]+$/.test(addr)) {
        pushMessage({ role: "ai", text: "❌ Invalid cell address. Example: C5, AA10. Please enter again." });
        setInput("");
        return;
      }

      setLoading(true);
      setLoadingMsg(`Placing result in ${addr}...`);
      try {
        await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const range = sheet.getRange(addr);
          range.load("values");
          await context.sync();
          const value = range.values?.[0]?.[0];
          if (value !== null && value !== "") {
            // instead of confirm, instruct to type 'overwrite'
            pushMessage({
              role: "ai",
              text: `⚠️ Cell ${addr} is occupied. Type 'overwrite ${addr}' in the input and press Enter to force overwrite.`,
            });
            setPendingPlacement({
              ...pendingPlacement,
              awaitingCell: false,
              awaitingOverwrite: { addr, chart: false },
            });
            setLoading(false);
            setLoadingMsg("Thinking...");
            setInput("");
            return;
          }

          // safe to write
          if (pendingPlacement.chart) {
            // charts: we keep chart insertion behavior - insert near used area
            await insertChartImageOnSheet(pendingPlacement.chart, null);
            pushMessage({ role: "ai", text: `✅ Chart placed (approx) near ${addr}.` });
          } else {
            await writeResult(pendingPlacement.note, false, addr);
            pushMessage({ role: "ai", text: `✅ Result placed at ${addr}.` });
          }
        });

        clearPendingPlacement();
      } catch (err: any) {
        pushMessage({ role: "ai", text: `⚠️ Failed to place: ${String(err?.message || err)}` });
      } finally {
        setLoading(false);
        setLoadingMsg("Thinking...");
        setInput("");
      }
      return;
    }

    // 3) If a general pendingPlacement exists (but not awaiting input), block new commands
    if (pendingPlacement) {
      pushMessage({ role: "ai", text: "⚠️ Finish placement before sending a new command." });
      return;
    }

    // 4) Normal command flow -> send to backend
    if (!input.trim()) return;
    const userText = input.trim();
    pushMessage({ role: "user", text: userText });
    setInput("");
    setLoading(true);
    setLoadingMsg("Thinking...");

    if (loadingTimerRef.current) window.clearTimeout(loadingTimerRef.current);
    loadingTimerRef.current = window.setTimeout(() => {
      setLoadingMsg("Still working — this may take a few seconds.");
    }, 2500);

    try {
      const excelData = await getExcelData();
      const payload = { user_message: userText, excel_data: excelData.values, chat_history: messages };

      const res = await fetch("https://excel-ai-backend-j0zh.onrender.com/chat", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });

      const data = await res.json();

      if (loadingTimerRef.current) {
        window.clearTimeout(loadingTimerRef.current);
        loadingTimerRef.current = null;
      }

      if (!data) {
        pushMessage({ role: "ai", text: "❌ No response from backend." });
        setLoading(false);
        return;
      }
      if (!data.ok) {
        pushMessage({ role: "ai", text: data.note || "❌ Backend returned an error." });
        setLoading(false);
        return;
      }

      const aiText = data.note || "✅ Done";
      const chart = data.chart || null;

      // push assistant preview with chart if present
      pushMessage({ role: "ai", text: aiText, chart });

      // suggest placement: if user typed 'sum <col>' then suggest that column
      const m = userText.match(/sum\s+(\w+)/i);
      setPendingPlacement({ note: aiText, chart, suggestedColumn: m ? m[1] : null, awaitingCell: false });
    } catch (e: any) {
      pushMessage({ role: "ai", text: `❌ Network/server error: ${String(e?.message || e)}` });
    } finally {
      setLoading(false);
      setLoadingMsg("Thinking...");
    }
  }

  const tooltipStyle = `
[data-tooltip]::after {
  content: attr(data-tooltip);
  position: absolute;
  bottom: calc(100% + 8px);
  left: 0;
  right: 0;
  margin: auto;

  width: max-content;
  max-width: 220px;

  background: #0b1220;
  color: #e5e7eb;
  padding: 4px 8px;
  border-radius: 6px;
  font-size: 12px;
  white-space: nowrap;
  text-align: center;

  opacity: 0;
  pointer-events: none;
  transition: opacity 0.15s ease;
  border: 1px solid #1f2937;
  z-index: 9999;
}

/* ✅ Special handling for Send button */
[data-tooltip-pos="right"]::after {
  left: auto;
  right: 0;
  transform: none;
}

[data-tooltip]:hover::after {
  opacity: 1;
}

`;


  if (bootLoading) {
    return (
      <div
        style={{
          height: "100vh",
          display: "flex",
          alignItems: "center",
          justifyContent: "center",
          background: "#0f1720",
        }}
      >
        <Loader />
      </div>
    );
  }

  // ----- UI -----
  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100vh", fontFamily: "Segoe UI, Roboto, system-ui", background: "#06111a" }}>
      <style>
        {`
      button {
        cursor: pointer;
      }
    `}
      </style>
      <style>{tooltipStyle}</style>
      {/* header */}
      <div
        style={{
          position: "sticky",
          top: 0,
          zIndex: 20,
          padding: 10,
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          background: "rgb(6,17,26)",
        }}
      >
        <h2 style={{ margin: 0, color: "#646464" }}>📊 Excel AI Assistant</h2>
        <button
          onClick={() => {
            window.open(
              `${window.location.origin}/assets/Excel_AI_Assistant_Examples.pdf`,
              "_blank"
            );
          }}

          style={{
            background: "transparent",
            border: "1px solid rgb(104, 119, 140)",
            color: "#e5e7eb",
            padding: "6px 10px",
            borderRadius: 8,
            cursor: "pointer",
            fontSize: 13,
          }}
        >
          Example
        </button>

        
      </div>

      {/* chat container */}
      <div ref={chatContainerRef} style={{ flex: 1, overflowY: "auto", padding: 12, background: "#06111a", minHeight: 0 }}>
        {messages.map((m, i) => (
          <div key={i} style={{ display: "flex", justifyContent: m.role === "user" ? "flex-end" : "flex-start", margin: "6px 0" }}>
            <div style={{ background: m.role === "user" ? "#16273a" : "#E9EBEE", color: m.role === "user" ? "#ffffff" : "#000000", padding: 10, borderRadius: 12, maxWidth: "78%", whiteSpace: "pre-wrap" }}>
              <div dangerouslySetInnerHTML={{ __html: m.text.replace(/\n/g, "<br/>") }} />
              {m.chart && (
                <div style={{ marginTop: 8 }}>
                  <img src={`data:image/png;base64,${m.chart}`} alt="chart preview" style={{ width: "100%", borderRadius: 8 }} />
                  <div style={{ marginTop: 6, display: "flex", gap: 8 }}>
                    <a
                      href={`data:image/png;base64,${m.chart}`}
                      download={`chart_${Date.now()}.png`}
                      style={{
                        padding: "6px 10px",
                        borderRadius: 8,
                        background: "#e6eef8",
                        textDecoration: "none",
                        color: "#0b1220",
                      }}
                    >
                      Download Chart
                    </a>
                  </div>
                </div>
              )}
            </div>
          </div>
        ))}
        {loading && (
          <div style={{ marginTop: 8, color: "#ffff" }}>
            <em>{loadingMsg}</em>
          </div>
        )}
      </div>

      {/* pendingUpload in-chat placement UI */}
      {pendingUpload && (
        <div style={{
          padding: 12,
          borderRadius: 16,
          background: "#0b1220",
          border: "2px solid rgb(104, 119, 140)",
          margin: "8px 0",
          maxWidth: "85%",
          boxShadow: "0 4px 12px rgba(0,0,0,0.25)"
}}>
          <div style={{ fontWeight: 600, marginBottom: 8, color: "#e5e7eb" }}>
            Where should I place the uploaded file <em>{pendingUpload.fileName}</em>?
          </div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <button
              onClick={async () => {
                // create new sheet name using filename (sanitized) + timestamp
                let base = (pendingUpload.fileName || "Uploaded").replace(/\.[^/.]+$/, "").replace(/[^\w\- ]/g, "");
                const newName = `${base}_${Date.now() % 10000}`;
                await createSheetFromUpload(newName, pendingUpload.headers, pendingUpload.rows);
                pushMessage({ role: "ai", text: `✅ Created sheet '${newName}' with uploaded data.` });
                setPendingUpload(null);
              }}
            >
              🗂️ Create new sheet
            </button>

            <button
              onClick={async () => {
                try {
                  const writeRes = await Excel.run(async (context) => {
                    const sheet = context.workbook.worksheets.getActiveWorksheet();
                    const used = sheet.getUsedRangeOrNullObject();
                    used.load("rowCount");
                    await context.sync();
                    return { isEmpty: (used && (used as any).isNullObject) || used.rowCount === 0 };
                  });
                  if (writeRes.isEmpty) {
                    await writeUploadToEmptySheet(pendingUpload.headers, pendingUpload.rows);
                    pushMessage({ role: "ai", text: "✅ Uploaded data written to active (empty) sheet." });
                  } else {
                    const appendRes = await appendUploadToActiveSheetVertical(pendingUpload.headers, pendingUpload.rows);
                    if (appendRes.ok) pushMessage({ role: "ai", text: `✅ Data appended: mode=${appendRes.mode}, start=${(appendRes as any).startCell}` });
                    else pushMessage({ role: "ai", text: `⚠️ Append failed: ${(appendRes as any).reason || "unknown"}` });
                  }
                } catch (err: any) {
                  pushMessage({ role: "ai", text: `⚠️ Append failed: ${String(err?.message || err)}` });
                } finally {
                  setPendingUpload(null);
                }
              }}
            >
              📎 Append below the data
            </button>

            <button
              onClick={async () => {
                try {
                  const appendRes = await appendUploadToActiveSheet(pendingUpload.headers, pendingUpload.rows);
                  if (appendRes.ok) pushMessage({ role: "ai", text: `✅ Data merged horizontally (mode: ${appendRes.mode}).` });
                  else pushMessage({ role: "ai", text: `⚠️ Merge failed: ${(appendRes as any).reason || "unknown"}` });
                } catch (err: any) {
                  pushMessage({ role: "ai", text: `⚠️ Merge failed: ${String(err?.message || err)}` });
                } finally {
                  setPendingUpload(null);
                }
              }}
            >
              🔀 Merge horizontally
            </button>

            <button
              onClick={() => {
                pushMessage({ role: "ai", text: "ℹ️ Upload skipped. File retained on server temporarily." });
                setPendingUpload(null);
              }}
            >
              🚫 Cancel
            </button>
          </div>
        </div>
      )}

      {/* placement UI (when pending for results) */}
      {pendingPlacement && (
        <div style={{
          padding: 12,
          borderRadius: 16,
          background: "#0b1220",
          border: "2px solid rgb(104, 119, 140)",
          margin: "8px 0",
          maxWidth: "85%",
          boxShadow: "0 4px 12px rgba(0,0,0,0.25)",
        }}>
          <div style={{ fontWeight: 600, marginBottom: 8, color: "#e5e7eb" }}>Where should I place the result?</div>

          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            {/* For charts: show Place in Excel (chart-only) and sheet options, but remove Place in specific cell */}
            {pendingPlacement.chart ? (
              <>
                <button
                  onClick={() => {
                    console.log("✅ PlaceChart clicked");
                    handlePlaceInExcelForChartOrBelow();
                  }}
                >
                  Place in Excel
                </button>
                <button
                  onClick={() => {
                    console.log("✅ AppendResults clicked");
                    handlePlaceOnResultsSheet(false);
                  }}
                >
                  Append to AI_Results sheet
                </button>
                <button
                  onClick={() => {
                    console.log("✅ NewResults clicked");
                    handlePlaceOnResultsSheet(true);
                  }}
                >
                  Create new AI_Results sheet
                </button>
                <button
                  onClick={() => {
                    console.log("✅ DontPlace clicked");
                    handleDontPlace();
                  }}
                >
                  Don't place
                </button>
              </>
            ) : (
              // Non-chart items: keep Place in specific cell + other options
              <>
                <button
                  onClick={() => {
                    console.log("✅ PlaceBelow clicked");
                    handlePlaceInExcelForChartOrBelow();
                  }}
                >
                  Place below data (same sheet)
                </button>

                <button
                  onClick={() => {
                    console.log("✅ PlaceInSpecificCell clicked");
                    handlePlaceInSpecificCell();
                  }}
                >
                  Place in specific cell
                </button>

                <button
                  onClick={() => {
                    console.log("✅ AppendResults clicked");
                    handlePlaceOnResultsSheet(false);
                  }}
                >
                  Append to AI_Results sheet
                </button>

                <button
                  onClick={() => {
                    console.log("✅ NewResults clicked");
                    handlePlaceOnResultsSheet(true);
                  }}
                >
                  Create new AI_Results sheet
                </button>

                <button
                  onClick={() => {
                    console.log("✅ DontPlace clicked");
                    handleDontPlace();
                  }}
                >
                  Don't place
                </button>
              </>
            )}
          </div>
        </div>
      )}

      {/* input row sticky bottom */}
      <div
        style={{
          position: "sticky",
          bottom: 13,
          padding: 6,
          margin: "0px 11px",
          borderRadius: "2rem",
          display: "flex",
          gap: 8,
          background: "rgb(22, 39, 58)",
          alignItems: "center",
        }}
      >
        {/* Upload button (icon) */}
        <label
          htmlFor="fileUpload"
          data-tooltip="Upload file"
          style={{
            cursor: "pointer",
            fontSize: 18,
            padding: "8px 10px",
            borderRadius: 10,
            background: "rgb(22, 39, 58)",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            position: "relative",
            fill: "#ffffff"
          }}
        >
          <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 16 16">
            <path d="M.5 9.9a.5.5 0 0 1 .5.5v2.5a1 1 0 0 0 1 1h12a1 1 0 0 0 1-1v-2.5a.5.5 0 0 1 1 0v2.5a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2v-2.5a.5.5 0 0 1 .5-.5" />
            <path d="M7.646 1.146a.5.5 0 0 1 .708 0l3 3a.5.5 0 0 1-.708.708L8.5 2.707V11.5a.5.5 0 0 1-1 0V2.707L5.354 4.854a.5.5 0 1 1-.708-.708z" />
          </svg>
        </label>
        <input
          id="fileUpload"
          type="file"
          accept=".xlsx,.xls,.csv"
          style={{ display: "none" }}
          onChange={async (e) => {
            // Non-blocking in-chat upload flow (no window.confirm / prompt)
            const file = e.target.files?.[0];
            if (!file) return;
            pushMessage({ role: "ai", text: `📤 Uploading ${file.name}...` });

            const formData = new FormData();
            formData.append("file", file);

            try {
              const res = await fetch("https://excel-ai-backend-j0zh.onrender.com/upload", { method: "POST", body: formData });
              const data = await res.json();
              if (!data || !data.ok) {
                pushMessage({ role: "ai", text: `⚠️ Upload failed: ${data?.detail || "unknown error"}` });
                return;
              }

              const parsed = data.parsed;
              const firstSheet = parsed.sheets && parsed.sheets.length > 0 ? parsed.sheets[0] : null;
              const headers = firstSheet ? firstSheet.headers : [];
              const rowCount = firstSheet ? firstSheet.row_count : 0;

              // fetch full uploaded content
              const fileReq = await fetch(`https://excel-ai-backend-j0zh.onrender.com/uploaded_data/${data.file_id}?sheet=0`);
              let full = null;
              if (fileReq.status === 200) {
                full = await fileReq.json();
              }

              const allRows = (full && full.rows_no_header) ? full.rows_no_header : [];

              setPendingUpload({
                fileId: data.file_id,
                fileName: data.original_name,
                headers,
                rows: allRows,
              });

              pushMessage({
                role: "ai",
                text: `✅ Uploaded ${data.original_name}. Detected ${parsed.type.toUpperCase()} with ${parsed.sheets.length} sheet(s). First sheet has ${rowCount} rows and headers: ${headers.join(", ")}`,
              });
            } catch (err) {
              pushMessage({ role: "ai", text: `❌ Upload error: ${String(err)}` });
            } finally {
              // reset input value so same file can be selected next time if needed
              try {
                (e.target as HTMLInputElement).value = "";
              } catch { }
            }
          }}
        />

        {/* Chat input */}
        <input
          value={input}
          onChange={(e) => setInput(e.target.value)}
          placeholder={
            pendingPlacement?.awaitingCell
              ? "Type cell address (e.g., C5) and press Enter"
              : pendingPlacement?.awaitingOverwrite
                ? `Type 'overwrite' to confirm overwrite ${pendingPlacement.awaitingOverwrite.addr}`
                : pendingPlacement
                  ? "Finish placement first (choose action above)"
                  : "Ask your Excel AI..."
          }
          onKeyDown={(e) => {
            if (e.key === "Enter") handleSend();
          }}
          // allow typing when awaiting a cell or awaiting overwrite; otherwise disable send while a pending non-awaiting placement exists
          disabled={!!pendingPlacement && !pendingPlacement.awaitingCell && !pendingPlacement.awaitingOverwrite}
          style={{ flex: 1, padding: 10, borderRadius: 20, color: "#ffffff", outline: "none", background: "rgb(22, 39, 58)", border: "none", fontSize: 14 }}
        />
        <button
          data-tooltip="Submit"
          data-tooltip-pos="right"
          onClick={() =>
            pendingPlacement && !pendingPlacement.awaitingCell && !pendingPlacement.awaitingOverwrite
              ? pushMessage({ role: "ai", text: "⚠️ Finish placement before sending a new command." })
              : handleSend()
          }
          style={{ padding: "6px 7px 0px 7px", borderRadius: 20, position: "relative", background: "#ffffffff", color: "#000000ff", border: "none", cursor: "pointer", fontSize: 18 }}
        >
          <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="24" height="24">
            <path fill="none" d="M0 0h24v24H0z"></path>
            <path fill="currentColor" d="M1.946 9.315c-.522-.174-.527-.455.01-.634l19.087-6.362c.529-.176.832.12.684.638l-5.454 19.086c-.15.529-.455.547-.679.045L12 14l6-8-8 6-8.054-2.685z"></path>
          </svg>
        </button>
      </div>
      <div style={{ color: "#fff", fontSize: "smaller"}}>
        <p style={{ margin: 2, display: "flex", justifyContent: "center"}}>
          Click the 
          <p onClick={() => {
            window.open(
              `${window.location.origin}/assets/Excel_AI_Assistant_Examples.pdf`,
              "_blank"
            );
          }}
          style={{margin: 0, padding: "0px 3px", cursor: "pointer", textDecoration: "underline"}}>
              Example 
          </p>
           button to see the manual
        </p>
      </div>
    </div>
  );
}
