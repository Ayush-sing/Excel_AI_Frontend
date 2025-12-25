// src/taskpane/components/App.tsx
import React, { useEffect, useRef, useState } from "react";

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
    const usedRange = sheet.getUsedRange();
    usedRange.load("values, address, rowCount, columnCount");
    await context.sync();
    return {
      values: usedRange.values,
      address: usedRange.address,
      rowCount: usedRange.rowCount,
      columnCount: usedRange.columnCount,
    };
  });
}

async function findPlacementCell(columnName: string) {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    usedRange.load("values, rowCount");
    await context.sync();

    const values = usedRange.values;
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
        const used = sheet.getUsedRange();
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

/* ---------------- React component ---------------- */

type Message = { role: "user" | "ai"; text: string; chart?: string | null };

export default function App(): JSX.Element {
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

  const [input, setInput] = useState("");
  const [pendingPlacement, setPendingPlacement] = useState<{
    note: string;
    chart?: string | null;
    suggestedColumn?: string | null;
    awaitingCell?: boolean; // when true, next user input is cell address
    awaitingOverwrite?: { sheet?: string | null; addr: string; chart?: boolean }; // when true, user must type 'overwrite'
  } | null>(null);

  const [loading, setLoading] = useState(false);
  const [loadingMsg, setLoadingMsg] = useState("Thinking...");
  const [darkMode, setDarkMode] = useState(false);
  const loadingTimerRef = useRef<number | null>(null);
  const chatContainerRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    document.body.style.background = darkMode ? "#0f1720" : "#f2f5f9";
  }, [darkMode]);

  useEffect(() => {
    // autoscroll to bottom when messages update
    if (chatContainerRef.current) {
      chatContainerRef.current.scrollTop = chatContainerRef.current.scrollHeight;
    }
  }, [messages]);

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
          const used = sheet.getUsedRange();
          used.load("rowCount");
          await context.sync();
          const row = (used.rowCount ?? 0) + 1;
          const putCell = `A${row + 1}`;
          const match = pendingPlacement.note.match(/[-+]?[0-9]*\.?[0-9]+/);
          const numericValue = match ? parseFloat(match[0]) : pendingPlacement.note;
          const range = sheet.getRange(putCell);
          range.values = [[numericValue]];
          range.format.font.bold = true;
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

      const res = await fetch("http://127.0.0.1:8000/chat", {
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

  // ----- UI -----
  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100vh", fontFamily: "Segoe UI, Roboto, system-ui" }}>
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
          background: darkMode ? "#0f1720" : "#f2f5f9",
        }}
      >
        <h2 style={{ margin: 0 }}>📊 Excel AI Assistant</h2>
        <div
          onClick={() => setDarkMode((d) => !d)}
          style={{ cursor: "pointer", padding: 6, borderRadius: 8, background: darkMode ? "#334155" : "#e2e8f0" }}
        >
          {darkMode ? "🌙 Dark" : "🌤️ Light"}
        </div>
      </div>

      {/* chat container */}
      <div ref={chatContainerRef} style={{ flex: 1, overflowY: "auto", padding: 12, background: darkMode ? "#06111a" : "#fff", minHeight: 0 }}>
        {messages.map((m, i) => (
          <div key={i} style={{ display: "flex", justifyContent: m.role === "user" ? "flex-end" : "flex-start", margin: "6px 0" }}>
            <div style={{ background: m.role === "user" ? "#DCF8C6" : "#E9EBEE", padding: 10, borderRadius: 12, maxWidth: "78%", whiteSpace: "pre-wrap" }}>
              <div>{m.text}</div>
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
          <div style={{ marginTop: 8 }}>
            <em>{loadingMsg}</em>
          </div>
        )}
      </div>

      {/* placement UI (when pending) */}
      {pendingPlacement && (
        <div style={{ padding: 12, borderRadius: 8, background: "#fff8e6", border: "1px solid #e6d8b6" }}>
          <div style={{ fontWeight: "bold", marginBottom: 8 }}>Where should I place the result?</div>

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
      <div style={{ position: "sticky", bottom: 0, padding: 10, display: "flex", gap: 8, background: darkMode ? "#0f1720" : "#f2f5f9" }}>
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
                  : "Ask your Excel AI... (one command at a time)"
          }
          onKeyDown={(e) => {
            if (e.key === "Enter") handleSend();
          }}
          // allow typing when awaiting a cell or awaiting overwrite; otherwise disable send while a pending non-awaiting placement exists
          disabled={!!pendingPlacement && !pendingPlacement.awaitingCell && !pendingPlacement.awaitingOverwrite}
          style={{ flex: 1, padding: 10, borderRadius: 20, border: "1px solid #ccc" }}
        />
        <button
          onClick={() =>
            pendingPlacement && !pendingPlacement.awaitingCell && !pendingPlacement.awaitingOverwrite
              ? pushMessage({ role: "ai", text: "⚠️ Finish placement first." })
              : handleSend()
          }
          style={{ padding: "10px 14px", borderRadius: 20, background: "#0078d7", color: "#fff", border: "none" }}
        >
          Send
        </button>
      </div>
    </div>
  );
}
