/**
 * ExcelTools - Quick utilities (NO AI needed)
 * 1. Clean Data: trim, remove duplicates, remove empty rows (preserves formatting)
 * 2. Quick Format: header, borders, number format, freeze
 */

import * as React from "react";
import { useState, useCallback } from "react";
import { Button, Card, Text, Spinner } from "@fluentui/react-components";
import {
  ArrowSync24Regular,
  CheckmarkCircle24Regular,
  ErrorCircle24Regular,
  Delete24Regular,
  PaintBrush24Regular,
} from "@fluentui/react-icons";

function colLetter(index) {
  let result = "";
  let idx = index;
  while (idx >= 0) {
    result = String.fromCharCode(65 + (idx % 26)) + result;
    idx = Math.floor(idx / 26) - 1;
  }
  return result;
}

/**
 * Safely extract the starting row number from a Range.address.
 * Address can be: "Sheet1!A1:D10", "'Báo cáo Q3'!C2:H15", "Data2024!B5:F20"
 * Must strip the sheet name first (may contain digits like Q3, 2024).
 */
function getStartRow(address) {
  const cellRef = address.includes("!") ? address.split("!").pop() : address;
  const m = cellRef.match(/\d+/);
  return m ? parseInt(m[0], 10) : 1;
}

const ExcelTools = () => {
  const [cleanLoading, setCleanLoading] = useState(false);
  const [cleanResult, setCleanResult] = useState(null);
  const [cleanError, setCleanError] = useState("");

  const [formatLoading, setFormatLoading] = useState(false);
  const [formatMsg, setFormatMsg] = useState({ type: "", text: "" });

  // ═══════════════════════════════════════════════════════
  // CLEAN DATA - preserves ALL formatting
  // Trim in-place, then delete empty/duplicate ROWS (shift up)
  // ═══════════════════════════════════════════════════════
  const handleCleanData = useCallback(async () => {
    setCleanLoading(true);
    setCleanResult(null);
    setCleanError("");

    try {
      const result = await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getUsedRange();
        range.load("values, rowCount, columnCount, columnIndex, address");
        await context.sync();

        if (range.rowCount < 2) throw new Error("Cần ít nhất 2 hàng dữ liệu");

        const values = range.values;
        const startRow = getStartRow(range.address);

        let trimCount = 0;

        // 1. Trim strings IN-PLACE (overwrite values, keeps formatting intact)
        const trimmed = values.map((row) =>
          row.map((cell) => {
            if (typeof cell === "string") {
              const t = cell.trim();
              if (t !== cell) trimCount++;
              return t;
            }
            return cell;
          })
        );

        range.values = trimmed;
        await context.sync();

        // 2. Find rows to DELETE (empty + duplicate) - skip header row
        const rowsToDelete = [];
        const seen = new Set();
        let emptyCount = 0;
        let dupCount = 0;

        for (let i = 1; i < trimmed.length; i++) {
          const row = trimmed[i];
          const isEmpty = row.every(
            (c) => c === null || c === undefined || c === ""
          );
          const key = JSON.stringify(row);
          const isDup = !isEmpty && seen.has(key);

          if (isEmpty) {
            emptyCount++;
            rowsToDelete.push(startRow + i);
          } else if (isDup) {
            dupCount++;
            rowsToDelete.push(startRow + i);
          }

          if (!isEmpty) seen.add(key);
        }

        // 3. Delete rows BOTTOM-TO-TOP (so indices don't shift)
        if (rowsToDelete.length > 0) {
          rowsToDelete.sort((a, b) => b - a); // descending
          for (const rowNum of rowsToDelete) {
            sheet.getRange(`${rowNum}:${rowNum}`).delete("Up");
          }
          await context.sync();
        }

        return {
          trimCount,
          emptyCount,
          dupCount,
          finalRows: trimmed.length - 1 - emptyCount - dupCount,
        };
      });

      setCleanResult(result);
    } catch (err) {
      setCleanError(err.message || "Lỗi dọn dữ liệu");
    } finally {
      setCleanLoading(false);
    }
  }, []);

  // ═══════════════════════════════════════════════════════
  // QUICK FORMAT - header, borders, auto-fit, numbers, freeze
  // ═══════════════════════════════════════════════════════
  const handleQuickFormat = useCallback(async () => {
    setFormatLoading(true);
    setFormatMsg({ type: "", text: "" });

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getUsedRange();
        range.load("values, rowCount, columnCount, columnIndex, address");
        await context.sync();

        if (range.rowCount < 1) throw new Error("Không có dữ liệu");

        const baseCol = range.columnIndex;
        const startRow = getStartRow(range.address);
        const endRow = startRow + range.rowCount - 1;
        const sL = colLetter(baseCol);
        const eL = colLetter(baseCol + range.columnCount - 1);

        // Header
        const headerRange = sheet.getRange(`${sL}${startRow}:${eL}${startRow}`);
        headerRange.format.font.bold = true;
        headerRange.format.fill.color = "#1e3a5f";
        headerRange.format.font.color = "#ffffff";
        headerRange.format.horizontalAlignment = "Center";
        headerRange.format.rowHeight = 28;

        // Borders
        const allRange = sheet.getRange(`${sL}${startRow}:${eL}${endRow}`);
        const borders = [
          "InsideHorizontal", "InsideVertical",
          "EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight",
        ];
        for (const b of borders) {
          try {
            allRange.format.borders.getItem(b).style = "Thin";
            allRange.format.borders.getItem(b).color = "#d1d5db";
          } catch (e) { /* skip */ }
        }

        // Auto-fit
        allRange.format.autofitColumns();

        // Number format for numeric columns (proper 2D array)
        const dataRows = range.values.slice(1);
        const dataRowCount = endRow - startRow; // rows of data (not header)

        for (let col = 0; col < range.columnCount; col++) {
          try {
            let allNum = true;
            let hasDecimals = false;
            let maxVal = 0;
            let nonEmpty = 0;

            for (const row of dataRows) {
              const val = row[col];
              if (val === null || val === undefined || val === "") continue;
              nonEmpty++;
              if (typeof val !== "number") { allNum = false; break; }
              if (val !== Math.floor(val)) hasDecimals = true;
              if (Math.abs(val) > maxVal) maxVal = Math.abs(val);
            }

            if (allNum && nonEmpty > 0 && maxVal >= 100) {
              const cL = colLetter(baseCol + col);
              const colRange = sheet.getRange(`${cL}${startRow + 1}:${cL}${endRow}`);
              const fmt = hasDecimals ? "#,##0.00" : "#,##0";
              // Build proper 2D array matching range size
              const fmtArr = [];
              for (let r = 0; r < dataRowCount; r++) {
                fmtArr.push([fmt]);
              }
              if (fmtArr.length > 0) {
                colRange.numberFormat = fmtArr;
              }
            }
          } catch (e) {
            // Skip column if format fails
          }
        }

        // Freeze header row (works even if data starts at row 5, etc.)
        try { sheet.freezePanes.unfreeze(); } catch (e) { /* no existing freeze */ }
        try {
          // freezeAt ensures the actual header row stays visible when scrolling
          sheet.freezePanes.freezeAt(headerRange);
        } catch (e) { /* ignore */ }

        await context.sync();
      });

      setFormatMsg({ type: "success", text: "Header, borders, số, auto-fit, freeze!" });
    } catch (err) {
      setFormatMsg({ type: "error", text: err.message || "Lỗi format" });
    } finally {
      setFormatLoading(false);
    }
  }, []);

  return (
    <div className="page-container">
      <div className="page-header">
        <h2 className="page-title">Công cụ nhanh</h2>
        <p className="page-subtitle">Dọn dữ liệu & format bảng</p>
      </div>

      {/* ═══════ CLEAN DATA ═══════ */}
      <Card className="card">
        <div className="d-flex align-items-center gap-8 mb-8">
          <Delete24Regular style={{ color: "#ef4444" }} />
          <Text weight="semibold" size={400}>Dọn dữ liệu</Text>
        </div>
        <Text size={200} style={{ color: "#6b7280", marginBottom: "12px" }} className="d-block">
          Trim khoảng trắng, xóa hàng trùng, xóa hàng trống (giữ nguyên định dạng)
        </Text>

        <Button
          appearance="primary"
          icon={cleanLoading ? <Spinner size="tiny" /> : <Delete24Regular />}
          onClick={handleCleanData}
          disabled={cleanLoading}
          style={{ width: "100%", background: "#ef4444", border: "none", color: "white" }}
        >
          {cleanLoading ? "Đang dọn..." : "Dọn dữ liệu"}
        </Button>

        {cleanResult && (
          <div className="action-feedback action-feedback--success" style={{ flexDirection: "column", alignItems: "flex-start" }}>
            <div className="d-flex align-items-center gap-8">
              <CheckmarkCircle24Regular />
              <Text size={200} weight="semibold">Dọn xong!</Text>
            </div>
            <div style={{ paddingLeft: "32px", marginTop: "4px" }}>
              {cleanResult.trimCount > 0 && <Text size={200} className="d-block">Trim {cleanResult.trimCount} ô</Text>}
              {cleanResult.dupCount > 0 && <Text size={200} className="d-block">Xóa {cleanResult.dupCount} hàng trùng</Text>}
              {cleanResult.emptyCount > 0 && <Text size={200} className="d-block">Xóa {cleanResult.emptyCount} hàng trống</Text>}
              {cleanResult.trimCount === 0 && cleanResult.dupCount === 0 && cleanResult.emptyCount === 0 && (
                <Text size={200}>Dữ liệu đã sạch!</Text>
              )}
              <Text size={200} style={{ color: "#6b7280" }}>Còn {cleanResult.finalRows} hàng</Text>
            </div>
          </div>
        )}
        {cleanError && (
          <div className="action-feedback action-feedback--error">
            <ErrorCircle24Regular />
            <Text size={200}>{cleanError}</Text>
          </div>
        )}
      </Card>

      {/* ═══════ QUICK FORMAT ═══════ */}
      <Card className="card">
        <div className="d-flex align-items-center gap-8 mb-8">
          <PaintBrush24Regular style={{ color: "#3b82f6" }} />
          <Text weight="semibold" size={400}>Format nhanh</Text>
        </div>
        <Text size={200} style={{ color: "#6b7280", marginBottom: "12px" }} className="d-block">
          Header đậm, borders, auto-fit, định dạng số, freeze header
        </Text>

        <Button
          appearance="primary"
          icon={formatLoading ? <Spinner size="tiny" /> : <PaintBrush24Regular />}
          onClick={handleQuickFormat}
          disabled={formatLoading}
          style={{ width: "100%", background: "#3b82f6", border: "none", color: "white" }}
        >
          {formatLoading ? "Đang format..." : "Format bảng đẹp"}
        </Button>

        {formatMsg.text && (
          <div className={`action-feedback action-feedback--${formatMsg.type}`}>
            {formatMsg.type === "success" ? <CheckmarkCircle24Regular /> : <ErrorCircle24Regular />}
            <Text size={200}>{formatMsg.text}</Text>
          </div>
        )}
      </Card>
    </div>
  );
};

export default ExcelTools;
