/**
 * DataAnalyzer - AI Analysis + Chart/PivotTable Builder
 *
 * Flow: AI phân tích → gợi ý chart/pivot → user chỉnh sửa → tạo
 * User có toàn quyền chọn: loại chart, cột nhãn, cột giá trị
 */

import * as React from "react";
import { useState, useCallback } from "react";
import { Button, Card, Spinner, Text, Badge } from "@fluentui/react-components";
import {
  DataUsage24Regular,
  ArrowUp24Regular,
  ArrowDown24Regular,
  Lightbulb24Regular,
  Warning24Regular,
  ChartMultiple24Regular,
  Sparkle24Filled,
  TableSimple24Regular,
  CheckmarkCircle24Regular,
  ErrorCircle24Regular,
  WifiOff24Regular,
  ArrowSync24Regular,
} from "@fluentui/react-icons";

import {
  analyzeExcelData,
  getExcelContext,
  createPivotTableInExcel,
} from "../../services/apiService";

import ModelSelector from "./ModelSelector";

// ── Constants ──
const CHART_TYPES = [
  { id: "column", label: "Cột", excel: "ColumnClustered" },
  { id: "bar", label: "Ngang", excel: "BarClustered" },
  { id: "line", label: "Đường", excel: "Line" },
  { id: "pie", label: "Tròn", excel: "Pie" },
  { id: "area", label: "Vùng", excel: "Area" },
];

function colLetter(index) {
  let result = "";
  let idx = index;
  while (idx >= 0) {
    result = String.fromCharCode(65 + (idx % 26)) + result;
    idx = Math.floor(idx / 26) - 1;
  }
  return result;
}

const DataAnalyzer = ({ disabled = false, onRequestComplete }) => {
  const [analysis, setAnalysis] = useState(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState("");
  const [selectedModel, setSelectedModel] = useState(null);
  const [isOffline, setIsOffline] = useState(!navigator.onLine);

  // Chart builder state
  const [chartColumns, setChartColumns] = useState([]);
  const [selectedChartType, setSelectedChartType] = useState("column");
  const [selectedLabelCol, setSelectedLabelCol] = useState(null);
  const [selectedValueCols, setSelectedValueCols] = useState([]);
  const [chartCreating, setChartCreating] = useState(false);
  const [chartSuccess, setChartSuccess] = useState("");
  const [chartError, setChartError] = useState("");

  // Pivot builder state
  const [pivotRowFields, setPivotRowFields] = useState([]);
  const [pivotValueFields, setPivotValueFields] = useState([]);
  const [pivotCreating, setPivotCreating] = useState(false);
  const [pivotSuccess, setPivotSuccess] = useState("");
  const [pivotError, setPivotError] = useState("");

  // Network
  React.useEffect(() => {
    const on = () => setIsOffline(false);
    const off = () => setIsOffline(true);
    window.addEventListener("online", on);
    window.addEventListener("offline", off);
    return () => {
      window.removeEventListener("online", on);
      window.removeEventListener("offline", off);
    };
  }, []);

  // ═══════════════════════════════════════════════════════
  // ANALYZE - call AI, then populate builders from response
  // ═══════════════════════════════════════════════════════
  const handleAnalyze = useCallback(async () => {
    if (disabled) { setError("Bạn đã hết lượt sử dụng!"); return; }
    if (!navigator.onLine) { setError("Không có kết nối mạng!"); return; }

    setIsLoading(true);
    setError("");
    setAnalysis(null);
    setChartSuccess("");
    setChartError("");
    setPivotSuccess("");
    setPivotError("");

    try {
      let excelContext;
      try {
        excelContext = await getExcelContext();
      } catch (e) {
        throw new Error("Không thể đọc dữ liệu từ Excel.");
      }

      if (!excelContext || !excelContext.sampleData || excelContext.sampleData.length === 0) {
        throw new Error("Không có dữ liệu để phân tích!");
      }
      if (excelContext.rowCount < 2) {
        throw new Error("Cần ít nhất 2 hàng dữ liệu.");
      }

      const result = await analyzeExcelData(excelContext, selectedModel);
      if (!result || typeof result !== "object") {
        throw new Error("Kết quả phân tích không hợp lệ.");
      }

      setAnalysis(result);

      // ── Populate Chart Builder from AI suggestion + actual data ──
      if (excelContext.columns) {
        const cols = excelContext.columns.map((c, i) => ({
          name: c.name,
          index: i,
          isNumeric: c.type === "number",
          isId: /^(stt|id|mã|ma|no\.?|#)$/i.test(c.name),
        }));
        setChartColumns(cols);

        const textCols = cols.filter((c) => !c.isNumeric && !c.isId);
        const numCols = cols.filter((c) => c.isNumeric && !c.isId);

        // Chart type from AI
        if (result.chartSuggestion?.type) {
          const aiType = result.chartSuggestion.type.toLowerCase();
          const match = CHART_TYPES.find((t) => t.id === aiType);
          setSelectedChartType(match ? match.id : "column");
        }

        // Columns: auto-detect
        setSelectedLabelCol(textCols.length > 0 ? textCols[0].index : null);
        setSelectedValueCols(numCols.map((c) => c.index));

        // Pivot fields from AI or auto
        if (result.pivotSuggestion?.rowFields?.length > 0) {
          setPivotRowFields(result.pivotSuggestion.rowFields);
        } else {
          setPivotRowFields(textCols.slice(0, 1).map((c) => c.name));
        }
        if (result.pivotSuggestion?.valueFields?.length > 0) {
          setPivotValueFields(result.pivotSuggestion.valueFields);
        } else {
          setPivotValueFields(numCols.slice(0, 3).map((c) => c.name));
        }
      }

      if (onRequestComplete) onRequestComplete();
    } catch (err) {
      if (err.message?.includes("Failed to fetch") || err.message?.includes("NetworkError")) {
        setError("Lỗi kết nối mạng!");
      } else if (err.message?.includes("timeout")) {
        setError("Phân tích quá lâu. Thử lại với ít dữ liệu hơn.");
      } else {
        setError(err.message || "Đã xảy ra lỗi!");
      }
    } finally {
      setIsLoading(false);
    }
  }, [disabled, selectedModel, onRequestComplete]);

  // ═══════════════════════════════════════════════════════
  // CREATE CHART - from user's selections (not raw AI)
  // ═══════════════════════════════════════════════════════
  const toggleValueCol = (colIndex) => {
    if (selectedValueCols.includes(colIndex)) {
      setSelectedValueCols(selectedValueCols.filter((i) => i !== colIndex));
    } else {
      if (selectedChartType === "pie") {
        setSelectedValueCols([colIndex]);
      } else {
        setSelectedValueCols([...selectedValueCols, colIndex]);
      }
    }
  };

  const handleCreateChart = useCallback(async () => {
    if (selectedValueCols.length === 0) {
      setChartError("Chọn ít nhất 1 cột giá trị!");
      return;
    }

    setChartCreating(true);
    setChartSuccess("");
    setChartError("");

    try {
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRange();
        usedRange.load("rowCount, columnCount, columnIndex, address");
        await context.sync();

        const totalRows = usedRange.rowCount;
        const baseCol = usedRange.columnIndex;
        // Safe: strip sheet name (may contain digits like "Data2024", "Q3")
        const cellRef = usedRange.address.includes("!") ? usedRange.address.split("!").pop() : usedRange.address;
        const startRow = parseInt((cellRef.match(/\d+/) || ["1"])[0], 10);

        // Build range from selected columns
        const allCols = [];
        if (selectedLabelCol !== null) allCols.push(selectedLabelCol);
        allCols.push(...selectedValueCols);
        allCols.sort((a, b) => a - b);

        const isContiguous = allCols.every(
          (c, i) => i === 0 || c === allCols[i - 1] + 1
        );

        let dataRange;
        if (isContiguous) {
          const sL = colLetter(baseCol + allCols[0]);
          const eL = colLetter(baseCol + allCols[allCols.length - 1]);
          dataRange = sheet.getRange(`${sL}${startRow}:${eL}${startRow + totalRows - 1}`);
        } else {
          dataRange = usedRange;
        }

        const typeObj = CHART_TYPES.find((t) => t.id === selectedChartType);
        const seriesBy = selectedLabelCol !== null ? "Columns" : "Auto";

        const chart = sheet.charts.add(typeObj?.excel || "ColumnClustered", dataRange, seriesBy);

        const valNames = chartColumns
          .filter((c) => selectedValueCols.includes(c.index))
          .map((c) => c.name)
          .join(", ");
        chart.title.text = analysis?.chartSuggestion?.title || `${typeObj?.label || "Biểu đồ"} - ${valNames}`;
        chart.title.format.font.size = 13;
        chart.title.format.font.bold = true;
        chart.height = 320;
        chart.width = 480;

        const posCol = colLetter(baseCol + usedRange.columnCount + 1);
        chart.setPosition(`${posCol}2`);

        try { chart.legend.position = "Bottom"; } catch (e) { /* */ }
        await context.sync();

        if (selectedChartType === "pie") {
          try {
            const series = chart.series.getItemAt(0);
            series.hasDataLabels = true;
            await context.sync();
            series.dataLabels.showPercentage = true;
            series.dataLabels.showCategoryName = true;
            series.dataLabels.showValue = false;
            await context.sync();
          } catch (e) { /* */ }
        }
      });

      setChartSuccess("Đã tạo biểu đồ!");
    } catch (err) {
      setChartError(err.message || "Lỗi tạo biểu đồ");
    } finally {
      setChartCreating(false);
    }
  }, [selectedChartType, selectedLabelCol, selectedValueCols, chartColumns, analysis]);

  // ═══════════════════════════════════════════════════════
  // CREATE PIVOT - from user's field selections
  // ═══════════════════════════════════════════════════════
  const togglePivotRow = (name) => {
    setPivotRowFields((prev) =>
      prev.includes(name) ? prev.filter((f) => f !== name) : [...prev, name]
    );
  };

  const togglePivotValue = (name) => {
    setPivotValueFields((prev) =>
      prev.includes(name) ? prev.filter((f) => f !== name) : [...prev, name]
    );
  };

  const handleCreatePivotTable = useCallback(async () => {
    if (pivotRowFields.length === 0 && pivotValueFields.length === 0) {
      setPivotError("Chọn ít nhất 1 trường!");
      return;
    }

    setPivotCreating(true);
    setPivotSuccess("");
    setPivotError("");

    try {
      const result = await createPivotTableInExcel({
        name: "PivotTable",
        rowFields: pivotRowFields,
        valueFields: pivotValueFields,
      });
      setPivotSuccess(`Đã tạo PivotTable "${result.name}" trong sheet mới!`);
    } catch (err) {
      setPivotError(err.message || "Không thể tạo PivotTable");
    } finally {
      setPivotCreating(false);
    }
  }, [pivotRowFields, pivotValueFields]);

  // ── Helpers ──
  const getTrendIcon = (type) => {
    if (type === "positive") return <ArrowUp24Regular style={{ color: "#10b981" }} />;
    if (type === "negative") return <ArrowDown24Regular style={{ color: "#ef4444" }} />;
    return <ArrowUp24Regular style={{ color: "#f59e0b" }} />;
  };

  const getTrendClass = (type) => {
    if (type === "positive") return "trend-item trend-positive";
    if (type === "negative") return "trend-item trend-negative";
    return "trend-item trend-neutral";
  };

  const textCols = chartColumns.filter((c) => !c.isNumeric && !c.isId);
  const numCols = chartColumns.filter((c) => c.isNumeric && !c.isId);
  const isPie = selectedChartType === "pie";

  return (
    <div className="page-container">
      <div className="page-header">
        <h2 className="page-title">
          <DataUsage24Regular style={{ color: "#3b82f6" }} /> BangAI Analyzer
        </h2>
        <p className="page-subtitle">BangAI phân tích → bạn chọn biểu đồ & PivotTable</p>
      </div>

      {/* Offline Warning */}
      {isOffline && (
        <Card className="card alert alert--warning">
          <div className="d-flex align-items-center gap-8">
            <WifiOff24Regular style={{ color: "#b45309" }} />
            <Text style={{ color: "#92400e" }}>Không có kết nối mạng.</Text>
          </div>
        </Card>
      )}

      {/* Analyze Button */}
      <Card className="card">
        <div style={{ display: "flex", gap: "8px", alignItems: "center" }}>
          {!isLoading ? (
            <Button
              appearance="primary"
              icon={<Sparkle24Filled />}
              onClick={handleAnalyze}
              disabled={isOffline}
              style={{ flex: 1, background: isOffline ? "#9ca3af" : "var(--primary-gradient)", border: "none", color: "white" }}
            >
              Phân tích dữ liệu
            </Button>
          ) : (
            <Button appearance="secondary" onClick={() => { setIsLoading(false); setError("Đã hủy"); }} style={{ flex: 1 }}>
              <Spinner size="tiny" style={{ marginRight: "8px" }} />
              Đang phân tích... (Hủy)
            </Button>
          )}
          <ModelSelector onModelChange={setSelectedModel} />
        </div>
        <Text size={200} className="d-block mt-12" style={{ color: "#6b7280" }}>
          AI phân tích dữ liệu, gợi ý biểu đồ & PivotTable — bạn tự chọn
        </Text>
      </Card>

      {/* Error */}
      {error && (
        <Card className="card alert alert--error">
          <div className="d-flex align-items-center gap-8">
            <ErrorCircle24Regular style={{ color: "#991b1b", flexShrink: 0 }} />
            <div>
              <Text style={{ color: "#991b1b" }}>{error}</Text>
              {error.includes("kết nối") && (
                <Button appearance="subtle" size="small" icon={<ArrowSync24Regular />} onClick={handleAnalyze} style={{ marginTop: "8px" }}>
                  Thử lại
                </Button>
              )}
            </div>
          </div>
        </Card>
      )}

      {/* Loading */}
      {isLoading && (
        <Card className="loading-card">
          <Spinner size="large" />
          <Text className="loading-card__text">AI đang phân tích...</Text>
        </Card>
      )}

      {/* ═══════ ANALYSIS RESULTS ═══════ */}
      {analysis && (
        <>
          {/* Summary */}
          <Card className="card card--primary">
            <Text weight="semibold" size={400} className="d-block mb-8" style={{ color: "#047857" }}>Tóm tắt</Text>
            <Text size={300} style={{ color: "#065f46", lineHeight: "1.6" }}>{analysis.summary}</Text>
          </Card>

          {/* Key Metrics */}
          {analysis.keyMetrics?.length > 0 && (
            <Card className="card">
              <Text weight="semibold" size={400} className="d-block mb-12">Chỉ số quan trọng</Text>
              <div className="metrics-grid">
                {analysis.keyMetrics.map((m, i) => (
                  <div key={i} className="metric-card">
                    <Text size={200} className="d-block" style={{ color: "#6b7280", marginBottom: "8px" }}>{m.label}</Text>
                    <Text size={400} weight="semibold" className="d-block" style={{ color: "#111827" }}>{m.value}</Text>
                  </div>
                ))}
              </div>
            </Card>
          )}

          {/* Trends */}
          {analysis.trends?.length > 0 && (
            <Card className="card">
              <Text weight="semibold" size={400} className="d-block mb-12">Xu hướng</Text>
              {analysis.trends.map((t, i) => (
                <div key={i} className={getTrendClass(t.type)}>
                  {getTrendIcon(t.type)}
                  <Text size={300} className="trend-content">{t.description}</Text>
                </div>
              ))}
            </Card>
          )}

          {/* Insights */}
          {analysis.insights?.length > 0 && (
            <Card className="card">
              <Text weight="semibold" size={400} className="d-block mb-12">Phát hiện</Text>
              <ul className="insights-list">
                {analysis.insights.map((ins, i) => (
                  <li key={i} className="insight-item">
                    <Lightbulb24Regular style={{ color: "#10b981" }} className="flex-shrink-0" />
                    <Text size={300} className="flex-1 line-height-16">{ins}</Text>
                  </li>
                ))}
              </ul>
            </Card>
          )}

          {/* Recommendations */}
          {analysis.recommendations?.length > 0 && (
            <Card className="card">
              <Text weight="semibold" size={400} className="d-block mb-12">Đề xuất</Text>
              <ul className="recommendations-list">
                {analysis.recommendations.map((rec, i) => (
                  <li key={i} className="recommendation-item">
                    <Text size={300} weight="semibold" style={{ color: "#047857" }} className="flex-shrink-0">{i + 1}.</Text>
                    <Text size={300} className="flex-1 line-height-16">{rec}</Text>
                  </li>
                ))}
              </ul>
            </Card>
          )}

          {/* Warnings */}
          {analysis.warnings?.length > 0 && (
            <Card className="card">
              <Text weight="semibold" size={400} className="d-block mb-12">Cảnh báo</Text>
              <ul className="warnings-list">
                {analysis.warnings.map((w, i) => (
                  <li key={i} className="warning-item">
                    <Warning24Regular style={{ color: "#f59e0b" }} className="flex-shrink-0" />
                    <Text size={300} className="flex-1 line-height-16">{w}</Text>
                  </li>
                ))}
              </ul>
            </Card>
          )}

          {/* ═══════ CHART BUILDER ═══════ */}
          {chartColumns.length > 0 && (
            <Card className="card card--chart">
              <div className="chart-section-header">
                <ChartMultiple24Regular style={{ color: "#3b82f6" }} />
                <Text weight="semibold" size={400}>Tạo biểu đồ</Text>
                {analysis.chartSuggestion?.type && (
                  <Badge appearance="filled" color="informative" size="small">
                    AI: {analysis.chartSuggestion.type}
                  </Badge>
                )}
              </div>

              {analysis.chartSuggestion?.description && (
                <Text size={200} style={{ color: "#1e40af", lineHeight: "1.5", marginBottom: "12px" }}>
                  {analysis.chartSuggestion.description}
                </Text>
              )}

              {/* Chart Type */}
              <Text size={200} weight="semibold" className="d-block mb-8" style={{ color: "#374151" }}>Loại biểu đồ</Text>
              <div className="chip-row">
                {CHART_TYPES.map((t) => (
                  <button
                    key={t.id}
                    className={`chip ${selectedChartType === t.id ? "chip--green" : ""}`}
                    onClick={() => {
                      setSelectedChartType(t.id);
                      if (t.id === "pie" && selectedValueCols.length > 1) {
                        setSelectedValueCols([selectedValueCols[0]]);
                      }
                    }}
                  >
                    {t.label}
                  </button>
                ))}
              </div>

              {/* Label Column */}
              <Text size={200} weight="semibold" className="d-block mt-12 mb-8" style={{ color: "#374151" }}>Cột nhãn (trục X)</Text>
              <div className="chip-row">
                <button
                  className={`chip ${selectedLabelCol === null ? "chip--green" : ""}`}
                  onClick={() => setSelectedLabelCol(null)}
                >
                  Không
                </button>
                {textCols.map((c) => (
                  <button
                    key={c.index}
                    className={`chip ${selectedLabelCol === c.index ? "chip--green" : ""}`}
                    onClick={() => setSelectedLabelCol(c.index)}
                  >
                    {c.name}
                  </button>
                ))}
              </div>

              {/* Value Columns */}
              <Text size={200} weight="semibold" className="d-block mt-12 mb-8" style={{ color: "#374151" }}>
                Cột giá trị (trục Y)
                {isPie && <span style={{ color: "#9ca3af", fontWeight: 400 }}> — tối đa 1</span>}
              </Text>
              <div className="chip-row">
                {numCols.map((c) => (
                  <button
                    key={c.index}
                    className={`chip ${selectedValueCols.includes(c.index) ? "chip--blue" : ""}`}
                    onClick={() => toggleValueCol(c.index)}
                  >
                    {selectedValueCols.includes(c.index) ? "✓ " : ""}{c.name}
                  </button>
                ))}
              </div>

              <Button
                appearance="primary"
                icon={chartCreating ? <Spinner size="tiny" /> : <ChartMultiple24Regular />}
                onClick={handleCreateChart}
                disabled={chartCreating || selectedValueCols.length === 0}
                style={{ marginTop: "16px", width: "100%", background: "var(--primary-gradient)", border: "none", color: "white" }}
              >
                {chartCreating ? "Đang tạo..." : "Tạo biểu đồ"}
              </Button>

              {chartSuccess && (
                <div className="action-feedback action-feedback--success">
                  <CheckmarkCircle24Regular />
                  <Text size={200}>{chartSuccess}</Text>
                </div>
              )}
              {chartError && (
                <div className="action-feedback action-feedback--error">
                  <ErrorCircle24Regular />
                  <Text size={200}>{chartError}</Text>
                </div>
              )}
            </Card>
          )}

          {/* ═══════ PIVOT BUILDER ═══════ */}
          {chartColumns.length > 0 && (
            <Card className="card card--pivot">
              <div className="chart-section-header">
                <TableSimple24Regular style={{ color: "#3b82f6" }} />
                <Text weight="semibold" size={400}>PivotTable</Text>
                {analysis.pivotSuggestion?.recommended && (
                  <Badge appearance="filled" color="informative" size="small">AI đề xuất</Badge>
                )}
              </div>

              {analysis.pivotSuggestion?.description && (
                <Text size={200} style={{ color: "#1e40af", lineHeight: "1.5", marginBottom: "12px" }}>
                  {analysis.pivotSuggestion.description}
                </Text>
              )}

              {/* Row fields (group by) */}
              <Text size={200} weight="semibold" className="d-block mb-8" style={{ color: "#374151" }}>Hàng (nhóm theo)</Text>
              <div className="chip-row">
                {textCols.map((c) => (
                  <button
                    key={c.index}
                    className={`chip ${pivotRowFields.includes(c.name) ? "chip--green" : ""}`}
                    onClick={() => togglePivotRow(c.name)}
                  >
                    {pivotRowFields.includes(c.name) ? "✓ " : ""}{c.name}
                  </button>
                ))}
                {textCols.length === 0 && (
                  <Text size={200} style={{ color: "#9ca3af" }}>Không có cột text</Text>
                )}
              </div>

              {/* Value fields */}
              <Text size={200} weight="semibold" className="d-block mt-12 mb-8" style={{ color: "#374151" }}>Giá trị (tính toán)</Text>
              <div className="chip-row">
                {numCols.map((c) => (
                  <button
                    key={c.index}
                    className={`chip ${pivotValueFields.includes(c.name) ? "chip--blue" : ""}`}
                    onClick={() => togglePivotValue(c.name)}
                  >
                    {pivotValueFields.includes(c.name) ? "✓ " : ""}{c.name}
                  </button>
                ))}
              </div>

              <Button
                appearance="primary"
                icon={pivotCreating ? <Spinner size="tiny" /> : <TableSimple24Regular />}
                onClick={handleCreatePivotTable}
                disabled={pivotCreating || (pivotRowFields.length === 0 && pivotValueFields.length === 0)}
                style={{ marginTop: "16px", width: "100%", background: "var(--primary-gradient)", border: "none", color: "white" }}
              >
                {pivotCreating ? "Đang tạo..." : "Tạo PivotTable"}
              </Button>

              {pivotSuccess && (
                <div className="action-feedback action-feedback--success">
                  <CheckmarkCircle24Regular />
                  <Text size={200}>{pivotSuccess}</Text>
                </div>
              )}
              {pivotError && (
                <div className="action-feedback action-feedback--error">
                  <ErrorCircle24Regular />
                  <Text size={200}>{pivotError}</Text>
                </div>
              )}
            </Card>
          )}
        </>
      )}
    </div>
  );
};

export default DataAnalyzer;
