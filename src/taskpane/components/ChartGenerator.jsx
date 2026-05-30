/**
 * ChartGenerator Component - AI Chart Creation
 *
 * Flow:
 * 1. User mô tả biểu đồ muốn tạo
 * 2. AI phân tích dữ liệu Excel → gợi ý chart type, data range, title
 * 3. User nhấn "Chèn biểu đồ vào Excel" → Office.js tạo chart trực tiếp
 */

import * as React from "react";
import { useState, useCallback } from "react";
import {
  Button,
  Card,
  Field,
  Textarea,
  Spinner,
  Text,
  Switch,
  Badge,
} from "@fluentui/react-components";
import {
  ChartMultiple24Regular,
  CheckmarkCircle24Regular,
  ErrorCircle24Regular,
  Sparkle24Filled,
} from "@fluentui/react-icons";

// API Service
import {
  getExcelContext,
  createChartInExcel,
  analyzeExcelData,
} from "../../services/apiService";

import ModelSelector from "./ModelSelector";

const ChartGenerator = ({ disabled = false, onRequestComplete }) => {
  const [description, setDescription] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState("");
  const [useContext, setUseContext] = useState(true);
  const [selectedModel, setSelectedModel] = useState(null);

  // AI suggestion result
  const [suggestion, setSuggestion] = useState(null);

  // Chart insertion state
  const [chartInserting, setChartInserting] = useState(false);
  const [chartSuccess, setChartSuccess] = useState("");
  const [chartError, setChartError] = useState("");

  // Cached context
  const [cachedContext, setCachedContext] = useState(null);

  const examplePrompts = [
    "Tạo biểu đồ cột so sánh doanh thu các tháng",
    "Biểu đồ tròn thể hiện tỷ lệ chi phí",
    "Vẽ biểu đồ đường xu hướng lợi nhuận",
    "Tạo biểu đồ kết hợp doanh thu và lợi nhuận",
  ];

  /**
   * Gọi AI phân tích dữ liệu và gợi ý biểu đồ
   */
  const handleGenerate = useCallback(async () => {
    if (!description.trim()) return;
    if (disabled) {
      setError("Bạn đã hết lượt sử dụng!");
      return;
    }
    if (!navigator.onLine) {
      setError("Không có kết nối mạng!");
      return;
    }

    setIsLoading(true);
    setError("");
    setSuggestion(null);
    setChartSuccess("");
    setChartError("");

    try {
      let excelContext = null;
      if (useContext) {
        try {
          excelContext = await getExcelContext();
          setCachedContext(excelContext);
        } catch (e) {
          console.warn("Could not get Excel context:", e);
        }
      }

      if (!excelContext || !excelContext.sampleData || excelContext.sampleData.length === 0) {
        throw new Error("Không có dữ liệu trong Excel. Vui lòng nhập dữ liệu trước.");
      }

      // Gọi AI analysis - nó sẽ trả về chartSuggestion
      const result = await analyzeExcelData(excelContext, selectedModel);

      if (result && result.chartSuggestion && result.chartSuggestion.type) {
        // Ghi đè title nếu user mô tả cụ thể
        const chartSugg = { ...result.chartSuggestion };

        // Nếu user mô tả loại chart cụ thể, thử detect từ prompt
        const desc = description.toLowerCase();
        if (desc.includes("tròn") || desc.includes("pie")) {
          chartSugg.type = "pie";
        } else if (desc.includes("đường") || desc.includes("line") || desc.includes("xu hướng")) {
          chartSugg.type = "line";
        } else if (desc.includes("cột") || desc.includes("column") || desc.includes("bar")) {
          chartSugg.type = "column";
        } else if (desc.includes("thanh ngang") || desc.includes("bar")) {
          chartSugg.type = "bar";
        } else if (desc.includes("vùng") || desc.includes("area")) {
          chartSugg.type = "area";
        } else if (desc.includes("scatter") || desc.includes("phân tán")) {
          chartSugg.type = "scatter";
        } else if (desc.includes("donut") || desc.includes("vành")) {
          chartSugg.type = "doughnut";
        }

        // Dùng title từ user nếu có mô tả ý nghĩa
        if (description.trim().length > 10) {
          chartSugg.title = chartSugg.title || description.trim().slice(0, 60);
        }

        setSuggestion(chartSugg);
      } else {
        setError("AI không đề xuất được biểu đồ phù hợp. Thử mô tả khác hoặc kiểm tra dữ liệu.");
      }

      if (onRequestComplete) onRequestComplete();
    } catch (err) {
      if (err.message?.includes("Failed to fetch") || err.message?.includes("NetworkError")) {
        setError("Lỗi kết nối mạng!");
      } else {
        setError(err.message || "Đã xảy ra lỗi!");
      }
    } finally {
      setIsLoading(false);
    }
  }, [description, disabled, useContext, selectedModel, onRequestComplete]);

  /**
   * Chèn biểu đồ vào Excel
   */
  const handleInsertChart = useCallback(async () => {
    if (!suggestion) return;

    setChartInserting(true);
    setChartSuccess("");
    setChartError("");

    try {
      const result = await createChartInExcel(suggestion, cachedContext);
      setChartSuccess(`Đã tạo biểu đồ "${result.title}" (${result.chartType}) thành công!`);
    } catch (err) {
      setChartError(err.message || "Không thể tạo biểu đồ");
    } finally {
      setChartInserting(false);
    }
  }, [suggestion, cachedContext]);

  const handleCancel = () => {
    setIsLoading(false);
    setError("Đã hủy");
  };

  return (
    <div className="page-container">
      <div className="page-header">
        <h2 className="page-title">
          <ChartMultiple24Regular /> AI Chart Generator
        </h2>
        <p className="page-subtitle">Mô tả biểu đồ bạn muốn, AI sẽ tạo trực tiếp trong Excel</p>
      </div>

      <Card className="card">
        <Field label="Mô tả biểu đồ bạn muốn tạo">
          <Textarea
            placeholder="VD: Biểu đồ tròn thể hiện tỷ lệ chi phí theo danh mục..."
            rows={3}
            value={description}
            onChange={(e) => setDescription(e.target.value)}
          />
        </Field>

        {/* Example chips */}
        <div className="example-chips" style={{ marginTop: "12px" }}>
          {examplePrompts.map((ex, idx) => (
            <div key={idx} className="chip" onClick={() => setDescription(ex)}>
              {ex}
            </div>
          ))}
        </div>

        {/* Context Toggle */}
        <div style={{ display: "flex", alignItems: "center", gap: "8px", marginTop: "12px" }}>
          <Switch checked={useContext} onChange={(e, data) => setUseContext(data.checked)} />
          <Text size={200} style={{ color: "#6b7280" }}>
            Sử dụng dữ liệu Excel (Khuyên dùng)
          </Text>
        </div>

        {/* Generate Button */}
        <div style={{ display: "flex", gap: "8px", alignItems: "center", marginTop: "16px" }}>
          {!isLoading ? (
            <Button
              appearance="primary"
              icon={<Sparkle24Filled />}
              onClick={handleGenerate}
              disabled={!description.trim()}
              style={{ flex: 1, background: "#10b981", border: "none", color: "white" }}
            >
              Tạo biểu đồ
            </Button>
          ) : (
            <Button appearance="secondary" onClick={handleCancel} style={{ flex: 1 }}>
              <Spinner size="tiny" style={{ marginRight: "8px" }} />
              Đang phân tích... (Hủy)
            </Button>
          )}
          <ModelSelector onModelChange={setSelectedModel} />
        </div>
      </Card>

      {/* Error */}
      {error && (
        <Card className="card alert alert--error">
          <div className="d-flex align-items-center gap-8">
            <ErrorCircle24Regular style={{ color: "#991b1b", flexShrink: 0 }} />
            <Text style={{ color: "#991b1b" }}>{error}</Text>
          </div>
        </Card>
      )}

      {/* Chart Suggestion Result */}
      {suggestion && (
        <Card className="card card--chart">
          <div className="chart-section-header">
            <Text weight="semibold" size={400}>
              Đề xuất biểu đồ
            </Text>
            <ChartMultiple24Regular style={{ color: "#10b981" }} />
          </div>

          <div className="chart-suggestion-detail">
            <div className="chart-detail-row">
              <Text size={200} style={{ color: "#6b7280" }}>Loại biểu đồ:</Text>
              <Text size={200} weight="semibold">
                {suggestion.type?.charAt(0).toUpperCase() + suggestion.type?.slice(1)}
              </Text>
            </div>
            {suggestion.dataRange && (
              <div className="chart-detail-row">
                <Text size={200} style={{ color: "#6b7280" }}>Vùng dữ liệu:</Text>
                <Text size={200} weight="semibold">{suggestion.dataRange}</Text>
              </div>
            )}
            <div className="chart-detail-row">
              <Text size={200} style={{ color: "#6b7280" }}>Tiêu đề:</Text>
              <Text size={200} weight="semibold">{suggestion.title}</Text>
            </div>
          </div>

          {suggestion.description && (
            <Text size={200} style={{ color: "#065f46", lineHeight: "1.5", marginTop: "8px" }}>
              {suggestion.description}
            </Text>
          )}

          {/* Insert button */}
          <Button
            appearance="primary"
            icon={chartInserting ? <Spinner size="tiny" /> : <ChartMultiple24Regular />}
            onClick={handleInsertChart}
            disabled={chartInserting}
            style={{
              marginTop: "12px",
              width: "100%",
              background: "#3b82f6",
              border: "none",
              color: "white",
            }}
          >
            {chartInserting ? "Đang tạo..." : "Chèn biểu đồ vào Excel"}
          </Button>

          {/* Success */}
          {chartSuccess && (
            <div className="action-feedback action-feedback--success">
              <CheckmarkCircle24Regular />
              <Text size={200}>{chartSuccess}</Text>
            </div>
          )}

          {/* Error */}
          {chartError && (
            <div className="action-feedback action-feedback--error">
              <ErrorCircle24Regular />
              <Text size={200}>{chartError}</Text>
            </div>
          )}
        </Card>
      )}

      {/* Empty State */}
      {!suggestion && !isLoading && !error && (
        <div className="empty-state">
          <ChartMultiple24Regular className="empty-state__icon" />
          <Text>Biểu đồ đề xuất sẽ xuất hiện ở đây</Text>
        </div>
      )}
    </div>
  );
};

export default ChartGenerator;
