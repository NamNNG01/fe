/**
 * VBAGenerator Component - AI VBA/Macro Code Generation
 *
 * Tính năng:
 * - Mô tả bằng tiếng Việt → AI viết code VBA
 * - Copy code to clipboard
 * - Hiển thị hướng dẫn sử dụng
 */

import * as React from "react";
import { useState } from "react";
import { Button, Card, Field, Textarea, Spinner, Text, Switch } from "@fluentui/react-components";
import {
  Code24Regular,
  Copy24Regular,
  Checkmark24Regular,
  Eye24Regular,
  Warning24Regular,
} from "@fluentui/react-icons";

// API Service
import { generateVBACode, getExcelContext } from "../../services/apiService";

import ModelSelector from "./ModelSelector";

const VBAGenerator = ({ disabled = false, onRequestComplete }) => {
  const [description, setDescription] = useState("");
  const [result, setResult] = useState(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState("");
  const [copied, setCopied] = useState(false);
  const [useContext, setUseContext] = useState(true);
  const [contextInfo, setContextInfo] = useState(null);
  const [selectedModel, setSelectedModel] = useState(null);

  const examplePrompts = [
    "Tô màu các hàng chẵn màu xanh nhạt",
    "Xóa các hàng trống trong sheet",
    "Tự động format số tiền thành tiền VND",
    "Sắp xếp dữ liệu theo cột A từ A-Z",
    "Copy dữ liệu từ sheet này sang sheet mới",
  ];

  /**
   * Generate VBA code - gọi Backend API
   */
  const handleGenerate = async () => {
    if (!description.trim()) return;

    if (disabled) {
      setError("Bạn đã hết lượt sử dụng!");
      return;
    }

    setIsLoading(true);
    setError("");
    setResult(null);
    setContextInfo(null);

    try {
      // Get Excel context nếu enabled
      let excelContext = null;
      if (useContext) {
        try {
          excelContext = await getExcelContext();
          setContextInfo(excelContext);
          console.log("📊 Excel context:", excelContext);
        } catch (ctxErr) {
          console.warn("⚠️ Could not get Excel context:", ctxErr);
        }
      }

      // Gọi API
      const vbaResult = await generateVBACode(description, excelContext, selectedModel);
      setResult(vbaResult);

      // Notify parent to refresh credits
      if (onRequestComplete) {
        onRequestComplete();
      }
    } catch (err) {
      if (err.name === "AbortError") {
        setError("Đã hủy yêu cầu");
      } else if (err.message?.includes("Failed to fetch") || err.message?.includes("NetworkError")) {
        setError("Lỗi kết nối mạng! Kiểm tra kết nối internet và thử lại.");
      } else if (err.message?.includes("timeout") || err.message?.includes("Timeout")) {
        setError("Yêu cầu quá thời gian. Vui lòng thử lại với mô tả ngắn gọn hơn.");
      } else {
        setError(err.message || "Đã xảy ra lỗi không xác định!");
      }
    } finally {
      setIsLoading(false);
    }
  };

  /**
   * Cancel pending request
   */
  const handleCancel = () => {
    setIsLoading(false);
    setError("Đã hủy yêu cầu");
  };

  /**
   * Copy VBA code to clipboard
   */
  const handleCopy = () => {
    if (result?.code) {
      navigator.clipboard.writeText(result.code);
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    }
  };

  const handleExampleClick = (exampleText) => {
    setDescription(exampleText);
  };

  return (
    <div className="page-container">
      <div className="page-header">
        <h2 className="page-title">
          <Code24Regular /> VBA/Macro Generator
        </h2>
        <p className="page-subtitle">Mô tả bằng tiếng Việt, AI sẽ viết code VBA cho bạn</p>
      </div>

      <Card className="card">
        <Field label="Mô tả macro bạn muốn tạo">
          <Textarea
            placeholder="VD: Tô màu các hàng có giá trị âm màu đỏ, tự động format số tiền..."
            rows={4}
            value={description}
            onChange={(e) => setDescription(e.target.value)}
          />
        </Field>

        {/* Context Toggle */}
        <div style={{ display: "flex", alignItems: "center", gap: "8px", marginTop: "12px" }}>
          <Switch checked={useContext} onChange={(e, data) => setUseContext(data.checked)} />
          <Text size={200} style={{ color: "#6b7280" }}>
            Sử dụng ngữ cảnh Excel (cột, dữ liệu mẫu)
          </Text>
        </div>

        {/* Button row with Model Selector on the RIGHT */}
        <div style={{ display: "flex", gap: "8px", alignItems: "center", marginTop: "16px" }}>
          {!isLoading ? (
            <Button
              appearance="primary"
              icon={<Code24Regular />}
              onClick={handleGenerate}
              disabled={!description.trim()}
              style={{
                flex: 1,
                background: "#10b981", // Green color
                border: "none",
                color: "white",
              }}
            >
              Tạo VBA Code
            </Button>
          ) : (
            <Button appearance="secondary" onClick={handleCancel} style={{ flex: 1 }}>
              <Spinner size="tiny" style={{ marginRight: "8px" }} />
              Đang tạo... (Hủy)
            </Button>
          )}

          {/* Model Selector - compact on the right */}
          <ModelSelector onModelChange={setSelectedModel} />
        </div>

        <div className="mt-16">
          <Text size={200} className="d-block mb-8">
            Ví dụ nhanh:
          </Text>
          <div className="example-chips">
            {examplePrompts.map((ex, idx) => (
              <div key={idx} className="chip" onClick={() => handleExampleClick(ex)}>
                {ex}
              </div>
            ))}
          </div>
        </div>
      </Card>

      {/* Error Message */}
      {error && <div className="alert alert--error">{error}</div>}

      {/* Display Context Info */}
      {contextInfo && (
        <Card className="card context-info-card">
          <Text weight="semibold" size={300} className="context-info-title">
            📊 AI đã phân tích Excel của bạn:
          </Text>
          <Text size={200} className="context-info-content">
            • Sheet: <strong>{contextInfo.sheetName}</strong>
            <br />• Dữ liệu: {contextInfo.rowCount} hàng × {contextInfo.columnCount} cột
          </Text>
        </Card>
      )}

      {/* VBA Result */}
      {result && (
        <div className="vba-result-container animate-fade-in">
          <Card className="card mb-16">
            <div className="d-flex justify-between align-center mb-12">
              <Text weight="semibold" size={500} style={{ color: "#107C10" }}>
                ✅ Đã tạo code xong!
              </Text>
              <Button
                appearance="primary"
                icon={copied ? <Checkmark24Regular /> : <Copy24Regular />}
                onClick={handleCopy}
                size="large"
                className={copied ? "btn-success" : ""}
              >
                {copied ? "Đã copy" : "Copy Code"}
              </Button>
            </div>

            <Text size={300} className="mb-12 d-block">
              {result.description || "Dưới đây là đoạn code VBA để thực hiện yêu cầu của bạn."}
            </Text>

            {/* VBA Code Box */}
            <div className="vba-code-box">
              <div className="vba-code-header">VBA Module</div>
              <pre className="vba-pre">{result.code}</pre>
            </div>
          </Card>

          {/* User Guide */}
          <Card className="card guide-card">
            <Text weight="bold" size={400} className="mb-12 d-block">
              🚀 Hướng dẫn chạy nhanh:
            </Text>

            <div className="guide-steps">
              <div className="guide-step">
                <div className="step-number">1</div>
                <div className="guide-step-content">
                  <Text weight="semibold">Mở cửa sổ code</Text>
                  <div className="shortcut-box">
                    <span className="key">Alt</span> + <span className="key">F11</span>
                  </div>
                </div>
              </div>

              <div className="guide-step">
                <div className="step-number">2</div>
                <div className="guide-step-content">
                  <Text weight="semibold">Tạo module mới</Text>
                  <Text size={200}>
                    Menu <b>Insert</b> → <b>Module</b>
                  </Text>
                </div>
              </div>

              <div className="guide-step">
                <div className="step-number">3</div>
                <div className="guide-step-content">
                  <Text weight="semibold">Dán & Chạy</Text>
                  <div className="shortcut-box">
                    <span className="key">Ctrl</span> + <span className="key">V</span>
                  </div>
                  <div className="shortcut-box mt-4">
                    <span className="key">F5</span> để chạy
                  </div>
                </div>
              </div>
            </div>

            {/* Custom AI Steps/Warnings */}
            {(result.howToUse || result.warnings) && (
              <div className="ai-notes mt-16">
                {result.howToUse && result.howToUse.length > 0 && (
                  <div className="mb-8">
                    <Text weight="semibold">💡 Lưu ý từ AI:</Text>
                    <ul className="msg-list">
                      {result.howToUse.map((step, idx) => (
                        <li key={idx}>{step}</li>
                      ))}
                    </ul>
                  </div>
                )}

                {result.warnings && result.warnings.length > 0 && (
                  <div>
                    <Text weight="semibold" style={{ color: "#d13438" }}>
                      ⚠️ Cảnh báo:
                    </Text>
                    <ul className="msg-list warning">
                      {result.warnings.map((warn, idx) => (
                        <li key={idx}>{warn}</li>
                      ))}
                    </ul>
                  </div>
                )}
              </div>
            )}
          </Card>
        </div>
      )}

      {/* Empty State */}
      {!result && !isLoading && !error && (
        <div className="empty-state">
          <Code24Regular className="empty-state__icon" />
          <Text>VBA code của bạn sẽ xuất hiện ở đây</Text>
        </div>
      )}
    </div>
  );
};

export default VBAGenerator;
