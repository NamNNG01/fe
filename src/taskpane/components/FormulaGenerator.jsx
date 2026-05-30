/**
 * FormulaGenerator Component - AI Formula Generation
 *
 * REFACTORED:
 * - Loại bỏ makeStyles, inline styles → CSS classes
 * - Sử dụng apiService: generateExcelFormula, insertFormulaToExcel
 * - Frontend CHỈ handle UI state + API calls
 * - Business logic (validation, AI processing) → Backend
 */

import * as React from "react";
import { useState } from "react";
import { Button, Card, Field, Textarea, Spinner, Text, Switch } from "@fluentui/react-components";
import {
  Sparkle24Regular,
  Copy24Regular,
  Checkmark24Regular,
  Send24Filled,
  Eye24Regular,
} from "@fluentui/react-icons";

// API Service
import {
  generateExcelFormula,
  getExcelContext,
  insertFormulaToExcel,
  cancelAIRequest,
} from "../../services/apiService";

import ModelSelector from "./ModelSelector";

const FormulaGenerator = ({ disabled = false, onRequestComplete }) => {
  const [prompt, setPrompt] = useState("");
  const [formula, setFormula] = useState("");
  const [explanation, setExplanation] = useState("");
  const [example, setExample] = useState("");
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState("");
  const [copied, setCopied] = useState(false);
  const [useContext, setUseContext] = useState(true);
  const [contextInfo, setContextInfo] = useState(null);
  const [insertSuccess, setInsertSuccess] = useState(false);
  const [currentAbortController, setCurrentAbortController] = useState(null);
  const [selectedModel, setSelectedModel] = useState(null);

  // Bang-Bot Emotion Logic
  const getBotEmotion = () => {
    if (isLoading) return "thinking";
    if (error) return "sad";
    if (insertSuccess || copied) return "happy";
    if (formula) return "excited";
    return "idle";
  };

  const BangBot = ({ emotion }) => {
    const isThinking = emotion === "thinking";
    const isHappy = emotion === "happy" || emotion === "excited";
    const isSad = emotion === "sad";

    return (
      <div className="bang-bot-container">
        <svg width="80" height="80" viewBox="0 0 100 100" className="bang-bot-svg">
          {/* Body */}
          <rect x="20" y="30" width="60" height="50" rx="15" fill="var(--primary-color)" />
          <rect x="25" y="35" width="50" height="40" rx="10" fill="white" opacity="0.2" />

          {/* Eyes Background */}
          <rect x="30" y="45" width="40" height="20" rx="10" fill="#1e293b" />

          {/* Eyes */}
          {isThinking ? (
            <g className="bot-eye-thinking">
              <path d="M40 55 A10 10 0 1 1 60 55" stroke="#38bdf8" strokeWidth="3" fill="none" />
            </g>
          ) : isSad ? (
            <g>
              <path d="M35 50 L45 60 M45 50 L35 60" stroke="#f43f5e" strokeWidth="3" />
              <path d="M55 50 L65 60 M65 50 L55 60" stroke="#f43f5e" strokeWidth="3" />
            </g>
          ) : isHappy ? (
            <g>
              <path d="M35 60 Q40 50 45 60" stroke="#4ade80" strokeWidth="3" fill="none" />
              <path d="M55 60 Q60 50 65 60" stroke="#4ade80" strokeWidth="3" fill="none" />
            </g>
          ) : (
            <g className="bot-eye-blink">
              <circle cx="40" cy="55" r="4" fill="#38bdf8" />
              <circle cx="60" cy="55" r="4" fill="#38bdf8" />
            </g>
          )}

          {/* Mouth */}
          <path
            d={isHappy ? "M40 70 Q50 80 60 70" : isSad ? "M40 75 Q50 65 60 75" : "M45 72 L55 72"}
            stroke={isHappy ? "#4ade80" : isSad ? "#f43f5e" : "white"}
            strokeWidth="3"
            fill="none"
            className="bot-mouth"
          />

          {/* Antena */}
          <line x1="50" y1="30" x2="50" y2="15" stroke="var(--primary-color)" strokeWidth="4" />
          <circle cx="50" cy="15" r="5" fill={isThinking ? "#f59e0b" : "var(--primary-light)"}>
            {isThinking && <animate attributeName="opacity" values="1;0;1" dur="0.5s" repeatCount="indefinite" />}
          </circle>
        </svg>
      </div>
    );
  };

  const examplePrompts = [
    "Tổng doanh thu nếu khu vực là 'Miền Bắc'",
    "Tìm giá trị trùng lặp giữa cột A và cột B",
    "Tách họ và tên từ chuỗi văn bản",
    "Trích xuất số từ ô chứa cả chữ và số",
  ];

  /**
   * Generate formula - gọi Backend API
   * TODO BACKEND: POST /api/formula/generate
   */
  const handleGenerate = async () => {
    if (!prompt.trim()) return;

    if (disabled) {
      setError("Hết hạn sử dụng. Vui lòng nâng cấp tài khoản!");
      return;
    }

    setIsLoading(true);
    setError("");
    setFormula("");
    setExplanation("");
    setExample("");
    setContextInfo(null);

    try {
      // Get Excel context nếu enabled
      let excelContext = null;
      if (useContext) {
        try {
          excelContext = await getExcelContext();
          setContextInfo(excelContext);
          console.log("📊 Context Synced:", excelContext);
        } catch (ctxErr) {
          console.warn("⚠️ Context Offline:", ctxErr);
          // Continue without context if it fails
        }
      }

      // Gọi API qua apiService (auto handles auth, base URL, etc.)
      const result = await generateExcelFormula(prompt, excelContext, selectedModel);

      // Xử lý trường hợp AI trả về formula rỗng (yêu cầu không rõ ràng)
      if (!result.formula || result.formula.trim() === "") {
        // Hiển thị explanation như một warning/info message
        setError(result.explanation || "Trí tuệ nhân tạo cần thêm thông tin. Hãy mô tả chi tiết hơn.");
        setFormula("");
        setExplanation("");
        setExample("");
      } else {
        setFormula(result.formula);
        setExplanation(result.explanation);
        setExample(result.example || "");
      }

      // Notify parent to refresh credits
      if (onRequestComplete) {
        onRequestComplete();
      }
    } catch (err) {
      if (err.name === "AbortError") {
        setError("Yêu cầu đã bị hủy.");
      } else if (err.message?.includes("Failed to fetch") || err.message?.includes("NetworkError")) {
        setError("Lỗi mạng! Kiểm tra kết nối và thử lại.");
      } else if (err.message?.includes("timeout") || err.message?.includes("Timeout")) {
        setError("Hệ thống quá tải. Vui lòng thử lại sau giây lát.");
      } else {
        setError(err.message || "Đã xảy ra lỗi ngoài ý muốn.");
      }
    } finally {
      setIsLoading(false);
      setCurrentAbortController(null);
    }
  };

  /**
   * Cancel pending request - KISS: just reset loading state
   */
  const handleCancel = () => {
    setIsLoading(false);
    setError("Đã dừng quá trình xử lý AI.");
  };

  /**
   * Copy formula to clipboard
   */
  const handleCopy = () => {
    navigator.clipboard.writeText(formula);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  /**
   * Insert formula vào Excel - Client-side Excel API
   */
  const handleInsertToExcel = async () => {
    if (!formula) return;

    try {
      await insertFormulaToExcel(formula);
      setError("");
      setInsertSuccess(true);
      setTimeout(() => setInsertSuccess(false), 3000);
    } catch (err) {
      setError("❌ Không thể chèn vào bảng tính: " + err.message);
    }
  };

  const handleExampleClick = (exampleText) => {
    setPrompt(exampleText);
  };

  return (
    <div className="page-container">
      <div className="page-header">
        <BangBot emotion={getBotEmotion()} />
        <h2 className="page-title logo-container-float">
          <Sparkle24Regular style={{ color: "var(--primary-color)", filter: "drop-shadow(0 0 8px var(--primary-light))" }} /> BangAI Formula
        </h2>
        <p className="page-subtitle">
          Khởi tạo sức mạnh Excel cùng BangAI
        </p>
      </div>

      <Card className="card glow-card">
        <Field label="Bạn muốn làm gì với bảng tính này?">
          <Textarea
            placeholder="Mô tả ý tưởng của bạn... AI sẽ lo phần còn lại."
            rows={4}
            value={prompt}
            onChange={(e) => setPrompt(e.target.value)}
            className="glass-panel"
          />
        </Field>

        {/* Context Toggle */}
        <div style={{ display: "flex", alignItems: "center", gap: "8px", marginTop: "16px" }}>
          <Switch checked={useContext} onChange={(e, data) => setUseContext(data.checked)} />
          <Text size={200} style={{ color: "var(--text-secondary)", fontWeight: "600" }}>
            Đọc ngữ cảnh thông minh (Smart Context)
          </Text>
        </div>

        {/* Button row with Model Selector on the RIGHT */}
        <div style={{ display: "flex", gap: "10px", alignItems: "center", marginTop: "24px" }}>
          {!isLoading ? (
            <Button
              appearance="primary"
              icon={<Sparkle24Regular />}
              onClick={handleGenerate}
              disabled={!prompt.trim()}
              className="btn-premium"
              style={{
                flex: 1,
                height: "48px",
                fontSize: "16px",
                fontWeight: "700",
                borderRadius: "14px"
              }}
            >
              Phân tích & Tạo
            </Button>
          ) : (
            <Button appearance="secondary" onClick={handleCancel} style={{ flex: 1, height: "48px", borderRadius: "14px" }}>
              <Spinner size="tiny" style={{ marginRight: "8px" }} />
              Đang suy nghĩ... (Dừng)
            </Button>
          )}

          {/* Model Selector - compact on the right */}
          <ModelSelector onModelChange={setSelectedModel} />
        </div>

        <div className="mt-16">
          <Text size={200} className="d-block mb-8" style={{ fontWeight: "700", color: "var(--primary-dark)", letterSpacing: "0.5px" }}>
            ⚡ Ý TƯỞNG NHANH:
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
      {error && <div className="alert alert--error" style={{ borderRadius: "14px", border: "none", boxShadow: "0 4px 12px rgba(239, 68, 68, 0.1)" }}>{error}</div>}

      {/* Success Message */}
      {insertSuccess && (
        <div className="alert alert--success" style={{ borderRadius: "14px", border: "none", boxShadow: "0 4px 12px rgba(16, 185, 129, 0.1)" }}>
          <Checkmark24Regular />
          <Text weight="semibold">✅ Tuyệt vời! Công thức đã được áp dụng.</Text>
        </div>
      )}

      {/* Display Context Info */}
      {contextInfo && (
        <Card className="card glass-panel" style={{ border: "none", marginTop: "20px", boxShadow: "var(--shadow-md)" }}>
          <Text weight="bold" size={300} style={{ color: "var(--primary-dark)", display: "flex", alignItems: "center", gap: "8px" }}>
            <Eye24Regular /> 🧠 PHÂN TÍCH NGỮ CẢNH:
          </Text>
          <Text size={200} className="context-info-content" style={{ marginTop: "10px", lineHeight: "1.6" }}>
            • Đang xử lý: <strong>{contextInfo.sheetName}</strong>
            <br />• Quy mô: <span style={{ color: "var(--primary-color)", fontWeight: "600" }}>{contextInfo.rowCount} hàng</span> | <span style={{ color: "var(--primary-color)", fontWeight: "600" }}>{contextInfo.columnCount} cột</span>
            {contextInfo.selectedCell && (
              <>
                <br />• Vị trí đích: <code style={{ background: "#eef2ff", padding: "2px 6px", borderRadius: "4px" }}>{contextInfo.selectedCell.address}</code>
              </>
            )}
          </Text>
        </Card>
      )}

      {/* Formula Result */}
      {formula && (
        <Card className="card glow-card" style={{ marginTop: "24px" }}>
          <Text weight="bold" size={400} className="d-block mb-12" style={{ color: "var(--primary-dark)", letterSpacing: "-0.02em" }}>
            GIẢI PHÁP TỪ AI:
          </Text>

          <div className="formula-box-premium">{formula}</div>

          <div className="button-group" style={{ marginTop: "20px" }}>
            <Button
              appearance="secondary"
              icon={copied ? <Checkmark24Regular /> : <Copy24Regular />}
              onClick={handleCopy}
              style={{ borderRadius: "12px", height: "40px" }}
            >
              {copied ? "Xong!" : "Sao chép"}
            </Button>
            <Button
              appearance="primary"
              icon={<Send24Filled />}
              onClick={handleInsertToExcel}
              className="btn-premium"
              style={{
                borderRadius: "12px",
                height: "40px",
                flex: 1
              }}
            >
              Áp dụng vào bảng
            </Button>
          </div>

          {/* Explanation */}
          {explanation && (
            <div className="explanation-box glass-panel" style={{ marginTop: "20px", borderRadius: "16px" }}>
              <span className="explanation-box__title" style={{ color: "var(--primary-dark)", fontWeight: "700" }}>💡 LOGIC XỬ LÝ:</span>
              <Text size={300} className="explanation-box__content" style={{ color: "#334155", fontStyle: "italic" }}>
                "{explanation}"
              </Text>
            </div>
          )}
        </Card>
      )}

      {/* Empty State */}
      {!formula && !isLoading && !error && (
        <div className="empty-state" style={{ marginTop: "24px", borderRadius: "16px", padding: "40px 20px" }}>
          <Sparkle24Regular className="empty-state__icon" style={{ fontSize: "40px", marginBottom: "16px" }} />
          <Text style={{ fontWeight: "500" }}>Sẵn sàng hỗ trợ ý tưởng của bạn</Text>
        </div>
      )}
    </div>
  );
};

export default FormulaGenerator;
