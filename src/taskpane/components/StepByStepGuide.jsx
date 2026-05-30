/**
 * StepByStepGuide Component - AI Step-by-Step Guide
 *
 * ENHANCED:
 * - Excel context support for personalized guidance
 * - Visual cues (emoji icons)
 * - Cell highlighting for interactive mode
 */

import * as React from "react";
import { useState } from "react";
import { Button, Card, Field, Textarea, Spinner, Text, Switch } from "@fluentui/react-components";
import {
  Sparkle24Regular,
  Lightbulb24Regular,
  Warning24Regular,
  CheckmarkCircle24Regular,
  ChevronRight24Regular,
  ChevronLeft24Regular,
  TargetEdit24Regular,
} from "@fluentui/react-icons";

// API Service
import { generateStepByStep, cancelAIRequest, getExcelContext } from "../../services/apiService";

// Model Selector
import ModelSelector from "./ModelSelector";

const StepByStepGuide = ({ disabled = false, onRequestComplete }) => {
  const [task, setTask] = useState("");
  const [taskName, setTaskName] = useState("");
  const [steps, setSteps] = useState([]);
  const [currentStep, setCurrentStep] = useState(0);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState("");
  const [currentAbortController, setCurrentAbortController] = useState(null);
  const [selectedModel, setSelectedModel] = useState(null);
  const [useContext, setUseContext] = useState(true); // Default ON
  const [difficulty, setDifficulty] = useState(null);
  const [estimatedTime, setEstimatedTime] = useState(null);

  const exampleTasks = [
    "Tạo biểu đồ cột từ dữ liệu",
    "Sử dụng VLOOKUP",
    "Tạo Pivot Table",
    "Conditional Formatting",
  ];

  /**
   * Select cell in Excel using Office.js
   */
  const highlightCell = async (address) => {
    if (!address) return;
    try {
      await Excel.run(async (context) => {
        const range = context.workbook.worksheets.getActiveWorksheet().getRange(address);
        range.select();
        await context.sync();
      });
    } catch (err) {
      console.log("Could not select cell:", err);
    }
  };

  /**
   * Generate step-by-step guide - gọi Backend API
   */
  const handleGenerate = async () => {
    if (!task.trim()) return;

    if (disabled) {
      setError("Bạn đã hết lượt sử dụng!");
      return;
    }

    setIsLoading(true);
    setError("");
    setSteps([]);
    setCurrentStep(0);
    setDifficulty(null);
    setEstimatedTime(null);

    try {
      // Get Excel context if enabled
      let excelContext = null;
      if (useContext) {
        try {
          excelContext = await getExcelContext();
        } catch (e) {
          console.log("Could not get Excel context:", e);
        }
      }

      // Call API with context
      const result = await generateStepByStep(task, excelContext, selectedModel);
      setTaskName(result.taskName);
      setSteps(result.steps);
      setDifficulty(result.difficulty);
      setEstimatedTime(result.estimatedTime);

      // Notify parent to refresh credits
      if (onRequestComplete) {
        onRequestComplete();
      }
    } catch (err) {
      if (err.name === "AbortError") {
        setError("Đã hủy hướng dẫn");
      } else if (err.message?.includes("Failed to fetch") || err.message?.includes("NetworkError")) {
        setError("Lỗi kết nối mạng! Kiểm tra kết nối internet và thử lại.");
      } else if (err.message?.includes("timeout") || err.message?.includes("Timeout")) {
        setError("Yêu cầu quá thời gian. Vui lòng thử lại với mô tả ngắn gọn hơn.");
      } else {
        setError(err.message || "Đã xảy ra lỗi không xác định!");
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
    setError("Đã hủy yêu cầu");
  };

  const handleExampleClick = (exampleText) => {
    setTask(exampleText);
  };

  const handleNext = () => {
    if (currentStep < steps.length - 1) {
      setCurrentStep(currentStep + 1);
    } else {
      setCurrentStep(steps.length); // Completion state
    }
  };

  const handlePrev = () => {
    if (currentStep > 0) {
      setCurrentStep(currentStep - 1);
    }
  };

  const handleReset = () => {
    setCurrentStep(0);
  };

  const progress = steps.length > 0 ? ((currentStep + 1) / steps.length) * 100 : 0;

  return (
    <div className="page-container">
      <div className="page-header">
        <h2 className="page-title">
          <Lightbulb24Regular /> Hướng Dẫn Step by Step
        </h2>
        <p className="page-subtitle">
          Mô tả task bạn muốn thực hiện, AI sẽ hướng dẫn từng bước chi tiết
        </p>
      </div>

      <Card className="card">
        <Field label="Mô tả task của bạn">
          <Textarea
            placeholder="VD: Tôi muốn tạo một biểu đồ cột để hiển thị doanh thu theo tháng..."
            rows={4}
            value={task}
            onChange={(e) => setTask(e.target.value)}
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
              icon={<Sparkle24Regular />}
              onClick={handleGenerate}
              disabled={!task.trim()}
              style={{
                flex: 1,
                background: "#10b981",
                border: "none",
                color: "white",
              }}
            >
              Tạo hướng dẫn
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
            {exampleTasks.map((ex, idx) => (
              <div key={idx} className="chip" onClick={() => handleExampleClick(ex)}>
                {ex}
              </div>
            ))}
          </div>
        </div>
      </Card>

      {/* Task Metadata */}
      {taskName && (
        <Card className="card" style={{ background: "#f0fdf4", borderLeft: "4px solid #10b981" }}>
          <Text weight="semibold" size={400}>
            {taskName}
          </Text>
        </Card>
      )}

      {/* Error */}
      {error && <div className="alert alert--error">{error}</div>}

      {/* Steps Display */}
      {steps.length > 0 && currentStep < steps.length && (
        <>
          {/* Progress Bar */}
          <div className="progress-container mb-16">
            <div className="progress-header">
              <Text weight="semibold">
                Bước {currentStep + 1} / {steps.length}
              </Text>
              <Text>{Math.round(progress)}%</Text>
            </div>
            <div className="progress-bar">
              <div className="progress-fill" style={{ width: `${progress}%` }} />
            </div>
          </div>

          {/* Current Step */}
          <Card className="stepper-card">
            <div className="step-header">
              <div className="step-number">{currentStep + 1}</div>
              <div className="step-header__content">
                <h3 className="step-title">{steps[currentStep].title}</h3>
                <p className="step-description">{steps[currentStep].description}</p>
              </div>
            </div>

            {/* Details */}
            <div className="mb-16">
              <Text weight="semibold" className="d-block mb-12">
                Chi tiết thực hiện:
              </Text>
              <ul className="details-list">
                {steps[currentStep].details.map((detail, idx) => (
                  <li key={idx} className="detail-item">
                    <span className="detail-bullet">▸</span>
                    <Text size={300}>{detail}</Text>
                  </li>
                ))}
              </ul>
            </div>

            {/* Highlight Cell Button */}
            {steps[currentStep].cellToHighlight && (
              <Button
                appearance="outline"
                icon={<TargetEdit24Regular />}
                onClick={() => highlightCell(steps[currentStep].cellToHighlight)}
                style={{
                  marginBottom: "16px",
                  borderColor: "#10b981",
                  color: "#10b981",
                }}
              >
                Highlight ô {steps[currentStep].cellToHighlight}
              </Button>
            )}

            {/* Tips */}
            {steps[currentStep].tips && (
              <div className="tips-box">
                <div className="tips-box__header">
                  <Lightbulb24Regular style={{ color: "#3b82f6" }} className="flex-shrink-0" />
                  <div>
                    <span className="tips-box__title">💡 Mẹo hữu ích:</span>
                    <Text size={300} className="tips-box__content">
                      {steps[currentStep].tips}
                    </Text>
                  </div>
                </div>
              </div>
            )}

            {/* Warning */}
            {steps[currentStep].warning && (
              <div className="warning-box">
                <div className="warning-box__header">
                  <Warning24Regular style={{ color: "#d97706" }} className="flex-shrink-0" />
                  <div>
                    <span className="warning-box__title">⚠️ Lưu ý:</span>
                    <Text size={300} className="warning-box__content">
                      {steps[currentStep].warning}
                    </Text>
                  </div>
                </div>
              </div>
            )}

            {/* Navigation */}
            <div className="step-navigation">
              <Button
                appearance="secondary"
                icon={<ChevronLeft24Regular />}
                onClick={handlePrev}
                disabled={currentStep === 0}
              >
                Quay lại
              </Button>
              <Button
                appearance="primary"
                icon={<ChevronRight24Regular />}
                iconPosition="after"
                onClick={handleNext}
                className="btn-primary"
              >
                {currentStep === steps.length - 1 ? "Hoàn thành" : "Bước tiếp theo"}
              </Button>
            </div>
          </Card>
        </>
      )}

      {/* Completion */}
      {steps.length > 0 && currentStep === steps.length && (
        <div
          className="completion-card"
          style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: "12px" }}
        >
          <CheckmarkCircle24Regular className="completion-card__icon" />
          <Text size={500} weight="semibold">
            Hoàn thành!
          </Text>
          <Button appearance="primary" onClick={handleReset} className="btn-primary">
            Xem lại từ đầu
          </Button>
        </div>
      )}

      {/* Empty State */}
      {!steps.length && !isLoading && !error && (
        <div className="empty-state">
          <Lightbulb24Regular className="empty-state__icon" />
          <Text size={400} className="d-block mb-8">
            Hướng dẫn chi tiết sẽ xuất hiện ở đây
          </Text>
        </div>
      )}
    </div>
  );
};

export default StepByStepGuide;
