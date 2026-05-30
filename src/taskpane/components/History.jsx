/**
 * History Component - AI History Viewer
 * Redesigned to match FormulaGenerator/DataAnalyzer UI patterns
 */

import * as React from "react";
import { useState, useEffect } from "react";
import {
  Button,
  Card,
  Text,
  Spinner,
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
} from "@fluentui/react-components";
import {
  History24Regular,
  Delete24Regular,
  Sparkle24Regular,
  DataUsage24Regular,
  Lightbulb24Regular,
  Code24Regular,
  Copy24Regular,
  Dismiss24Regular,
} from "@fluentui/react-icons";

// API Service
import { getAIHistory, deleteAIHistory } from "../../services/apiService";

const History = () => {
  const [selectedItem, setSelectedItem] = useState(null);
  const [histories, setHistories] = useState([]);
  const [loading, setLoading] = useState(false);
  const [filter, setFilter] = useState("all");
  const [page, setPage] = useState(1);
  const [totalPages, setTotalPages] = useState(1);
  const [error, setError] = useState("");

  useEffect(() => {
    fetchHistory();
  }, [filter, page]);

  const fetchHistory = async () => {
    setLoading(true);
    setError("");
    try {
      const typeFilter = filter !== "all" ? filter : null;
      const data = await getAIHistory(typeFilter, page, 10);
      setHistories(data.histories);
      setTotalPages(data.pagination.totalPages);
    } catch (error) {
      console.error("Error fetching history:", error);
      setError(error.message || "Không thể tải lịch sử");
    } finally {
      setLoading(false);
    }
  };

  const handleDelete = async (id) => {
    try {
      await deleteAIHistory(id);
      fetchHistory(); // Refresh list
    } catch (error) {
      console.error("Error deleting history:", error);
      setError("Không thể xóa lịch sử");
    }
  };

  const formatDate = (dateString) => {
    const date = new Date(dateString);
    return date.toLocaleString("vi-VN", {
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit",
    });
  };

  const getTypeIcon = (type) => {
    switch (type) {
      case "formula":
        return <Sparkle24Regular />;
      case "analysis":
        return <DataUsage24Regular />;
      case "guide":
        return <Lightbulb24Regular />;
      case "vba":
        return <Code24Regular />;
      default:
        return <History24Regular />;
    }
  };

  const getTypeLabel = (type) => {
    switch (type) {
      case "formula":
        return "Công thức";
      case "analysis":
        return "Phân tích";
      case "guide":
        return "Hướng dẫn";
      case "vba":
        return "VBA/Macro";
      default:
        return type;
    }
  };

  const getResultPreview = (item) => {
    if (!item.output || !item.output.result) return "Không có kết quả";

    const result = item.output.result;

    if (item.type === "formula") {
      return result.formula || "Không có công thức";
    }

    if (item.type === "analysis") {
      return result.summary || "Không có tóm tắt";
    }

    if (item.type === "guide") {
      return result.taskName || "Không có tiêu đề";
    }

    if (item.type === "vba") {
      return result.macroName ? `${result.macroName} - ${result.description}` : "VBA Code";
    }

    return "Xem chi tiết";
  };

  return (
    <div className="page-container">
      <div className="page-header">
        <h2 className="page-title">
          <History24Regular /> Lịch sử AI
        </h2>
        <p className="page-subtitle">Xem lại các yêu cầu AI đã thực hiện</p>
      </div>

      {/* Filter Buttons */}
      <Card className="card">
        <Text weight="semibold" size={300} className="d-block mb-12">
          Lọc theo loại:
        </Text>
        <div className="button-group">
          <Button
            appearance={filter === "all" ? "primary" : "secondary"}
            onClick={() => setFilter("all")}
            size="small"
          >
            Tất cả
          </Button>
          <Button
            appearance={filter === "formula" ? "primary" : "secondary"}
            onClick={() => setFilter("formula")}
            icon={<Sparkle24Regular />}
            size="small"
          >
            Công thức
          </Button>
          <Button
            appearance={filter === "analysis" ? "primary" : "secondary"}
            onClick={() => setFilter("analysis")}
            icon={<DataUsage24Regular />}
            size="small"
          >
            Phân tích
          </Button>
          <Button
            appearance={filter === "guide" ? "primary" : "secondary"}
            onClick={() => setFilter("guide")}
            icon={<Lightbulb24Regular />}
            size="small"
          >
            Hướng dẫn
          </Button>
          <Button
            appearance={filter === "vba" ? "primary" : "secondary"}
            onClick={() => setFilter("vba")}
            icon={<Code24Regular />}
            size="small"
          >
            VBA
          </Button>
        </div>
      </Card>

      {/* Error */}
      {error && (
        <div className="alert alert--error">
          <Text>{error}</Text>
        </div>
      )}

      {/* Loading */}
      {loading && (
        <Card className="loading-card">
          <Spinner size="large" />
          <Text className="loading-card__text">Đang tải lịch sử...</Text>
        </Card>
      )}

      {/* History List */}
      {!loading && histories.length === 0 && (
        <div className="empty-state">
          <History24Regular className="empty-state__icon" />
          <Text>Chưa có lịch sử {filter !== "all" ? getTypeLabel(filter).toLowerCase() : ""}</Text>
        </div>
      )}

      {!loading && histories.length > 0 && (
        <>
          {histories.map((item) => (
            <Card
              key={item._id}
              className="card"
              style={{ cursor: "pointer" }}
              onClick={() => setSelectedItem(item)}
            >
              <div style={{ display: "flex", alignItems: "flex-start", gap: "12px" }}>
                <div style={{ color: "#0078d4", flexShrink: 0 }}>{getTypeIcon(item.type)}</div>
                <div style={{ flex: 1 }}>
                  {/* Header */}
                  <div
                    style={{
                      display: "flex",
                      justifyContent: "space-between",
                      alignItems: "center",
                      marginBottom: "8px",
                    }}
                  >
                    <div>
                      <Text weight="semibold" size={300} style={{ color: "#0078d4" }}>
                        {getTypeLabel(item.type)}
                      </Text>
                      <Text size={200} style={{ color: "#6b7280", marginLeft: "8px" }}>
                        {formatDate(item.createdAt)}
                      </Text>
                    </div>
                    <Button
                      appearance="subtle"
                      icon={<Delete24Regular />}
                      onClick={(e) => {
                        e.stopPropagation(); // Prevent opening dialog
                        handleDelete(item._id);
                      }}
                      size="small"
                    />
                  </div>

                  {/* Prompt */}
                  <Text size={300} className="d-block mb-8" style={{ color: "#111827" }}>
                    <strong>Yêu cầu:</strong> {item.input.prompt}
                  </Text>

                  {/* Result Preview */}
                  <div
                    style={{
                      background: "#f3f4f6",
                      padding: "8px 12px",
                      borderRadius: "4px",
                      borderLeft: "3px solid #0078d4",
                    }}
                  >
                    <Text size={200} style={{ color: "#374151" }}>
                      {getResultPreview(item)}
                    </Text>
                  </div>
                </div>
              </div>
            </Card>
          ))}

          {/* Pagination */}
          {totalPages > 1 && (
            <Card className="card">
              <div
                style={{
                  display: "flex",
                  justifyContent: "center",
                  alignItems: "center",
                  gap: "12px",
                }}
              >
                <Button
                  appearance="secondary"
                  onClick={() => setPage(page - 1)}
                  disabled={page === 1}
                  size="small"
                >
                  Trước
                </Button>
                <Text size={300}>
                  Trang {page} / {totalPages}
                </Text>
                <Button
                  appearance="secondary"
                  onClick={() => setPage(page + 1)}
                  disabled={page === totalPages}
                  size="small"
                >
                  Sau
                </Button>
              </div>
            </Card>
          )}
        </>
      )}
      {/* Detail Dialog */}
      <Dialog open={!!selectedItem} onOpenChange={(e, data) => !data.open && setSelectedItem(null)}>
        <DialogSurface style={{ maxWidth: "600px", width: "90%" }}>
          <DialogBody>
            <DialogTitle
              action={
                <Button
                  appearance="subtle"
                  icon={<Dismiss24Regular />}
                  onClick={() => setSelectedItem(null)}
                />
              }
            >
              <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                {selectedItem && getTypeIcon(selectedItem.type)}
                {selectedItem && getTypeLabel(selectedItem.type)}
              </div>
            </DialogTitle>

            <DialogContent>
              {selectedItem && (
                <div style={{ marginTop: "16px" }}>
                  <Text weight="semibold" className="d-block mb-8">
                    Yêu cầu gốc:
                  </Text>
                  <div
                    className="mb-16"
                    style={{ padding: "12px", background: "#f3f4f6", borderRadius: "8px" }}
                  >
                    {selectedItem.input.prompt}
                  </div>

                  {selectedItem.type === "vba" && selectedItem.output?.result?.code && (
                    <>
                      <div className="d-flex justify-between align-center mb-8">
                        <Text weight="semibold">VBA Code:</Text>
                        <Button
                          appearance="outline"
                          icon={<Copy24Regular />}
                          size="small"
                          onClick={() => {
                            navigator.clipboard.writeText(selectedItem.output.result.code);
                          }}
                        >
                          Copy
                        </Button>
                      </div>
                      <div className="vba-code-box">
                        <pre className="vba-pre" style={{ maxHeight: "300px" }}>
                          {selectedItem.output.result.code}
                        </pre>
                      </div>
                    </>
                  )}

                  {selectedItem.type === "formula" && (
                    <>
                      <Text weight="semibold" className="d-block mb-8">
                        Công thức:
                      </Text>
                      <div className="formula-box">
                        {selectedItem.output?.result?.formula || "N/A"}
                      </div>
                      <Text className="d-block mt-12">
                        {selectedItem.output?.result?.explanation}
                      </Text>
                    </>
                  )}

                  {selectedItem.type === "analysis" && (
                    <>
                      <Text weight="semibold" className="d-block mb-8">
                        Kết quả phân tích:
                      </Text>
                      <Text>{selectedItem.output?.result?.summary || "N/A"}</Text>
                    </>
                  )}

                  {selectedItem.type === "guide" && (
                    <>
                      <Text weight="semibold" className="d-block mb-8">
                        {selectedItem.output?.result?.taskName}
                      </Text>
                      <ul>
                        {selectedItem.output?.result?.steps?.map((step, idx) => (
                          <li key={idx}>
                            <strong>{step.title}:</strong> {step.description}
                          </li>
                        ))}
                      </ul>
                    </>
                  )}
                </div>
              )}
            </DialogContent>

            <DialogActions>
              <Button appearance="secondary" onClick={() => setSelectedItem(null)}>
                Đóng
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
};

export default History;
