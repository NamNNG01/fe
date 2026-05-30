/**
 * TemplateGallery Component - Thư viện mẫu Excel & Hướng dẫn AI
 *
 * Cung cấp:
 * - Các mẫu Excel phổ biến với hướng dẫn sử dụng
 * - Tips sử dụng AI hiệu quả
 * - Quick actions: insert template data, apply formatting
 */

import * as React from "react";
import { useState, useCallback } from "react";
import {
  Button,
  Card,
  Text,
  Spinner,
  Input,
} from "@fluentui/react-components";
import {
  DocumentBulletList24Regular,
  Search24Regular,
  ArrowDownload24Regular,
  Lightbulb24Regular,
  BookOpen24Regular,
  Star24Filled,
  ChartMultiple24Regular,
  TableSimple24Regular,
  Calculator24Regular,
  CheckmarkCircle24Regular,
  ErrorCircle24Regular,
  Sparkle24Regular,
  ClipboardPaste24Regular,
  Info24Regular,
} from "@fluentui/react-icons";

// ============================================================================
// TEMPLATE DATA
// ============================================================================

const TEMPLATE_CATEGORIES = [
  { id: "all", label: "Tất cả", icon: <DocumentBulletList24Regular /> },
  { id: "finance", label: "Tài chính", icon: <Calculator24Regular /> },
  { id: "data", label: "Dữ liệu", icon: <ChartMultiple24Regular /> },
  { id: "project", label: "Dự án", icon: <TableSimple24Regular /> },
  { id: "tips", label: "AI Tips", icon: <Lightbulb24Regular /> },
];

const TEMPLATES = [
  // === FINANCE ===
  {
    id: "budget",
    category: "finance",
    title: "Quản lý Ngân sách",
    description: "Theo dõi thu chi, phân loại chi tiêu, tổng hợp theo tháng",
    tags: ["SUM", "SUMIF", "Conditional Formatting"],
    difficulty: "Dễ",
    headers: ["Ngày", "Mô tả", "Loại", "Thu nhập", "Chi tiêu", "Số dư"],
    sampleData: [
      ["01/01/2025", "Lương tháng 1", "Thu nhập", 15000000, "", ""],
      ["02/01/2025", "Tiền nhà", "Nhà ở", "", 5000000, ""],
      ["03/01/2025", "Ăn uống", "Sinh hoạt", "", 500000, ""],
      ["05/01/2025", "Xăng xe", "Đi lại", "", 300000, ""],
      ["10/01/2025", "Freelance", "Thu nhập", 3000000, "", ""],
    ],
    formulas: [
      { cell: "F2", formula: '=IF(D2<>"", D2, 0) - IF(E2<>"", E2, 0) + IF(ROW()>2, F1, 0)' },
    ],
    aiPrompt: "Tính tổng thu nhập, tổng chi tiêu và số dư cuối tháng",
  },
  {
    id: "invoice",
    category: "finance",
    title: "Hóa đơn Bán hàng",
    description: "Mẫu hóa đơn với tính năng tự động tổng, thuế VAT, chiết khấu",
    tags: ["VLOOKUP", "IF", "TEXT"],
    difficulty: "Trung bình",
    headers: ["STT", "Sản phẩm", "Đơn giá", "Số lượng", "Thành tiền", "Ghi chú"],
    sampleData: [
      [1, "Laptop Dell XPS 15", 25000000, 2, "", ""],
      [2, "Chuột Logitech MX", 1500000, 5, "", ""],
      [3, "Bàn phím cơ", 2000000, 3, "", ""],
      [4, "Màn hình 27 inch", 8000000, 2, "", ""],
    ],
    formulas: [{ cell: "E2", formula: "=C2*D2" }],
    aiPrompt: "Tính thành tiền cho mỗi sản phẩm, tổng cộng, VAT 10% và tổng thanh toán",
  },
  // === DATA ===
  {
    id: "sales-dashboard",
    category: "data",
    title: "Dashboard Doanh số",
    description: "Phân tích doanh số theo sản phẩm, vùng miền, nhân viên",
    tags: ["SUMIFS", "PivotTable", "Chart"],
    difficulty: "Nâng cao",
    headers: ["Ngày", "Nhân viên", "Vùng", "Sản phẩm", "Số lượng", "Doanh thu"],
    sampleData: [
      ["01/01/2025", "Nguyễn Văn A", "Miền Bắc", "Sản phẩm X", 10, 5000000],
      ["02/01/2025", "Trần Thị B", "Miền Nam", "Sản phẩm Y", 8, 4000000],
      ["03/01/2025", "Lê Văn C", "Miền Trung", "Sản phẩm X", 15, 7500000],
      ["05/01/2025", "Nguyễn Văn A", "Miền Bắc", "Sản phẩm Z", 5, 3000000],
      ["07/01/2025", "Phạm Thị D", "Miền Nam", "Sản phẩm Y", 12, 6000000],
    ],
    formulas: [],
    aiPrompt:
      "Phân tích doanh số theo vùng miền, tìm nhân viên có doanh thu cao nhất, tạo PivotTable và biểu đồ",
  },
  {
    id: "student-grades",
    category: "data",
    title: "Bảng điểm Học sinh",
    description: "Quản lý điểm, xếp loại tự động, thống kê lớp học",
    tags: ["AVERAGE", "IF", "RANK", "COUNTIF"],
    difficulty: "Dễ",
    headers: ["STT", "Họ tên", "Toán", "Văn", "Anh", "TB", "Xếp loại"],
    sampleData: [
      [1, "Nguyễn Văn An", 8, 7, 9, "", ""],
      [2, "Trần Thị Bình", 9, 8, 8, "", ""],
      [3, "Lê Minh Cường", 6, 7, 5, "", ""],
      [4, "Phạm Hồng Dương", 7, 9, 8, "", ""],
      [5, "Hoàng Mai Linh", 10, 8, 9, "", ""],
    ],
    formulas: [
      { cell: "F2", formula: "=AVERAGE(C2:E2)" },
      {
        cell: "G2",
        formula: '=IF(F2>=8,"Giỏi",IF(F2>=6.5,"Khá",IF(F2>=5,"Trung bình","Yếu")))',
      },
    ],
    aiPrompt: "Tính điểm trung bình, xếp loại, và thống kê số học sinh mỗi loại",
  },
  // === PROJECT ===
  {
    id: "project-tracker",
    category: "project",
    title: "Quản lý Dự án",
    description: "Theo dõi tiến độ task, người phụ trách, deadline, trạng thái",
    tags: ["Conditional Formatting", "DATEDIF", "TODAY"],
    difficulty: "Trung bình",
    headers: ["Task", "Người phụ trách", "Bắt đầu", "Deadline", "Trạng thái", "Tiến độ (%)"],
    sampleData: [
      ["Thiết kế UI", "Nguyễn A", "01/01/2025", "15/01/2025", "Hoàn thành", 100],
      ["Phát triển Backend", "Trần B", "10/01/2025", "30/01/2025", "Đang làm", 60],
      ["Testing", "Lê C", "25/01/2025", "10/02/2025", "Chưa bắt đầu", 0],
      ["Deploy", "Phạm D", "10/02/2025", "15/02/2025", "Chưa bắt đầu", 0],
    ],
    formulas: [],
    aiPrompt: "Tính số ngày còn lại cho mỗi task, highlight task sắp trễ deadline",
  },
  {
    id: "inventory",
    category: "project",
    title: "Quản lý Kho hàng",
    description: "Theo dõi tồn kho, nhập xuất, cảnh báo hết hàng",
    tags: ["SUMIFS", "IF", "Conditional Formatting"],
    difficulty: "Trung bình",
    headers: ["Mã SP", "Tên sản phẩm", "Tồn kho", "Nhập thêm", "Xuất", "Tồn cuối", "Trạng thái"],
    sampleData: [
      ["SP001", "Laptop Dell", 50, 20, 15, "", ""],
      ["SP002", "Chuột không dây", 200, 0, 80, "", ""],
      ["SP003", "USB 32GB", 5, 100, 3, "", ""],
      ["SP004", "Tai nghe Bluetooth", 0, 50, 10, "", ""],
    ],
    formulas: [
      { cell: "F2", formula: "=C2+D2-E2" },
      { cell: "G2", formula: '=IF(F2<=10,"Sắp hết",IF(F2<=0,"Hết hàng","Còn hàng"))' },
    ],
    aiPrompt: "Tính tồn kho cuối, cảnh báo sản phẩm sắp hết, tạo biểu đồ tồn kho",
  },
  // === AI TIPS ===
  {
    id: "tip-formula",
    category: "tips",
    title: "Cách viết prompt tạo công thức",
    description: "Hướng dẫn mô tả yêu cầu để AI tạo công thức chính xác nhất",
    tags: ["Best Practice", "Prompt Engineering"],
    difficulty: "Tips",
    isTip: true,
    tipContent: [
      {
        title: "Mô tả RÕ RÀNG mục đích",
        examples: [
          '❌ "Tính cột C"',
          '✅ "Tính tổng doanh thu cột C cho các dòng có trạng thái là Đã thanh toán"',
        ],
      },
      {
        title: "Đề cập tên cột/cell cụ thể",
        examples: [
          '❌ "Tìm giá trị lớn nhất"',
          '✅ "Tìm giá trị lớn nhất trong cột Doanh thu (D2:D100)"',
        ],
      },
      {
        title: "Bật ngữ cảnh Excel",
        examples: [
          "Bật switch 'Sử dụng ngữ cảnh Excel' để AI tự đọc cấu trúc bảng",
          "AI sẽ dùng đúng tên cột, range, và Named Tables",
        ],
      },
      {
        title: "Nêu điều kiện rõ ràng",
        examples: [
          '✅ "Đếm số đơn hàng có giá trị > 1 triệu VÀ trạng thái là Hoàn thành"',
          '✅ "Tính trung bình điểm Toán cho học sinh có xếp loại Giỏi"',
        ],
      },
    ],
  },
  {
    id: "tip-analysis",
    category: "tips",
    title: "Tips phân tích dữ liệu hiệu quả",
    description: "Cách chuẩn bị data và sử dụng AI Analyzer tối ưu",
    tags: ["Data Analysis", "Best Practice"],
    difficulty: "Tips",
    isTip: true,
    tipContent: [
      {
        title: "Chuẩn bị dữ liệu tốt",
        examples: [
          "Header rõ ràng ở hàng đầu tiên (không merge cells)",
          "Không có hàng/cột trống xen giữa dữ liệu",
          "Định dạng nhất quán (ngày tháng, số, text)",
        ],
      },
      {
        title: "Dùng Named Table (Ctrl+T)",
        examples: [
          "Chọn vùng dữ liệu → Ctrl+T → Đặt tên bảng",
          "AI sẽ tự nhận diện cấu trúc Table chính xác hơn",
        ],
      },
      {
        title: "Tận dụng Chart & PivotTable",
        examples: [
          "Sau khi phân tích, click 'Tạo biểu đồ' để tạo chart ngay trong Excel",
          "Click 'Tạo PivotTable' để tổng hợp dữ liệu tự động",
        ],
      },
    ],
  },
  {
    id: "tip-vba",
    category: "tips",
    title: "Hướng dẫn sử dụng VBA Generator",
    description: "Cách mô tả macro VBA để AI tạo code chính xác",
    tags: ["VBA", "Automation"],
    difficulty: "Tips",
    isTip: true,
    tipContent: [
      {
        title: "Mô tả hành động cụ thể",
        examples: [
          '✅ "Tự động highlight các ô có giá trị âm trong cột D bằng màu đỏ"',
          '✅ "Gửi email từ danh sách trong cột A với nội dung ở cột B"',
        ],
      },
      {
        title: "Chỉ rõ phạm vi tác động",
        examples: [
          '"Áp dụng cho tất cả sheet trong workbook"',
          '"Chỉ xử lý vùng đang chọn (selection)"',
          '"Từ hàng 2 đến hàng cuối cùng có dữ liệu"',
        ],
      },
      {
        title: "Cách chạy VBA code",
        examples: [
          "1. Copy code từ kết quả",
          "2. Mở VBA Editor: Alt + F11",
          "3. Insert → Module",
          "4. Paste code vào module",
          "5. Chạy: F5 hoặc nút Run",
        ],
      },
    ],
  },
];

// ============================================================================
// COMPONENT
// ============================================================================

const TemplateGallery = () => {
  const [selectedCategory, setSelectedCategory] = useState("all");
  const [searchQuery, setSearchQuery] = useState("");
  const [expandedTemplate, setExpandedTemplate] = useState(null);
  const [insertingId, setInsertingId] = useState(null);
  const [insertSuccess, setInsertSuccess] = useState("");
  const [insertError, setInsertError] = useState("");

  // Filter templates
  const filteredTemplates = TEMPLATES.filter((t) => {
    const matchCategory = selectedCategory === "all" || t.category === selectedCategory;
    const matchSearch =
      !searchQuery ||
      t.title.toLowerCase().includes(searchQuery.toLowerCase()) ||
      t.description.toLowerCase().includes(searchQuery.toLowerCase()) ||
      t.tags.some((tag) => tag.toLowerCase().includes(searchQuery.toLowerCase()));
    return matchCategory && matchSearch;
  });

  /**
   * Insert template data vào Excel
   */
  const handleInsertTemplate = useCallback(async (template) => {
    if (template.isTip) return;

    setInsertingId(template.id);
    setInsertSuccess("");
    setInsertError("");

    try {
      await Excel.run(async (context) => {
        // LUÔN tạo sheet MỚI - không bao giờ ghi đè data người dùng
        const sheetName = template.title.slice(0, 28).replace(/[\\/*?:\[\]]/g, "");
        const newSheet = context.workbook.worksheets.add(sheetName);
        newSheet.activate();
        await context.sync();

        const lastColLetter = String.fromCharCode(64 + template.headers.length);

        // Insert headers
        const headerRange = newSheet.getRange(`A1:${lastColLetter}1`);
        headerRange.values = [template.headers];
        headerRange.format.font.bold = true;
        headerRange.format.fill.color = "#10b981";
        headerRange.format.font.color = "#ffffff";
        headerRange.format.horizontalAlignment = "Center";

        // Insert sample data
        if (template.sampleData && template.sampleData.length > 0) {
          const dataRange = newSheet.getRange(
            `A2:${lastColLetter}${template.sampleData.length + 1}`
          );
          dataRange.values = template.sampleData;
        }

        // Insert formulas
        if (template.formulas && template.formulas.length > 0) {
          for (const f of template.formulas) {
            const cell = newSheet.getRange(f.cell);
            cell.formulas = [[f.formula]];
          }
        }

        // Auto-fit columns
        const totalRows = (template.sampleData?.length || 0) + 1;
        const allRange = newSheet.getRange(`A1:${lastColLetter}${totalRows}`);
        allRange.format.autofitColumns();

        // Add borders
        allRange.format.borders.getItem("InsideHorizontal").style = "Thin";
        allRange.format.borders.getItem("InsideVertical").style = "Thin";
        allRange.format.borders.getItem("EdgeTop").style = "Thin";
        allRange.format.borders.getItem("EdgeBottom").style = "Thin";
        allRange.format.borders.getItem("EdgeLeft").style = "Thin";
        allRange.format.borders.getItem("EdgeRight").style = "Thin";

        await context.sync();
      });

      setInsertSuccess(`Đã tạo sheet "${template.title}" với mẫu dữ liệu!`);
      setTimeout(() => setInsertSuccess(""), 4000);
    } catch (error) {
      console.error("Insert template error:", error);
      setInsertError("Không thể chèn mẫu: " + error.message);
      setTimeout(() => setInsertError(""), 4000);
    } finally {
      setInsertingId(null);
    }
  }, []);

  return (
    <div className="page-container">
      <div className="page-header">
        <h2 className="page-title">
          <BookOpen24Regular /> Template Gallery
        </h2>
        <p className="page-subtitle">Mẫu Excel & Hướng dẫn sử dụng AI hiệu quả</p>
      </div>

      {/* Search */}
      <div className="template-search">
        <Input
          placeholder="Tìm kiếm mẫu..."
          contentBefore={<Search24Regular />}
          value={searchQuery}
          onChange={(e) => setSearchQuery(e.target.value)}
          style={{ width: "100%" }}
        />
      </div>

      {/* Category Tabs */}
      <div className="template-categories">
        {TEMPLATE_CATEGORIES.map((cat) => (
          <button
            key={cat.id}
            className={`template-category-btn ${
              selectedCategory === cat.id ? "template-category-btn--active" : ""
            }`}
            onClick={() => setSelectedCategory(cat.id)}
          >
            {cat.icon}
            <span>{cat.label}</span>
          </button>
        ))}
      </div>

      {/* Status Messages */}
      {insertSuccess && (
        <div className="action-feedback action-feedback--success" style={{ margin: "0 0 12px 0" }}>
          <CheckmarkCircle24Regular />
          <Text size={200}>{insertSuccess}</Text>
        </div>
      )}
      {insertError && (
        <div className="action-feedback action-feedback--error" style={{ margin: "0 0 12px 0" }}>
          <ErrorCircle24Regular />
          <Text size={200}>{insertError}</Text>
        </div>
      )}

      {/* Template Cards */}
      <div className="template-grid">
        {filteredTemplates.map((template) => (
          <Card
            key={template.id}
            className={`template-card ${
              expandedTemplate === template.id ? "template-card--expanded" : ""
            }`}
            onClick={() =>
              setExpandedTemplate(expandedTemplate === template.id ? null : template.id)
            }
          >
            <div className="template-card-header">
              <div className="template-card-title-row">
                <Text weight="semibold" size={300}>
                  {template.isTip ? "💡" : "📊"} {template.title}
                </Text>
              </div>
              <Text size={200} style={{ color: "#6b7280", marginTop: "4px" }}>
                {template.description}
              </Text>

              {/* Tags */}
              <div className="template-tags">
                {template.tags.map((tag, idx) => (
                  <span key={idx} className="template-tag">
                    {tag}
                  </span>
                ))}
              </div>
            </div>

            {/* Expanded Content */}
            {expandedTemplate === template.id && (
              <div className="template-card-body" onClick={(e) => e.stopPropagation()}>
                {/* Tips Content */}
                {template.isTip && template.tipContent && (
                  <div className="tip-content-list">
                    {template.tipContent.map((tip, idx) => (
                      <div key={idx} className="tip-item">
                        <Text weight="semibold" size={200} className="d-block mb-8">
                          {idx + 1}. {tip.title}
                        </Text>
                        <ul className="tip-examples">
                          {tip.examples.map((ex, i) => (
                            <li key={i}>
                              <Text size={200}>{ex}</Text>
                            </li>
                          ))}
                        </ul>
                      </div>
                    ))}
                  </div>
                )}

                {/* Template Content */}
                {!template.isTip && (
                  <>
                    {/* Preview: Headers */}
                    <div className="template-preview">
                      <Text size={200} weight="semibold" className="d-block mb-8">
                        Cấu trúc cột:
                      </Text>
                      <div className="template-headers-preview">
                        {template.headers.map((h, i) => (
                          <span key={i} className="template-header-chip">
                            {h}
                          </span>
                        ))}
                      </div>
                    </div>

                    {/* AI Prompt Suggestion */}
                    {template.aiPrompt && (
                      <div className="template-ai-prompt">
                        <div className="d-flex align-items-center gap-8 mb-8">
                          <Sparkle24Regular style={{ color: "#10b981" }} />
                          <Text size={200} weight="semibold" style={{ color: "#047857" }}>
                            Thử hỏi AI:
                          </Text>
                        </div>
                        <Text
                          size={200}
                          style={{
                            color: "#065f46",
                            fontStyle: "italic",
                            lineHeight: "1.6",
                          }}
                        >
                          "{template.aiPrompt}"
                        </Text>
                      </div>
                    )}

                    {/* Actions */}
                    <div className="template-actions">
                      <Button
                        appearance="primary"
                        icon={
                          insertingId === template.id ? (
                            <Spinner size="tiny" />
                          ) : (
                            <ClipboardPaste24Regular />
                          )
                        }
                        onClick={() => handleInsertTemplate(template)}
                        disabled={insertingId === template.id}
                        size="small"
                        style={{
                          background: "#10b981",
                          border: "none",
                          color: "white",
                          flex: 1,
                        }}
                      >
                        {insertingId === template.id
                          ? "Đang tạo sheet mới..."
                          : "Chèn mẫu (sheet mới)"}
                      </Button>
                    </div>
                  </>
                )}
              </div>
            )}
          </Card>
        ))}
      </div>

      {/* No Results */}
      {filteredTemplates.length === 0 && (
        <div className="empty-state">
          <Search24Regular className="empty-state__icon" />
          <Text>Không tìm thấy mẫu phù hợp</Text>
          <Text size={200} style={{ color: "#9ca3af", marginTop: "8px" }}>
            Thử từ khóa khác hoặc chọn danh mục
          </Text>
        </div>
      )}
    </div>
  );
};

export default TemplateGallery;
