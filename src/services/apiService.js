/**
 * API Service - Frontend API Client
 *
 * ARCHITECTURE:
 * - Frontend (React) CHỈ gọi API, KHÔNG xử lý business logic
 * - Backend (Node.js + Express) xử lý TẤT CẢ business logic bao gồm Gemini AI
 * - Frontend chỉ giữ UI state và hiển thị dữ liệu từ API
 *
 * AUTH FLOW:
 * - Bắt buộc login để sử dụng AI
 * - Free users: 10 lệnh miễn phí
 * - Hết credits = không sử dụng được nữa
 */

import { getExcelContext as getExcelContextFromService } from "./excelContextService.js";

// ============================================================================
// API CONFIGURATION
// ============================================================================

// 🔴 Cấu hình API Localhost
const LOCAL_API_URL = "http://localhost:3001";

// Backend API URL - Auto-detect local vs production
const API_BASE_URL =
  window.location.hostname === "localhost"
    ? `${LOCAL_API_URL}/api/v1`
    : "https://web-production-8b555d.up.railway.app/api/v1";

/**
 * Lấy JWT token từ localStorage
 */
function getAuthToken() {
  return localStorage.getItem("auth_token");
}

/**
 * Lưu JWT token
 */
export function setAuthToken(token) {
  localStorage.setItem("auth_token", token);
}

/**
 * Xóa JWT token (logout)
 */
export function clearAuthToken() {
  localStorage.removeItem("auth_token");
}

/**
 * Kiểm tra đã login chưa
 */
export function isLoggedIn() {
  const token = getAuthToken();
  return !!(token && token.trim());
}

/**
 * Generic API call wrapper với error handling và auth
 * ✅ AUTO-REFRESH: Automatically refresh token on 401 and retry
 */
let isRefreshing = false;
let refreshPromise = null;

async function apiCall(endpoint, options = {}, isRetry = false) {
  try {
    const token = getAuthToken();
    const headers = {
      "Content-Type": "application/json",
      "ngrok-skip-browser-warning": "true",
      ...options.headers,
    };

    // Thêm Authorization header nếu có token
    if (token) {
      headers["Authorization"] = `Bearer ${token}`;
    }

    const response = await fetch(`${API_BASE_URL}${endpoint}`, {
      headers,
      ...options,
    });

    // Handle 401 - Try to refresh token (but NOT for login/register endpoints)
    const isAuthEndpoint = endpoint.includes("/auth/login") || endpoint.includes("/auth/register");
    if (response.status === 401 && !isRetry && !isAuthEndpoint) {
      console.log("[API] 401 detected, attempting token refresh...");

      // Get refresh token
      const refreshToken = localStorage.getItem("refresh_token");

      if (refreshToken) {
        try {
          // Prevent multiple simultaneous refresh calls
          if (!isRefreshing) {
            isRefreshing = true;
            refreshPromise = fetch(`${API_BASE_URL}/auth/refresh`, {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({ refreshToken }),
            }).then(async (res) => {
              if (!res.ok) throw new Error("Refresh failed");
              return res.json();
            });
          }

          const refreshData = await refreshPromise;
          isRefreshing = false;

          if (refreshData.token) {
            console.log("[API] Token refreshed successfully!");
            setAuthToken(refreshData.token);
            if (refreshData.refreshToken) {
              localStorage.setItem("refresh_token", refreshData.refreshToken);
            }

            // Retry original request with new token
            return apiCall(endpoint, options, true);
          }
        } catch (refreshError) {
          console.error("[API] Token refresh failed:", refreshError);
          isRefreshing = false;
        }
      }

      // Refresh failed or no refresh token - clear auth
      clearAuthToken();
      localStorage.removeItem("refresh_token");
      throw new Error("Phiên đăng nhập hết hạn. Vui lòng đăng nhập lại!");
    }

    const data = await response.json();

    if (!response.ok) {
      if (response.status === 403) {
        throw new Error(data.message || "Hết lượt sử dụng. Vui lòng nâng cấp!");
      }
      throw new Error(data.message || `Lỗi API: ${response.status}`);
    }

    return data;
  } catch (error) {
    console.error(`API Call failed [${endpoint}]:`, error);
    throw error;
  }
}

// ============================================================================
// AUTHENTICATION (NEW API with OTP)
// ============================================================================

/**
 * Lưu refresh token
 */
export function setRefreshToken(token) {
  localStorage.setItem("refresh_token", token);
}

/**
 * Lấy refresh token
 */
function getRefreshToken() {
  return localStorage.getItem("refresh_token");
}

/**
 * Xóa tất cả tokens
 */
export function clearAllTokens() {
  localStorage.removeItem("auth_token");
  localStorage.removeItem("refresh_token");
}

/**
 * Đăng ký user mới (sends OTP email)
 * @returns { message, otpSent, email }
 */
export async function register(email, password, confirmPassword, name) {
  const data = await apiCall("/auth/register", {
    method: "POST",
    body: JSON.stringify({ email, password, confirmPassword, name }),
  });
  return data;
}

/**
 * Verify email OTP sau khi đăng ký
 * @returns { user, token, refreshToken }
 */
export async function verifyEmailOTP(email, otp) {
  const data = await apiCall("/auth/verify-email-otp", {
    method: "POST",
    body: JSON.stringify({ email, otp }),
  });

  // Lưu tokens
  if (data.token) {
    setAuthToken(data.token);
  }
  if (data.refreshToken) {
    setRefreshToken(data.refreshToken);
  }

  return data;
}

/**
 * Resend OTP
 * @returns { message, cooldownRemaining, resendsRemaining }
 */
export async function resendOTP(email, purpose = "signup") {
  return apiCall("/auth/resend-otp", {
    method: "POST",
    body: JSON.stringify({ email, purpose }),
  });
}

/**
 * Đăng nhập
 */
export async function login(email, password) {
  const data = await apiCall("/auth/login", {
    method: "POST",
    body: JSON.stringify({ email, password }),
  });

  // Lưu tokens
  if (data.token) {
    setAuthToken(data.token);
  }
  if (data.refreshToken) {
    setRefreshToken(data.refreshToken);
  }

  return data;
}

/**
 * Đăng xuất
 */
export async function logout() {
  try {
    const refreshToken = getRefreshToken();
    await apiCall("/auth/logout", {
      method: "POST",
      body: JSON.stringify({ refreshToken }),
    });
  } catch (error) {
    // Ignore logout errors
  } finally {
    clearAllTokens();
  }
}

/**
 * Forgot password - gửi OTP reset
 */
export async function forgotPassword(email) {
  return apiCall("/auth/forgot-password", {
    method: "POST",
    body: JSON.stringify({ email }),
  });
}

/**
 * Verify reset OTP
 * @returns { message, resetToken }
 */
export async function verifyResetOTP(email, otp) {
  return apiCall("/auth/verify-reset-otp", {
    method: "POST",
    body: JSON.stringify({ email, otp }),
  });
}

/**
 * Reset password với token
 */
export async function resetPassword(resetToken, newPassword, confirmPassword) {
  return apiCall("/auth/reset-password", {
    method: "POST",
    body: JSON.stringify({ resetToken, newPassword, confirmPassword }),
  });
}

/**
 * Refresh access token
 */
export async function refreshAccessToken() {
  const refreshToken = getRefreshToken();
  if (!refreshToken) {
    throw new Error("No refresh token");
  }

  const data = await apiCall("/auth/refresh", {
    method: "POST",
    body: JSON.stringify({ refreshToken }),
  });

  if (data.token) {
    setAuthToken(data.token);
  }
  if (data.refreshToken) {
    setRefreshToken(data.refreshToken);
  }

  return data;
}

/**
 * Lấy thông tin profile
 */
export async function getProfile() {
  // Add timestamp to bust cache
  const cacheBuster = `t=${Date.now()}`;
  return apiCall(`/auth/profile?${cacheBuster}`);
}

/**
 * Lấy thông tin credits
 */
export async function getCredits() {
  // Add timestamp to bust cache
  const cacheBuster = `t=${Date.now()}`;
  return apiCall(`/users/credits?${cacheBuster}`);
}

// ============================================================================
// PAYMENT
// ============================================================================

/**
 * Lấy danh sách giá
 */
export async function getPricing() {
  return apiCall("/payments/pricing");
}

/**
 * Tạo payment intent (QR code)
 * @returns { intent: { id, amount, transferCode, qrData, expiresAt } }
 */
export const createPaymentIntent = async (plan) => {
  const res = await fetch("/api/v1/payments/intents", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${localStorage.getItem("token")}`,
    },
    body: JSON.stringify({ plan }),
  });

  return res.json();
};

export const getPaymentStatus = async (id) => {
  const res = await fetch(`/api/v1/payments/intents/${id}`, {
    headers: {
      Authorization: `Bearer ${localStorage.getItem("token")}`,
    },
  });

  return res.json();
};

/**
 * Lấy trạng thái payment intent (polling)
 * ✅ FIX: Add cache-buster to prevent Office WebView caching
 */
export async function getPaymentIntentStatus(intentId) {
  // Add timestamp to bust cache in Office WebView
  const cacheBuster = `t=${Date.now()}`;
  return apiCall(`/payments/intents/${intentId}?${cacheBuster}`);
}

/**
 * Lấy lịch sử thanh toán
 */
export async function getPaymentHistory(page = 1, limit = 10) {
  return apiCall(`/payments/history?page=${page}&limit=${limit}`);
}

// ============================================================================
// AI FEATURES - GỌI BACKEND
// ============================================================================

/**
 * Generate Excel formula từ prompt
 * Gọi BE endpoint: POST /api/v1/ai/ask với type="formula"
 */
export async function generateExcelFormula(prompt, excelContext = null, model = null) {
  if (!prompt || !prompt.trim()) {
    throw new Error("Prompt không được rỗng!");
  }

  const data = await apiCall("/ai/ask", {
    method: "POST",
    body: JSON.stringify({
      type: "formula",
      prompt,
      excelContext,
      model,
    }),
  });

  // Backend trả về { result: {...}, cached, creditsRemaining }
  return data.result;
}

/**
 * Analyze Excel data
 * Gọi BE endpoint: POST /api/v1/ai/ask với type="analysis"
 */
export async function analyzeExcelData(excelContext, model = null) {
  if (!excelContext || !excelContext.sampleData || excelContext.sampleData.length === 0) {
    throw new Error("Không có dữ liệu để phân tích!");
  }

  const data = await apiCall("/ai/ask", {
    method: "POST",
    body: JSON.stringify({
      type: "analysis",
      prompt: "analyze", // BE cần prompt cho cache key
      excelContext,
      model,
    }),
  });

  return data.result;
}

/**
 * Generate step-by-step guide
 * Gọi BE endpoint: POST /api/v1/ai/ask với type="guide"
 */
export async function generateStepByStep(task, excelContext = null, model = null) {
  if (!task || !task.trim()) {
    throw new Error("Task description không được rỗng!");
  }

  const data = await apiCall("/ai/ask", {
    method: "POST",
    body: JSON.stringify({
      type: "guide",
      prompt: task,
      excelContext,
      model,
    }),
  });

  return data.result;
}

/**
 * Generate VBA/Macro code từ mô tả
 * Gọi BE endpoint: POST /api/v1/ai/ask với type="vba"
 */
export async function generateVBACode(description, excelContext = null, model = null) {
  if (!description || !description.trim()) {
    throw new Error("Mô tả macro không được rỗng!");
  }

  const data = await apiCall("/ai/ask", {
    method: "POST",
    body: JSON.stringify({
      type: "vba",
      prompt: description,
      excelContext,
      model,
    }),
  });

  // Backend trả về { result: {...}, cached, creditsRemaining }
  return data.result;
}

/**
 * Generate chart configuration
 * Gọi BE endpoint: POST /api/v1/ai/ask với type="chart"
 */
export async function generateChartSuggestion(prompt, excelContext = null, model = null) {
  if (!prompt || !prompt.trim()) {
    throw new Error("Mô tả biểu đồ không được rỗng!");
  }

  const data = await apiCall("/ai/ask", {
    method: "POST",
    body: JSON.stringify({
      type: "chart",
      prompt,
      excelContext,
      model,
    }),
  });

  return data.result;
}

/**
 * Insert chart vào Excel
 * NOTE: Client-side operation qua Excel API
 * @param {object} chartConfig - { chartType, dataRange, title, seriesBy }
 */
export async function insertChartToExcel(chartConfig) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // Select data range
      const range = sheet.getRange(chartConfig.dataRange);

      // Determine series by rows or columns
      const seriesBy =
        chartConfig.seriesBy === "rows" ? Excel.ChartSeriesBy.rows : Excel.ChartSeriesBy.columns;

      // Add chart
      const chart = sheet.charts.add(chartConfig.chartType, range, seriesBy);

      // Set title
      chart.title.text = chartConfig.title;

      // Positioning - resize and place nicely
      // Default: place near the data or at a standard position
      chart.top = 50;
      chart.left = 300;
      chart.height = 300;
      chart.width = 500;

      await context.sync();
    });
    return { success: true };
  } catch (error) {
    console.error("Insert chart error:", error);
    throw new Error("Không thể tạo biểu đồ: " + error.message);
  }
}

// ============================================================================
// AI HISTORY
// ============================================================================

/**
 * Lấy lịch sử AI
 */
export async function getAIHistory(type = null, page = 1, limit = 20) {
  let url = `/ai/history?page=${page}&limit=${limit}`;
  if (type) {
    url += `&type=${type}`;
  }
  return apiCall(url);
}

/**
 * Xóa một mục lịch sử
 */
export async function deleteAIHistory(id) {
  return apiCall(`/ai/history/${id}`, {
    method: "DELETE",
  });
}

/**
 * Hủy request AI đang pending
 */
export async function cancelAIRequest(requestId) {
  return apiCall(`/ai/cancel/${requestId}`, {
    method: "POST",
  });
}

// ============================================================================
// EXCEL CONTEXT SERVICE (Client-side - không qua BE)
// ============================================================================

/**
 * Get Excel context từ active worksheet
 * NOTE: Chạy ở CLIENT SIDE (Excel Add-in API)
 */
export async function getExcelContext() {
  return getExcelContextFromService();
}

/**
 * Insert formula vào Excel
 * NOTE: Client-side operation qua Excel API
 */
export async function insertFormulaToExcel(formula, targetCell = null) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = targetCell ? sheet.getRange(targetCell) : context.workbook.getSelectedRange();

      range.load("address");
      await context.sync();

      // Clean formula - loại bỏ các ký tự không hợp lệ
      let cleanFormula = formula.trim();

      // Đảm bảo formula bắt đầu bằng =
      if (!cleanFormula.startsWith("=")) {
        cleanFormula = "=" + cleanFormula;
      }

      // Fix escape issues - thay thế backslash sai
      cleanFormula = cleanFormula.replace(/\\\\/g, "\\");

      // Thử insert như formula trước
      try {
        range.formulas = [[cleanFormula]];
        await context.sync();
      } catch (formulaError) {
        console.warn("Formula insert failed, trying as value:", formulaError);
        // Nếu fail, thử insert như text công thức để user copy
        range.values = [[cleanFormula]];
        await context.sync();
      }

      range.format.autofitColumns();
      await context.sync();

      return { success: true, address: range.address };
    });
  } catch (error) {
    console.error("Insert formula error:", error);
    throw new Error("Không thể insert formula vào Excel: " + error.message);
  }
}

// ============================================================================
// EXCEL CHART CREATION (Client-side - Office.js)
// ============================================================================

/**
 * Map chart type string từ AI → Excel.ChartType enum string
 * Office.js ChartType enum: https://learn.microsoft.com/en-us/javascript/api/excel/excel.charttype
 */
function mapChartType(type) {
  if (!type) return "ColumnClustered";
  const normalized = type.toLowerCase().trim();
  const chartTypeMap = {
    column: "ColumnClustered",
    columnclustered: "ColumnClustered",
    columnstacked: "ColumnStacked",
    bar: "BarClustered",
    barclustered: "BarClustered",
    barstacked: "BarStacked",
    line: "Line",
    linemarkers: "LineMarkers",
    pie: "Pie",
    "3dpie": "Pie3D",
    area: "Area",
    scatter: "XYScatter",
    xyscatter: "XYScatter",
    doughnut: "Doughnut",
    radar: "Radar",
  };
  return chartTypeMap[normalized] || "ColumnClustered";
}

/**
 * Clean range address: loại bỏ sheet name prefix nếu có
 * "Sheet1!A1:D10" → "A1:D10"
 * "'My Sheet'!A1:D10" → "A1:D10"
 */
function cleanRangeAddress(address) {
  if (!address || typeof address !== "string") return null;
  const cleaned = address
    .replace(/^'?[^'!]+'?!/, "")
    .replace(/^[^!]+!/, "")
    .trim();
  if (/^[A-Z]+\d+(:[A-Z]+\d+)?$/i.test(cleaned)) {
    return cleaned;
  }
  return null;
}

/**
 * Phân tích values 2D array: tìm cột nào là text, cột nào là số
 * @param {Array<Array>} values - 2D array values from Excel range
 * @returns {{ textCols: number[], numCols: number[], headerRow: any[] }}
 */
function analyzeRangeData(values) {
  if (!values || values.length < 2) {
    return { textCols: [], numCols: [], idCols: [], headerRow: [] };
  }

  const headerRow = values[0];
  const dataRows = values.slice(1); // skip header
  const colCount = headerRow.length;
  const textCols = [];
  const numCols = [];
  const idCols = []; // Cột ID/STT - không nên dùng cho chart

  for (let col = 0; col < colCount; col++) {
    let numCount = 0;
    let textCount = 0;
    let totalNonEmpty = 0;
    const numericValues = [];

    for (let row = 0; row < dataRows.length; row++) {
      const val = dataRows[row][col];
      if (val === null || val === undefined || val === "") continue;
      totalNonEmpty++;
      if (typeof val === "number" && !isNaN(val)) {
        numCount++;
        numericValues.push(val);
      } else {
        textCount++;
      }
    }

    if (totalNonEmpty === 0) continue; // skip empty columns

    // Detect ID-like columns: sequential integers 1,2,3... or header contains STT/ID/Mã/No
    const headerName = String(headerRow[col] || "")
      .toLowerCase()
      .trim();
    const isIdHeader = /^(stt|id|mã|ma|no\.?|#|số tt|sốtt|ordinal)$/i.test(headerName);
    const isSequential =
      numericValues.length >= 2 &&
      numericValues.every((v, i) => i === 0 || v === numericValues[i - 1] + 1) &&
      numericValues[0] === 1;

    if (
      isIdHeader ||
      (isSequential && numCount === totalNonEmpty && numericValues.length === dataRows.length)
    ) {
      idCols.push(col);
      continue; // Skip ID columns entirely
    }

    if (numCount > textCount) {
      numCols.push(col);
    } else {
      textCols.push(col);
    }
  }

  return { textCols, numCols, idCols, headerRow };
}

/**
 * Convert 0-based column index → Excel column letter(s)
 * 0 → A, 1 → B, ..., 25 → Z, 26 → AA, 27 → AB, ...
 */
function getColLetter(index) {
  let result = "";
  let idx = index;
  while (idx >= 0) {
    result = String.fromCharCode(65 + (idx % 26)) + result;
    idx = Math.floor(idx / 26) - 1;
  }
  return result;
}

/**
 * Smart data range detection: tìm bảng dữ liệu đầu tiên chính xác
 * Xử lý sheet có nhiều bảng bằng cách:
 * 1. Ưu tiên Named Table (Ctrl+T) - đáng tin nhất
 * 2. Nếu user đang chọn vùng >= 2 hàng × 2 cột → dùng selection đó
 * 3. Phát hiện bảng liền mạch đầu tiên (dừng ở hàng/cột trống)
 * 4. Fallback: toàn bộ usedRange
 *
 * @returns {Excel.Range} Range đã load: address, rowCount, columnCount, values, columnIndex
 */
async function getSmartDataRange(context, sheet) {
  // ── 1. Named Tables (Ctrl+T) → most reliable ──
  try {
    const tables = sheet.tables;
    tables.load("items");
    await context.sync();

    if (tables.items.length > 0) {
      const table = tables.items[0];
      table.load("name");
      const tableRange = table.getRange();
      tableRange.load("address, rowCount, columnCount, values, columnIndex");
      await context.sync();
      console.log("[SmartRange] Using Named Table:", table.name, tableRange.address);
      return tableRange;
    }
  } catch (e) {
    console.warn("[SmartRange] Table check failed:", e.message);
  }

  // ── 2. User selection (nếu chọn vùng có ý nghĩa) ──
  try {
    const selection = context.workbook.getSelectedRange();
    selection.load("rowCount, columnCount, address, values, columnIndex");
    await context.sync();

    if (selection.rowCount >= 2 && selection.columnCount >= 2) {
      console.log("[SmartRange] Using user selection:", selection.address);
      return selection;
    }
  } catch (e) {
    console.warn("[SmartRange] Selection check failed:", e.message);
  }

  // ── 3. UsedRange + detect first contiguous table ──
  const usedRange = sheet.getUsedRange();
  usedRange.load("address, rowCount, columnCount, values, columnIndex");
  await context.sync();

  const values = usedRange.values;

  if (usedRange.rowCount <= 1 || usedRange.columnCount <= 1) {
    return usedRange;
  }

  // Walk header row: stop at first empty cell after non-empty ones
  let lastHeaderCol = -1;
  let seenHeader = false;
  for (let col = 0; col < values[0].length; col++) {
    const val = values[0][col];
    if (val !== null && val !== undefined && val !== "") {
      seenHeader = true;
      lastHeaderCol = col;
    } else if (seenHeader) {
      break; // Gap in headers → different table starts after this
    }
  }

  if (lastHeaderCol < 0) {
    console.log("[SmartRange] No headers found, using full usedRange");
    return usedRange;
  }

  // Walk data rows: stop at first fully-empty row (within header columns)
  let lastDataRow = 0;
  for (let row = 1; row < values.length; row++) {
    let hasData = false;
    for (let col = 0; col <= lastHeaderCol; col++) {
      const val = values[row][col];
      if (val !== null && val !== undefined && val !== "") {
        hasData = true;
        break;
      }
    }
    if (hasData) {
      lastDataRow = row;
    } else {
      break; // Empty row → end of first table
    }
  }

  const detectedCols = lastHeaderCol + 1;
  const detectedRows = lastDataRow + 1;

  // If detected block is smaller than usedRange → create sub-range
  if (detectedCols < values[0].length || detectedRows < values.length) {
    const baseColIdx = usedRange.columnIndex;
    const addrMatch = usedRange.address.match(/!?([A-Z]+)(\d+)/);
    const startRow = addrMatch ? parseInt(addrMatch[2]) : 1;

    const startColLetter = getColLetter(baseColIdx);
    const endColLetter = getColLetter(baseColIdx + lastHeaderCol);
    const rangeAddr = `${startColLetter}${startRow}:${endColLetter}${startRow + lastDataRow}`;

    console.log(
      "[SmartRange] Detected first table:",
      rangeAddr,
      `(${detectedRows}r × ${detectedCols}c of ${usedRange.rowCount}r × ${usedRange.columnCount}c)`
    );

    const subRange = sheet.getRange(rangeAddr);
    subRange.load("address, rowCount, columnCount, values, columnIndex");
    await context.sync();
    return subRange;
  }

  console.log("[SmartRange] Single table, using full usedRange:", usedRange.address);
  return usedRange;
}

/**
 * Tạo biểu đồ trực tiếp trong Excel từ chart suggestion
 *
 * LOGIC CHỦ ĐẠO:
 * 1. Smart Range Detection - tìm bảng dữ liệu đầu tiên
 * 2. Phân tích cột text vs số, bỏ ID/STT
 * 3. Build range tối ưu: 1 cột label + cột số phù hợp
 *
 * @param {object} chartSuggestion - { type, title, description, dataRange? }
 * @param {object} excelContext - Context từ Excel (usedRange, columns...)
 */
export async function createChartInExcel(chartSuggestion, excelContext = null) {
  if (!chartSuggestion || !chartSuggestion.type) {
    throw new Error("Không có gợi ý biểu đồ phù hợp!");
  }

  const chartType = String(chartSuggestion.type).toLowerCase().trim();
  if (chartType === "null" || chartType === "none" || chartType === "") {
    throw new Error("AI không đề xuất loại biểu đồ cho dữ liệu này.");
  }

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      await context.sync();

      // ═══════════════════════════════════════════════════════
      // STEP 1: Smart Range Detection
      //   Tự tìm bảng dữ liệu đầu tiên (Named Table > Selection > First block)
      //   Xử lý đúng sheet có nhiều bảng
      // ═══════════════════════════════════════════════════════
      const liveUsedRange = await getSmartDataRange(context, sheet);

      console.log("[Chart] Sheet:", sheet.name);
      console.log(
        "[Chart] Smart range:",
        liveUsedRange.address,
        `(${liveUsedRange.rowCount}r × ${liveUsedRange.columnCount}c)`
      );

      if (liveUsedRange.rowCount < 2) {
        throw new Error("Cần ít nhất 2 hàng (header + data) để tạo biểu đồ!");
      }

      // ═══════════════════════════════════════════════════════
      // STEP 2: Phân tích data thực tế trong usedRange
      // ═══════════════════════════════════════════════════════
      const allValues = liveUsedRange.values;
      const analysis = analyzeRangeData(allValues);

      console.log(
        "[Chart] Data analysis:",
        `textCols=[${analysis.textCols}]`,
        `numCols=[${analysis.numCols}]`,
        `idCols=[${analysis.idCols}]`,
        `headers=[${analysis.headerRow.slice(0, 8)}]`
      );

      if (analysis.numCols.length === 0) {
        throw new Error(
          "Không tìm thấy cột dữ liệu số nào! Biểu đồ cần ít nhất 1 cột chứa số. " +
            `Các cột hiện tại: ${analysis.headerRow.join(", ")}`
        );
      }

      // ═══════════════════════════════════════════════════════
      // STEP 3: Build range tối ưu cho biểu đồ
      //   - Bỏ qua cột ID/STT (không có giá trị phân tích)
      //   - Pie chart: chỉ 1 label + 1 value (quá nhiều series sẽ rối)
      //   - Các loại khác: 1 label + tất cả numeric cols
      // ═══════════════════════════════════════════════════════
      const startCol = liveUsedRange.columnIndex; // 0-based column index of usedRange
      const totalRows = liveUsedRange.rowCount;
      const isPieType = chartType === "pie" || chartType === "doughnut" || chartType === "3dpie";

      // Columns to include in chart (relative to usedRange start)
      let chartColIndices = [];

      // Include first text column as category labels
      if (analysis.textCols.length > 0) {
        chartColIndices.push(analysis.textCols[0]);
      }

      // For pie/doughnut: only 1 value column (the first numeric)
      // For others: all numeric columns
      if (isPieType) {
        chartColIndices.push(analysis.numCols[0]);
      } else {
        chartColIndices = chartColIndices.concat(analysis.numCols);
      }

      // Sort by column index to maintain order
      chartColIndices.sort((a, b) => a - b);

      console.log(
        "[Chart] Selected columns (relative):",
        chartColIndices,
        isPieType ? "(PIE - 1 value only)" : `(${analysis.numCols.length} value cols)`
      );

      // Build the actual Excel range
      let dataRange;
      let rangeDescription;

      if (chartColIndices.length >= 2 && liveUsedRange.columnCount > 2) {
        // Nhiều cột → chỉ lấy các cột cần thiết
        // Kiểm tra xem các cột có liên tiếp không
        const isContiguous = chartColIndices.every(
          (col, i) => i === 0 || col === chartColIndices[i - 1] + 1
        );

        if (isContiguous) {
          // Cột liên tiếp → dùng range liền
          const firstAbsCol = startCol + chartColIndices[0];
          const lastAbsCol = startCol + chartColIndices[chartColIndices.length - 1];
          const startColLetter = getColLetter(firstAbsCol);
          const endColLetter = getColLetter(lastAbsCol);
          // Determine actual start row from range address
          // Safe: strip sheet name before extracting row (sheet name may contain digits)
          const chartCellRef = liveUsedRange.address.includes("!")
            ? liveUsedRange.address.split("!").pop()
            : liveUsedRange.address;
          const chartStartRow = parseInt((chartCellRef.match(/\d+/) || ["1"])[0], 10);
          const rangeAddr = `${startColLetter}${chartStartRow}:${endColLetter}${chartStartRow + totalRows - 1}`;
          dataRange = sheet.getRange(rangeAddr);
          rangeDescription = rangeAddr;
        } else {
          // Cột không liên tiếp → dùng toàn bộ usedRange (an toàn hơn)
          dataRange = liveUsedRange;
          rangeDescription = liveUsedRange.address;
        }
      } else {
        // Ít cột hoặc đơn giản → dùng toàn bộ usedRange
        dataRange = liveUsedRange;
        rangeDescription = liveUsedRange.address;
      }

      dataRange.load("address, rowCount, columnCount");
      await context.sync();

      console.log(
        "[Chart] Final dataRange:",
        dataRange.address,
        `(${dataRange.rowCount}r × ${dataRange.columnCount}c)`
      );

      // ═══════════════════════════════════════════════════════
      // STEP 4: Tạo chart
      // ═══════════════════════════════════════════════════════
      const chartTypeName = mapChartType(chartType);
      console.log("[Chart] Type:", chartType, "→", chartTypeName);

      // SeriesBy: "Columns" khi có label ở cột đầu + data ở các cột tiếp
      // Mỗi cột số = 1 series, các hàng = categories
      const seriesBy = analysis.textCols.length > 0 ? "Columns" : "Auto";
      const chart = sheet.charts.add(chartTypeName, dataRange, seriesBy);

      // Title
      chart.title.text = chartSuggestion.title || "Biểu đồ dữ liệu";
      chart.title.format.font.size = 14;
      chart.title.format.font.bold = true;

      // Size
      chart.height = 320;
      chart.width = 500;

      // Position: đặt bên phải data
      const lastDataCol = startCol + liveUsedRange.columnCount;
      const posCol = getColLetter(lastDataCol + 1);
      chart.setPosition(`${posCol}2`);

      // Legend
      try {
        chart.legend.position = "Bottom";
        chart.legend.format.font.size = 10;
      } catch (e) {
        console.warn("[Chart] Legend error:", e.message);
      }

      await context.sync();

      // ═══════════════════════════════════════════════════════
      // STEP 5: Post-create formatting (needs sync first)
      // ═══════════════════════════════════════════════════════
      if (chartType === "pie" || chartType === "doughnut") {
        try {
          const series = chart.series.getItemAt(0);
          series.hasDataLabels = true;
          await context.sync();
          series.dataLabels.showPercentage = true;
          series.dataLabels.showCategoryName = true;
          series.dataLabels.showValue = false;
          series.dataLabels.separator = "\n";
          await context.sync();
        } catch (e) {
          console.warn("[Chart] Pie labels error:", e.message);
        }
      }

      console.log(
        "✅ Chart created:",
        chartSuggestion.title,
        "| type:",
        chartTypeName,
        "| range:",
        rangeDescription
      );
    });

    return { success: true, chartType: chartType, title: chartSuggestion.title };
  } catch (error) {
    console.error("❌ Create chart error:", error);
    const msg = error.message || String(error);
    if (msg.includes("InvalidArgument") || msg.includes("InvalidReference")) {
      throw new Error(
        "Vùng dữ liệu không hợp lệ. Đảm bảo bảng Excel có header ở hàng 1 và dữ liệu số ở các hàng tiếp theo."
      );
    }
    throw new Error("Không thể tạo biểu đồ: " + msg);
  }
}

// ============================================================================
// EXCEL PIVOT TABLE CREATION (Client-side - Office.js)
// ============================================================================

/**
 * Tạo PivotTable trong Excel
 * @param {object} pivotConfig - { name, rowFields, valueFields, filterFields, columnFields }
 * @param {object} excelContext - Context từ Excel
 */
export async function createPivotTableInExcel(pivotConfig = {}, excelContext = null) {
  try {
    const result = await Excel.run(async (context) => {
      const sourceSheet = context.workbook.worksheets.getActiveWorksheet();
      sourceSheet.load("name");
      await context.sync();

      // ═══════════════════════════════════════════════════════
      // STEP 1: Smart Range Detection (same logic as chart)
      // ═══════════════════════════════════════════════════════
      const liveRange = await getSmartDataRange(context, sourceSheet);

      console.log(
        "[Pivot] Source:",
        liveRange.address,
        `${liveRange.rowCount}r × ${liveRange.columnCount}c`
      );

      if (liveRange.rowCount < 2) {
        throw new Error("Cần ít nhất 2 hàng dữ liệu (1 header + 1 data) để tạo PivotTable!");
      }

      // ═══════════════════════════════════════════════════════
      // STEP 2: Validate headers - PivotTable CỨNG yêu cầu header rõ ràng
      // ═══════════════════════════════════════════════════════
      const allValues = liveRange.values;
      const headerRow = allValues[0];

      // Check for empty/duplicate headers (PivotTable sẽ fail nếu có)
      const realHeaders = [];
      const seenNames = new Set();
      for (let i = 0; i < headerRow.length; i++) {
        let name = headerRow[i];
        if (name === null || name === undefined || name === "") {
          name = `Column${getColLetter(i)}`;
        } else {
          name = String(name).trim();
        }
        // Deduplicate (PivotTable cần tên unique)
        let finalName = name;
        let suffix = 2;
        while (seenNames.has(finalName)) {
          finalName = `${name}_${suffix++}`;
        }
        seenNames.add(finalName);
        realHeaders.push(finalName);
      }

      console.log("[Pivot] Headers:", realHeaders);

      if (realHeaders.length === 0) {
        throw new Error("Không tìm thấy header. PivotTable cần hàng đầu tiên là tiêu đề cột.");
      }

      // Phân tích text vs number columns
      const dataRows = allValues.slice(1);
      const isIdColumn = (name) => /^(stt|id|mã|ma|no\.?|#|số tt)$/i.test(name);

      const textHeaders = [];
      const numHeaders = [];

      realHeaders.forEach((header, colIdx) => {
        if (isIdColumn(header)) return; // skip ID columns

        let numCount = 0;
        let textCount = 0;
        for (const row of dataRows) {
          const val = row[colIdx];
          if (val === null || val === undefined || val === "") continue;
          if (typeof val === "number" && !isNaN(val)) numCount++;
          else textCount++;
        }
        if (numCount > textCount) {
          numHeaders.push(header);
        } else if (textCount > 0) {
          textHeaders.push(header);
        }
      });

      console.log("[Pivot] Text cols:", textHeaders, "| Num cols:", numHeaders);

      if (numHeaders.length === 0) {
        throw new Error(
          "Cần ít nhất 1 cột chứa số để tạo PivotTable ý nghĩa. " +
            `Các cột: [${realHeaders.join(", ")}]`
        );
      }

      // ═══════════════════════════════════════════════════════
      // STEP 3: Fix headers nếu có empty/duplicate (ghi lại vào sheet)
      // ═══════════════════════════════════════════════════════
      const needsHeaderFix = headerRow.some(
        (h, i) => h === null || h === undefined || h === "" || String(h).trim() !== realHeaders[i]
      );

      let sourceRange = liveRange;
      if (needsHeaderFix) {
        console.log("[Pivot] Fixing headers in source range...");
        // Write cleaned headers back to the header row
        const addrMatch = liveRange.address.match(/!?([A-Z]+)(\d+)/);
        const startRow = addrMatch ? parseInt(addrMatch[2]) : 1;
        const startColIdx = liveRange.columnIndex;
        const headerAddr = `${getColLetter(startColIdx)}${startRow}:${getColLetter(startColIdx + realHeaders.length - 1)}${startRow}`;
        const headerRange = sourceSheet.getRange(headerAddr);
        headerRange.values = [realHeaders];
        await context.sync();
        // Re-fetch the range after fixing
        sourceRange = sourceSheet.getRange(liveRange.address.replace(/^.*!/, ""));
        sourceRange.load("address, rowCount, columnCount");
        await context.sync();
      }

      // ═══════════════════════════════════════════════════════
      // STEP 4: Create PivotTable on new sheet
      // ═══════════════════════════════════════════════════════
      const pivotSheetName = `Pivot_${Date.now().toString(36).slice(-5)}`;
      const pivotSheet = context.workbook.worksheets.add(pivotSheetName);
      const destRange = pivotSheet.getRange("A3");

      // Unique name to avoid conflicts
      const pivotName = `PT_${Date.now().toString(36)}`;

      console.log("[Pivot] Creating PivotTable:", pivotName, "from", sourceRange.address);

      const pivotTable = context.workbook.pivotTables.add(pivotName, sourceRange, destRange);
      await context.sync();

      // ═══════════════════════════════════════════════════════
      // STEP 5: Load available hierarchies and add fields
      // ═══════════════════════════════════════════════════════
      pivotTable.hierarchies.load("items");
      await context.sync();

      const availableFields = pivotTable.hierarchies.items.map((h) => h.name);
      console.log("[Pivot] Available fields:", availableFields);

      // Fuzzy field name matching
      function findField(requestedName) {
        if (!requestedName) return null;
        const exact = availableFields.find((f) => f === requestedName);
        if (exact) return exact;
        const lower = requestedName.toLowerCase().trim();
        const caseMatch = availableFields.find((f) => f.toLowerCase().trim() === lower);
        if (caseMatch) return caseMatch;
        const partial = availableFields.find(
          (f) => f.toLowerCase().includes(lower) || lower.includes(f.toLowerCase())
        );
        return partial || null;
      }

      // Determine row + value fields
      let rowFields = (pivotConfig.rowFields || []).map(findField).filter(Boolean);
      let valueFields = (pivotConfig.valueFields || []).map(findField).filter(Boolean);

      // Auto-detect if config doesn't match
      if (rowFields.length === 0) {
        rowFields = textHeaders.slice(0, 1).map(findField).filter(Boolean);
      }
      if (valueFields.length === 0) {
        valueFields = numHeaders.slice(0, 3).map(findField).filter(Boolean);
      }

      console.log("[Pivot] Adding rows:", rowFields, "| values:", valueFields);

      // Add row fields
      let addedCount = 0;
      for (const fieldName of rowFields) {
        try {
          pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem(fieldName));
          await context.sync();
          addedCount++;
        } catch (e) {
          console.warn(`[Pivot] Row field "${fieldName}" failed:`, e.message);
        }
      }

      // Add value fields (DON'T set summarizeBy immediately - it can cause errors)
      for (const fieldName of valueFields) {
        try {
          pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem(fieldName));
          await context.sync();
          addedCount++;
        } catch (e) {
          console.warn(`[Pivot] Value field "${fieldName}" failed:`, e.message);
        }
      }

      // Try to set summarizeBy AFTER all fields are added (less error-prone)
      try {
        pivotTable.dataHierarchies.load("items");
        await context.sync();
        for (const dh of pivotTable.dataHierarchies.items) {
          try {
            dh.summarizeBy = "Sum";
          } catch (e) {
            // Ignore - default aggregation is fine
          }
        }
        await context.sync();
      } catch (e) {
        console.warn("[Pivot] summarizeBy failed (using defaults):", e.message);
      }

      if (addedCount === 0) {
        throw new Error(
          "Không thể thêm trường nào vào PivotTable. " +
            `Headers: [${realHeaders.join(", ")}]. Available: [${availableFields.join(", ")}]`
        );
      }

      pivotSheet.activate();
      await context.sync();
      console.log("✅ PivotTable created:", pivotName, `(${addedCount} fields)`);

      return { name: pivotName, sheetName: pivotSheetName, fieldCount: addedCount };
    });

    return { success: true, name: result.name || "PivotTable" };
  } catch (error) {
    console.error("❌ Create PivotTable error:", error);
    const msg = error.message || String(error);
    if (msg.includes("InvalidArgument") || msg.includes("invalid or missing")) {
      throw new Error(
        "Dữ liệu không phù hợp cho PivotTable. Kiểm tra: " +
          "(1) Hàng đầu tiên phải là tiêu đề rõ ràng, " +
          "(2) Không có ô merge, " +
          "(3) Không có hàng trống xen giữa dữ liệu."
      );
    }
    throw new Error("Không thể tạo PivotTable: " + msg);
  }
}

/**
 * Auto-suggest PivotTable configuration dựa trên Excel context
 * @param {object} excelContext - Excel context data
 * @returns {object} - Suggested pivot config
 */
export function suggestPivotConfig(excelContext) {
  if (!excelContext || !excelContext.columns || excelContext.columns.length === 0) {
    return null;
  }

  // Skip ID-like columns (STT, ID, Mã, No, #...)
  const isIdColumn = (col) => {
    const name = String(col.name || "")
      .toLowerCase()
      .trim();
    return /^(stt|id|mã|ma|no\.?|#|số tt|sốtt|ordinal)$/i.test(name);
  };

  const textColumns = excelContext.columns.filter(
    (c) => c.type === "text" && c.hasData && !isIdColumn(c)
  );
  const numberColumns = excelContext.columns.filter(
    (c) => c.type === "number" && c.hasData && !isIdColumn(c)
  );
  const dateColumns = excelContext.columns.filter((c) => c.type === "date" && c.hasData);

  if (numberColumns.length === 0) {
    return null; // Cần ít nhất 1 cột số để tạo PivotTable ý nghĩa
  }

  return {
    rowFields: textColumns.slice(0, 2).map((c) => c.name),
    valueFields: numberColumns.slice(0, 3).map((c) => c.name),
    columnFields: dateColumns.length > 0 ? [dateColumns[0].name] : [],
    filterFields: textColumns.length > 2 ? [textColumns[2].name] : [],
  };
}

// ============================================================================
// DEPRECATED - Giữ lại cho backward compatibility nhưng sẽ xóa
// ============================================================================

/**
 * @deprecated Không còn cần API key vì BE dùng system key
 */
export function hasApiKey() {
  // Giờ chỉ cần check login
  return isLoggedIn();
}

/**
 * @deprecated Không còn cần
 */
export async function saveApiKey(apiKey) {
  console.warn("saveApiKey() is deprecated. System uses server-side API key.");
  return { success: true };
}

/**
 * @deprecated Không còn cần
 */
export async function clearApiKey() {
  console.warn("clearApiKey() is deprecated. System uses server-side API key.");
  return { success: true };
}

/**
 * @deprecated Không còn cần
 */
export function getApiKeyMasked() {
  return "System API Key";
}

// ============================================================================
// ADMIN APIs
// ============================================================================

/**
 * Lấy thống kê admin
 */
export async function getAdminStats() {
  return apiCall("/admin/stats");
}

/**
 * Lấy danh sách giao dịch admin
 */
export async function getAdminTransactions(page = 1, limit = 20, search = "", status = "") {
  let url = `/admin/transactions?page=${page}&limit=${limit}`;
  if (search) url += `&search=${encodeURIComponent(search)}`;
  if (status) url += `&status=${status}`;
  return apiCall(url);
}

/**
 * Lấy danh sách giao dịch chưa khớp
 */
export async function getAdminUnmatchedTransactions(limit = 50) {
  return apiCall(`/admin/transactions/unmatched?limit=${limit}`);
}

/**
 * Khớp giao dịch thủ công
 */
export async function manualMatchTransaction(transactionId, intentId) {
  return apiCall("/admin/transactions/manual-match", {
    method: "POST",
    body: JSON.stringify({ transactionId, intentId }),
  });
}

/**
 * Lấy danh sách người dùng admin
 */
export async function getAdminUsers(page = 1, limit = 20, search = "", plan = "") {
  let url = `/admin/users?page=${page}&limit=${limit}`;
  if (search) url += `&search=${encodeURIComponent(search)}`;
  if (plan) url += `&plan=${plan}`;
  return apiCall(url);
}

/**
 * Lấy logs webhook
 */
export async function getWebhookLogs(page = 1, limit = 50, status = "") {
  let url = `/admin/webhook-logs?page=${page}&limit=${limit}`;
  if (status) url += `&status=${status}`;
  return apiCall(url);
}

/**
 * Admin: Nâng cấp user lên Pro thủ công
 */
export async function adminUpgradeUser(userId, plan = "pro_monthly") {
  return apiCall(`/admin/users/${userId}/upgrade`, {
    method: "POST",
    body: JSON.stringify({ plan }),
  });
}

// ============================================================================
// EXPORT
// ============================================================================

export default {
  // Auth
  register,
  login,
  logout,
  isLoggedIn,
  getProfile,
  getCredits,

  // AI Features
  generateExcelFormula,
  analyzeExcelData,
  generateStepByStep,
  generateVBACode,
  generateChartSuggestion,
  insertChartToExcel,

  // AI History
  getAIHistory,
  deleteAIHistory,
  cancelAIRequest,

  // Excel Context (client-side)
  getExcelContext,
  insertFormulaToExcel,

  // Chart & PivotTable (client-side)
  createChartInExcel,
  createPivotTableInExcel,
  suggestPivotConfig,

  // Deprecated
  hasApiKey,
  saveApiKey,
  clearApiKey,
  getApiKeyMasked,

  // Admin
  getAdminStats,
  getAdminTransactions,
  getAdminUnmatchedTransactions,
  manualMatchTransaction,
  getAdminUsers,
  getWebhookLogs,
  adminUpgradeUser,
};
