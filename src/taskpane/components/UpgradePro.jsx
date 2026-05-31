/**
 * UpgradePro Component - Redirect to External Payment Page
 *
 * Tuân thủ chính sách Microsoft: KHÔNG xử lý thanh toán bên trong Add-in.
 * Chỉ hiển thị bảng giá → redirect ra trang web bên ngoài để thanh toán.
 * Polling credits sau khi redirect để detect upgrade.
 */

import * as React from "react";
import { useState, useEffect, useRef } from "react";
import {
  Button,
  Text,
  Spinner,
  DialogTitle,
  DialogContent,
  DialogActions,
  RadioGroup,
  Radio,
  Card,
} from "@fluentui/react-components";
import {
  Sparkle24Filled,
  CheckmarkCircle24Filled,
  Rocket24Regular,
  Shield24Regular,
  Star24Filled,
  Warning24Filled,
  Open24Regular,
} from "@fluentui/react-icons";

import { getCredits } from "../../services/apiService";

// Payment page URL
const PAYMENT_PAGE_URL =
  window.location.hostname === "localhost"
    ? "http://localhost:3000/pricing"
    : "https://fe-rho-lemon.vercel.app/pricing";
const UpgradePro = ({ onClose, currentPlan }) => {
  const [billingCycle, setBillingCycle] = useState("monthly");
  const [selectedPlan, setSelectedPlan] = useState("pro_monthly");
  const [status, setStatus] = useState("idle"); // idle, waiting, success
  const [error, setError] = useState("");

  const pollingRef = useRef(null);

  useEffect(() => {
    return () => {
      if (pollingRef.current) clearInterval(pollingRef.current);
    };
  }, []);

  const handleOpenPaymentPage = () => {
    // Get token from localStorage
    const token = localStorage.getItem("auth_token");
    if (!token) {
      setError("Vui lòng đăng nhập lại để tiếp tục.");
      return;
    }

    // Build URL with token and plan
    const plan = selectedPlan === "student" ? "student" : selectedPlan;
    let url = `${PAYMENT_PAGE_URL}?token=${encodeURIComponent(token)}&plan=${plan}`;

    // If running locally, pass the ngrok API URL so the pricing page calls local BE
    if (window.location.hostname === "localhost") {
      const localApiUrl = "https://broodier-unsistered-orlando.ngrok-free.dev/api/v1";
      url += `&api=${encodeURIComponent(localApiUrl)}`;
    }

    // Open in browser
    window.open(url, "_blank");

    // Start polling credits to detect upgrade
    setStatus("waiting");
    startCreditPolling();
  };

  const startCreditPolling = () => {
    if (pollingRef.current) clearInterval(pollingRef.current);

    let pollCount = 0;
    const maxPolls = 120; // 120 * 3s = 6 minutes

    pollingRef.current = setInterval(async () => {
      pollCount++;
      if (pollCount >= maxPolls) {
        clearInterval(pollingRef.current);
        pollingRef.current = null;
        setStatus("idle");
        return;
      }

      try {
        const creditsData = await getCredits();
        if (creditsData?.plan === "pro") {
          console.log("[UpgradePro] ✅ Pro upgrade detected via polling!");
          clearInterval(pollingRef.current);
          pollingRef.current = null;
          setStatus("success");

          // Trigger App refresh via localStorage
          localStorage.setItem("payment_success", "true");
          localStorage.setItem("payment_timestamp", Date.now().toString());

          // Auto-close after 2s
          setTimeout(() => {
            onClose?.();
          }, 2000);
        }
      } catch (err) {
        // Ignore polling errors, keep trying
      }
    }, 3000);
  };

  // ============== SUCCESS VIEW ==============
  if (status === "success") {
    return (
      <>
        <DialogTitle>
          <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
            <CheckmarkCircle24Filled style={{ color: "#10b981" }} />
            <span>Nâng cấp thành công!</span>
          </div>
        </DialogTitle>
        <DialogContent>
          <div
            style={{
              textAlign: "center",
              padding: "24px 0",
            }}
          >
            <CheckmarkCircle24Filled
              style={{ color: "#10b981", fontSize: "48px", marginBottom: "12px" }}
            />
            <Text
              weight="bold"
              size={500}
              style={{ display: "block", color: "#059669", marginBottom: "8px" }}
            >
              Tài khoản đã được nâng cấp Pro! 🎉
            </Text>
            <Text size={200} style={{ color: "#6b7280" }}>
              Đang cập nhật...
            </Text>
          </div>
        </DialogContent>
      </>
    );
  }

  // ============== PLAN SELECTION + REDIRECT VIEW ==============
  return (
    <>
      <DialogTitle>
        <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
          <Rocket24Regular style={{ color: "#3b82f6" }} />
          <span>Nâng cấp Pro</span>
        </div>
      </DialogTitle>

      <DialogContent>
        <Text style={{ marginBottom: "16px", display: "block" }}>
          Mở khóa toàn bộ tính năng của BangAI
        </Text>

        {/* Waiting Banner */}
        {status === "waiting" && (
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: "8px",
              padding: "10px 12px",
              background: "#eff6ff",
              borderRadius: "8px",
              marginBottom: "12px",
              border: "1px solid #bfdbfe",
            }}
          >
            <Spinner size="tiny" />
            <Text size={200} style={{ color: "#1e40af" }}>
              Đang chờ thanh toán... (trang thanh toán đã mở trên trình duyệt)
            </Text>
          </div>
        )}

        {/* Billing Cycle Toggle */}
        <div
          style={{
            display: "flex",
            justifyContent: "center",
            marginBottom: "16px",
            background: "#f3f4f6",
            padding: "4px",
            borderRadius: "8px",
            width: "fit-content",
            margin: "0 auto 16px auto",
          }}
        >
          <button
            onClick={() => {
              setBillingCycle("monthly");
              setSelectedPlan("pro_monthly");
            }}
            style={{
              padding: "6px 16px",
              borderRadius: "6px",
              border: "none",
              background: billingCycle === "monthly" ? "white" : "transparent",
              boxShadow: billingCycle === "monthly" ? "0 1px 3px rgba(0,0,0,0.1)" : "none",
              fontWeight: billingCycle === "monthly" ? "600" : "400",
              color: billingCycle === "monthly" ? "#111827" : "#6b7280",
              cursor: "pointer",
              transition: "all 0.2s",
            }}
          >
            Hàng tháng
          </button>
          <button
            onClick={() => {
              setBillingCycle("yearly");
              setSelectedPlan("pro_yearly");
            }}
            style={{
              padding: "6px 16px",
              borderRadius: "6px",
              border: "none",
              background: billingCycle === "yearly" ? "white" : "transparent",
              boxShadow: billingCycle === "yearly" ? "0 1px 3px rgba(0,0,0,0.1)" : "none",
              fontWeight: billingCycle === "yearly" ? "600" : "400",
              color: billingCycle === "yearly" ? "#111827" : "#6b7280",
              cursor: "pointer",
              transition: "all 0.2s",
              position: "relative",
            }}
          >
            Hàng năm
            <span
              style={{
                position: "absolute",
                top: "-8px",
                right: "-10px",
                background: "#ef4444",
                color: "white",
                fontSize: "9px",
                padding: "2px 4px",
                borderRadius: "4px",
                fontWeight: "bold",
              }}
            >
              -20%
            </span>
          </button>
        </div>

        <RadioGroup
          value={selectedPlan}
          onChange={(e, data) => setSelectedPlan(data.value)}
          style={{ marginBottom: "16px" }}
        >
          {/* Gói Pro */}
          <Card
            style={{
              width: "100%",
              padding: "12px",
              marginBottom: "8px",
              border:
                selectedPlan === (billingCycle === "monthly" ? "pro_monthly" : "pro_yearly")
                  ? "2px solid #3b82f6"
                  : "1px solid #e5e7eb",
              background:
                selectedPlan === (billingCycle === "monthly" ? "pro_monthly" : "pro_yearly")
                  ? "#eff6ff"
                  : "transparent",
              cursor: "pointer",
            }}
            onClick={() =>
              setSelectedPlan(billingCycle === "monthly" ? "pro_monthly" : "pro_yearly")
            }
          >
            <div
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
              }}
            >
              <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
                <Radio
                  value={billingCycle === "monthly" ? "pro_monthly" : "pro_yearly"}
                  checked={
                    selectedPlan === (billingCycle === "monthly" ? "pro_monthly" : "pro_yearly")
                  }
                />
                <div>
                  <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                    <Text weight="semibold">
                      {billingCycle === "monthly" ? "Gói Pro (Tháng)" : "Gói Pro (Năm)"}
                    </Text>
                    <span
                      style={{
                        background: "#3b82f6",
                        color: "white",
                        padding: "2px 6px",
                        borderRadius: "4px",
                        fontSize: "10px",
                        fontWeight: "600",
                      }}
                    >
                      PHỔ THÔNG
                    </span>
                  </div>
                  <Text size={200} style={{ color: "#6b7280" }}>
                    Thanh toán qua QR Banking
                  </Text>
                </div>
              </div>
              <div style={{ textAlign: "right" }}>
                <Text
                  weight="bold"
                  style={{ color: "#2563eb", fontSize: "18px", display: "block" }}
                >
                  {billingCycle === "monthly" ? "49.000 ₫" : "470.000 ₫"}
                </Text>
              </div>
            </div>
          </Card>

          {/* Gói Sinh viên */}
          <Card
            style={{
              width: "100%",
              padding: "12px",
              marginBottom: "8px",
              border: selectedPlan === "student" ? "2px solid #3b82f6" : "1px solid #e5e7eb",
              background: selectedPlan === "student" ? "#eff6ff" : "transparent",
              cursor: "pointer",
            }}
            onClick={() => setSelectedPlan("student")}
          >
            <div
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
              }}
            >
              <div style={{ display: "flex", alignItems: "center", gap: "12px" }}>
                <Radio value="student" checked={selectedPlan === "student"} />
                <div>
                  <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                    <Text weight="semibold">
                      {billingCycle === "monthly" ? "Gói Sinh Viên (Tháng)" : "Gói Sinh Viên (Năm)"}
                    </Text>
                    <span
                      style={{
                        background: "#3b82f6",
                        color: "white",
                        padding: "2px 6px",
                        borderRadius: "4px",
                        fontSize: "10px",
                        fontWeight: "600",
                      }}
                    >
                      ƯU ĐÃI
                    </span>
                  </div>
                  <Text size={200} style={{ color: "#6b7280" }}>
                    Liên hệ Fanpage để mua
                  </Text>
                </div>
              </div>
              <div style={{ textAlign: "right" }}>
                <Text
                  weight="bold"
                  style={{ color: "#3b82f6", fontSize: "18px", display: "block" }}
                >
                  {billingCycle === "monthly" ? "39.000 ₫" : "468.000 ₫"}
                </Text>
              </div>
            </div>
          </Card>
        </RadioGroup>

        {/* Features */}
        <div
          style={{
            background: "#f8fafc",
            padding: "12px",
            borderRadius: "8px",
            marginBottom: "16px",
          }}
        >
          <Text
            weight="semibold"
            size={200}
            style={{ display: "block", marginBottom: "8px", color: "#059669" }}
          >
            ✨ Quyền lợi Pro:
          </Text>
          <div style={{ display: "flex", alignItems: "center", gap: "8px", marginBottom: "4px" }}>
            <Sparkle24Filled style={{ color: "#10b981", fontSize: "16px" }} />
            <Text size={200}>Không giới hạn AI prompts</Text>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: "8px", marginBottom: "4px" }}>
            <Shield24Regular style={{ color: "#10b981", fontSize: "16px" }} />
            <Text size={200}>Phân tích dữ liệu nâng cao</Text>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
            <Star24Filled style={{ color: "#10b981", fontSize: "16px" }} />
            <Text size={200}>Hỗ trợ ưu tiên 24/7</Text>
          </div>
        </div>

        {/* Error */}
        {error && (
          <div
            style={{
              padding: "8px 12px",
              background: "#fee2e2",
              borderRadius: "6px",
              marginBottom: "12px",
              display: "flex",
              alignItems: "center",
              gap: "8px",
              color: "#991b1b",
            }}
          >
            <Warning24Filled style={{ fontSize: "16px" }} />
            <Text size={200}>{error}</Text>
          </div>
        )}
      </DialogContent>

      <DialogActions>
        <Button appearance="secondary" onClick={onClose}>
          Để sau
        </Button>
        {selectedPlan === "student" ? (
          <Button
            appearance="primary"
            onClick={() => window.open("https://www.facebook.com/ngxnam21/", "_blank")}
            icon={<Rocket24Regular />}
            style={{ background: "#3b82f6" }}
          >
            Liên hệ Fanpage
          </Button>
        ) : (
          <Button
            appearance="primary"
            onClick={handleOpenPaymentPage}
            disabled={status === "waiting"}
            icon={status === "waiting" ? <Spinner size="tiny" /> : <Open24Regular />}
            style={{ background: "#3b82f6" }}
          >
            {status === "waiting" ? "Đã mở trang thanh toán" : "Thanh toán ngay"}
          </Button>
        )}
      </DialogActions>
    </>
  );
};

export default UpgradePro;
