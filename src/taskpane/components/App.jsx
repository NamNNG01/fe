/**
 * App Component - Main Application Container
 *
 * AUTH FLOW:
 * - Bắt buộc login để sử dụng AI features
 * - Free users: 10 lệnh miễn phí
 * - Hết credits = không sử dụng được
 */

import * as React from "react";
import { useState, useEffect } from "react";
import PropTypes from "prop-types";
import {
  Tab,
  TabList,
  Button,
  Text,
  Spinner,
  Dialog,
  DialogSurface,
  DialogBody,
  Tooltip,
} from "@fluentui/react-components";
import {
  Sparkle24Regular,
  Lightbulb24Regular,
  DataUsage24Regular,
  SignOut24Regular,
  Person24Regular,
  History24Regular,
  Star24Regular,
  Settings24Regular,
  Code24Regular,
  BookOpen24Regular,
  Wrench24Regular,
} from "@fluentui/react-icons";

// Components
import Logo from "./Logo";
import FormulaGenerator from "./FormulaGenerator";
import StepByStepGuide from "./StepByStepGuide";
import DataAnalyzer from "./DataAnalyzer";
import AuthPage from "./AuthPage";
import History from "./History";
import UpgradePro from "./UpgradePro";
import AdminDashboard from "./admin/AdminDashboard";
import VBAGenerator from "./VBAGenerator";
import TemplateGallery from "./TemplateGallery";
import ExcelTools from "./ExcelTools";

// API Service
import { isLoggedIn, getProfile, getCredits, logout } from "../../services/apiService";

const App = (props) => {
  const { title } = props;
  const [selectedTab, setSelectedTab] = useState("formula");
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [isLoading, setIsLoading] = useState(true);
  const [user, setUser] = useState(null);
  const [credits, setCredits] = useState(null);
  const [showUpgradeDialog, setShowUpgradeDialog] = useState(false);
  const [isAdminViewOpen, setIsAdminViewOpen] = useState(false);
  const [windowWidth, setWindowWidth] = useState(window.innerWidth);

  // Check authentication on mount
  useEffect(() => {
    checkAuth();

    // Track window width for responsive tabs
    const handleResize = () => setWindowWidth(window.innerWidth);
    window.addEventListener("resize", handleResize);
    return () => window.removeEventListener("resize", handleResize);
  }, []);

  // Is compact mode (only icons)
  const isCompact = windowWidth < 480;

  // ✅ NEW LOGIC: Poll localStorage for payment success flag
  // This is more reliable than callbacks or events in Office iframe
  useEffect(() => {
    const checkPaymentSuccess = async () => {
      const flag = localStorage.getItem("payment_success");
      if (flag === "true") {
        console.log("[App] 🎉 Payment success flag detected! Refreshing data...");

        // Clear flag immediately to prevent duplicate processing
        localStorage.removeItem("payment_success");
        localStorage.removeItem("payment_timestamp");

        try {
          const [profileData, creditsData] = await Promise.all([getProfile(), getCredits()]);
          console.log("[App] Fresh data - plan:", creditsData?.plan);
          setUser(profileData);
          setCredits(creditsData);
          setShowUpgradeDialog(false);
        } catch (error) {
          console.error("[App] Failed to refresh after payment:", error);
        }
      }
    };

    // Check every 500ms
    const interval = setInterval(checkPaymentSuccess, 500);
    return () => clearInterval(interval);
  }, []);

  const checkAuth = async () => {
    setIsLoading(true);
    if (isLoggedIn()) {
      try {
        const [profileData, creditsData] = await Promise.all([getProfile(), getCredits()]);
        setUser(profileData);
        setCredits(creditsData);
        setIsAuthenticated(true);
      } catch (error) {
        console.error("Auth check failed:", error);
        setIsAuthenticated(false);
      }
    } else {
      setIsAuthenticated(false);
    }
    setIsLoading(false);
  };

  const handleLoginSuccess = async (userData) => {
    // ✅ CRITICAL: Update UI state IMMEDIATELY with data from login/verification response
    if (userData) {
      setUser(userData);
      setIsAuthenticated(true);
      setIsLoading(false);
    }

    // ✅ Then fetch latest data in background to sync
    try {
      const [profileData, creditsData] = await Promise.all([getProfile(), getCredits()]);
      setUser(profileData);
      setCredits(creditsData);
    } catch (error) {
      console.warn("Failed to fetch latest profile after login:", error);
      // User data from login response is still valid, so we don't clear auth state
      // Just fetch credits if possible
      try {
        const creditsData = await getCredits();
        setCredits(creditsData);
      } catch (creditsError) {
        console.warn("Failed to fetch credits:", creditsError);
      }
    }
  };

  const handleLogout = async () => {
    await logout();
    setIsAuthenticated(false);
    setUser(null);
    setCredits(null);
  };

  const handleTabChange = (_, data) => setSelectedTab(data.value);

  // Refresh credits after AI request
  const refreshCredits = async () => {
    try {
      const creditsData = await getCredits();
      setCredits(creditsData);
    } catch (error) {
      console.error("Failed to refresh credits:", error);
    }
  };

  // Loading state
  if (isLoading) {
    return (
      <div className="app-root">
        <div className="app-header">
          <Logo size={28} color="white" textColor="white" />
        </div>
        <div className="loading-container">
          <Spinner size="large" />
          <Text>Đang tải...</Text>
        </div>
      </div>
    );
  }

  // Not authenticated - show login
  if (!isAuthenticated) {
    return (
      <div className="app-root">
        <div className="app-header">
          <Logo size={28} color="white" textColor="white" />
          <p className="app-header__subtitle">Trợ lý Excel thông minh</p>
        </div>
        <AuthPage onLoginSuccess={handleLoginSuccess} />
      </div>
    );
  }

  // Authenticated - show main app
  return (
    <div className="app-root">
      {/* Header */}
      <div className="app-header">
        <Logo size={32} color="white" textColor="white" />
        <p className="app-header__subtitle">Trợ lý Excel AI của nhóm BangAI</p>
      </div>

      {/* User Info Bar */}
      <div className="user-info-bar">
        <div className="user-info">
          <Person24Regular />
          <Text weight="semibold">{user?.name || "User"}</Text>
        </div>
        <div className="credits-info">
          {credits?.plan === "pro" || user?.role === "admin" ? (
            <Text
              className="credits-badge credits-badge--pro"
              title={
                credits?.endDate
                  ? `Hết hạn: ${new Date(credits.endDate).toLocaleDateString("vi-VN")}`
                  : user?.role === "admin"
                    ? "Quyền Admin vô hạn"
                    : ""
              }
            >
              PRO{" "}
              {user?.role === "admin"
                ? "∞"
                : credits?.endDate
                  ? `đến ${new Date(credits.endDate).toLocaleDateString("vi-VN")}`
                  : "∞"}
            </Text>
          ) : (
            <Button
              appearance="subtle"
              size="small"
              icon={<Star24Regular />}
              onClick={() => setShowUpgradeDialog(true)}
              className="credits-badge"
            >
              {credits?.credits ?? 0} lượt | Nâng cấp
            </Button>
          )}
        </div>
        <div className="user-actions">
          {user?.role === "admin" && (
            <Tooltip content="Quản trị hệ thống" relationship="label" positioning="below">
              <Button
                appearance="subtle"
                size="small"
                icon={<Settings24Regular />}
                onClick={() => setIsAdminViewOpen(true)}
                aria-label="Quản trị hệ thống"
              />
            </Tooltip>
          )}
          <Tooltip content="Đăng xuất" relationship="label" positioning="below">
            <Button
              appearance="subtle"
              size="small"
              icon={<SignOut24Regular />}
              onClick={handleLogout}
              aria-label="Đăng xuất"
            />
          </Tooltip>
        </div>
      </div>

      {isAdminViewOpen && (
        <div className="admin-overlay">
          <AdminDashboard onClose={() => setIsAdminViewOpen(false)} />
        </div>
      )}

      {/* No Credits Warning with Upgrade Button */}
      {user?.role !== "admin" && credits?.plan !== "pro" && credits?.credits <= 0 && (
        <div className="upgrade-banner">
          <Text className="upgrade-banner__text">⚠️ Bạn đã hết lượt sử dụng miễn phí!</Text>
          <Button
            appearance="primary"
            size="small"
            icon={<Star24Regular />}
            onClick={() => setShowUpgradeDialog(true)}
          >
            Nâng cấp Pro
          </Button>
        </div>
      )}

      {/* Fluent UI Tabs - compact vertical layout (icon top, label bottom) */}
      <div className={`tabs-container ${isCompact ? "tabs-container--compact" : ""}`}>
        <TabList selectedValue={selectedTab} onTabSelect={handleTabChange} size="small">
          <Tooltip content="Công thức AI" relationship="label" positioning="below">
            <Tab value="formula" icon={<Sparkle24Regular />}>
              {!isCompact && <span>Công thức</span>}
            </Tab>
          </Tooltip>
          <Tooltip content="Phân tích dữ liệu" relationship="label" positioning="below">
            <Tab value="analyzer" icon={<DataUsage24Regular />}>
              {!isCompact && <span>Phân tích</span>}
            </Tab>
          </Tooltip>
          <Tooltip content="Hướng dẫn sử dụng" relationship="label" positioning="below">
            <Tab value="stepbystep" icon={<Lightbulb24Regular />}>
              {!isCompact && <span>Hướng dẫn</span>}
            </Tab>
          </Tooltip>
          <Tooltip content="Tạo mã VBA/Macro" relationship="label" positioning="below">
            <Tab value="vba" icon={<Code24Regular />}>
              {!isCompact && <span>VBA</span>}
            </Tab>
          </Tooltip>
          <Tooltip content="Công cụ Excel" relationship="label" positioning="below">
            <Tab value="tools" icon={<Wrench24Regular />}>
              {!isCompact && <span>Công cụ</span>}
            </Tab>
          </Tooltip>
          <Tooltip content="Mẫu Excel" relationship="label" positioning="below">
            <Tab value="templates" icon={<BookOpen24Regular />}>
              {!isCompact && <span>Mẫu</span>}
            </Tab>
          </Tooltip>
          <Tooltip content="Lịch sử yêu cầu" relationship="label" positioning="below">
            <Tab value="history" icon={<History24Regular />}>
              {!isCompact && <span>Lịch sử</span>}
            </Tab>
          </Tooltip>
        </TabList>
      </div>

      {/* Content */}
      <div className="content-container">
        {selectedTab === "formula" && (
          <FormulaGenerator
            disabled={user?.role !== "admin" && credits?.plan !== "pro" && credits?.credits <= 0}
            onRequestComplete={refreshCredits}
          />
        )}
        {selectedTab === "analyzer" && (
          <DataAnalyzer
            disabled={user?.role !== "admin" && credits?.plan !== "pro" && credits?.credits <= 0}
            onRequestComplete={refreshCredits}
          />
        )}
        {selectedTab === "stepbystep" && (
          <StepByStepGuide
            disabled={user?.role !== "admin" && credits?.plan !== "pro" && credits?.credits <= 0}
            onRequestComplete={refreshCredits}
          />
        )}
        {selectedTab === "vba" && (
          <VBAGenerator
            disabled={user?.role !== "admin" && credits?.plan !== "pro" && credits?.credits <= 0}
            onRequestComplete={refreshCredits}
          />
        )}
        {selectedTab === "tools" && <ExcelTools />}
        {selectedTab === "templates" && <TemplateGallery />}
        {selectedTab === "history" && <History />}
      </div>

      {/* Upgrade Pro Dialog */}
      <Dialog open={showUpgradeDialog} onOpenChange={(e, data) => setShowUpgradeDialog(data.open)}>
        <DialogSurface style={{ maxWidth: "480px" }}>
          <DialogBody>
            <UpgradePro onClose={() => setShowUpgradeDialog(false)} currentPlan={credits?.plan} />
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
};

App.propTypes = {
  title: PropTypes.string,
};

export default App;
