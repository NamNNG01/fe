import * as React from "react";
import { useState, useEffect } from "react";
import {
  Tab,
  TabList,
  Text,
  Button,
  Spinner,
  MessageBar,
  MessageBarBody,
  MessageBarTitle,
} from "@fluentui/react-components";
import {
  Board24Regular,
  Receipt24Regular,
  People24Regular,
  DocumentText24Regular,
  ArrowLeft24Regular,
} from "@fluentui/react-icons";
import AdminOverview from "./AdminOverview";
import AdminTransactions from "./AdminTransactions";
import AdminUsers from "./AdminUsers";
import WebhookLogs from "./WebhookLogs";
import { getAdminStats } from "../../../services/apiService";

const AdminDashboard = ({ onClose }) => {
  const [selectedTab, setSelectedTab] = useState("overview");
  const [stats, setStats] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);

  useEffect(() => {
    fetchStats();
  }, []);

  const fetchStats = async () => {
    try {
      setLoading(true);
      const data = await getAdminStats();
      setStats(data);
      setError(null);
    } catch (err) {
      console.error("Failed to fetch admin stats:", err);
      setError("Không thể tải thông tin thống kê. Kiểm tra quyền Admin.");
    } finally {
      setLoading(false);
    }
  };

  const handleTabChange = (event, data) => {
    setSelectedTab(data.value);
  };

  if (loading) {
    return (
      <div className="admin-container loading">
        <Spinner label="Đang tải dữ liệu quản trị..." />
      </div>
    );
  }

  return (
    <div className="admin-container">
      <div className="admin-header">
        <Button appearance="subtle" icon={<ArrowLeft24Regular />} onClick={onClose}>
          Quay lại
        </Button>
        <Text size={500} weight="semibold">
          Admin Control Center
        </Text>
      </div>

      {error && (
        <MessageBar intent="error">
          <MessageBarBody>
            <MessageBarTitle>Lỗi</MessageBarTitle>
            {error}
          </MessageBarBody>
        </MessageBar>
      )}

      <div className="admin-tabs">
        <TabList selectedValue={selectedTab} onTabSelect={handleTabChange} vertical>
          <Tab value="overview" icon={<Board24Regular />}>
            Tổng quan
          </Tab>
          <Tab value="transactions" icon={<Receipt24Regular />}>
            Giao dịch
          </Tab>
          <Tab value="users" icon={<People24Regular />}>
            Người dùng
          </Tab>
          <Tab value="logs" icon={<DocumentText24Regular />}>
            Webhook Logs
          </Tab>
        </TabList>

        <div className="admin-content">
          {selectedTab === "overview" && <AdminOverview stats={stats} onRefresh={fetchStats} />}
          {selectedTab === "transactions" && <AdminTransactions />}
          {selectedTab === "users" && <AdminUsers />}
          {selectedTab === "logs" && <WebhookLogs />}
        </div>
      </div>
    </div>
  );
};

export default AdminDashboard;
