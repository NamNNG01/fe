import * as React from "react";
import { Card, Text, CardHeader } from "@fluentui/react-components";
import {
  Money24Filled,
  People24Filled,
  Warning24Filled,
  CheckmarkCircle24Filled,
} from "@fluentui/react-icons";

const StatCard = ({ title, value, icon, color, description }) => (
  <Card className="stat-card">
    <CardHeader
      header={
        <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
          <div style={{ color: color }}>{icon}</div>
          <Text weight="semibold">{title}</Text>
        </div>
      }
    />
    <div style={{ padding: "0 12px 12px" }}>
      <Text size={700} weight="bold">
        {value}
      </Text>
      <div style={{ marginTop: "4px" }}>
        <Text size={200} appearance="subtle">
          {description}
        </Text>
      </div>
    </div>
  </Card>
);

const AdminOverview = ({ stats, onRefresh }) => {
  if (!stats) return null;

  return (
    <div className="admin-overview">
      <div className="admin-section-header">
        <Text size={400} weight="semibold">
          Hệ thống đang hoạt động ổn định
        </Text>
      </div>

      <div className="stats-grid">
        <StatCard
          title="Doanh thu (VND)"
          value={stats.revenue?.total?.toLocaleString() || "0"}
          icon={<Money24Filled />}
          color="#28a745"
          description={`Hôm nay: +${stats.revenue?.today?.toLocaleString()}`}
        />
        <StatCard
          title="Người dùng Pro"
          value={stats.users?.pro || "0"}
          icon={<People24Filled />}
          color="#0078d4"
          description={`Tổng cộng: ${stats.users?.total}`}
        />
        <StatCard
          title="Giao dịch chưa khớp"
          value={stats.alerts?.unmatched || "0"}
          icon={<Warning24Filled />}
          color={stats.alerts?.unmatched > 0 ? "#d13438" : "#999"}
          description="Cần xử lý thủ công ngay"
        />
        <StatCard
          title="Tỷ lệ thành công"
          value="98.5%"
          icon={<CheckmarkCircle24Filled />}
          color="#107c10"
          description="Dựa trên 30 ngày qua"
        />
      </div>

      {/* Gợi ý hành động */}
      {stats.alerts?.unmatched > 0 && (
        <div className="admin-alert-banner">
          <Warning24Filled />
          <Text>
            Có {stats.alerts.unmatched} giao dịch chưa được đối soát. Vui lòng kiểm tra tab Giao
            dịch.
          </Text>
        </div>
      )}
    </div>
  );
};

export default AdminOverview;
