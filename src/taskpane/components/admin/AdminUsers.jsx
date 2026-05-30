import * as React from "react";
import { useState, useEffect } from "react";
import {
  Table,
  TableHeader,
  TableRow,
  TableHeaderCell,
  TableBody,
  TableCell,
  Text,
  Input,
  Dropdown,
  Option,
  Badge,
  Button,
  Spinner,
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Label,
} from "@fluentui/react-components";
import {
  Search24Regular,
  ArrowSync24Regular,
  ArrowUp24Regular,
  ChevronLeft24Regular,
  ChevronRight24Regular,
} from "@fluentui/react-icons";
import { getAdminUsers, adminUpgradeUser } from "../../../services/apiService";

const AdminUsers = () => {
  const [users, setUsers] = useState([]);
  const [loading, setLoading] = useState(true);
  const [page, setPage] = useState(1);
  const [totalPages, setTotalPages] = useState(1);
  const [totalUsers, setTotalUsers] = useState(0);
  const [search, setSearch] = useState("");
  const [plan, setPlan] = useState("all");

  // Upgrade dialog
  const [showUpgradeDialog, setShowUpgradeDialog] = useState(false);
  const [selectedUser, setSelectedUser] = useState(null);
  const [upgradePlan, setUpgradePlan] = useState("pro_monthly");
  const [upgradeLoading, setUpgradeLoading] = useState(false);

  useEffect(() => {
    fetchUsers();
  }, [page, plan]);

  const fetchUsers = async (overridePage) => {
    try {
      setLoading(true);
      const queryPlan = plan === "all" ? "" : plan;
      const usePage = overridePage || page;
      const res = await getAdminUsers(usePage, 20, search, queryPlan);
      setUsers(res.users || []);
      setTotalPages(res.pagination?.pages || 1);
      setTotalUsers(res.pagination?.total || 0);
    } catch (err) {
      console.error("Failed to fetch users:", err);
    } finally {
      setLoading(false);
    }
  };

  const handleUpgrade = async () => {
    if (!selectedUser) return;

    try {
      setUpgradeLoading(true);
      await adminUpgradeUser(selectedUser._id, upgradePlan);
      setShowUpgradeDialog(false);
      setSelectedUser(null);
      setUpgradePlan("pro_monthly");
      fetchUsers();
      alert(`Đã nâng cấp ${selectedUser.email} lên Pro!`);
    } catch (err) {
      console.error("Upgrade failed:", err);
      alert(err.message || "Lỗi khi nâng cấp user");
    } finally {
      setUpgradeLoading(false);
    }
  };

  const getPlanBadge = (plan) => {
    switch (plan) {
      case "pro":
        return (
          <Badge appearance="filled" color="brand">
            PRO
          </Badge>
        );
      default:
        return <Badge appearance="outline">FREE</Badge>;
    }
  };

  return (
    <div className="admin-users">
      <div className="table-controls">
        <Input
          placeholder="Tìm theo email hoặc tên..."
          contentBefore={<Search24Regular />}
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter") {
              setPage(1);
              fetchUsers(1);
            }
          }}
        />
        <Dropdown
          placeholder="Gói dịch vụ"
          value={plan}
          onOptionSelect={(e, data) => {
            setPlan(data.optionValue);
            setPage(1);
          }}
        >
          <Option value="all">Tất cả</Option>
          <Option value="free">FREE</Option>
          <Option value="pro">PRO</Option>
        </Dropdown>
        <Button icon={<ArrowSync24Regular />} onClick={fetchUsers} />
      </div>

      <div className="table-scroll-container">
        <Table size="small">
          <TableHeader>
            <TableRow>
              <TableHeaderCell>Tên</TableHeaderCell>
              <TableHeaderCell>Email</TableHeaderCell>
              <TableHeaderCell>Gói</TableHeaderCell>
              <TableHeaderCell>Hết hạn</TableHeaderCell>
              <TableHeaderCell>Trạng thái</TableHeaderCell>
              <TableHeaderCell>Ngày tham gia</TableHeaderCell>
              <TableHeaderCell>Hành động</TableHeaderCell>
            </TableRow>
          </TableHeader>
          <TableBody>
            {loading ? (
              <TableRow>
                <TableCell colspan={7} style={{ textAlign: "center", padding: "40px" }}>
                  <Spinner label="Đang tải người dùng..." />
                </TableCell>
              </TableRow>
            ) : users.length === 0 ? (
              <TableRow>
                <TableCell colspan={7} style={{ textAlign: "center", padding: "40px" }}>
                  Không tìm thấy người dùng nào
                </TableCell>
              </TableRow>
            ) : (
              users.map((user) => (
                <TableRow key={user._id}>
                  <TableCell>{user.name}</TableCell>
                  <TableCell>
                    <div
                      style={{
                        maxWidth: "160px",
                        overflow: "hidden",
                        textOverflow: "ellipsis",
                        whiteSpace: "nowrap",
                      }}
                      title={user.email}
                    >
                      {user.email}
                    </div>
                  </TableCell>
                  <TableCell style={{ minWidth: "55px" }}>
                    {getPlanBadge(user.subscription?.plan)}
                  </TableCell>
                  <TableCell>
                    {user.subscription?.endDate
                      ? new Date(user.subscription.endDate).toLocaleDateString("vi-VN")
                      : "---"}
                  </TableCell>
                  <TableCell>
                    <Badge appearance="tint" color={user.isActive ? "success" : "danger"}>
                      {user.isActive ? "Active" : "Blocked"}
                    </Badge>
                  </TableCell>
                  <TableCell>{new Date(user.createdAt).toLocaleDateString("vi-VN")}</TableCell>
                  <TableCell>
                    {user.subscription?.plan !== "pro" && (
                      <Button
                        size="small"
                        appearance="primary"
                        icon={<ArrowUp24Regular />}
                        onClick={() => {
                          setSelectedUser(user);
                          setUpgradePlan("pro_monthly");
                          setShowUpgradeDialog(true);
                        }}
                      >
                        Upgrade
                      </Button>
                    )}
                    {user.subscription?.plan === "pro" &&
                      user.subscription?.endDate &&
                      new Date(user.subscription.endDate) < new Date() && (
                        <Button
                          size="small"
                          appearance="outline"
                          icon={<ArrowUp24Regular />}
                          onClick={() => {
                            setSelectedUser(user);
                            setUpgradePlan("pro_monthly");
                            setShowUpgradeDialog(true);
                          }}
                        >
                          Gia hạn
                        </Button>
                      )}
                  </TableCell>
                </TableRow>
              ))
            )}
          </TableBody>
        </Table>
      </div>

      {/* Pagination */}
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          padding: "12px 0",
          borderTop: "1px solid var(--border-color)",
          marginTop: "8px",
        }}
      >
        <Text size={200} style={{ color: "var(--text-secondary)" }}>
          Tổng: {totalUsers} người dùng • Trang {page}/{totalPages}
        </Text>
        <div style={{ display: "flex", gap: "4px" }}>
          <Button
            size="small"
            appearance="subtle"
            icon={<ChevronLeft24Regular />}
            disabled={page <= 1}
            onClick={() => setPage((p) => Math.max(1, p - 1))}
          />
          <Button
            size="small"
            appearance="subtle"
            icon={<ChevronRight24Regular />}
            disabled={page >= totalPages}
            onClick={() => setPage((p) => Math.min(totalPages, p + 1))}
          />
        </div>
      </div>

      {/* Upgrade Pro Dialog */}
      <Dialog open={showUpgradeDialog} onOpenChange={(e, d) => setShowUpgradeDialog(d.open)}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Nâng cấp lên Pro</DialogTitle>
            <DialogContent>
              <div
                style={{ display: "flex", flexDirection: "column", gap: "12px", marginTop: "10px" }}
              >
                <div>
                  <Text weight="semibold">Người dùng: </Text>
                  <Text>
                    {selectedUser?.name} ({selectedUser?.email})
                  </Text>
                </div>

                <Label required>Chọn gói</Label>
                <Dropdown
                  value={upgradePlan === "pro_monthly" ? "Pro Hàng Tháng" : "Pro Hàng Năm"}
                  onOptionSelect={(e, data) => setUpgradePlan(data.optionValue)}
                >
                  <Option value="pro_monthly">Pro Hàng Tháng (+1 tháng)</Option>
                  <Option value="pro_yearly">Pro Hàng Năm (+1 năm)</Option>
                </Dropdown>

                <Text size={200} italic style={{ color: "var(--text-secondary)" }}>
                  * Hệ thống sẽ set plan = "pro", startDate = hôm nay, endDate theo gói đã chọn.
                </Text>
              </div>
            </DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => setShowUpgradeDialog(false)}>
                Hủy
              </Button>
              <Button
                appearance="primary"
                onClick={handleUpgrade}
                disabled={upgradeLoading}
                icon={upgradeLoading ? <Spinner size="tiny" /> : <ArrowUp24Regular />}
              >
                {upgradeLoading ? "Đang xử lý..." : "Xác nhận Upgrade"}
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
};

export default AdminUsers;
