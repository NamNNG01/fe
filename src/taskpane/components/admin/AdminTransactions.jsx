import * as React from "react";
import { useState, useEffect } from "react";
import {
  Table,
  TableHeader,
  TableRow,
  TableHeaderCell,
  TableBody,
  TableCell,
  TableSelectionCell,
  Text,
  Button,
  Input,
  Dropdown,
  Option,
  Badge,
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Label,
  Spinner,
} from "@fluentui/react-components";
import {
  Search24Regular,
  ArrowSync24Regular,
  CheckmarkCircle24Regular,
  Warning24Regular,
} from "@fluentui/react-icons";
import { getAdminTransactions, manualMatchTransaction } from "../../../services/apiService";

const AdminTransactions = () => {
  const [data, setData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [page, setPage] = useState(1);
  const [search, setSearch] = useState("");
  const [status, setStatus] = useState("all");
  const [showMatchDialog, setShowMatchDialog] = useState(false);
  const [selectedTx, setSelectedTx] = useState(null);
  const [manualIntentId, setManualIntentId] = useState("");
  const [matchLoading, setMatchLoading] = useState(false);

  useEffect(() => {
    fetchTransactions();
  }, [page, status]);

  const fetchTransactions = async () => {
    try {
      setLoading(true);
      const queryStatus = status === "all" ? "" : status;
      const res = await getAdminTransactions(page, 20, search, queryStatus);
      setData(res.transactions);
    } catch (err) {
      console.error("Failed to fetch transactions:", err);
    } finally {
      setLoading(false);
    }
  };

  const handleManualMatch = async () => {
    if (!manualIntentId || !selectedTx) return;

    try {
      setMatchLoading(true);
      await manualMatchTransaction(selectedTx._id, manualIntentId);
      setShowMatchDialog(false);
      fetchTransactions(); // Refresh list
      alert("Khớp giao dịch thành công!");
    } catch (err) {
      console.error("Manual match failed:", err);
      alert(err.message || "Lỗi khi khớp giao dịch");
    } finally {
      setMatchLoading(false);
    }
  };

  const getStatusBadge = (status) => {
    switch (status) {
      case "matched":
        return (
          <Badge appearance="filled" color="success">
            Paid
          </Badge>
        );
      case "unmatched":
        return (
          <Badge appearance="filled" color="warning">
            Unmatched
          </Badge>
        );
      case "failed":
        return (
          <Badge appearance="filled" color="danger">
            Failed
          </Badge>
        );
      default:
        return <Badge appearance="outline">{status}</Badge>;
    }
  };

  return (
    <div className="admin-transactions">
      <div className="table-controls">
        <Input
          placeholder="Tìm mã transferCode hoặc nội dung..."
          contentBefore={<Search24Regular />}
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && fetchTransactions()}
        />
        <Dropdown
          placeholder="Trạng thái"
          value={status}
          onOptionSelect={(e, data) => setStatus(data.optionValue)}
        >
          <Option value="all">Tất cả</Option>
          <Option value="matched">Paid</Option>
          <Option value="unmatched">Unmatched</Option>
        </Dropdown>
        <Button icon={<ArrowSync24Regular />} onClick={fetchTransactions} />
      </div>

      <div className="table-scroll-container">
        <Table size="small">
          <TableHeader>
            <TableRow>
              <TableHeaderCell>Mã giao dịch</TableHeaderCell>
              <TableHeaderCell>Thời gian</TableHeaderCell>
              <TableHeaderCell>Số tiền</TableHeaderCell>
              <TableHeaderCell>Nội dung</TableHeaderCell>
              <TableHeaderCell>Mã match</TableHeaderCell>
              <TableHeaderCell>Trạng thái</TableHeaderCell>
              <TableHeaderCell>Hành động</TableHeaderCell>
            </TableRow>
          </TableHeader>
          <TableBody>
            {loading ? (
              <TableRow>
                <TableCell colspan={7} style={{ textAlign: "center", padding: "40px" }}>
                  <Spinner label="Đang tải giao dịch..." />
                </TableCell>
              </TableRow>
            ) : data.length === 0 ? (
              <TableRow>
                <TableCell colspan={7} style={{ textAlign: "center", padding: "40px" }}>
                  Không có giao dịch nào
                </TableCell>
              </TableRow>
            ) : (
              data.map((tx) => (
                <TableRow key={tx.providerTxId}>
                  <TableCell title={tx.providerTxId}>
                    {tx.providerTxId.substring(0, 8)}...
                  </TableCell>
                  <TableCell>{new Date(tx.createdAt).toLocaleString("vi-VN")}</TableCell>
                  <TableCell weight="bold">{tx.amount?.toLocaleString()}đ</TableCell>
                  <TableCell title={tx.description}>
                    <div
                      style={{
                        maxWidth: "150px",
                        overflow: "hidden",
                        textOverflow: "ellipsis",
                        whiteSpace: "nowrap",
                      }}
                    >
                      {tx.description}
                    </div>
                  </TableCell>
                  <TableCell>{tx.transferCode || "---"}</TableCell>
                  <TableCell>{getStatusBadge(tx.status)}</TableCell>
                  <TableCell>
                    {tx.status === "unmatched" && (
                      <Button
                        size="small"
                        appearance="primary"
                        onClick={() => {
                          setSelectedTx(tx);
                          setShowMatchDialog(true);
                        }}
                      >
                        Match
                      </Button>
                    )}
                  </TableCell>
                </TableRow>
              ))
            )}
          </TableBody>
        </Table>
      </div>

      {/* Manual Match Dialog */}
      <Dialog open={showMatchDialog} onOpenChange={(e, d) => setShowMatchDialog(d.open)}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Đối soát giao dịch thủ công</DialogTitle>
            <DialogContent>
              <div
                style={{ display: "flex", flexDirection: "column", gap: "12px", marginTop: "10px" }}
              >
                <div className="tx-details-mini">
                  <Text>
                    Giao dịch: <b>{selectedTx?.amount?.toLocaleString()}đ</b>
                  </Text>
                  <Text size={200} block>
                    Nội dung: {selectedTx?.description}
                  </Text>
                </div>

                <Label required>Nhập Payment Intent ID hoặc mã TransferCode</Label>
                <Input
                  placeholder="Ví dụ: 65a... hoặc EOAI-ABC123"
                  value={manualIntentId}
                  onChange={(e) => setManualIntentId(e.target.value)}
                />
                <Text size={200} italic appearance="subtle">
                  * Hệ thống sẽ tìm Intent tương ứng và nâng cấp User sang gói Pro.
                </Text>
              </div>
            </DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => setShowMatchDialog(false)}>
                Hủy
              </Button>
              <Button
                appearance="primary"
                onClick={handleManualMatch}
                disabled={!manualIntentId || matchLoading}
                icon={matchLoading ? <Spinner size="tiny" /> : <CheckmarkCircle24Regular />}
              >
                Match & Upgrade
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
};

export default AdminTransactions;
