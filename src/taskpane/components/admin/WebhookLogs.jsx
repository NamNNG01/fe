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
  Button,
  Badge,
  Spinner,
  Accordion,
  AccordionItem,
  AccordionHeader,
  AccordionPanel,
} from "@fluentui/react-components";
import { ArrowSync24Regular, Code24Regular } from "@fluentui/react-icons";
import { getWebhookLogs } from "../../../services/apiService";

const WebhookLogs = () => {
  const [logs, setLogs] = useState([]);
  const [loading, setLoading] = useState(true);
  const [page, setPage] = useState(1);

  useEffect(() => {
    fetchLogs();
  }, [page]);

  const fetchLogs = async () => {
    try {
      setLoading(true);
      const res = await getWebhookLogs(page, 50);
      setLogs(res.logs || []);
    } catch (err) {
      console.error("Failed to fetch webhook logs:", err);
    } finally {
      setLoading(false);
    }
  };

  const getProcessingBadge = (status) => {
    switch (status) {
      case "processed":
        return (
          <Badge appearance="filled" color="success">
            Processed
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
    <div className="admin-logs">
      <div className="table-controls">
        <Button icon={<ArrowSync24Regular />} onClick={fetchLogs}>
          Refresh Logs
        </Button>
      </div>

      <div className="logs-container">
        {loading ? (
          <div style={{ textAlign: "center", padding: "40px" }}>
            <Spinner label="Đang tải logs..." />
          </div>
        ) : logs.length === 0 ? (
          <div style={{ textAlign: "center", padding: "40px" }}>Chưa có log nào</div>
        ) : (
          <Accordion multiple collapsible>
            {logs.map((log) => (
              <AccordionItem key={log._id} value={log._id}>
                <AccordionHeader expandIconPosition="end">
                  <div
                    style={{ display: "flex", gap: "20px", width: "100%", alignItems: "center" }}
                  >
                    <Text weight="semibold" style={{ minWidth: "160px" }}>
                      {new Date(log.createdAt).toLocaleString("vi-VN")}
                    </Text>
                    <div style={{ flex: 1 }}>
                      {getProcessingBadge(log.processingStatus)}
                      <Badge appearance="tint" style={{ marginLeft: "8px" }}>
                        {log.signatureStatus === "verified"
                          ? "✅ Signature"
                          : "⚠️ Sig " + log.signatureStatus}
                      </Badge>
                    </div>
                    <Text appearance="subtle" size={100}>
                      {log.responseTime}ms
                    </Text>
                  </div>
                </AccordionHeader>
                <AccordionPanel>
                  <div className="log-detail">
                    <div className="log-block">
                      <Text weight="bold" block>
                        Body (Payload):
                      </Text>
                      <pre className="code-block">{JSON.stringify(log.body, null, 2)}</pre>
                    </div>
                    {log.results && (
                      <div className="log-block">
                        <Text weight="bold" block>
                          Results:
                        </Text>
                        <pre className="code-block">{JSON.stringify(log.results, null, 2)}</pre>
                      </div>
                    )}
                    {log.error && (
                      <div className="log-block error">
                        <Text weight="bold" block>
                          Error:
                        </Text>
                        <Text font="monospace" color="danger">
                          {log.error}
                        </Text>
                      </div>
                    )}
                  </div>
                </AccordionPanel>
              </AccordionItem>
            ))}
          </Accordion>
        )}
      </div>
    </div>
  );
};

export default WebhookLogs;
