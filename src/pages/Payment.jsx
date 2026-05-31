import { useEffect, useState } from "react";
import axios from "axios";
import { useParams } from "react-router-dom";

export default function Payment() {
  const { id } = useParams();

  const [intent, setIntent] = useState(null);
  const [status, setStatus] = useState("pending");
  const [loading, setLoading] = useState(true);

  const fetchStatus = async () => {
    try {
      const res = await axios.get(
        `https://web-production-8b555d.up.railway.app/api/v1/payments/intents/${id}`,
        {
          headers: {
            Authorization: `Bearer ${localStorage.getItem("token")}`,
          },
        }
      );

      setIntent(res.data);
      setStatus(res.data.status);
      setLoading(false);

      if (res.data.status === "paid") {
        clearInterval(window.paymentInterval);
      }
    } catch (err) {
      console.log(err);
    }
  };

  useEffect(() => {
    fetchStatus();

    window.paymentInterval = setInterval(fetchStatus, 3000);

    return () => clearInterval(window.paymentInterval);
  }, [id]);

  if (loading) return <div>Loading...</div>;

  if (!intent) return <div>Không tìm thấy payment</div>;

  return (
    <div style={{ textAlign: "center", maxWidth: 500, margin: "auto" }}>
      <h2>Thanh toán</h2>

      <img src={intent.qrData.qrCodeUrl} alt="QR" style={{ width: 250 }} />

      <p>
        <b>Số tiền:</b> {intent.amount.toLocaleString()} VND
      </p>

      <p>
        <b>Mã chuyển khoản:</b> {intent.transferCode}
      </p>

      <p>
        <b>Trạng thái:</b> {status}
      </p>

      {status === "paid" && <h3 style={{ color: "green" }}>Thanh toán thành công 🎉</h3>}
    </div>
  );
}
