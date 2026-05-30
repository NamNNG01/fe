/**
 * ApiKeySetup Component - API Key Configuration
 *
 * REFACTORED:
 * - Loại bỏ makeStyles, inline styles → CSS classes
 * - Sử dụng apiService thay vì geminiService trực tiếp
 * - Frontend CHỈ handle UI state, logic validation ở backend
 */

import * as React from "react";
import { useState, useEffect } from "react";
import { Button, Field, Input, Card, Text } from "@fluentui/react-components";
import {
  Key24Regular,
  CheckmarkCircle24Regular,
  Dismiss24Regular,
  Info24Regular,
} from "@fluentui/react-icons";

// API Service
import { saveApiKey, clearApiKey, hasApiKey, getApiKeyMasked } from "../../services/apiService";

const ApiKeySetup = ({ onKeySet }) => {
  const [apiKey, setApiKey] = useState("");
  const [hasKey, setHasKey] = useState(false);
  const [maskedKey, setMaskedKey] = useState("");
  const [message, setMessage] = useState("");

  useEffect(() => {
    checkApiKey();
  }, []);

  const checkApiKey = () => {
    const keyExists = hasApiKey();
    setHasKey(keyExists);
    if (keyExists) {
      setMaskedKey(getApiKeyMasked());
    }
  };

  const handleSaveKey = async () => {
    try {
      // TODO BACKEND: Validation sẽ được handle ở backend API
      await saveApiKey(apiKey);
      setMessage("✅ API key đã được lưu!");
      setApiKey("");
      checkApiKey();

      // Notify parent component
      if (onKeySet) {
        onKeySet();
      }

      setTimeout(() => setMessage(""), 3000);
    } catch (error) {
      setMessage("❌ " + error.message);
    }
  };

  const handleClearKey = async () => {
    try {
      await clearApiKey();
      setHasKey(false);
      setMaskedKey("");
      setMessage("API key đã được xóa");
      setTimeout(() => setMessage(""), 3000);
    } catch (error) {
      setMessage("❌ " + error.message);
    }
  };

  return (
    <div className="api-key-container">
      <Card className="card">
        <div className="api-key-header">
          <Key24Regular className="api-key-header__icon" />
          <h2 className="api-key-header__title">Gemini API Key Setup</h2>
        </div>

        {hasKey ? (
          <div className="api-key-success-box">
            <CheckmarkCircle24Regular style={{ fontSize: "24px", color: "#10b981" }} />
            <div className="api-key-success-box__content">
              <Text weight="semibold" className="api-key-success-box__title d-block mb-4">
                API Key đã được cấu hình
              </Text>
              <Text size={200} className="api-key-success-box__key">
                {maskedKey}
              </Text>
            </div>
            <Button appearance="subtle" icon={<Dismiss24Regular />} onClick={handleClearKey}>
              Xóa
            </Button>
          </div>
        ) : (
          <>
            <Field
              label="Nhập Gemini API Key của bạn"
              className="form-field"
              hint="Key sẽ được lưu an toàn trong trình duyệt"
            >
              <Input
                type="password"
                value={apiKey}
                onChange={(e) => setApiKey(e.target.value)}
                placeholder="AIzaSy..."
                contentAfter={
                  apiKey && (
                    <Dismiss24Regular onClick={() => setApiKey("")} style={{ cursor: "pointer" }} />
                  )
                }
              />
            </Field>

            <div className="button-group">
              <Button
                appearance="primary"
                icon={<CheckmarkCircle24Regular />}
                onClick={handleSaveKey}
                disabled={!apiKey.trim()}
                className="btn-primary"
              >
                Lưu API Key
              </Button>
            </div>
          </>
        )}

        {message && (
          <div className="mt-16">
            <Text>{message}</Text>
          </div>
        )}
      </Card>

      <div className="api-key-info-box">
        <div className="api-key-info-title">
          <Info24Regular />
          <span>Làm sao để lấy API Key?</span>
        </div>
        <Text className="api-key-info-text">
          1. Truy cập{" "}
          <a
            href="https://aistudio.google.com/app/apikey"
            target="_blank"
            rel="noopener noreferrer"
            className="api-key-link"
          >
            Google AI Studio
          </a>
        </Text>
        <Text className="api-key-info-text">2. Đăng nhập bằng tài khoản Google</Text>
        <Text className="api-key-info-text">3. Click "Create API Key" hoặc "Get API Key"</Text>
        <Text className="api-key-info-text">4. Copy key và paste vào ô bên trên</Text>
        <Text className="api-key-info-text api-key-info-highlight">
          💡 API key miễn phí với 15 requests/phút
        </Text>
      </div>
    </div>
  );
};

export default ApiKeySetup;
