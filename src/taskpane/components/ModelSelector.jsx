/**
 * ModelSelector Component - Premium AI Model Selection
 *
 * Uses React Portal to render dropdown OUTSIDE the normal DOM flow
 * This prevents clipping by parent containers
 */

import * as React from "react";
import { useState, useEffect, useRef } from "react";
import { createPortal } from "react-dom";

// Models: Nhanh (2.5), Cân bằng (3.0 Flash), Pro (3.0 Pro)
const AI_MODELS = [
  {
    id: "gemini-2.5-flash",
    name: "Nhanh",
    description: "Phản hồi tức thì",
    isDefault: true,
  },
  {
    id: "gemini-3-flash-preview",
    name: "Cân bằng",
    description: "Tốc độ & chất lượng",
    isDefault: false,
  },
  {
    id: "gemini-3-pro-preview",
    name: "Pro",
    description: "Phân tích sâu",
    isDefault: false,
  },
];

const STORAGE_KEY = "eoffice_ai_model";

const ModelSelector = ({ onModelChange }) => {
  const [selectedModel, setSelectedModel] = useState(() => {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved && AI_MODELS.find((m) => m.id === saved)) {
      return saved;
    }
    return AI_MODELS.find((m) => m.isDefault)?.id || AI_MODELS[0].id;
  });
  const [isOpen, setIsOpen] = useState(false);
  const [menuPosition, setMenuPosition] = useState({ top: 0, left: 0 });
  const triggerRef = useRef(null);

  const currentModel = AI_MODELS.find((m) => m.id === selectedModel);

  useEffect(() => {
    if (onModelChange) {
      onModelChange(selectedModel);
    }
  }, [selectedModel]);

  // Calculate menu position when opening
  useEffect(() => {
    if (isOpen && triggerRef.current) {
      const rect = triggerRef.current.getBoundingClientRect();
      setMenuPosition({
        top: rect.bottom + 4,
        left: rect.right - 160, // Align right edge
      });
    }
  }, [isOpen]);

  // Close on click outside
  useEffect(() => {
    const handleClickOutside = (e) => {
      if (triggerRef.current && !triggerRef.current.contains(e.target)) {
        // Check if click is inside the portal menu
        const menu = document.getElementById("model-selector-portal");
        if (menu && !menu.contains(e.target)) {
          setIsOpen(false);
        }
      }
    };

    const handleScroll = () => setIsOpen(false);

    document.addEventListener("mousedown", handleClickOutside);
    window.addEventListener("scroll", handleScroll, true);

    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
      window.removeEventListener("scroll", handleScroll, true);
    };
  }, []);

  const handleSelect = (modelId) => {
    setSelectedModel(modelId);
    localStorage.setItem(STORAGE_KEY, modelId);
    setIsOpen(false);
  };

  // Portal menu component
  const DropdownMenu = () => {
    if (!isOpen) return null;

    return createPortal(
      <div
        id="model-selector-portal"
        style={{
          position: "fixed",
          top: menuPosition.top,
          left: menuPosition.left,
          width: "160px",
          background: "white",
          borderRadius: "8px",
          boxShadow: "0 8px 30px rgba(0,0,0,0.15), 0 0 1px rgba(0,0,0,0.1)",
          border: "1px solid #e5e7eb",
          zIndex: 99999, // Very high z-index
          overflow: "hidden",
          animation: "dropdownFadeIn 0.15s ease",
        }}
      >
        {AI_MODELS.map((model, idx) => {
          const isSelected = selectedModel === model.id;
          return (
            <div
              key={model.id}
              onClick={() => handleSelect(model.id)}
              style={{
                padding: "10px 12px",
                cursor: "pointer",
                background: isSelected ? "#f0fdf4" : "transparent",
                borderBottom: idx < AI_MODELS.length - 1 ? "1px solid #f3f4f6" : "none",
                transition: "background 0.1s ease",
              }}
              onMouseEnter={(e) => {
                if (!isSelected) e.currentTarget.style.background = "#f9fafb";
              }}
              onMouseLeave={(e) => {
                e.currentTarget.style.background = isSelected ? "#f0fdf4" : "transparent";
              }}
            >
              <div
                style={{
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                }}
              >
                <span
                  style={{
                    fontWeight: isSelected ? 600 : 500,
                    fontSize: "13px",
                    color: isSelected ? "#059669" : "#374151",
                  }}
                >
                  {model.name}
                </span>
                {isSelected && (
                  <svg width="14" height="14" viewBox="0 0 14 14" fill="none">
                    <path
                      d="M3 7L6 10L11 4"
                      stroke="#059669"
                      strokeWidth="2"
                      strokeLinecap="round"
                    />
                  </svg>
                )}
              </div>
              <div
                style={{
                  fontSize: "11px",
                  color: "#9ca3af",
                  marginTop: "2px",
                }}
              >
                {model.description}
              </div>
            </div>
          );
        })}
      </div>,
      document.body
    );
  };

  return (
    <>
      {/* Trigger Button */}
      <button
        ref={triggerRef}
        onClick={() => setIsOpen(!isOpen)}
        style={{
          display: "flex",
          alignItems: "center",
          gap: "4px",
          padding: "6px 10px",
          background: isOpen ? "#f3f4f6" : "transparent",
          border: "1px solid #e5e7eb",
          borderRadius: "6px",
          cursor: "pointer",
          fontSize: "13px",
          fontWeight: 500,
          color: "#374151",
          transition: "all 0.15s ease",
          whiteSpace: "nowrap",
          width: "90px",
          justifyContent: "center",
        }}
      >
        {currentModel?.name}
        <svg
          width="12"
          height="12"
          viewBox="0 0 12 12"
          fill="none"
          style={{
            transition: "transform 0.2s ease",
            transform: isOpen ? "rotate(180deg)" : "rotate(0deg)",
          }}
        >
          <path d="M3 4.5L6 7.5L9 4.5" stroke="#6b7280" strokeWidth="1.5" strokeLinecap="round" />
        </svg>
      </button>

      {/* Portal Menu - rendered to document.body */}
      <DropdownMenu />

      {/* Animation keyframes */}
      <style>{`
        @keyframes dropdownFadeIn {
          from { 
            opacity: 0; 
            transform: translateY(-4px); 
          }
          to { 
            opacity: 1; 
            transform: translateY(0); 
          }
        }
      `}</style>
    </>
  );
};

export default ModelSelector;
export { AI_MODELS, STORAGE_KEY };
