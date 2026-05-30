import * as React from "react";

// Logo icon component - có thể dùng riêng hoặc với text
const LogoIcon = ({ size = 24, color = "#10b981" }) => {
  return (
    <svg 
      width={size} 
      height={size} 
      viewBox="0 0 24 24" 
      fill="none"
      style={{ flexShrink: 0 }}
    >
      {/* Stylized "E" shape representing documents/office layers */}
      <rect x="4" y="4" width="16" height="3" rx="1" fill={color} opacity="0.9"/>
      <rect x="4" y="10" width="12" height="3" rx="1" fill={color} opacity="0.7"/>
      <rect x="4" y="16" width="14" height="3" rx="1" fill={color} opacity="0.5"/>
      
      {/* AI spark indicator */}
      <circle cx="19" cy="6" r="2" fill="#fbbf24" opacity="0.9"/>
      <path 
        d="M19 4L19.5 5.5L21 6L19.5 6.5L19 8L18.5 6.5L17 6L18.5 5.5L19 4Z" 
        fill="white" 
        opacity="0.9"
      />
    </svg>
  );
};

export default LogoIcon;

