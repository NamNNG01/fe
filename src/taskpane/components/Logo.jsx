import * as React from "react";


const Logo = () => {
  return (
    <div className="logo group">
      <div className="logo__icon-box group-hover-rotate">
        {/* BangAI Logo Icon - Modern Sparkles */}
        <svg
          xmlns="http://www.w3.org/2000/svg"
          width="24"
          height="24"
          viewBox="0 0 24 24"
          fill="none"
          stroke="currentColor"
          strokeWidth="2"
          strokeLinecap="round"
          strokeLinejoin="round"
          className="text-white"
        >
          <path d="m12 3-1.912 5.813a2 2 0 0 1-1.275 1.275L3 12l5.813 1.912a2 2 0 0 1 1.275 1.275L12 21l1.912-5.813a2 2 0 0 1 1.275-1.275L21 12l-5.813-1.912a2 2 0 0 1-1.275-1.275L12 3Z" />
          <path d="M5 3v4" />
          <path d="M7 5H3" />
          <path d="M21 17v4" />
          <path d="M23 19h-4" />
        </svg>
      </div>
      <div className="logo__text">
        Bang<span>AI</span>
      </div>
    </div>
  );
};

export default Logo;