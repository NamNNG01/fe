import * as React from "react";
import logoImg from "../../images/BangAIlogo.jpg";

const Logo = ({ size = 32 }) => {
  return (
    <div className="logo group">
      <div className="logo__icon-box group-hover-rotate" style={{ width: size + 12, height: size + 12, padding: 0, overflow: "hidden" }}>
        {/* BangAI Logo Official Image */}
        <img
          src={logoImg}
          alt="BangAI Logo"
          style={{
            width: "100%",
            height: "100%",
            objectFit: "cover",
            borderRadius: "10px"
          }}
        />
      </div>
      <div className="logo__text">
        Bang<span>AI</span>
      </div>
    </div>
  );
};

export default Logo;
