import React from "react";
import { useMsal } from "@azure/msal-react";
import "../styles/Login.css";

const Login: React.FC = () => {
  const { instance } = useMsal();

  const handleAzureLogin = async () => {
    await instance.loginRedirect({
      scopes: ["openid", "profile", "email"],
      prompt: "select_account", // ✅ user selects account
    });
  };

  return (
    <div className="login-container">
      <div className="login-box">
        <h2>Login</h2>

        <button className="azure-btn" onClick={handleAzureLogin}>
          Login with Microsoft
        </button>
      </div>
    </div>
  );
};

export default Login;

