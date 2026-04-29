import { Routes, Route, Navigate } from "react-router-dom";
import { useMsal, useIsAuthenticated } from "@azure/msal-react";
import { InteractionStatus } from "@azure/msal-browser";

import Login from "./pages/Login";
import Home from "./pages/Home";
import Data from "./pages/Data";

export default function App() {
  const { inProgress } = useMsal();
  const isAuthenticated = useIsAuthenticated();

  // ✅ WAIT until MSAL finishes redirect/login
  if (
    inProgress === InteractionStatus.Startup ||
    inProgress === InteractionStatus.HandleRedirect
  ) {
    return null; // or loading spinner
  }

  return (
    <Routes>
      {/* LOGIN */}
      <Route
        path="/"
        element={
          isAuthenticated ? <Navigate to="/home" replace /> : <Login />
        }
      />

      {/* HOME */}
      <Route
        path="/home"
        element={
          isAuthenticated ? <Home /> : <Navigate to="/" replace />
        }
      />

      {/* ✅ DATA : ALL COMPANIES */}
      <Route
        path="/data"
        element={
          isAuthenticated ? <Data /> : <Navigate to="/" replace />
        }
      />

      {/* ✅ DATA : COMPANY SPECIFIC */}
      <Route
        path="/data/:company"
        element={
          isAuthenticated ? <Data /> : <Navigate to="/" replace />
        }
      />
    </Routes>
  );
}
