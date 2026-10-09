import React from "react";
import { createRoot } from "react-dom/client";
import * as microsoftTeams from "@microsoft/teams-js";
import { App } from "./App.jsx";
import "./styles.css";

async function start() {
  await microsoftTeams.app.initialize();
  const token = await microsoftTeams.authentication.getAuthToken();
  createRoot(document.getElementById("root")).render(
    <React.StrictMode>
      <App token={token} teams={microsoftTeams} />
    </React.StrictMode>
  );
}

start().catch((error) => {
  createRoot(document.getElementById("root")).render(
    <main className="page">
      <h1>Power Platform COE</h1>
      <p role="alert">
        Teams sign-in failed. Open this tab inside Teams and verify the app SSO
        configuration. {error.message}
      </p>
    </main>
  );
});
