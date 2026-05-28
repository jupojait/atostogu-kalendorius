import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { PublicClientApplication } from "@azure/msal-browser";
import { MsalProvider } from "@azure/msal-react";

import "./index.css";
import App from "./App.jsx";

const msalInstance = new PublicClientApplication({
  auth: {
    clientId: "2b5e1bf6-1a2d-4658-899c-05f9dce12378",
    authority: "https://login.microsoftonline.com/1e98a3e5-7ec4-427a-ae35-10b2b9eade6a",
    redirectUri: "http://localhost:5173"
  },
  cache: {
    cacheLocation: "localStorage"
  }
});

await msalInstance.initialize();

createRoot(document.getElementById("root")).render(
  <StrictMode>
    <MsalProvider instance={msalInstance}>
      <App />
    </MsalProvider>
  </StrictMode>
);
redirectUri: window.location.origin