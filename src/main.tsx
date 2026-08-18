import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import App from "./App";
import { lockDown } from "./lockdown";
import "./styles.css";

function mount() {
  if (import.meta.env.PROD) lockDown();
  createRoot(document.getElementById("root")!).render(
    <StrictMode>
      <App />
    </StrictMode>
  );
}

// Office.onReady never fires in a plain browser tab, so fall back after a short wait
// to keep `npm run dev` usable outside Excel.
if (typeof Office !== "undefined" && Office.onReady) {
  Office.onReady().then(mount);
} else {
  mount();
}
