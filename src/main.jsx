import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import SOWParser from "./SOWParser.jsx";

createRoot(document.getElementById("root")).render(
  <StrictMode>
    <SOWParser />
  </StrictMode>
);
