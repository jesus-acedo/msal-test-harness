import { createRoot } from "react-dom/client";
import { App } from "./test-harness/App";

const container = document.getElementById("msal-test-harness");
const root = createRoot(container);
root.render(<App />);
