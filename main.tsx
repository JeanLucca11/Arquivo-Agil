import { createRoot } from "react-dom/client";
import Home from "./app/page";
import "./app/globals.css";
const root = document.getElementById("root");
if (!root) throw new Error("Não foi encontrado o ponto de entrada da aplicação.");
createRoot(root).render(<Home />);
