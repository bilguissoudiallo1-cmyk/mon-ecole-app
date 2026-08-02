import "./storage.js";
import React from "react";
import { createRoot } from "react-dom/client";
import EcoleApp from "./App.jsx";

const container = document.getElementById("root");
const root = createRoot(container);
root.render(<EcoleApp />);

// Enregistrement du service worker pour le fonctionnement hors-ligne
if ("serviceWorker" in navigator) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("./service-worker.js").catch((err) => {
      console.error("Échec de l'enregistrement du service worker :", err);
    });
  });
}
