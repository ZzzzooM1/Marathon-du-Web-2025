/**
 * Ã¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢Â
 *  VERIDEM — Site du projet DIGIDEM
 *  Université de Clermont-Auvergne × Partenaires Erasmus+
 * Ã¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢Â
 *
 *  PALETTE :
 *  • Marine profond (#0c1b33) — Autorité, profondeur, confiance
 *  • Bleu institution (#1a3a5c) — Sérieux, académique
 *  • Or / Doré (#c8a45e) — Excellence, prestige, label européen
 *  • Or clair (#e8d5a3) — Accents chauds, élégance
 *  • Blanc crème (#faf6ee) — Douceur, lisibilité, papier noble
 *  • Ivoire (#f2ece0) — Fonds secondaires
 *
 *  TYPOGRAPHIE :
 *  • Titres : "Playfair Display" (serif)
 *  • Corps : "Source Sans 3" (sans-serif humaniste)
 */

import { useState, useEffect, useRef } from "react";
import ReportDashboard from "./ReportDashboard";

const API_URL = process.env.REACT_APP_API_URL || "http://localhost:5000";

/* Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬ Google Fonts Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬ */
const fontLink = document.createElement("link");
fontLink.href =
  "https://fonts.googleapis.com/css2?family=Playfair+Display:ital,wght@0,400;0,600;0,700;0,800;1,400&family=Source+Sans+3:ital,wght@0,300;0,400;0,500;0,600;0,700;1,400&display=swap";
fontLink.rel = "stylesheet";
document.head.appendChild(fontLink);

/* Ã¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢Â
   CONSTANTES DESIGN
   Ã¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢Â */
const C = {
  navy: "#0c1b33",
  navyLight: "#1a3a5c",
  navyMid: "#132847",
  gold: "#c8a45e",
  goldLight: "#e8d5a3",
  goldPale: "#f0e4c8",
  cream: "#faf6ee",
  ivory: "#f2ece0",
  warmWhite: "#fffdf7",
  text: "#2c3e50",
  textLight: "#5a6c7d",
  textMuted: "#8a9aab",
  border: "#e0d8c8",
  borderLight: "#ece6d8",
  success: "#2e7d5b",
  error: "#b53a3a",
  warning: "#c8873e",
  cyan: "#1e91a6",
};

const FONT = {
  serif: "'Playfair Display', Georgia, 'Times New Roman', serif",
  sans: "'Source Sans 3', 'Segoe UI', Tahoma, sans-serif",
};

const LANG = {
  fr: "FR",
  en: "EN",
};

/* Ã¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢Â
   LOGO SVG
   Ã¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢Â */
function VeridemLogo({ height = 32, color = "#1e91a6" }) {
  return (
    <svg viewBox="0 0 1749.94 594.43" height={height} xmlns="http://www.w3.org/2000/svg">
      <g fill={color}>
        <path d="M1069.73,421.99c3.33-5.94,5-13.12,5-21.52s-1.67-15.77-5-21.64c-3.33-5.86-7.97-10.3-13.91-13.32-5.94-3.01-12.88-4.52-20.81-4.52h-23.54v78.95h23.54c7.93,0,14.86-1.5,20.81-4.52,5.94-3.01,10.58-7.49,13.91-13.44h0Z"/>
        <path d="M693.66,394.29c3.72-3.49,5.59-8.24,5.59-14.27s-1.86-10.78-5.59-14.27c-3.73-3.49-9.47-5.23-17.24-5.23h-20.45v39h20.45c7.77,0,13.51-1.74,17.24-5.23Z"/>
        <polygon points="671.25 441.84 655.96 441.84 655.96 482.19 698.47 482.19 671.25 441.84"/>
        <path d="M0,0v482.19h213.03l-69.75-164.95h60.4l39.61,97.14,40.76-97.14h55.17l-69.75,164.95h122.78v-164.95h135.3v42.33h-80.14v19.5h70.39v40.43h-70.39v21.88h83.23v40.82h69.2v-164.95h77.76c25.52,0,44.94,5.63,58.26,16.88,13.32,11.26,19.97,26.56,19.97,45.9,0,12.84-3.09,23.9-9.27,33.17-5.4,8.1-12.77,14.54-22.09,19.31l33.91,49.7h66.18v-164.95h56.12v164.95h74.67v-164.95h82.04c18.55,0,34.87,3.37,48.99,10.11,14.11,6.74,25.12,16.29,33.05,28.65,7.93,12.37,11.89,27.19,11.89,44.47s-3.96,31.91-11.89,44.35c-7.93,12.45-18.95,22.04-33.05,28.77-8.69,4.15-18.22,7.01-28.59,8.6h141.54v-164.95h135.3v42.33h-80.14v19.5h70.39v40.43h-70.39v21.88h83.23v40.82h69.2v-164.95h46.13l54.28,88.59,52.73-88.59h46.13l.47,164.95h143.28V0H0Z"/>
        <polygon points="1555.13 407.77 1518.92 468.72 1494.19 468.72 1458.04 410.4 1458.04 482.19 1555.53 482.19 1555.13 407.77"/>
      </g>
      <g fill={color}>
        <path d="M591.84,594v-50.62h21.84c5.64,0,10.02,1.13,13.12,3.4,3.11,2.27,4.66,5.47,4.66,9.62,0,2.75-.64,5.06-1.92,6.94-1.28,1.88-3.01,3.3-5.21,4.27-2.19.96-4.59,1.45-7.19,1.45l1.23-1.74c3.23,0,6.01.5,8.35,1.48,2.34.99,4.16,2.46,5.46,4.41,1.3,1.95,1.95,4.38,1.95,7.27,0,4.29-1.58,7.62-4.74,9.98s-7.92,3.54-14.28,3.54h-23.28.01ZM597.19,589.59h17.86c4.43,0,7.83-.76,10.2-2.28,2.36-1.52,3.54-3.94,3.54-7.27s-1.18-5.68-3.54-7.19c-2.36-1.52-5.76-2.28-10.2-2.28h-18.44v-4.41h16.7c4.05,0,7.19-.77,9.44-2.31,2.24-1.54,3.36-3.83,3.36-6.87s-1.12-5.39-3.36-6.91c-2.24-1.52-5.39-2.28-9.44-2.28h-16.13v41.8h.01Z"/>
        <path d="M669.64,594v-18.8l1.23,3.33-21.48-35.14h5.71l18.87,30.88h-3.04l18.87-30.88h5.35l-21.48,35.14,1.23-3.33v18.8h-5.28.02Z"/>
        <path d="M751.36,594v-50.62h20.61c5.4,0,10.13,1.07,14.21,3.22,4.07,2.15,7.24,5.12,9.51,8.93s3.4,8.2,3.4,13.16-1.13,9.35-3.4,13.16-5.44,6.79-9.51,8.93c-4.07,2.15-8.81,3.22-14.21,3.22h-20.61ZM756.71,589.37h14.97c4.48,0,8.38-.88,11.68-2.64s5.87-4.19,7.7-7.3,2.75-6.69,2.75-10.74-.92-7.69-2.75-10.77-4.4-5.51-7.7-7.27-7.19-2.64-11.68-2.64h-14.97v41.36h0Z"/>
        <path d="M825.55,594v-50.62h5.35v50.62h-5.35Z"/>
        <path d="M884.05,594.43c-3.86,0-7.41-.64-10.67-1.92-3.25-1.28-6.07-3.07-8.46-5.39s-4.24-5.04-5.57-8.17-1.99-6.56-1.99-10.27.66-7.13,1.99-10.27c1.33-3.13,3.19-5.86,5.6-8.17s5.24-4.11,8.5-5.39,6.81-1.92,10.66-1.92,7.39.6,10.59,1.81c3.21,1.21,5.94,3.04,8.21,5.5l-3.33,3.4c-2.12-2.12-4.46-3.64-7.01-4.56-2.56-.91-5.3-1.37-8.24-1.37-3.13,0-6.01.52-8.64,1.55-2.63,1.04-4.92,2.51-6.87,4.41-1.95,1.91-3.47,4.12-4.56,6.65-1.08,2.53-1.63,5.31-1.63,8.35s.54,5.76,1.63,8.32c1.08,2.56,2.6,4.79,4.56,6.69,1.95,1.91,4.24,3.38,6.87,4.41,2.63,1.04,5.48,1.55,8.57,1.55,2.89,0,5.63-.43,8.21-1.3s4.95-2.34,7.12-4.41l3.04,4.05c-2.41,2.12-5.23,3.72-8.46,4.81-3.23,1.09-6.6,1.63-10.12,1.63h0ZM902.63,588l-5.13-.65v-18.66h5.13v19.31Z"/>
        <path d="M932.21,594v-50.62h5.35v50.62h-5.35Z"/>
        <path d="M968.51,594v-50.62h20.61c5.4,0,10.14,1.07,14.21,3.22,4.07,2.15,7.24,5.12,9.51,8.93s3.4,8.2,3.4,13.16-1.13,9.35-3.4,13.16-5.44,6.79-9.51,8.93c-4.07,2.15-8.81,3.22-14.21,3.22h-20.61ZM973.86,589.37h14.97c4.48,0,8.38-.88,11.68-2.64s5.87-4.19,7.7-7.3,2.75-6.69,2.75-10.74-.92-7.69-2.75-10.77-4.4-5.51-7.7-7.27-7.19-2.64-11.68-2.64h-14.97v41.36h0Z"/>
        <path d="M1042.7,594v-50.62h34.71v4.63h-29.36v41.36h30.44v4.63h-35.79,0ZM1047.47,570.57v-4.56h26.75v4.56h-26.75Z"/>
        <path d="M1105.54,594v-50.62h4.41l23.14,39.48h-2.31l22.92-39.48h4.41v50.62h-5.13v-42.59h1.23l-21.12,36.16h-2.53l-21.26-36.16h1.37v42.59h-5.13Z"/>
      </g>
    </svg>
  );
}

/* Ã¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢Â
   DONNÉES
   Ã¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢Â */
const PARTNERS = {
  fr: [
    { abbr: "UCA", name: "Université Clermont Auvergne", country: "France", flag: "🇫🇷", role: "Organisation coordinatrice du projet", lead: true },
    { abbr: "LMS", name: "Les Militants des Savoirs", country: "France", flag: "🇫🇷", role: "Contenu pédagogique, éducation citoyenne et lutte contre la désinformation" },
    { abbr: "SEHA", name: "Sürdürülebilir Eğitim ve Sağlık Akademisi Derneği", country: "Turquie", flag: "🇹🇷", role: "Prévention de la violence numérique et promotion de la santé durable" },
    { abbr: "IPS", name: "QUALIFY JUST — IT Solutions and Consulting LDA", country: "Portugal", flag: "🇵🇹", role: "Technologie éducative via la réalité virtuelle" },
    { abbr: "UoM", name: "Universita ta Malta", country: "Malte", flag: "🇲🇹", role: "Programmes éducatifs et gestion de projets européens en éducation" },
    { abbr: "UNICAL", name: "Universita della Calabria", country: "Italie", flag: "🇮🇹", role: "Sciences sociales et technologies, outils numériques éducatifs" },
    { abbr: "Insight", name: "Insight S.A.", country: "Luxembourg", flag: "🇱🇺", role: "Citoyenneté active et formation numérique, dissémination" },
  ],
  en: [
    { abbr: "UCA", name: "Universite Clermont Auvergne", country: "France", flag: "🇫🇷", role: "Project coordinating organisation", lead: true },
    { abbr: "LMS", name: "Les Militants des Savoirs", country: "France", flag: "🇫🇷", role: "Pedagogical content, civic education and anti-disinformation work" },
    { abbr: "SEHA", name: "Surdurulebilir Egitim ve Saglik Akademisi Dernegi", country: "Turkey", flag: "🇹🇷", role: "Digital violence prevention and sustainable health promotion" },
    { abbr: "IPS", name: "QUALIFY JUST - IT Solutions and Consulting LDA", country: "Portugal", flag: "🇵🇹", role: "Educational technology through virtual reality" },
    { abbr: "UoM", name: "Universita ta Malta", country: "Malta", flag: "🇲🇹", role: "Educational programmes and European project management in education" },
    { abbr: "UNICAL", name: "Universita della Calabria", country: "Italy", flag: "🇮🇹", role: "Social sciences and technology, digital educational tools" },
    { abbr: "Insight", name: "Insight S.A.", country: "Luxembourg", flag: "🇱🇺", role: "Active citizenship, digital training and dissemination" },
  ],
};

const STATS = {
  fr: [
    { value: "86 %", label: "des Européens exposés aux fake news chaque semaine", icon: "◈" },
    { value: "×6", label: "vitesse de propagation de la désinformation vs la vérité", icon: "◆" },
    { value: "400", label: "images et textes annotés dans notre corpus de recherche", icon: "◇" },
    { value: "67 %", label: "des 18-24 ans ne vérifient pas avant de partager", icon: "◈" },
  ],
  en: [
    { value: "86%", label: "of Europeans are exposed to fake news every week", icon: "◈" },
    { value: "x6", label: "faster spread of disinformation than truth", icon: "◆" },
    { value: "400", label: "annotated images and texts in our research corpus", icon: "◇" },
    { value: "67%", label: "of 18-24 year-olds do not verify before sharing", icon: "◈" },
  ],
};

function normalizeLookup(value) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

/* Ã¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢Â
   COMPOSANTS UTILITAIRES
   Ã¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢Â */

function Ornament({ width = 60, style: s = {} }) {
  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 8, ...s }}>
      <div style={{ width: width * 0.35, height: 1, background: `linear-gradient(90deg, transparent, ${C.gold})` }} />
      <div style={{ width: 6, height: 6, border: `1px solid ${C.gold}`, transform: "rotate(45deg)" }} />
      <div style={{ width: width * 0.35, height: 1, background: `linear-gradient(270deg, transparent, ${C.gold})` }} />
    </div>
  );
}

function Label({ children, variant = "gold" }) {
  const styles = {
    gold: { color: C.gold, borderColor: `${C.gold}55`, bg: `${C.gold}0c` },
    navy: { color: C.navyLight, borderColor: `${C.navyLight}33`, bg: `${C.navyLight}08` },
    success: { color: C.success, borderColor: `${C.success}44`, bg: `${C.success}0c` },
    error: { color: C.error, borderColor: `${C.error}44`, bg: `${C.error}0c` },
    cyan: { color: C.cyan, borderColor: `${C.cyan}44`, bg: `${C.cyan}0c` },
  };
  const v = styles[variant] || styles.gold;
  return (
    <span style={{ display: "inline-block", padding: "5px 14px", borderRadius: 4, fontSize: 11, fontWeight: 600, fontFamily: FONT.sans, letterSpacing: 1.5, textTransform: "uppercase", color: v.color, border: `1px solid ${v.borderColor}`, background: v.bg }}>
      {children}
    </span>
  );
}

function FadeIn({ children, delay = 0, style: s = {} }) {
  const ref = useRef(null);
  const [visible, setVisible] = useState(false);
  useEffect(() => {
    const el = ref.current;
    if (!el) return;
    const obs = new IntersectionObserver(([entry]) => {
      if (entry.isIntersecting) { setTimeout(() => setVisible(true), delay); obs.unobserve(el); }
    }, { threshold: 0.15 });
    obs.observe(el);
    return () => obs.disconnect();
  }, [delay]);
  return (
    <div ref={ref} style={{ opacity: visible ? 1 : 0, transform: visible ? "translateY(0)" : "translateY(24px)", transition: `opacity 0.7s ease ${delay}ms, transform 0.7s ease ${delay}ms`, ...s }}>
      {children}
    </div>
  );
}

/* Ã¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢Â
   NAVIGATION
   Ã¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢Â */
function Nav({ currentPage, setPage, lang, setLang }) {
  const [scrolled, setScrolled] = useState(false);
  const [menuOpen, setMenuOpen] = useState(false);
  useEffect(() => {
    const fn = () => setScrolled(window.scrollY > 20);
    window.addEventListener("scroll", fn);
    return () => window.removeEventListener("scroll", fn);
  }, []);

  const links = [
    { id: "home", label: lang === "fr" ? "Accueil" : "Home" },
    { id: "visu", label: lang === "fr" ? "Visualisation" : "Visualization" },
    { id: "demo", label: lang === "fr" ? "Démonstration" : "Demo" },
  ];

  return (
    <nav style={{ position: "fixed", top: 0, left: 0, right: 0, zIndex: 1000, height: 68, display: "flex", alignItems: "center", justifyContent: "space-between", padding: "0 clamp(16px, 4vw, 40px)", background: scrolled ? `${C.cream}f5` : `${C.cream}00`, backdropFilter: scrolled ? "blur(16px)" : "none", borderBottom: scrolled ? `1px solid ${C.border}` : "1px solid transparent", transition: "all 0.4s ease", fontFamily: FONT.sans }}>

      <div style={{ display: "flex", alignItems: "center", gap: 12, cursor: "pointer" }} onClick={() => setPage("home")}>
        <VeridemLogo height={28} color={C.gold} />
      </div>

      <div style={{ display: "flex", gap: 2 }} className="nav-desktop">
        {links.map((l) => (
          <button key={l.id} onClick={() => { setPage(l.id); setMenuOpen(false); }}
            style={{ background: "none", border: "none", padding: "10px 20px", fontFamily: FONT.sans, fontWeight: currentPage === l.id ? 600 : 400, fontSize: 14, color: currentPage === l.id ? C.navy : C.textLight, cursor: "pointer", position: "relative", transition: "color 0.3s" }}>
            {l.label}
            <div style={{ position: "absolute", bottom: 4, left: 20, right: 20, height: 2, background: C.gold, borderRadius: 1, transform: currentPage === l.id ? "scaleX(1)" : "scaleX(0)", transition: "transform 0.3s ease" }} />
          </button>
        ))}
      </div>

      <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, fontSize: 12, color: C.textMuted }} className="nav-badge">
          <svg width="24" height="16" viewBox="0 0 24 16">
            <rect width="24" height="16" rx="2" fill="#003399" />
            {Array.from({ length: 12 }, (_, i) => {
              const a = (i * 30 - 90) * (Math.PI / 180);
              const cx = 12 + 5 * Math.cos(a), cy = 8 + 5 * Math.sin(a);
              return <polygon key={i} points={`${cx},${cy-1.2} ${cx+0.4},${cy-0.3} ${cx+1.2},${cy-0.3} ${cx+0.5},${cy+0.3} ${cx+0.8},${cy+1.2} ${cx},${cy+0.6} ${cx-0.8},${cy+1.2} ${cx-0.5},${cy+0.3} ${cx-1.2},${cy-0.3} ${cx-0.4},${cy-0.3}`} fill="#FFCC00" />;
            })}
          </svg>
          <span style={{ fontFamily: FONT.sans, fontWeight: 500 }}>
            {lang === "fr" ? "Cofinancé par Erasmus+" : "Co-funded by Erasmus+"}
          </span>
        </div>

        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          {Object.keys(LANG).map((code) => (
            <button
              key={code}
              onClick={() => setLang(code)}
              style={{
                background: lang === code ? `${C.gold}18` : "transparent",
                border: lang === code ? `1px solid ${C.gold}` : `1px solid ${C.borderLight}`,
                color: lang === code ? C.navy : C.textMuted,
                borderRadius: 4,
                padding: "6px 10px",
                fontFamily: FONT.sans,
                fontSize: 11,
                fontWeight: 700,
                letterSpacing: 0.8,
                cursor: "pointer",
              }}
            >
              {LANG[code]}
            </button>
          ))}
        </div>
      </div>

      <button className="nav-mobile-btn" onClick={() => setMenuOpen(!menuOpen)} style={{ background: "none", border: "none", cursor: "pointer", padding: 8, display: "none" }}>
        <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke={C.navy} strokeWidth="2">
          {menuOpen ? <path d="M18 6L6 18M6 6l12 12" /> : <path d="M3 12h18M3 6h18M3 18h18" />}
        </svg>
      </button>

      {menuOpen && (
        <div style={{ position: "absolute", top: 68, left: 0, right: 0, background: C.cream, borderBottom: `1px solid ${C.border}`, padding: "8px 16px", animation: "fadeSlideIn 0.3s ease" }}>
          <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
            {Object.keys(LANG).map((code) => (
              <button
                key={code}
                onClick={() => setLang(code)}
                style={{
                  background: lang === code ? `${C.gold}18` : "transparent",
                  border: lang === code ? `1px solid ${C.gold}` : `1px solid ${C.borderLight}`,
                  color: lang === code ? C.navy : C.textMuted,
                  borderRadius: 4,
                  padding: "8px 12px",
                  fontFamily: FONT.sans,
                  fontSize: 11,
                  fontWeight: 700,
                  letterSpacing: 0.8,
                  cursor: "pointer",
                }}
              >
                {LANG[code]}
              </button>
            ))}
          </div>
          {links.map((l) => (
            <button key={l.id} onClick={() => { setPage(l.id); setMenuOpen(false); }}
              style={{ display: "block", width: "100%", textAlign: "left", background: currentPage === l.id ? `${C.gold}12` : "none", border: "none", padding: "14px 16px", fontFamily: FONT.sans, fontWeight: currentPage === l.id ? 600 : 400, fontSize: 15, color: currentPage === l.id ? C.navy : C.textLight, cursor: "pointer", borderRadius: 6 }}>
              {l.label}
            </button>
          ))}
        </div>
      )}
    </nav>
  );
}

/* Ã¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢Â
   PAGE ACCUEIL
   Ã¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢Â */
const HERO_BG = "data:image/jpeg;base64,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";

function HomePage({ setPage, lang }) {
  const isEn = lang === "en";
  const [hv, setHv] = useState(false);
  useEffect(() => { setTimeout(() => setHv(true), 150); }, []);

  return (
    <div style={{ paddingTop: 68, background: C.cream }}>
      {/* HERO */}
      <section style={{ minHeight: "90vh", background: C.navy, backgroundImage: `url(${HERO_BG})`, backgroundSize: "cover", backgroundPosition: "center", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", textAlign: "center", padding: "60px 32px 80px", position: "relative", overflow: "hidden" }}>
        {/* Léger voile sombre pour lisibilité du texte sur l'image */}
        <div style={{ position: "absolute", inset: 0, background: `linear-gradient(180deg, ${C.navy}88 0%, ${C.navy}55 40%, ${C.navy}77 100%)`, zIndex: 1 }} />

        <div style={{ position: "relative", zIndex: 3, display: "flex", flexDirection: "column", alignItems: "center" }}>
        <div style={{ opacity: hv ? 1 : 0, transform: hv ? "translateY(0)" : "translateY(16px)", transition: "all 0.8s ease 0.1s" }}>
          <Label>{isEn ? "Web Marathon 2026 - Universite Clermont Auvergne" : "Marathon du Web 2026 — Université Clermont Auvergne"}</Label>
        </div>

        <div style={{ opacity: hv ? 1 : 0, transform: hv ? "translateY(0)" : "translateY(16px)", transition: "all 0.9s ease 0.2s", margin: "28px 0 0" }}>
          <VeridemLogo height={48} color={C.gold} />
        </div>

        <h1 style={{ fontFamily: FONT.serif, fontWeight: 700, fontSize: "clamp(28px, 5vw, 58px)", lineHeight: 1.15, color: C.cream, margin: "24px 0 0", maxWidth: 720, opacity: hv ? 1 : 0, transform: hv ? "translateY(0)" : "translateY(24px)", transition: "all 1s ease 0.25s" }}>
          {isEn ? "Seeing is no longer enough." : "Voir ne suffit plus."}<br /><span style={{ color: C.gold }}>{isEn ? "Learn to verify." : "Apprenez à vérifier."}</span>
        </h1>

        <div style={{ opacity: hv ? 1 : 0, transition: "opacity 1s ease 0.5s", margin: "24px 0" }}><Ornament width={80} /></div>

        <p style={{ fontFamily: FONT.sans, fontSize: "clamp(14px, 1.6vw, 17px)", fontWeight: 300, color: C.goldLight, maxWidth: 560, lineHeight: 1.75, opacity: hv ? 1 : 0, transform: hv ? "translateY(0)" : "translateY(20px)", transition: "all 1s ease 0.45s" }}>
          {isEn ? "In the age of deepfakes and generative AI, every European citizen needs a critical eye on information. The DIGIDEM project gives you the tools." : "À l'ère des deepfakes et de l'IA générative, chaque citoyen européen doit développer un regard critique sur l'information. Le projet DIGIDEM vous donne les outils."}
        </p>

        <div style={{ display: "flex", gap: 16, marginTop: 32, flexWrap: "wrap", justifyContent: "center", opacity: hv ? 1 : 0, transform: hv ? "translateY(0)" : "translateY(20px)", transition: "all 1s ease 0.6s" }}>
          <button onClick={() => setPage("visu")} style={{ background: C.gold, color: C.navy, border: "none", borderRadius: 4, padding: "15px 36px", fontFamily: FONT.sans, fontWeight: 600, fontSize: 14, letterSpacing: 0.5, cursor: "pointer", transition: "all 0.3s", boxShadow: `0 4px 20px ${C.gold}33` }}
            onMouseEnter={(e) => { e.target.style.background = C.goldLight; e.target.style.boxShadow = `0 6px 28px ${C.gold}55`; }}
            onMouseLeave={(e) => { e.target.style.background = C.gold; e.target.style.boxShadow = `0 4px 20px ${C.gold}33`; }}
          >{isEn ? "Explore the data" : "Explorer les données"}</button>
          <button onClick={() => setPage("demo")} style={{ background: "transparent", color: C.goldLight, border: `1px solid ${C.gold}44`, borderRadius: 4, padding: "15px 36px", fontFamily: FONT.sans, fontWeight: 500, fontSize: 14, letterSpacing: 0.5, cursor: "pointer", transition: "all 0.3s" }}
            onMouseEnter={(e) => (e.target.style.borderColor = C.gold)}
            onMouseLeave={(e) => (e.target.style.borderColor = `${C.gold}44`)}
          >{isEn ? "Try the AI image detector" : "Tester le détecteur d'images IA"}</button>
        </div>
        </div>
      </section>

      {/* VIDEO */}
      <section style={{ padding: "64px 32px", maxWidth: 900, margin: "0 auto" }}>
        <FadeIn>
          <div style={{ textAlign: "center", marginBottom: 40 }}>
            <Label variant="navy">{isEn ? "Video overview" : "Présentation vidéo"}</Label>
            <h2 style={{ fontFamily: FONT.serif, fontWeight: 700, fontSize: "clamp(22px, 3vw, 34px)", color: C.navy, margin: "16px 0 0" }}>
              {isEn ? "Discover the project " : "Découvrez le projet "}<span style={{ color: C.gold, fontStyle: "italic" }}>DIGIDEM</span>
            </h2>
            <Ornament width={50} style={{ marginTop: 16 }} />
          </div>
        </FadeIn>
        <FadeIn delay={150}>
          <a href="https://www.youtube.com/watch?v=Ja2VEBRqq0A" target="_blank" rel="noopener noreferrer"
            style={{ display: "block", position: "relative", borderRadius: 8, overflow: "hidden", border: `1px solid ${C.border}`, boxShadow: `0 12px 48px ${C.navy}08`, background: C.navy, aspectRatio: "16/9", cursor: "pointer", textDecoration: "none" }}>
            {/* YouTube thumbnail */}
            <img
              src="https://img.youtube.com/vi/SanVY3U2f8A/maxresdefault.jpg"
              alt={isEn ? "DIGIDEM - project overview" : "DIGIDEM — Présentation du projet"}
              style={{ width: "100%", height: "100%", objectFit: "cover", display: "block", opacity: 0.85, transition: "opacity 0.3s" }}
              onMouseEnter={(e) => (e.target.style.opacity = 1)}
              onMouseLeave={(e) => (e.target.style.opacity = 0.85)}
              onError={(e) => { e.target.src = "https://img.youtube.com/vi/SanVY3U2f8A/hqdefault.jpg"; }}
            />
            {/* Dark overlay */}
            <div style={{ position: "absolute", inset: 0, background: "linear-gradient(180deg, rgba(12,27,51,0.15) 0%, rgba(12,27,51,0.5) 100%)", pointerEvents: "none" }} />
            {/* Play button */}
            <div style={{ position: "absolute", top: "50%", left: "50%", transform: "translate(-50%, -50%)", width: 80, height: 80, borderRadius: "50%", background: `${C.gold}ee`, display: "flex", alignItems: "center", justifyContent: "center", boxShadow: `0 8px 32px ${C.navy}44`, transition: "transform 0.3s, background 0.3s" }}>
              <svg width="32" height="32" viewBox="0 0 24 24" fill={C.navy} style={{ marginLeft: 3 }}>
                <polygon points="6,3 20,12 6,21" />
              </svg>
            </div>
            {/* Badge */}
            <div style={{ position: "absolute", bottom: 16, left: 16, background: `${C.navy}cc`, backdropFilter: "blur(8px)", borderRadius: 4, padding: "6px 14px", fontFamily: FONT.sans, fontSize: 12, fontWeight: 600, color: C.goldLight, letterSpacing: 0.5 }}>
              {isEn ? "Watch on YouTube" : "Regarder sur YouTube"}
            </div>
            {/* Duration badge */}
            <div style={{ position: "absolute", bottom: 16, right: 16, background: `${C.navy}cc`, backdropFilter: "blur(8px)", borderRadius: 4, padding: "6px 12px", fontFamily: FONT.sans, fontSize: 11, color: C.goldLight }}>
              {isEn ? "DIGIDEM - Web Marathon" : "DIGIDEM — Marathon du Web"}
            </div>
          </a>
        </FadeIn>
      </section>

      {/* CHIFFRES CLÉS */}
      <section style={{ padding: "80px 32px", maxWidth: 960, margin: "0 auto" }}>
        <FadeIn>
          <div style={{ textAlign: "center", marginBottom: 56 }}>
            <Label variant="navy">{isEn ? "Why it matters" : "L'urgence en chiffres"}</Label>
            <h2 style={{ fontFamily: FONT.serif, fontWeight: 700, fontSize: "clamp(24px, 3.5vw, 38px)", color: C.navy, margin: "20px 0 0" }}>
              {isEn ? "Disinformation, a democratic challenge" : "La désinformation, un défi démocratique"}
            </h2>
            <Ornament width={50} style={{ marginTop: 20 }} />
          </div>
        </FadeIn>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 24 }}>
          {STATS[lang].map((stat, i) => (
            <FadeIn key={i} delay={i * 120}>
              <div
                style={{ background: C.warmWhite, border: `1px solid ${C.borderLight}`, borderRadius: 6, padding: "32px 24px", textAlign: "center", transition: "all 0.4s ease", cursor: "default" }}
                onMouseEnter={(e) => { e.currentTarget.style.borderColor = C.gold; e.currentTarget.style.boxShadow = `0 8px 32px ${C.gold}15`; e.currentTarget.style.transform = "translateY(-3px)"; }}
                onMouseLeave={(e) => { e.currentTarget.style.borderColor = C.borderLight; e.currentTarget.style.boxShadow = "none"; e.currentTarget.style.transform = "translateY(0)"; }}
              >
                <div style={{ color: C.gold, fontSize: 18, marginBottom: 8, opacity: 0.6 }}>{stat.icon}</div>
                <div style={{ fontFamily: FONT.serif, fontWeight: 700, fontSize: 40, color: C.navy, lineHeight: 1 }}>{stat.value}</div>
                <p style={{ fontFamily: FONT.sans, fontSize: 13, color: C.textLight, lineHeight: 1.65, marginTop: 12 }}>{stat.label}</p>
              </div>
            </FadeIn>
          ))}
        </div>
      </section>

      {/* MISSION */}
      <section style={{ background: C.ivory, borderTop: `1px solid ${C.borderLight}`, borderBottom: `1px solid ${C.borderLight}`, padding: "80px 32px" }}>
        <div style={{ maxWidth: 800, margin: "0 auto", textAlign: "center" }}>
          <FadeIn>
            <Label>{isEn ? "Our mission" : "Notre mission"}</Label>
            <h2 style={{ fontFamily: FONT.serif, fontWeight: 700, fontSize: "clamp(22px, 3.2vw, 34px)", color: C.navy, margin: "20px 0 0", lineHeight: 1.3 }}>
              {isEn ? "Training the next generation " : "Former la prochaine génération "}<span style={{ color: C.gold, fontStyle: "italic" }}>{isEn ? "of critical digital citizens" : "de citoyens numériques critiques"}</span>
            </h2>
            <Ornament width={50} style={{ margin: "24px auto" }} />
            <p style={{ fontFamily: FONT.sans, fontSize: 16, color: C.textLight, lineHeight: 1.85, maxWidth: 620, margin: "0 auto" }}>
              DIGIDEM est un projet européen Erasmus+ réunissant sept partenaires de cinq pays autour d'un objectif commun : développer un programme de formation innovant visant à renforcer les compétences numériques et l'engagement civique. Notre approche allie analyse computationnelle d'images, détection par IA, et ateliers pratiques lors du Marathon du Web.
            </p>
          </FadeIn>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))", gap: 20, marginTop: 48 }}>
            {[
              { icon: <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke={C.gold} strokeWidth="1.5"><path d="M22 10v6M2 10l10-5 10 5-10 5z" /><path d="M6 12v5c3 3 9 3 12 0v-5" /></svg>, title: "Éduquer", desc: "Développer l'esprit critique face aux médias numériques par la pratique et l'analyse de cas réels de désinformation." },
              { icon: <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke={C.gold} strokeWidth="1.5"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z" /></svg>, title: "Outiller", desc: "Fournir des instruments d'analyse par intelligence artificielle (ViT, Florence-2, OCR) accessibles et transparents." },
              { icon: <svg width="28" height="28" viewBox="0 0 24 24" fill="none" stroke={C.gold} strokeWidth="1.5"><circle cx="12" cy="12" r="10" /><line x1="2" y1="12" x2="22" y2="12" /><path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z" /></svg>, title: "Fédérer", desc: "Unir universités et associations européennes dans un consortium de recherche-action contre la désinformation." },
            ].map((p, i) => (
              <FadeIn key={i} delay={i * 100}>
                <div style={{ background: C.warmWhite, border: `1px solid ${C.borderLight}`, borderRadius: 6, padding: "32px 24px", textAlign: "left", transition: "border-color 0.3s" }}
                  onMouseEnter={(e) => (e.currentTarget.style.borderColor = C.gold)}
                  onMouseLeave={(e) => (e.currentTarget.style.borderColor = C.borderLight)}>
                  <div style={{ marginBottom: 16 }}>{p.icon}</div>
                  <h3 style={{ fontFamily: FONT.serif, fontWeight: 700, fontSize: 20, color: C.navy, marginBottom: 8 }}>{p.title}</h3>
                  <p style={{ fontFamily: FONT.sans, fontSize: 14, color: C.textLight, lineHeight: 1.7, margin: 0 }}>{p.desc}</p>
                </div>
              </FadeIn>
            ))}
          </div>
        </div>
      </section>

      {/* PARTENAIRES */}
      <section style={{ padding: "80px 32px", maxWidth: 960, margin: "0 auto" }}>
        <FadeIn>
          <div style={{ textAlign: "center", marginBottom: 48 }}>
            <Label>{isEn ? "European consortium" : "Consortium européen"}</Label>
            <h2 style={{ fontFamily: FONT.serif, fontWeight: 700, fontSize: "clamp(22px, 3.5vw, 34px)", color: C.navy, margin: "20px 0 0" }}>{isEn ? "Seven partners, five countries, one shared goal" : "Sept partenaires, cinq pays, un objectif commun"}</h2>
            <Ornament width={50} style={{ marginTop: 20 }} />
          </div>
        </FadeIn>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(250px, 1fr))", gap: 16 }}>
          {PARTNERS[lang].map((p, i) => (
            <FadeIn key={i} delay={i * 60}>
              <div
                style={{ background: C.warmWhite, border: p.lead ? `1.5px solid ${C.gold}` : `1px solid ${C.borderLight}`, borderRadius: 6, padding: "28px 20px", position: "relative", transition: "all 0.3s" }}
                onMouseEnter={(e) => { e.currentTarget.style.transform = "translateY(-2px)"; e.currentTarget.style.boxShadow = `0 6px 24px ${C.navy}08`; }}
                onMouseLeave={(e) => { e.currentTarget.style.transform = "translateY(0)"; e.currentTarget.style.boxShadow = "none"; }}>
                {p.lead && <div style={{ position: "absolute", top: -8, left: 20, background: C.gold, color: C.navy, fontSize: 9, fontWeight: 700, fontFamily: FONT.sans, padding: "2px 10px", borderRadius: 3, letterSpacing: 1, textTransform: "uppercase" }}>{isEn ? "Coordinator" : "Coordinateur"}</div>}
                <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 12 }}>
                  <span style={{ fontSize: 24 }}>{p.flag}</span>
                  <div>
                    <div style={{ fontFamily: FONT.serif, fontWeight: 700, fontSize: 16, color: C.navy }}>{p.abbr}</div>
                    <div style={{ fontFamily: FONT.sans, fontSize: 11, color: C.textMuted, fontStyle: "italic" }}>{p.country}</div>
                  </div>
                </div>
                <div style={{ fontFamily: FONT.sans, fontWeight: 600, fontSize: 13, color: C.navy, lineHeight: 1.3, marginBottom: 8 }}>{p.name}</div>
                <p style={{ fontFamily: FONT.sans, fontSize: 12, color: C.textLight, lineHeight: 1.6, margin: 0 }}>{p.role}</p>
              </div>
            </FadeIn>
          ))}
        </div>
      </section>

      {/* FOOTER */}
      <footer style={{ padding: "40px 32px", textAlign: "center", borderTop: `1px solid ${C.borderLight}`, background: C.ivory }}>
        <div style={{ marginBottom: 16 }}>
          <VeridemLogo height={22} color={C.gold} />
        </div>
        <Ornament width={40} style={{ marginBottom: 16 }} />
        <p style={{ fontFamily: FONT.sans, fontSize: 12, color: C.textMuted, lineHeight: 1.7, maxWidth: 700, margin: "0 auto" }}>
          © 2024–2027 DIGIDEM — Projet co-financé par le programme Erasmus+ de l'Union européenne (KA220-HED-3743E85D).<br />
          Coordinateur : Université Clermont Auvergne — Laboratoire Communication et Sociétés (ComSocs) — Responsable : Nawel CHAOUNI.<br />
          Les contenus de ce site n'engagent que leurs auteurs et la Commission européenne ne saurait être tenue responsable de l'usage qui pourrait en être fait.
        </p>
      </footer>
    </div>
  );
}

/* Ã¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢Â
   PAGE VISUALISATION
   Ã¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢Â */
function VisuPage({ lang }) {
  return <ReportDashboard lang={lang} colors={C} fonts={FONT} Label={Label} Ornament={Ornament} FadeIn={FadeIn} />;
}

/* ===========================================
   PAGE DEMONSTRATION IA
   =========================================== */
function DemoPage({ lang }) {
  const isEn = lang === "en";
  const [file, setFile] = useState(null);
  const [preview, setPreview] = useState(null);
  const [analyzing, setAnalyzing] = useState(false);
  const [progress, setProgress] = useState(0);
  const [result, setResult] = useState(null);
  const [dragOver, setDragOver] = useState(false);
  const [activeResultTab, setActiveResultTab] = useState("verdict");
  const fileInputRef = useRef(null);

  const handleFile = (f) => {
    if (!f || !f.type.startsWith("image/")) return;
    setFile(f); setResult(null); setActiveResultTab("verdict");
    const reader = new FileReader();
    reader.onload = (e) => setPreview(e.target.result);
    reader.readAsDataURL(f);
  };

  /* Ã¢â€â‚¬Ã¢â€â‚¬ Générateurs fidèles au format réel du pipeline Florence-2 + ViT Ã¢â€â‚¬Ã¢â€â‚¬ */
  const pick = (arr) => arr[Math.floor(Math.random() * arr.length)];
  const randInt = (min, max) => min + Math.floor(Math.random() * (max - min + 1));
  const randFloat = (min, max) => +(min + Math.random() * (max - min)).toFixed(4);

  /* OCR : simule le texte brut que Florence-2 <OCR> extrairait d'un screenshot/document */
  const generateOCR = () => {
    const templates = [
      () => [
        pick(["@news_alert_eu", "@realinfo24", "@factcheck_fr", "@europe_truth", "@breakingnow_", "@urgent_info_eu"]),
        pick(["15h", "8h", "22h", "3h", "11h", "6h", "19h"]),
        "",
        pick([
          "BREAKING: Scientists confirm link between 5G and virus spread",
          "URGENT: EU votes to ban all diesel cars by 2025",
          "EXCLUSIVE: Leaked documents reveal secret vaccine deal",
          "ALERT: Study shows wifi causes brain tumors in children",
          "SHOCKING: President caught with lobbyists in secret meeting",
          "NEW: AI-generated images used to manipulate election results",
        ]),
        "",
        pick(["Share before they delete this!", "They don't want you to see this!", "EXPOSED — spread the word!", "MSM won't show you this!"]),
        `${pick(["12", "45", "89", "156", "234", "567"])}K ${pick(["Retweets", "Shares", "Likes"])} ${pick(["1.2", "3.4", "5.6", "8.9"])}K ${pick(["Comments", "Replies"])}`,
        "",
        pick(["#WakeUp #Truth", "#FakeNews #Alert", "#Exposed #Share", "#NWO #Resist", "#BigPharma #Truth"]),
      ].join("\n"),

      () => [
        pick(["InfoVérité360.com", "SantéLibre24.fr", "AlerteEurope.net", "RéveilCitoyen.info", "TruthReport247.com"]),
        pick(["Accueil", "Home"]) + " > " + pick(["Actualités", "Santé", "Politique", "Société"]) + " > " + pick(["Article", "Dossier", "Exclusif"]),
        "",
        pick([
          "Une étude confidentielle prouve que les ondes 5G sont dangereuses",
          "La vérité sur les vaccins que Big Pharma vous cache",
          "Scandale : la Commission européenne ment sur le climat",
          "EXCLUSIF : les preuves de fraude électorale en Europe",
          "Alerte santé : le café provoque le cancer du cerveau",
        ]),
        "",
        `Par ${pick(["Dr. Martin Dupont", "Prof. S. Martinez", "La Rédaction", "Sources anonymes"])}`,
        `Publié le ${randInt(1,28)}/${randInt(1,12)}/${pick(["2024", "2025"])}`,
        `Lu ${randInt(12, 456)} ${randInt(100, 999)} fois`,
        "",
        pick([
          "Partagez cet article avant qu'il ne soit censuré !",
          "Les médias mainstream refusent d'en parler.",
          "Faites tourner, la vérité doit éclater.",
        ]),
      ].join("\n"),

      () => [
        pick(["COMMUNIQUÉ DE PRESSE", "PRESS RELEASE", "NOTE D'INFORMATION"]),
        pick(["Ministère de l'Intérieur", "European Commission", "Préfecture du Puy-de-Dôme", "Mairie de Clermont-Ferrand", "Parlement européen"]),
        `Réf. ${pick(["COM", "INF", "PRE", "DOC"])}-${randInt(2024,2025)}-${randInt(1000,9999)}`,
        `${randInt(1,28)}/${randInt(1,12)}/${pick(["2024", "2025"])}`,
        "",
        pick([
          "Objet : Mise au point concernant les informations circulant sur les réseaux sociaux",
          "Objet : Résultats officiels du scrutin",
          "Subject: Clarification regarding recent media reports",
          "Objet : Bilan des actions de lutte contre la désinformation",
        ]),
        "",
        pick([
          `Le ${pick(["ministère", "service de communication"])} dément formellement les allégations relayées en ligne.`,
          "Les résultats publiés sont conformes aux données officielles.",
          "The European Commission reaffirms its commitment to fighting disinformation.",
        ]),
        "",
        `Contact : ${pick(["presse@gov.fr", "info@europa.eu", "communication@prefet63.fr"])}`,
      ].join("\n"),

      () => {
        /* Simule un screenshot de réseau social avec texte mélangé */
        const user = pick(["Marie L.", "Jean-Pierre M.", "Anonymous", "Patrick D.", "Sophie K.", "Mohamed A."]);
        return [
          pick(["Facebook", "Twitter", "Instagram", "Telegram"]),
          `${user} ${pick(["a partagé un lien", "shared a post", "a publié"])}`,
          pick(["Il y a 2h", "3 hours ago", "Hier", "Il y a 5 min"]),
          "",
          pick([
            "Regardez cette photo !!! C'est la PREUVE que tout est manipulé",
            "Je savais que c'était faux. Voilà la preuve en image.",
            "Mon ami médecin m'a envoyé ça. Les gens doivent savoir.",
            "RT si vous pensez qu'on nous ment !!",
            "Personne n'en parle mais c'est ÉNORME",
            "Attention arnaque ! Cette image a été générée par l'IA",
          ]),
          "",
          `${pick(["👍", "❤️", "😮"])} ${randInt(2,89)}K   💬 ${randInt(100,9999)}   ↔️ ${randInt(1,45)}K`,
          "",
          `${pick(["Voir les commentaires", "View comments", "Voir plus"])}`,
        ].join("\n");
      },

      () => "(aucun texte détecté)",
    ];
    return pick(templates)();
  };

  /* DESCRIPTION : simule la sortie de Florence-2 <MORE_DETAILED_CAPTION> — anglais, factuel, court */
  const generateDescription = (isAI) => {
    const scene = pick([
      "The image shows a screenshot of a social media post",
      "The image is a photograph of a person standing in front of a building",
      "The image contains a news article headline with a large photograph",
      "The image shows a document with text and a photograph attached",
      "The image depicts a group of people in an outdoor setting",
      "The image is a screenshot from a messaging application",
      "The image shows a political poster with text overlay",
      "The image contains a chart or infographic with data",
      "The image is a photograph of a public event with crowds",
      "The image shows a portrait of a person against a plain background",
      "The image depicts a landscape with buildings in the background",
      "The image shows a split view with two photographs side by side",
    ]);
    const details = pick([
      "with text in French visible at the top and bottom of the frame",
      "with bold red and white text overlaid on the image",
      "featuring a headline in large font and a smaller caption below",
      "with a logo in the upper left corner and a timestamp",
      "containing several lines of text and social media interaction buttons",
      "with a watermark visible in the lower right corner",
      "featuring multiple people and text annotations",
      "with a dark background and bright text elements",
      "including a URL and publication date at the bottom",
      "with hashtags and emoji visible in the text",
    ]);
    const ending = isAI
      ? pick([
          "The photograph within the image appears to have been artificially generated, with unnatural skin textures and inconsistent lighting.",
          "The main photograph shows signs of AI generation, including distorted background elements and overly smooth surfaces.",
          "The embedded image contains visual artifacts consistent with AI-generated content, such as warped text and blurred edges.",
          "The visual content shows characteristics typical of diffusion model outputs, with unnaturally perfect symmetry.",
        ])
      : pick([
          "The photograph appears to be a genuine capture with natural lighting and camera noise consistent with a mobile phone.",
          "The image shows typical characteristics of an authentic photograph, including natural depth of field and motion blur.",
          "The embedded photograph appears genuine, with consistent EXIF-like noise patterns and natural color gradients.",
          "The visual content shows no significant indicators of artificial generation.",
        ]);
    return `${scene}, ${details}. ${ending}`;
  };

  /* PHOTOS : simule la sortie de detect_photos + predict_ai_or_real — format fidèle au pipeline */
  const generatePhotos = (isAI, sourceFile = file) => {
    const count = randInt(1, 3);
    /* Captions style Florence-2 <CAPTION> : courts, factuels, anglais */
    const captions = [
      "a person wearing a suit standing in front of a building",
      "a group of people standing in a room",
      "a close up of a person's face",
      "a man in a white coat standing in front of a microphone",
      "a large building with a flag on top",
      "a screenshot of a web page with text",
      "a crowd of people holding signs",
      "a woman sitting at a desk with a computer",
      "a city street with cars and buildings",
      "two people shaking hands in front of cameras",
      "a document with text and official stamps",
      "a landscape with mountains in the background",
      "a person holding a phone showing a social media post",
      "a chart showing data with colored bars",
      "a close up of a newspaper headline",
    ];
    const photos = [];
    const used = [];
    const baseName = sourceFile?.name?.split(".")[0] || "image";
    for (let n = 0; n < count; n++) {
      const w = pick([320, 467, 512, 640, 768, 800, 1024]);
      const h = pick([240, 350, 384, 480, 576, 600, 768]);
      let cap = pick(captions);
      while (used.includes(cap) && used.length < captions.length) cap = pick(captions);
      used.push(cap);

      const probAI = isAI ? randFloat(0.65, 0.98) : randFloat(0.01, 0.18);
      const probHuman = +(1 - probAI).toFixed(4);
      const verdict = probAI > probHuman ? "IA" : "RÉEL";
      const confidence = Math.max(probAI, probHuman);

      photos.push({
        filename: `${baseName}_photo_${String(n + 1).padStart(2, "0")}.png`,
        size: `${w}×${h}`,
        caption: cap,
        verdict,
        confidence,
        probAI,
        probHuman,
      });
    }
    return photos;
  };

  const inferMockVerdict = (selectedFile) => {
    const fileKey = normalizeLookup(selectedFile?.name);
    const aiHints = ["ai", "ia", "generated", "genere", "deepfake", "midjourney", "dalle", "sdxl", "stable diffusion"];
    const realHints = ["photo", "camera", "iphone", "screenshot", "capture", "scan", "dsc", "img"];

    if (aiHints.some((hint) => fileKey.includes(hint))) return true;
    if (realHints.some((hint) => fileKey.includes(hint))) return false;
    return Math.random() >= 0.5;
  };

  const generateMockZones = (isAI) => {
    const zoneRanges = isAI
      ? [
          ["Artefacts de génération", [68, 94]],
          ["Cohérence des textures", [54, 88]],
          ["Patterns spectraux", [49, 84]],
          ["Contours et transitions", [42, 79]],
        ]
      : [
          ["Artefacts de génération", [8, 28]],
          ["Cohérence des textures", [12, 34]],
          ["Patterns spectraux", [10, 30]],
          ["Contours et transitions", [14, 36]],
        ];

    return zoneRanges.map(([label, [min, max]]) => ({
      label,
      score: randInt(min, max),
    }));
  };

  const buildMockAnalysis = (selectedFile) => {
    const isAI = inferMockVerdict(selectedFile);
    const strongestScore = randFloat(74, 97);
    const probAI = +(isAI ? strongestScore : 100 - strongestScore).toFixed(1);
    const probReal = +(100 - probAI).toFixed(1);
    const ocrText = generateOCR();
    const description = generateDescription(isAI);
    const photos = generatePhotos(isAI, selectedFile);
    const totalAI = photos.filter((photo) => photo.verdict === "IA").length;
    const fileFormat = selectedFile?.type?.split("/")[1]?.toUpperCase() || "UNKNOWN";
    const fileSizeMo = selectedFile?.size ? `${(selectedFile.size / (1024 * 1024)).toFixed(2)} Mo` : "Inconnue";

    return {
      isAI,
      confidence: Math.round(Math.max(probAI, probReal)),
      probAI,
      probReal,
      ocrText,
      ocrCharCount: ocrText.length,
      description,
      photos,
      totalPhotos: photos.length,
      totalAI,
      totalReal: photos.length - totalAI,
      metadata: [
        { key: "Nom du fichier", val: selectedFile?.name || "image" },
        { key: "Format", val: fileFormat },
        { key: "Taille du fichier", val: fileSizeMo },
        { key: "Mode d'analyse", val: "Simulation locale - serveur Python indisponible" },
        { key: "Photos détectées", val: `${photos.length} photo(s) simulée(s)` },
        { key: "Texte OCR", val: ocrText === "(aucun texte détecté)" ? "Aucun texte" : `${ocrText.length} caractères extraits` },
      ],
      zones: generateMockZones(isAI),
    };
  };

  const runAnalysis = async () => {
    setAnalyzing(true); setProgress(0); setResult(null);
 
    // Labels de progression (affichés pendant l'attente)
    const progressSteps = [
      { p: 5,  l: "Envoi de l'image au serveur…" },
      { p: 15, l: "Chargement des modèles (ViT + Florence-2)…" },
      { p: 25, l: "Détection des photos contenues (OpenCV)…" },
      { p: 40, l: "Extraction du texte par OCR (Florence-2 <OCR>)…" },
      { p: 55, l: "Génération de la description (Florence-2 <CAPTION>)…" },
      { p: 70, l: "Classification IA vs Réel (ViT AI-image-detector)…" },
      { p: 80, l: "Analyse des zones et contours (OpenCV)…" },
      { p: 90, l: "Finalisation du rapport…" },
    ];
 
    // Progression animée pendant l'appel réseau
    let step = 0;
    const progressInterval = setInterval(() => {
      if (step < progressSteps.length) {
        setProgress(progressSteps[step].p);
        step++;
      }
    }, 2000);
 
    try {
      // Envoyer l'image au serveur Python
      const formData = new FormData();
      formData.append("file", file);
 
      const response = await fetch(`${API_URL}/analyze`, {
        method: "POST",
        body: formData,
      });
 
      clearInterval(progressInterval);
 
      if (!response.ok) {
        throw new Error(`Erreur serveur : ${response.status}`);
      }
 
      // Le JSON retourné a le MÊME format que la simulation
      const data = await response.json();
 
      setProgress(100);
      setTimeout(() => {
        setResult(data);
        setAnalyzing(false);
        setActiveResultTab("verdict");
      }, 500);
 
    } catch (error) {
      clearInterval(progressInterval);
      console.error("Erreur :", error);

      const isNetworkError =
        error instanceof TypeError ||
        /failed to fetch|networkerror|load failed/i.test(String(error?.message || ""));

      if (!isNetworkError) {
        setAnalyzing(false);
        setProgress(0);
        alert(
          "Erreur : le serveur Python a renvoyé une erreur.\n\n" +
          "Vérifiez la console du backend dans le dossier 'pipeline', puis réessayez."
        );
        return;
      }

      const fallbackResult = buildMockAnalysis(file);
      setProgress(100);
      setTimeout(() => {
        setResult(fallbackResult);
        setAnalyzing(false);
        setActiveResultTab("verdict");
      }, 500);
    }
  };

  const analysisLabels = [
    { t: 0, l: "Chargement des modèles (ViT + Florence-2)…" },
    { t: 8, l: "Pré-traitement de l'image…" },
    { t: 18, l: "Détection des photos contenues (OpenCV + Florence-2 grounding)…" },
    { t: 30, l: "Extraction du texte par OCR (Florence-2 <OCR>)…" },
    { t: 45, l: "Génération de la description (Florence-2 <CAPTION>)…" },
    { t: 58, l: "Classification IA vs Réel (ViT AI-image-detector)…" },
    { t: 72, l: "Analyse des zones et contours (OpenCV)…" },
    { t: 85, l: "Calcul des scores de confiance…" },
    { t: 95, l: "Génération du rapport complet…" },
  ];
  const currentLabel = analysisLabels.filter((s) => progress >= s.t).pop();

  const riskLevel = (score) => {
    if (score >= 70) return { label: "Élevé", color: C.error };
    if (score >= 40) return { label: "Modéré", color: C.warning };
    return { label: "Faible", color: C.success };
  };

  const resetDemo = () => { setFile(null); setPreview(null); setResult(null); setAnalyzing(false); setProgress(0); setActiveResultTab("verdict"); };

  const resultTabs = [
    { id: "verdict", label: isEn ? "AI verdict" : "Verdict IA", icon: "◈" },
    { id: "ocr", label: isEn ? "OCR text" : "Texte OCR", icon: "◆" },
    { id: "description", label: isEn ? "Description" : "Description", icon: "◇" },
    { id: "photos", label: isEn ? "Detected photos" : "Photos détectées", icon: "◆" },
    { id: "details", label: isEn ? "Technical details" : "Détails techniques", icon: "◈" },
  ];

  return (
    <div style={{ paddingTop: 68, minHeight: "100vh", background: C.cream }}>
      <div style={{ maxWidth: 720, margin: "0 auto", padding: "48px 24px" }}>
        <div style={{ textAlign: "center", marginBottom: 40 }}>
          <Label>{isEn ? "Full pipeline" : "Pipeline complet"}</Label>
          <h1 style={{ fontFamily: FONT.serif, fontWeight: 700, fontSize: "clamp(24px, 4vw, 36px)", color: C.navy, margin: "16px 0 0" }}>
            {isEn ? "Image analysis " : "Analyse d'image "}<span style={{ color: C.gold, fontStyle: "italic" }}>{isEn ? "with artificial intelligence" : "par intelligence artificielle"}</span>
          </h1>
          <Ornament width={50} style={{ margin: "20px auto" }} />
          <p style={{ fontFamily: FONT.sans, fontSize: 15, color: C.textLight, maxWidth: 560, margin: "0 auto", lineHeight: 1.75 }}>
            {isEn ? "Upload an image to simulate the full DIGIDEM pipeline: text extraction (OCR), automatic description, photo detection and AI-vs-real classification." : "Soumettez une image pour simuler le pipeline complet DIGIDEM : extraction de texte (OCR), description automatique, détection de photos et classification IA vs Réel."}
          </p>
        </div>

        {/* Pipeline steps */}
        <FadeIn>
          <div style={{ background: C.warmWhite, border: `1px solid ${C.borderLight}`, borderRadius: 8, padding: "20px 24px", marginBottom: 32 }}>
            <div style={{ display: "flex", flexWrap: "wrap", justifyContent: "center", gap: 6 }}>
              {[
                { n: "1", label: isEn ? "Photo detection" : "Détection photos", sub: "OpenCV + Florence-2" },
                { n: "2", label: "OCR", sub: "Florence-2 <OCR>" },
                { n: "3", label: isEn ? "Description" : "Description", sub: "Florence-2 <CAPTION>" },
                { n: "4", label: isEn ? "AI detection" : "Détection IA", sub: "ViT Classifier" },
              ].map((step, i) => (
                <div key={i} style={{ display: "flex", alignItems: "center", gap: 6 }}>
                  <div style={{ textAlign: "center" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                      <span style={{ width: 24, height: 24, borderRadius: "50%", background: C.navy, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: FONT.sans, fontSize: 11, fontWeight: 700, color: C.gold }}>{step.n}</span>
                      <span style={{ fontFamily: FONT.sans, fontSize: 13, fontWeight: 600, color: C.navy }}>{step.label}</span>
                    </div>
                    <span style={{ fontFamily: FONT.sans, fontSize: 10, color: C.textMuted, marginLeft: 30, display: "block" }}>{step.sub}</span>
                  </div>
                  {i < 3 && <span style={{ color: C.gold, margin: "0 4px", fontSize: 14, marginBottom: 14 }}>→</span>}
                </div>
              ))}
            </div>
          </div>
        </FadeIn>

        {/* Drop zone */}
        {!preview && (
          <div
            onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
            onDragLeave={() => setDragOver(false)}
            onDrop={(e) => { e.preventDefault(); setDragOver(false); handleFile(e.dataTransfer.files[0]); }}
            onClick={() => fileInputRef.current?.click()}
            style={{ border: `2px dashed ${dragOver ? C.gold : C.border}`, borderRadius: 8, padding: "64px 32px", textAlign: "center", cursor: "pointer", background: dragOver ? `${C.gold}06` : C.warmWhite, transition: "all 0.3s" }}>
            <input ref={fileInputRef} type="file" accept="image/*" style={{ display: "none" }} onChange={(e) => handleFile(e.target.files[0])} />
            <svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke={dragOver ? C.gold : C.textMuted} strokeWidth="1.2" style={{ marginBottom: 16, transition: "stroke 0.3s" }}>
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" /><polyline points="17 8 12 3 7 8" /><line x1="12" y1="3" x2="12" y2="15" />
            </svg>
            <p style={{ fontFamily: FONT.serif, fontWeight: 600, fontSize: 18, color: C.navy, margin: "0 0 6px" }}>{isEn ? "Drop an image here" : "Déposez une image ici"}</p>
            <p style={{ fontFamily: FONT.sans, fontSize: 13, color: C.textMuted }}>{isEn ? "or click to choose one - PNG, JPEG and WebP supported" : "ou cliquez pour sélectionner — formats PNG, JPEG, WebP acceptés"}</p>
          </div>
        )}

        {preview && (
          <div>
            {/* Image preview */}
            <div style={{ position: "relative", borderRadius: 8, overflow: "hidden", border: `1px solid ${C.border}`, marginBottom: 24 }}>
              <img src={preview} alt={isEn ? "Submitted image" : "Image soumise"} style={{ width: "100%", maxHeight: 380, objectFit: "cover", display: "block" }} />
              {result?.isAI && <div style={{ position: "absolute", inset: 0, background: `radial-gradient(ellipse at 38% 32%, ${C.error}40 0%, transparent 45%), radial-gradient(ellipse at 62% 58%, ${C.warning}30 0%, transparent 38%)`, pointerEvents: "none", animation: "heatPulse 3s ease-in-out infinite alternate" }} />}
              {analyzing && <div style={{ position: "absolute", inset: 0, overflow: "hidden", pointerEvents: "none" }}><div style={{ position: "absolute", left: 0, right: 0, height: 2, background: `linear-gradient(90deg, transparent, ${C.gold}, transparent)`, boxShadow: `0 0 12px ${C.gold}88`, animation: "scanDown 2.5s ease-in-out infinite" }} /></div>}
              <div style={{ position: "absolute", bottom: 10, left: 10, background: `${C.navy}cc`, backdropFilter: "blur(8px)", borderRadius: 4, padding: "4px 10px", fontFamily: FONT.sans, fontSize: 11, color: C.goldLight }}>{file?.name} — {(file?.size / 1024).toFixed(0)} Ko</div>
            </div>

            {/* Launch button */}
            {!result && !analyzing && (
              <button onClick={runAnalysis} style={{ width: "100%", background: C.navy, color: C.gold, border: "none", borderRadius: 4, padding: "16px", fontFamily: FONT.sans, fontWeight: 600, fontSize: 15, cursor: "pointer", letterSpacing: 0.5, boxShadow: `0 4px 16px ${C.navy}22`, marginBottom: 20 }}>
                {isEn ? "Run the full pipeline" : "Lancer le pipeline complet"}
              </button>
            )}

            {/* Progress */}
            {analyzing && (
              <div style={{ background: C.warmWhite, border: `1px solid ${C.border}`, borderRadius: 6, padding: 24, marginBottom: 20 }}>
                <p style={{ fontFamily: FONT.sans, fontSize: 13, color: C.navyLight, margin: "0 0 14px" }}>
                  <span style={{ color: C.gold, marginRight: 8 }}>◆</span>{currentLabel?.l}
                </p>
                <div style={{ height: 4, background: C.borderLight, borderRadius: 2, overflow: "hidden" }}>
                  <div style={{ height: "100%", width: `${progress}%`, background: `linear-gradient(90deg, ${C.gold}, ${C.goldLight})`, borderRadius: 2, transition: "width 0.5s ease" }} />
                </div>
                <p style={{ fontFamily: FONT.sans, fontSize: 11, color: C.textMuted, marginTop: 8, textAlign: "right" }}>{progress} %</p>
              </div>
            )}

            {/* Ã¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢Â RESULTS Ã¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢Â */}
            {result && (
              <div>
                {/* Result tabs */}
                <div style={{ display: "flex", gap: 4, marginBottom: 20, flexWrap: "wrap" }}>
                  {resultTabs.map((tab) => (
                    <button key={tab.id} onClick={() => setActiveResultTab(tab.id)}
                      style={{
                        flex: "1 1 auto", minWidth: 100,
                        background: activeResultTab === tab.id ? C.warmWhite : "transparent",
                        border: activeResultTab === tab.id ? `1.5px solid ${C.gold}` : `1px solid ${C.borderLight}`,
                        borderRadius: 6, padding: "10px 12px",
                        fontFamily: FONT.sans, fontWeight: activeResultTab === tab.id ? 600 : 400,
                        fontSize: 12, color: activeResultTab === tab.id ? C.navy : C.textMuted,
                        cursor: "pointer", transition: "all 0.3s",
                        boxShadow: activeResultTab === tab.id ? `0 2px 8px ${C.gold}12` : "none",
                      }}>
                      <span style={{ color: C.gold, marginRight: 4 }}>{tab.icon}</span>{tab.label}
                    </button>
                  ))}
                </div>

                {/* Ã¢â€â‚¬Ã¢â€â‚¬ TAB: VERDICT Ã¢â€â‚¬Ã¢â€â‚¬ */}
                {activeResultTab === "verdict" && (
                  <div style={{ animation: "fadeSlideIn 0.3s ease" }}>
                    <div style={{ background: C.warmWhite, border: `1.5px solid ${result.isAI ? C.error : C.success}44`, borderRadius: 6, padding: "32px 28px", textAlign: "center", marginBottom: 20 }}>
                      <Label variant={result.isAI ? "error" : "success"}>{result.isAI ? "Génération artificielle probable" : "Image authentique probable"}</Label>
                      <div style={{ fontFamily: FONT.serif, fontWeight: 700, fontSize: 52, color: C.navy, margin: "16px 0 4px", lineHeight: 1 }}>
                        {result.confidence}<span style={{ fontSize: 24, color: C.textMuted }}> %</span>
                      </div>
                      <p style={{ fontFamily: FONT.sans, fontSize: 13, color: C.textMuted, margin: 0 }}>Indice de confiance du classifieur ViT</p>
                    </div>

                    {/* Score bars */}
                    <div style={{ background: C.warmWhite, border: `1px solid ${C.border}`, borderRadius: 6, padding: 24, marginBottom: 16 }}>
                      <h4 style={{ fontFamily: FONT.serif, fontWeight: 700, fontSize: 15, color: C.navy, margin: "0 0 18px" }}>Scores de classification (ViT)</h4>
                      {[
                        { label: "Score IA (artificial)", value: result.probAI, color: C.error },
                        { label: "Score Réel (human)", value: result.probReal, color: C.success },
                      ].map((bar, i) => (
                        <div key={i} style={{ marginBottom: i === 0 ? 14 : 0 }}>
                          <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
                            <span style={{ fontFamily: FONT.sans, fontSize: 13, color: C.textLight }}>{bar.label}</span>
                            <span style={{ fontFamily: FONT.sans, fontSize: 13, fontWeight: 600, color: bar.color }}>{bar.value} %</span>
                          </div>
                          <div style={{ height: 8, background: C.borderLight, borderRadius: 4, overflow: "hidden" }}>
                            <div style={{ height: "100%", width: `${bar.value}%`, background: `linear-gradient(90deg, ${bar.color}, ${bar.color}88)`, borderRadius: 4, transition: "width 1.2s ease" }} />
                          </div>
                        </div>
                      ))}
                      <p style={{ fontFamily: FONT.sans, fontSize: 11, color: C.textMuted, marginTop: 14, margin: "14px 0 0" }}>
                        Modèle : <em>umm-maybe/AI-image-detector</em> — Architecture : Vision Transformer (ViT)
                      </p>
                    </div>

                    {/* Zone analysis */}
                    <div style={{ background: C.warmWhite, border: `1px solid ${C.border}`, borderRadius: 6, padding: 24 }}>
                      <h4 style={{ fontFamily: FONT.serif, fontWeight: 700, fontSize: 15, color: C.navy, margin: "0 0 16px" }}>Analyse par zone (OpenCV)</h4>
                      {result.zones.map((z, i) => {
                        const risk = riskLevel(z.score);
                        return (
                          <div key={i} style={{ marginBottom: i < result.zones.length - 1 ? 16 : 0 }}>
                            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6 }}>
                              <span style={{ fontFamily: FONT.sans, fontSize: 13, color: C.text }}>{z.label}</span>
                              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                                <span style={{ fontFamily: FONT.sans, fontSize: 10, fontWeight: 600, color: risk.color, textTransform: "uppercase", letterSpacing: 1 }}>{risk.label}</span>
                                <span style={{ fontFamily: FONT.sans, fontSize: 13, fontWeight: 600, color: C.navy }}>{z.score} %</span>
                              </div>
                            </div>
                            <div style={{ height: 4, background: C.borderLight, borderRadius: 2, overflow: "hidden" }}>
                              <div style={{ height: "100%", width: `${z.score}%`, background: risk.color, borderRadius: 2, transition: "width 1.2s ease" }} />
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                )}

                {/* Ã¢â€â‚¬Ã¢â€â‚¬ TAB: OCR Ã¢â€â‚¬Ã¢â€â‚¬ */}
                {activeResultTab === "ocr" && (
                  <div style={{ animation: "fadeSlideIn 0.3s ease" }}>
                    <div style={{ background: C.warmWhite, border: `1px solid ${C.border}`, borderRadius: 6, padding: 24, marginBottom: 16 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
                        <h4 style={{ fontFamily: FONT.serif, fontWeight: 700, fontSize: 15, color: C.navy, margin: 0 }}>Texte extrait par OCR</h4>
                        <Label variant="navy">Florence-2 &lt;OCR&gt;</Label>
                      </div>
                      <div style={{ background: `${C.navy}08`, border: `1px solid ${C.borderLight}`, borderRadius: 6, padding: 20 }}>
                        <pre style={{ fontFamily: "'Source Code Pro', 'Courier New', monospace", fontSize: 13, color: C.text, lineHeight: 1.75, margin: 0, whiteSpace: "pre-wrap", wordBreak: "break-word" }}>
                          {result.ocrText}
                        </pre>
                      </div>
                      <div style={{ display: "flex", justifyContent: "space-between", marginTop: 14 }}>
                        <span style={{ fontFamily: FONT.sans, fontSize: 12, color: C.textMuted }}>{result.ocrCharCount} caractères extraits</span>
                        <span style={{ fontFamily: FONT.sans, fontSize: 12, color: C.textMuted }}>Modèle : Florence-2-large</span>
                      </div>
                    </div>
                    <div style={{ background: `${C.gold}08`, border: `1px solid ${C.gold}33`, borderRadius: 6, padding: 16 }}>
                      <p style={{ fontFamily: FONT.sans, fontSize: 12, color: C.textLight, lineHeight: 1.7, margin: 0 }}>
                        <strong style={{ color: C.navy }}>Comment ça marche —</strong> Florence-2 utilise la tÃƒÂ¢che <code style={{ background: `${C.navy}0a`, padding: "1px 5px", borderRadius: 3, fontSize: 11 }}>&lt;OCR&gt;</code> pour extraire tout le texte visible dans l'image, y compris les titres, légendes, noms d'utilisateurs et URLs. Ce texte est ensuite analysé pour détecter des marqueurs typiques de désinformation.
                      </p>
                    </div>
                  </div>
                )}

                {/* Ã¢â€â‚¬Ã¢â€â‚¬ TAB: DESCRIPTION Ã¢â€â‚¬Ã¢â€â‚¬ */}
                {activeResultTab === "description" && (
                  <div style={{ animation: "fadeSlideIn 0.3s ease" }}>
                    <div style={{ background: C.warmWhite, border: `1px solid ${C.border}`, borderRadius: 6, padding: 24, marginBottom: 16 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
                        <h4 style={{ fontFamily: FONT.serif, fontWeight: 700, fontSize: 15, color: C.navy, margin: 0 }}>Description générée automatiquement</h4>
                        <Label variant="navy">Florence-2 &lt;CAPTION&gt;</Label>
                      </div>

                      {/* Global description */}
                      <div style={{ background: `${C.navy}06`, border: `1px solid ${C.borderLight}`, borderRadius: 6, padding: 20, marginBottom: 16 }}>
                        <div style={{ fontFamily: FONT.sans, fontSize: 11, fontWeight: 600, color: C.gold, letterSpacing: 1, textTransform: "uppercase", marginBottom: 8 }}>Description globale (MORE_DETAILED_CAPTION)</div>
                        <p style={{ fontFamily: FONT.sans, fontSize: 14, color: C.text, lineHeight: 1.75, margin: 0, fontStyle: "italic" }}>
                          "{result.description}"
                        </p>
                      </div>

                      {/* Per-photo descriptions */}
                      <div style={{ fontFamily: FONT.sans, fontSize: 11, fontWeight: 600, color: C.gold, letterSpacing: 1, textTransform: "uppercase", marginBottom: 10 }}>Descriptions par photo extraite (CAPTION)</div>
                      {result.photos.map((photo, i) => (
                        <div key={i} style={{ display: "flex", gap: 12, alignItems: "flex-start", padding: "12px 0", borderBottom: i < result.photos.length - 1 ? `1px solid ${C.borderLight}` : "none" }}>
                          <div style={{ width: 40, height: 40, borderRadius: 4, background: C.navy, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                            <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke={C.gold} strokeWidth="1.5"><rect x="3" y="3" width="18" height="18" rx="2" /><circle cx="8.5" cy="8.5" r="1.5" /><path d="M21 15l-5-5L5 21" /></svg>
                          </div>
                          <div>
                            <div style={{ fontFamily: FONT.sans, fontSize: 12, fontWeight: 600, color: C.navy }}>{photo.filename} <span style={{ fontWeight: 400, color: C.textMuted }}>({photo.size})</span></div>
                            <p style={{ fontFamily: FONT.sans, fontSize: 13, color: C.textLight, lineHeight: 1.6, margin: "4px 0 0" }}>"{photo.caption}"</p>
                          </div>
                        </div>
                      ))}
                    </div>
                    <div style={{ background: `${C.gold}08`, border: `1px solid ${C.gold}33`, borderRadius: 6, padding: 16 }}>
                      <p style={{ fontFamily: FONT.sans, fontSize: 12, color: C.textLight, lineHeight: 1.7, margin: 0 }}>
                        <strong style={{ color: C.navy }}>Comment ça marche —</strong> Florence-2 génère une description textuelle de l'image via les tÃƒÂ¢ches <code style={{ background: `${C.navy}0a`, padding: "1px 5px", borderRadius: 3, fontSize: 11 }}>&lt;CAPTION&gt;</code> et <code style={{ background: `${C.navy}0a`, padding: "1px 5px", borderRadius: 3, fontSize: 11 }}>&lt;MORE_DETAILED_CAPTION&gt;</code>. Chaque photo extraite reçoit sa propre description.
                      </p>
                    </div>
                  </div>
                )}

                {/* Ã¢â€â‚¬Ã¢â€â‚¬ TAB: PHOTOS DÉTECTÉES Ã¢â€â‚¬Ã¢â€â‚¬ */}
                {activeResultTab === "photos" && (
                  <div style={{ animation: "fadeSlideIn 0.3s ease" }}>
                    <div style={{ background: C.warmWhite, border: `1px solid ${C.border}`, borderRadius: 6, padding: 24, marginBottom: 16 }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
                        <h4 style={{ fontFamily: FONT.serif, fontWeight: 700, fontSize: 15, color: C.navy, margin: 0 }}>Photos détectées et extraites</h4>
                        <Label variant="navy">OpenCV + Florence-2</Label>
                      </div>
                      <p style={{ fontFamily: FONT.sans, fontSize: 13, color: C.textMuted, margin: "0 0 20px" }}>
                        {result.photos.length} photo(s) détectée(s) dans le document via contours OpenCV et Florence-2 grounding.
                      </p>

                      {result.photos.map((photo, i) => (
                        <div key={i} style={{ background: `${C.navy}05`, border: `1px solid ${C.borderLight}`, borderRadius: 6, padding: 20, marginBottom: i < result.photos.length - 1 ? 12 : 0 }}>
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
                            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                              <div style={{ width: 36, height: 36, borderRadius: 4, background: C.navy, display: "flex", alignItems: "center", justifyContent: "center" }}>
                                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke={C.gold} strokeWidth="1.5"><rect x="3" y="3" width="18" height="18" rx="2" /><circle cx="8.5" cy="8.5" r="1.5" /><path d="M21 15l-5-5L5 21" /></svg>
                              </div>
                              <div>
                                <div style={{ fontFamily: FONT.sans, fontSize: 14, fontWeight: 600, color: C.navy }}>{photo.filename}</div>
                                <div style={{ fontFamily: FONT.sans, fontSize: 11, color: C.textMuted }}>{photo.size} px</div>
                              </div>
                            </div>
                            <Label variant={photo.verdict === "IA" ? "error" : "success"}>
                              {photo.verdict === "IA" ? "IA détectée" : "Image réelle"}
                            </Label>
                          </div>

                          <div style={{ display: "flex", gap: 20, flexWrap: "wrap" }}>
                            <div style={{ flex: 1, minWidth: 150 }}>
                              <div style={{ fontFamily: FONT.sans, fontSize: 11, color: C.textMuted, marginBottom: 4 }}>Description (Florence-2)</div>
                              <p style={{ fontFamily: FONT.sans, fontSize: 13, color: C.text, lineHeight: 1.6, margin: 0, fontStyle: "italic" }}>{photo.caption}</p>
                            </div>
                            <div style={{ minWidth: 140 }}>
                              <div style={{ fontFamily: FONT.sans, fontSize: 11, color: C.textMuted, marginBottom: 8 }}>Verdict : <strong style={{ color: photo.verdict === "IA" ? C.error : C.success }}>{photo.verdict === "IA" ? "🤖 IA" : "📰 RÉEL"}</strong> ({(photo.confidence * 100).toFixed(1)}%)</div>
                              <div style={{ marginBottom: 6 }}>
                                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 3 }}>
                                  <span style={{ fontFamily: FONT.sans, fontSize: 11, color: C.textMuted }}>Score IA</span>
                                  <span style={{ fontFamily: FONT.sans, fontSize: 11, fontWeight: 600, color: C.error }}>{(photo.probAI * 100).toFixed(1)}%</span>
                                </div>
                                <div style={{ height: 4, background: C.borderLight, borderRadius: 2, overflow: "hidden" }}>
                                  <div style={{ height: "100%", width: `${photo.probAI * 100}%`, background: C.error, borderRadius: 2 }} />
                                </div>
                              </div>
                              <div>
                                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 3 }}>
                                  <span style={{ fontFamily: FONT.sans, fontSize: 11, color: C.textMuted }}>Score Réel</span>
                                  <span style={{ fontFamily: FONT.sans, fontSize: 11, fontWeight: 600, color: C.success }}>{(photo.probHuman * 100).toFixed(1)}%</span>
                                </div>
                                <div style={{ height: 4, background: C.borderLight, borderRadius: 2, overflow: "hidden" }}>
                                  <div style={{ height: "100%", width: `${photo.probHuman * 100}%`, background: C.success, borderRadius: 2 }} />
                                </div>
                              </div>
                            </div>
                          </div>
                        </div>
                      ))}
                    </div>
                    <div style={{ background: `${C.gold}08`, border: `1px solid ${C.gold}33`, borderRadius: 6, padding: 16 }}>
                      <p style={{ fontFamily: FONT.sans, fontSize: 12, color: C.textLight, lineHeight: 1.7, margin: 0 }}>
                        <strong style={{ color: C.navy }}>Comment ça marche —</strong> Les photos sont détectées via deux méthodes combinées : les contours OpenCV (Canny + findContours) et le Florence-2 grounding (<code style={{ background: `${C.navy}0a`, padding: "1px 5px", borderRadius: 3, fontSize: 11 }}>&lt;CAPTION_TO_PHRASE_GROUNDING&gt;</code> avec les requêtes "a photograph", "a photo image"). Les boîtes sont fusionnées puis chaque photo extraite est classifiée par le ViT.
                      </p>
                    </div>
                  </div>
                )}

                {/* Ã¢â€â‚¬Ã¢â€â‚¬ TAB: DÉTAILS TECHNIQUES Ã¢â€â‚¬Ã¢â€â‚¬ */}
                {activeResultTab === "details" && (
                  <div style={{ animation: "fadeSlideIn 0.3s ease" }}>
                    <div style={{ background: C.warmWhite, border: `1px solid ${C.border}`, borderRadius: 6, padding: 24, marginBottom: 16 }}>
                      <h4 style={{ fontFamily: FONT.serif, fontWeight: 700, fontSize: 15, color: C.navy, margin: "0 0 16px" }}>Métadonnées analysées</h4>
                      {result.metadata.map((m, i) => (
                        <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 0", borderBottom: i < result.metadata.length - 1 ? `1px solid ${C.borderLight}` : "none" }}>
                          <span style={{ fontFamily: FONT.sans, fontSize: 13, color: C.textMuted }}>{m.key}</span>
                          <span style={{ fontFamily: FONT.sans, fontSize: 13, color: C.text, fontWeight: 500, textAlign: "right" }}>{m.val}</span>
                        </div>
                      ))}
                    </div>

                    {/* Modèles utilisés */}
                    <div style={{ background: C.warmWhite, border: `1px solid ${C.border}`, borderRadius: 6, padding: 24, marginBottom: 16 }}>
                      <h4 style={{ fontFamily: FONT.serif, fontWeight: 700, fontSize: 15, color: C.navy, margin: "0 0 16px" }}>Modèles du pipeline</h4>
                      {[
                        { name: "Florence-2-large", role: "OCR, description, détection de photos (grounding)", task: "<OCR>, <CAPTION>, <MORE_DETAILED_CAPTION>, <CAPTION_TO_PHRASE_GROUNDING>" },
                        { name: "umm-maybe/AI-image-detector", role: "Classification IA vs Réel", task: "image-classification (ViT)" },
                        { name: "OpenCV", role: "Détection de contours, extraction de photos", task: "Canny, findContours, boundingRect" },
                      ].map((model, i) => (
                        <div key={i} style={{ padding: "14px 0", borderBottom: i < 2 ? `1px solid ${C.borderLight}` : "none" }}>
                          <div style={{ fontFamily: FONT.sans, fontSize: 14, fontWeight: 600, color: C.navy }}>{model.name}</div>
                          <div style={{ fontFamily: FONT.sans, fontSize: 12, color: C.textLight, marginTop: 2 }}>{model.role}</div>
                          <div style={{ fontFamily: FONT.sans, fontSize: 11, color: C.textMuted, marginTop: 4 }}>
                            TÃƒÂ¢ches : <code style={{ background: `${C.navy}0a`, padding: "1px 5px", borderRadius: 3, fontSize: 10 }}>{model.task}</code>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                )}

                {/* Note + reset */}
                <div style={{ background: `${C.gold}08`, border: `1px solid ${C.gold}33`, borderRadius: 6, padding: 20, marginTop: 20, marginBottom: 20 }}>
                  <p style={{ fontFamily: FONT.sans, fontSize: 12, color: C.textLight, lineHeight: 1.75, margin: 0 }}>
                    <strong style={{ color: C.navy }}>Note scientifique —</strong> Ce module est une simulation pédagogique illustrant le pipeline complet de détection DIGIDEM. En production, l'image est traitée par Florence-2 (OCR + description + détection de photos), OpenCV (contours) et le classifieur ViT (<em>umm-maybe/AI-image-detector</em>). Les résultats présentés ici sont simulés à des fins de démonstration.
                  </p>
                </div>

                <button onClick={resetDemo} style={{ width: "100%", background: "transparent", color: C.navy, border: `1px solid ${C.border}`, borderRadius: 4, padding: "14px", fontFamily: FONT.sans, fontWeight: 600, fontSize: 14, cursor: "pointer", transition: "border-color 0.3s" }}
                  onMouseEnter={(e) => (e.target.style.borderColor = C.gold)}
                  onMouseLeave={(e) => (e.target.style.borderColor = C.border)}>
                  {isEn ? "Analyse another image" : "Soumettre une autre image"}
                </button>
              </div>
            )}
          </div>
        )}

        <FadeIn>
          <div style={{ marginTop: 48, background: C.warmWhite, border: `1px solid ${C.border}`, borderRadius: 6, padding: 28 }}>
            <h4 style={{ fontFamily: FONT.serif, fontWeight: 700, fontSize: 16, color: C.navy, margin: "0 0 16px" }}>{isEn ? "About the DIGIDEM pipeline" : "À propos du pipeline DIGIDEM"}</h4>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(130px, 1fr))", gap: 16 }}>
              {[{ n: "400", l: isEn ? "Images + annotated texts" : "Images + textes annotés" }, { n: "3", l: isEn ? "Combined AI models" : "Modèles IA combinés" }, { n: "4", l: isEn ? "Pipeline stages" : "Étapes du pipeline" }, { n: "CC BY", l: isEn ? "Open licence" : "Licence ouverte" }].map((d, i) => (
                <div key={i} style={{ textAlign: "center", padding: "8px 0" }}>
                  <div style={{ fontFamily: FONT.serif, fontWeight: 700, fontSize: 24, color: C.navy }}>{d.n}</div>
                  <div style={{ fontFamily: FONT.sans, fontSize: 12, color: C.textMuted, marginTop: 2 }}>{d.l}</div>
                </div>
              ))}
            </div>
          </div>
        </FadeIn>
      </div>
    </div>
  );
}

/* Ã¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢Â
   APPLICATION PRINCIPALE
   Ã¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢ÂÃ¢â€¢Â */
export default function DigidemApp() {
  const [page, setPage] = useState("home");
  const [lang, setLang] = useState("fr");
  useEffect(() => { window.scrollTo({ top: 0, behavior: "smooth" }); }, [page]);

  return (
    <div style={{ minHeight: "100vh", background: C.cream }}>
      <style>{`
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { background: ${C.cream}; }
        @keyframes fadeSlideIn { from { opacity: 0; transform: translateY(16px); } to { opacity: 1; transform: translateY(0); } }
        @keyframes scanDown { 0% { top: -2px; } 100% { top: 100%; } }
        @keyframes heatPulse { 0% { opacity: 0.5; } 100% { opacity: 0.85; } }
        ::-webkit-scrollbar { width: 6px; }
        ::-webkit-scrollbar-track { background: ${C.cream}; }
        ::-webkit-scrollbar-thumb { background: ${C.border}; border-radius: 3px; }
        ::-webkit-scrollbar-thumb:hover { background: ${C.gold}88; }
        ::selection { background: ${C.gold}33; color: ${C.navy}; }
        @media (max-width: 768px) {
          .nav-desktop { display: none !important; }
          .nav-badge { display: none !important; }
          .nav-mobile-btn { display: block !important; }
        }
        @media (min-width: 769px) {
          .nav-mobile-btn { display: none !important; }
        }
      `}</style>
      <Nav currentPage={page} setPage={setPage} lang={lang} setLang={setLang} />
      {page === "home" && <HomePage setPage={setPage} lang={lang} />}
      {page === "visu" && <VisuPage lang={lang} />}
      {page === "demo" && <DemoPage lang={lang} />}
    </div>
  );
}










