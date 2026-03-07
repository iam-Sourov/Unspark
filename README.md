# Unspark

Unspark is a powerful, client-side web application built with Next.js, React, Tailwind CSS, and PptxGenJS. It allows users to instantly convert complex, styled HTML components into polished, downloadable Microsoft PowerPoint (.pptx) presentations—directly from the browser!

## ✨ Features
- **Live Markdown & HTML Editing**: Paste any HTML structure directly into the source boxes and get real-time previews.
- **Pixel-Perfect Fidelity**: Complex components (like Flexbox layouts, background colors, custom images, absolute positioning) are flawlessly captured and processed.
- **Dynamic Cross-Origin Support**: Automatically mitigates CORS restrictions when trying to fetch external SVGs/Fonts (like FontAwesome or Google Fonts) directly onto your presentation template engine.
- **Instant Client-Side Generation**: Entirely private and localized. All PowerPoint rendering logic is run natively inside your device's browser relying on `pptxgenjs` and `html2canvas` — no proprietary backend needed.
- **Chart.js & Custom Web Components**: Equipped with strategic cloning hooks and rendering delays that perfectly replicate canvas animations, complex graphs, and web-fonts inside your templates! 

## 🚀 Getting Started

First, run the development server:

```bash
npm run dev
# or
yarn dev
# or
pnpm dev
# or
bun dev
```

Open [http://localhost:3000](http://localhost:3000) with your browser to see the outcome.

## 🛠️ Tech Stack
- **Framework**: [Next.js](https://nextjs.org/)
- **Styling**: [Tailwind CSS](https://tailwindcss.com/)
- **Core Compiler**: [PptxGenJS](https://gitbrent.github.io/PptxGenJS/)
- **Visual Capture**: [html2canvas](https://html2canvas.hertzen.com/)

## 🎨 How to Use
1. Clone the repository and run `npm install`.
2. Start the local server `npm run dev`.
3. Click '+ Add Slide' to generate a new slide.
4. Input HTML into the "HTML Sources" editor. 
5. See changes natively in the isolated **Live Preview** iframe.
6. Click the brilliant blue **Download PPTX** button to safely convert and export!
