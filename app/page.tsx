"use client";

import React, { useRef, useState } from "react";
import pptxgen from "pptxgenjs";
import html2canvas from "html2canvas";

export default function Page() {
  const containerRef = useRef<HTMLDivElement>(null);
  const [isGenerating, setIsGenerating] = useState(false);

  // Default initial slides
  const [slidesHtml, setSlidesHtml] = useState<string[]>([
    `<div class="slide-container" style="background-color: #ffffff;">\n  <div style="padding: 64px; display: flex; flex-direction: column; gap: 24px; min-height: 500px;">\n    <h1 style="color: #1e3a8a; font-size: 48px; font-weight: 800;">Welcome to Unspark</h1>\n    <p style="color: #374151; font-size: 20px;">This tool extracts HTML components and natively constructs PowerPoint presentations.</p>\n    <p style="color: #4b5563; font-size: 16px; font-weight: bold; background-color: #f3f4f6; padding: 16px; border-radius: 8px; display: inline-block;">\n      You can edit this HTML directly in the source boxes!\n    </p>\n  </div>\n</div>`,
    `<div class="slide-container" style="background-color: #1e293b;">\n  <div style="padding: 64px; text-align: center; min-height: 500px;">\n    <h2 style="color: #f8fafc; font-size: 36px; font-weight: bold; margin-bottom: 24px;">Dynamic Image Handling</h2>\n    <img src="https://images.unsplash.com/photo-1551033406-611cf9a28f67?ixlib=rb-4.0.3&auto=format&fit=crop&w=600&q=80" alt="Code screenshot"\n      style="border-radius: 8px; max-width: 100%; display: inline-block;" width="600" crossOrigin="anonymous" />\n  </div>\n</div>`
  ]);

  const handleHtmlChange = (index: number, value: string) => {
    const newSlides = [...slidesHtml];
    newSlides[index] = value;
    setSlidesHtml(newSlides);
  };

  const addSlide = () => {
    setSlidesHtml([
      ...slidesHtml,
      `<div class="slide-container" style="background-color: #ffffff;">\n  <div style="padding: 64px; min-height: 500px;">\n    <h1 style="color: #000000; font-size: 32px;">New Slide</h1>\n    <p style="color: #4b5563; font-size: 16px;">Add your content here...</p>\n  </div>\n</div>`
    ]);
  };

  const removeSlide = (index: number) => {
    const newSlides = slidesHtml.filter((_, i) => i !== index);
    setSlidesHtml(newSlides);
  };

  // Helper to ensure cross-origin IFrames can be strictly read by html2canvas for Pseudo-Elements (FontAwesome)
  const processHtmlForCors = (rawHtml: string) => {
    let processed = rawHtml;
    // Add crossorigin="anonymous" to <link> tags if not present, and specifically replace FontAwesome CSS with JS (SVG)
    processed = processed.replace(/<link([^>]+)>/gi, (match, p1) => {
      if (p1.includes('font-awesome') && p1.includes('css/all.min.css')) {
        // Swap FontAwesome CSS to JS so it uses SVGs instead of Web Fonts (which html2canvas struggles with)
        const newSrc = p1.match(/href="([^"]+)"/)?.[1]?.replace('css/all.min.css', 'js/all.min.js') || 'https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/js/all.min.js';
        return `<script src="${newSrc}" crossorigin="anonymous"></script>`;
      }

      if (p1.toLowerCase().includes('rel="stylesheet"') && !p1.toLowerCase().includes('crossorigin')) {
        return `<link${p1} crossorigin="anonymous">`;
      }
      return match;
    });
    // Add crossorigin="anonymous" to <img> tags if not present
    processed = processed.replace(/<img([^>]+)>/gi, (match, p1) => {
      if (!p1.toLowerCase().includes('crossorigin')) {
        return `<img${p1} crossorigin="anonymous">`;
      }
      return match;
    });
    return processed;
  };

  const handleDownloadPptx = async () => {
    if (!containerRef.current) return;
    setIsGenerating(true);

    try {
      const pptx = new pptxgen();
      pptx.layout = "LAYOUT_16x9"; // 10 x 5.625 inches standard size

      // Grab all slide iframes from the DOM
      const iframes = Array.from(document.querySelectorAll<HTMLIFrameElement>(".slide-preview-iframe"));

      for (const iframe of iframes) {
        const iframeDoc = iframe.contentDocument || iframe.contentWindow?.document;
        if (!iframeDoc) continue;

        const pptxSlide = pptx.addSlide();

        try {
          // Explicitly wait for all custom web-fonts (like FontAwesome/Google Fonts) to finish downloading
          if (iframeDoc.fonts && iframeDoc.fonts.ready) {
            await iframeDoc.fonts.ready;
          }

          // Gather all images in the iframe and ensure they are fully loaded
          const iframeImages = Array.from(iframeDoc.images);
          await Promise.all(
            iframeImages.map((img) => {
              if (img.complete) return Promise.resolve();
              return new Promise((resolve) => {
                img.onload = resolve;
                img.onerror = resolve;
              });
            })
          );

          // Give an extended buffer to permit embedded Javascript blocks (Chart.js / Tailwind script) to finish their execution
          await new Promise((resolve) => setTimeout(resolve, 1500));

          // Use html2canvas to create an exact visual replica inside iframe
          // target .slide-container if exists, otherwise body
          const targetEl = iframeDoc.querySelector('.slide-container') || iframeDoc.body;

          const canvas = await html2canvas(targetEl as HTMLElement, {
            scale: 2, // 2x scale for sharper images / text
            useCORS: true, // Crucial for loading external Unsplash/Genspark images
            allowTaint: false,
            backgroundColor: null, // Allow transparent backgrounds if any
            logging: false,
            windowWidth: 1280, // Fixes responsive Chart.js components expanding/condensing unexpectedly upon clone
            windowHeight: 720,
            onclone: async (cloneDoc) => {
              if (cloneDoc.fonts && cloneDoc.fonts.ready) {
                await cloneDoc.fonts.ready;
              }
              // Required secondary delay because the clone processes external stylesheets (like fontawesome URLs) fresh again
              await new Promise(r => setTimeout(r, 1000));
            }
          });

          // Convert canvas buffer to a base64 encoded PNG
          const imgData = canvas.toDataURL("image/png");

          // Map the exactly crafted visual image onto the slide's entire background/foreground
          pptxSlide.addImage({
            data: imgData,
            x: 0, // start top left
            y: 0, // start top left
            w: 10, // 10 inches wide
            h: 5.625, // 5.625 inches tall
          });
        } catch (err) {
          console.error("Failed to capture slide with html2canvas", err);
        }
      }

      await pptx.writeFile({ fileName: "converted_presentation.pptx" });

    } catch (error) {
      console.error("Error generating PPTX:", error);
      alert("An error occurred during PPTX conversion. Check console for details.");
    } finally {
      setIsGenerating(false);
    }
  };

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col font-sans">
      <header className="bg-white shadow-sm border-b px-8 py-4 flex items-center justify-between sticky top-0 z-20">
        <div>
          <h1 className="text-xl font-bold text-gray-800">Unspark</h1>
          <p className="text-sm text-gray-500">Convert standard web structures into PowerPoint slides</p>
        </div>
        <button
          onClick={handleDownloadPptx}
          disabled={isGenerating}
          className="bg-blue-600 hover:bg-blue-700 text-white px-5 py-2.5 rounded-md font-medium shadow transition-colors disabled:opacity-50 disabled:cursor-not-allowed flex items-center gap-2"
        >
          {isGenerating ? (
            <>
              <svg className="w-5 h-5 animate-spin text-white" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
                <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
              </svg>
              Generating...
            </>
          ) : (
            <>
              <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" /><polyline points="7 10 12 15 17 10" /><line x1="12" x2="12" y1="15" y2="3" /></svg>
              Download PPTX
            </>
          )}
        </button>
      </header>

      <main className="flex-1 p-8">
        <div className="max-w-7xl mx-auto grid grid-cols-1 lg:grid-cols-2 gap-12">

          {/* Left Column: Editor */}
          <div>
            <div className="flex justify-between items-center mb-6">
              <h2 className="text-lg font-semibold flex items-center gap-2 text-gray-700">
                <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="text-blue-500"><path d="M14.5 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7.5L14.5 2z" /><polyline points="14 2 14 8 20 8" /><line x1="16" x2="8" y1="13" y2="13" /><line x1="16" x2="8" y1="17" y2="17" /><line x1="10" x2="8" y1="9" y2="9" /></svg>
                HTML Sources
              </h2>
              <button
                onClick={addSlide}
                className="bg-blue-50 text-blue-600 hover:bg-blue-100 px-4 py-2 rounded-md font-medium text-sm transition-colors border border-blue-200 shadow-sm"
              >
                + Add Slide
              </button>
            </div>

            <div className="flex flex-col gap-6">
              {slidesHtml.map((html, idx) => (
                <div key={idx} className="border border-gray-200 rounded-xl bg-white shadow-sm overflow-hidden flex flex-col focus-within:ring-2 focus-within:ring-blue-500 transition-shadow">
                  <div className="bg-gray-50 border-b border-gray-200 px-4 py-3 flex justify-between items-center">
                    <span className="font-semibold text-gray-700 text-sm">Slide {idx + 1}</span>
                    <button
                      onClick={() => removeSlide(idx)}
                      disabled={slidesHtml.length === 1}
                      className="text-red-500 hover:text-red-700 text-xs font-semibold disabled:opacity-50 disabled:cursor-not-allowed uppercase tracking-wider"
                    >
                      Remove
                    </button>
                  </div>
                  <textarea
                    value={html}
                    onChange={(e) => handleHtmlChange(idx, e.target.value)}
                    className="w-full h-56 p-4 font-mono text-sm text-gray-800 bg-white focus:outline-none resize-y"
                    placeholder="<div class='slide-container'>...</div>"
                    spellCheck="false"
                  />
                </div>
              ))}
            </div>
          </div>

          {/* Right Column: Live Preview */}
          <div>
            <div className="sticky top-28">
              <h2 className="text-lg font-semibold mb-6 flex items-center gap-2 text-gray-700">
                <svg xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="text-blue-500"><rect width="18" height="18" x="3" y="3" rx="2" ry="2" /><line x1="3" x2="21" y1="9" y2="9" /><line x1="9" x2="9" y1="21" y2="9" /></svg>
                Live Preview
              </h2>

              <div
                ref={containerRef}
                className="flex flex-col gap-8 custom-scrollbar overflow-y-auto pr-2 pb-8"
                style={{ maxHeight: "calc(100vh - 10rem)", "--tw-scrollbar-width": "thin" } as any}
              >
                {slidesHtml.map((html, idx) => (
                  <div
                    key={idx}
                    className="border border-gray-200 shadow-xl rounded-xl overflow-hidden bg-white shrink-0 relative"
                    style={{ containerType: 'inline-size', width: '100%', aspectRatio: '16/9' }}
                  >
                    {/* The scaling ensures the 1280x720 slides fit nicely on the screen without massive scrolling */}
                    <iframe
                      className="slide-preview-iframe absolute top-0 left-0 pointer-events-none"
                      style={{
                        width: '1280px',
                        height: '720px',
                        border: 'none',
                        transform: 'scale(calc(100cqw / 1280))',
                        transformOrigin: 'top left'
                      }}
                      srcDoc={processHtmlForCors(html)}
                      title={`Slide ${idx + 1}`}
                    />
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>
      </main>
    </div>
  );
}
