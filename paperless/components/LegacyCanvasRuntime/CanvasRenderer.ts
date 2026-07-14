import type { LegacyRuntimeField, DebugFlags } from "./LegacyCanvasRuntime";

export function renderFrame(
  ctx: CanvasRenderingContext2D,
  backgroundImage: HTMLImageElement | null,
  fields: LegacyRuntimeField[],
  editingFieldId: string | null,
  debug: DebugFlags,
  pageWidth: number,
  pageHeight: number
): void {
  const canvasEl = ctx.canvas;

  // Stage 8: renderFrame() called
  console.log("[Stage 8] renderFrame() ENTERED");
  console.log("[Stage 8] Parameters:", {
    canvasWidth: canvasEl.width,
    canvasHeight: canvasEl.height,
    pageWidth,
    pageHeight,
    hasBackgroundImage: backgroundImage !== null,
    fieldCount: fields.length,
  });

  if (canvasEl.width !== pageWidth) {
    console.warn(`[Stage 8] Canvas width mismatch: HTML=${canvasEl.width}, param=${pageWidth}`);
  }
  if (canvasEl.height !== pageHeight) {
    console.warn(`[Stage 8] Canvas height mismatch: HTML=${canvasEl.height}, param=${pageHeight}`);
  }

  // Stage 9: Verify backgroundImage before drawImage
  if (backgroundImage) {
    console.log("[Stage 9] backgroundImage is non-null. Checking properties:");
    console.log("[Stage 9]   .width:", backgroundImage.width);
    console.log("[Stage 9]   .height:", backgroundImage.height);
    console.log("[Stage 9]   .naturalWidth:", backgroundImage.naturalWidth);
    console.log("[Stage 9]   .naturalHeight:", backgroundImage.naturalHeight);
    console.log("[Stage 9]   .complete:", backgroundImage.complete);
    console.log("[Stage 9]   .currentSrc:", backgroundImage.currentSrc);
    console.log("[Stage 9]   .nodeName:", backgroundImage.nodeName);
    console.log("[Stage 9]   .tagName:", (backgroundImage as HTMLElement).tagName);

    if (backgroundImage.naturalWidth === 0) {
      console.error("[Stage 9] FATAL: backgroundImage.naturalWidth is 0! Image has no decoded data.");
    }

    // Check if decode() method is available and use it
    if (typeof (backgroundImage as any).decode === "function") {
      console.log("[Stage 9] Calling backgroundImage.decode() to ensure image is fully decoded...");
      (backgroundImage as any).decode()
        .then(() => {
          console.log("[Stage 9] backgroundImage.decode() resolved successfully");
          console.log("[Stage 9] After decode - .width:", backgroundImage.width, ".height:", backgroundImage.height);
        })
        .catch((decodeErr: any) => {
          console.error("[Stage 9] backgroundImage.decode() FAILED:", decodeErr);
        });
    } else {
      console.log("[Stage 9] backgroundImage.decode() not available on this browser");
    }
  } else {
    console.log("[Stage 9] backgroundImage is null - will fill white");
  }

  // Canvas state check
  console.log("[Stage 9] Canvas state before any drawing:");
  console.log("[Stage 9]   globalAlpha:", ctx.globalAlpha);
  console.log("[Stage 9]   globalCompositeOperation:", ctx.globalCompositeOperation);
  const t = ctx.getTransform();
  console.log("[Stage 9]   transform:", { a: t.a, b: t.b, c: t.c, d: t.d, e: t.e, f: t.f });

  // Stage 8: Render order step by step

  // Step 1: clearRect
  console.log("[Stage 8] Step 1: ctx.clearRect(0, 0,", pageWidth, ",", pageHeight, ")");
  ctx.clearRect(0, 0, pageWidth, pageHeight);
  console.log("[Stage 8]   clearRect completed");

  // Step 2: drawImage (background)
  if (backgroundImage) {
    console.log("[Stage 9] Step 2: About to execute ctx.drawImage(backgroundImage, 0, 0,", pageWidth, ",", pageHeight, ")");
    console.log("[Stage 9]   Arguments: image=", backgroundImage, "dx=0, dy=0, dw=", pageWidth, "dh=", pageHeight);

    try {
      ctx.drawImage(backgroundImage, 0, 0, pageWidth, pageHeight);
      console.log("[Stage 9] ctx.drawImage EXECUTED SUCCESSFULLY (no exception thrown)");
    } catch (drawError) {
      console.error("[Stage 9] ctx.drawImage THREW EXCEPTION:", drawError);
    }

    // Stage 11: Immediate pixel check after drawImage
    console.log("[Stage 11] Checking first 4 pixels IMMEDIATELY after drawImage:");
    try {
      const pixels = ctx.getImageData(0, 0, 4, 4);
      console.log("[Stage 11]   4x4 pixel block RGBA values:");
      for (let y = 0; y < 4; y++) {
        const row: number[] = [];
        for (let x = 0; x < 4; x++) {
          const idx = (y * 4 + x) * 4;
          row.push(pixels.data[idx], pixels.data[idx + 1], pixels.data[idx + 2], pixels.data[idx + 3]);
        }
        console.log(`[Stage 11]   Row ${y}:`, row.join(","));
      }

      let allWhite = true;
      let hasRed = false;
      for (let i = 0; i < 16; i++) {
        const idx = i * 4;
        if (pixels.data[idx] !== 255 || pixels.data[idx + 1] !== 255 || pixels.data[idx + 2] !== 255) {
          allWhite = false;
        }
        if (pixels.data[idx] > 200 && pixels.data[idx + 1] < 80 && pixels.data[idx + 2] < 80) {
          hasRed = true;
        }
      }
      console.log("[Stage 11]   All white:", allWhite);
      console.log("[Stage 11]   Has red pixels:", hasRed);

      if (allWhite) {
        console.warn("[Stage 11] *** CRITICAL: All 16 checked pixels are white immediately after drawImage!");
        console.warn("[Stage 11] *** drawImage did NOT actually paint any non-white pixels.");
        console.warn("[Stage 11] *** This means drawImage appears to succeed but produces no visible output.");
      } else if (hasRed) {
        console.log("[Stage 11] *** Red pixels found! The red diagnostic rect IS rendering.");
        console.log("[Stage 11] *** If red is visible but background image is NOT, the image draws behind the red rect.");
      } else if (!allWhite) {
        console.log("[Stage 11] *** Non-white pixels found (not red, not white). Background might be rendering!");
      }
    } catch (e) {
      console.error("[Stage 11] getImageData failed immediately after drawImage:", e);
    }
  } else {
    console.log("[Stage 9] Step 2: No backgroundImage, filling white");
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, pageWidth, pageHeight);
  }

  // Step 3: Post-drawImage pixel check at known content position
  if (backgroundImage) {
    try {
      // Check at content boundary area (x=850 known from PNG analysis)
      const contentPixels = ctx.getImageData(850, 100, 4, 4);
      let contentAllWhite = true;
      for (let i = 0; i < 16; i++) {
        const idx = i * 4;
        if (contentPixels.data[idx] !== 255 || contentPixels.data[idx + 1] !== 255 || contentPixels.data[idx + 2] !== 255) {
          contentAllWhite = false;
        }
      }
      if (!contentAllWhite) {
        console.log("[Stage 11] Background image content detected at (850,100) on canvas.");
      } else {
        console.warn("[Stage 11] WARNING: No non-white pixels at (850,100)!");
      }
    } catch (e) {
      console.error("[Stage 11] getImageData at (850,100) failed:", e);
    }
  }

  // Remaining render steps (unchanged from original)
  if (debug.showPageBounds) {
    ctx.strokeStyle = "#ff0000";
    ctx.lineWidth = 2;
    ctx.strokeRect(0, 0, pageWidth, pageHeight);
  }

  if (debug.showOriginMarker) {
    ctx.strokeStyle = "#ff0000";
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(0, -10);
    ctx.lineTo(0, 10);
    ctx.moveTo(-10, 0);
    ctx.lineTo(10, 0);
    ctx.stroke();
  }

  if (debug.showPixelGrid) {
    ctx.strokeStyle = "rgba(200, 200, 255, 0.3)";
    ctx.lineWidth = 0.5;
    const gridSize = 10;
    for (let x = gridSize; x < pageWidth; x += gridSize) {
      ctx.beginPath();
      ctx.moveTo(x, 0);
      ctx.lineTo(x, pageHeight);
      ctx.stroke();
    }
    for (let y = gridSize; y < pageHeight; y += gridSize) {
      ctx.beginPath();
      ctx.moveTo(0, y);
      ctx.lineTo(pageWidth, y);
      ctx.stroke();
    }
  }

  if (debug.showFieldRectangles) {
    for (const f of fields) {
      const isEditing = f.id === editingFieldId;

      ctx.fillStyle = isEditing
        ? "rgba(0, 200, 0, 0.15)"
        : "rgba(255, 212, 0, 0.25)";
      ctx.strokeStyle = isEditing ? "#00cc00" : "#cc9900";
      ctx.lineWidth = isEditing ? 2 : 1;

      ctx.fillRect(f.leftPx, f.topPx, f.widthPx, f.heightPx);
      ctx.strokeRect(f.leftPx, f.topPx, f.widthPx, f.heightPx);
    }
  }

  if (debug.showHitTestRegions) {
    for (const f of fields) {
      ctx.strokeStyle = "rgba(0, 100, 200, 0.5)";
      ctx.lineWidth = 1;
      ctx.setLineDash([4, 4]);
      ctx.strokeRect(f.leftPx - 2, f.topPx - 2, f.widthPx + 4, f.heightPx + 4);
      ctx.setLineDash([]);
    }
  }

  if (debug.showFieldIds) {
    ctx.font = "10px monospace";
    for (const f of fields) {
      ctx.fillStyle = "rgba(0,0,0,0.7)";
      ctx.fillText(f.id, f.leftPx + 2, f.topPx + 12);
    }
  }

  if (debug.showCoordinates) {
    ctx.font = "9px monospace";
    ctx.fillStyle = "rgba(100,0,0,0.6)";
    for (const f of fields) {
      const coord = `(${f.leftPx.toFixed(0)},${f.topPx.toFixed(0)}) ${f.widthPx.toFixed(0)}x${f.heightPx.toFixed(0)}`;
      ctx.fillText(coord, f.leftPx + 2, f.topPx + f.heightPx - 4);
    }
  }

  console.log("[Stage 8] renderFrame() COMPLETE");
}
