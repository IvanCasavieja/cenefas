import fs from "node:fs/promises";
import path from "node:path";

import PptxGenJS from "pptxgenjs";

const outputPath = path.resolve(process.cwd(), "examples", "template-demo.pptx");

async function ensureDirectory(targetPath) {
  await fs.mkdir(path.dirname(targetPath), {
    recursive: true
  });
}

async function createExampleTemplate() {
  await ensureDirectory(outputPath);

  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "Codex";
  pptx.subject = "Template demo para generador de cenefas";
  pptx.title = "Template demo de cenefas";
  pptx.company = "Local";
  pptx.lang = "es-UY";

  const slide = pptx.addSlide();
  slide.background = {
    color: "F7F0E5"
  };

  slide.addShape(pptx.ShapeType.rect, {
    x: 0.4,
    y: 0.35,
    w: 12.4,
    h: 6.3,
    line: {
      color: "5A4134",
      pt: 1.25
    },
    fill: {
      color: "F6D8DE"
    },
    radius: 0.08
  });

  slide.addShape(pptx.ShapeType.rect, {
    x: 0.4,
    y: 0.35,
    w: 3.2,
    h: 6.3,
    line: {
      color: "5A4134",
      pt: 1.25
    },
    fill: {
      color: "C5D3D9"
    }
  });

  slide.addText("COD : {{COD}}", {
    x: 4.1,
    y: 0.65,
    w: 7.7,
    h: 0.4,
    fontFace: "Aptos",
    fontSize: 18,
    bold: true,
    color: "4A3227"
  });

  slide.addText("{{DESC}}", {
    x: 4.1,
    y: 1.2,
    w: 7.8,
    h: 1.8,
    fontFace: "Aptos Display",
    fontSize: 28,
    bold: true,
    color: "4A3227",
    valign: "mid",
    fit: "shrink"
  });

  slide.addText("PRECIO REGULAR", {
    x: 4.1,
    y: 3.45,
    w: 3.2,
    h: 0.35,
    fontFace: "Aptos",
    fontSize: 16,
    bold: true,
    color: "6A5349"
  });

  slide.addText("U$S {{PR}}", {
    x: 4.1,
    y: 3.78,
    w: 3.8,
    h: 0.6,
    fontFace: "Aptos",
    fontSize: 24,
    bold: true,
    color: "4A3227"
  });

  slide.addText("OFERTA", {
    x: 4.1,
    y: 4.5,
    w: 2.4,
    h: 0.45,
    fontFace: "Aptos Display",
    fontSize: 24,
    bold: true,
    color: "9D4F2F"
  });

  slide.addText("U$S {{PO}}", {
    x: 4.1,
    y: 4.95,
    w: 4.2,
    h: 1,
    fontFace: "Aptos Display",
    fontSize: 34,
    bold: true,
    color: "9D4F2F"
  });

  slide.addText("Imagen del producto", {
    x: 0.85,
    y: 2.05,
    w: 2.3,
    h: 0.5,
    align: "center",
    fontFace: "Aptos",
    fontSize: 16,
    bold: true,
    color: "4A3227"
  });

  slide.addText("{{COD}}", {
    x: 1.05,
    y: 2.65,
    w: 1.9,
    h: 0.45,
    align: "center",
    fontFace: "Aptos",
    fontSize: 18,
    bold: true,
    color: "4A3227"
  });

  await pptx.writeFile({
    fileName: outputPath
  });

  console.log(`Template demo creado en ${outputPath}`);
}

createExampleTemplate().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});

