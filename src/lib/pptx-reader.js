import path from "node:path";
import JSZip from "jszip";

import { AppError, ensure } from "./errors.js";
import { extractPlaceholders } from "./placeholder-utils.js";
import {
  findDescendants,
  getAttributeByLocalName,
  getChildElements,
  getLocalName,
  parseXml
} from "./pptx-xml.js";

const SLIDE_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";

function resolveRelativeTarget(baseFile, target) {
  if (target.startsWith("/")) {
    return target.replace(/^\/+/, "");
  }

  return path.posix.normalize(path.posix.join(path.posix.dirname(baseFile), target));
}

export async function loadPptx(buffer) {
  try {
    return await JSZip.loadAsync(buffer);
  } catch (error) {
    throw new AppError("No se pudo abrir el archivo PPTX. Verificá que sea un .pptx válido.", 400, {
      cause: error.message
    });
  }
}

export async function readZipText(zip, zipPath) {
  const file = zip.file(zipPath);
  ensure(file, `No se encontró el archivo interno "${zipPath}" dentro del PPTX.`, 400);
  return file.async("string");
}

export function extractParagraphText(paragraphNode) {
  let text = "";

  for (const child of getChildElements(paragraphNode)) {
    const localName = getLocalName(child);

    if (localName === "r" || localName === "fld") {
      const textNode = getChildElements(child, "t")[0];
      text += textNode?.textContent ?? "";
      continue;
    }

    if (localName === "br") {
      text += "\n";
      continue;
    }

    if (localName === "tab") {
      text += "\t";
    }
  }

  return text;
}

export function analyzeSlideXml(xmlText) {
  const document = parseXml(xmlText);
  const textBodies = findDescendants(document, "txBody");
  const slidePlaceholders = [];
  const seenPlaceholders = new Set();
  const paragraphPreview = [];

  for (const textBody of textBodies) {
    const paragraphs = findDescendants(textBody, "p");

    for (const paragraph of paragraphs) {
      const paragraphText = extractParagraphText(paragraph);

      if (paragraphText.trim()) {
        paragraphPreview.push(paragraphText.trim());
      }

      for (const placeholder of extractPlaceholders(paragraphText)) {
        if (seenPlaceholders.has(placeholder.key)) {
          continue;
        }

        seenPlaceholders.add(placeholder.key);
        slidePlaceholders.push(placeholder);
      }
    }
  }

  return {
    placeholders: slidePlaceholders,
    previewText: paragraphPreview.slice(0, 3).join(" / ").slice(0, 180)
  };
}

export async function getOrderedSlides(zip) {
  const presentationXml = await readZipText(zip, "ppt/presentation.xml");
  const presentationRelsXml = await readZipText(zip, "ppt/_rels/presentation.xml.rels");

  const presentationDocument = parseXml(presentationXml);
  const relationshipsDocument = parseXml(presentationRelsXml);

  const slideRelationships = new Map();

  for (const relationship of findDescendants(relationshipsDocument, "Relationship")) {
    const relationshipId = getAttributeByLocalName(relationship, "Id");
    const relationshipType = getAttributeByLocalName(relationship, "Type");
    const target = getAttributeByLocalName(relationship, "Target");

    if (relationshipId && relationshipType === SLIDE_RELATIONSHIP_TYPE && target) {
      slideRelationships.set(relationshipId, resolveRelativeTarget("ppt/presentation.xml", target));
    }
  }

  const slideIdList = findDescendants(presentationDocument, "sldId");
  const orderedSlides = slideIdList
    .map((slideIdNode, index) => {
      const relationshipId =
        slideIdNode.getAttribute("r:id") ??
        slideIdNode.getAttributeNS?.("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
      const slidePath = relationshipId ? slideRelationships.get(relationshipId) : null;

      if (!slidePath) {
        return null;
      }

      return {
        index: index + 1,
        relationshipId,
        path: slidePath
      };
    })
    .filter(Boolean);

  ensure(orderedSlides.length > 0, "El PPTX no contiene diapositivas utilizables.");

  return orderedSlides;
}

export async function analyzePptxTemplate(buffer) {
  const zip = await loadPptx(buffer);
  const orderedSlides = await getOrderedSlides(zip);
  const allPlaceholders = [];
  const seenPlaceholders = new Set();
  const slides = [];

  for (const slide of orderedSlides) {
    const slideXml = await readZipText(zip, slide.path);
    const analysis = analyzeSlideXml(slideXml);

    for (const placeholder of analysis.placeholders) {
      if (seenPlaceholders.has(placeholder.key)) {
        continue;
      }

      seenPlaceholders.add(placeholder.key);
      allPlaceholders.push(placeholder);
    }

    slides.push({
      index: slide.index,
      path: slide.path,
      previewText: analysis.previewText,
      placeholders: analysis.placeholders
    });
  }

  return {
    slideCount: orderedSlides.length,
    placeholders: allPlaceholders,
    slides
  };
}
