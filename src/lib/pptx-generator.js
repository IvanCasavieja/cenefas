import path from "node:path";

import { AppError, ensure } from "./errors.js";
import { replacePlaceholdersInText } from "./placeholder-utils.js";
import { analyzeSlideXml, getOrderedSlides, loadPptx, readZipText } from "./pptx-reader.js";
import {
  XML_NAMESPACES,
  findDescendants,
  getAttributeByLocalName,
  getChildElements,
  getLocalName,
  isElement,
  parseXml,
  serializeXml,
  setTextContentWithSpace
} from "./pptx-xml.js";

const CONTENT_TYPE_SLIDE =
  "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const SLIDE_RELATIONSHIP_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const FORBIDDEN_CLONED_RELATIONSHIP_TYPES = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide",
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors"
]);

function getSlideRelsPath(slidePath) {
  return path.posix.join(path.posix.dirname(slidePath), "_rels", `${path.posix.basename(slidePath)}.rels`);
}

function getNextSlideNumber(zip) {
  let maxSlideNumber = 0;

  for (const fileName of Object.keys(zip.files)) {
    const match = fileName.match(/^ppt\/slides\/slide(\d+)\.xml$/);
    if (!match) {
      continue;
    }

    maxSlideNumber = Math.max(maxSlideNumber, Number(match[1]));
  }

  return maxSlideNumber + 1;
}

function buildGenerationPlan(slides) {
  const groups = [];

  for (const slide of slides) {
    const type = slide.hasPlaceholders ? "template" : "static";
    const lastGroup = groups[groups.length - 1];

    if (!lastGroup || lastGroup.type !== type) {
      groups.push({
        type,
        slides: [slide]
      });
      continue;
    }

    lastGroup.slides.push(slide);
  }

  return groups;
}

function getTextSegments(paragraphNode) {
  const segments = [];

  for (const child of getChildElements(paragraphNode)) {
    const localName = getLocalName(child);

    if (localName === "r" || localName === "fld") {
      const textNode = getChildElements(child, "t")[0];

      if (textNode) {
        segments.push({
          type: "text",
          node: textNode,
          value: textNode.textContent ?? ""
        });
      }

      continue;
    }

    if (localName === "br") {
      segments.push({
        type: "break",
        node: child,
        value: "\n"
      });
      continue;
    }

    if (localName === "tab") {
      segments.push({
        type: "literal",
        node: child,
        value: "\t"
      });
    }
  }

  return segments;
}

function getRunPropertiesSource(paragraphNode) {
  for (const child of getChildElements(paragraphNode)) {
    const localName = getLocalName(child);

    if (localName === "r" || localName === "fld") {
      const runProperties = getChildElements(child, "rPr")[0];

      if (runProperties) {
        return runProperties;
      }
    }
  }

  return getChildElements(paragraphNode, "endParaRPr")[0] ?? null;
}

function insertBeforeTailNode(paragraphNode, newNode) {
  for (const child of getChildElements(paragraphNode)) {
    const localName = getLocalName(child);

    if (localName === "endParaRPr" || localName === "extLst") {
      paragraphNode.insertBefore(newNode, child);
      return;
    }
  }

  paragraphNode.appendChild(newNode);
}

function createRunNode(document, textValue, runPropertiesTemplate) {
  const runNode = document.createElementNS(XML_NAMESPACES.drawing, "a:r");

  if (runPropertiesTemplate) {
    runNode.appendChild(runPropertiesTemplate.cloneNode(true));
  }

  const textNode = document.createElementNS(XML_NAMESPACES.drawing, "a:t");
  setTextContentWithSpace(textNode, textValue);
  runNode.appendChild(textNode);

  return runNode;
}

function rebuildParagraphWithText(paragraphNode, desiredText) {
  const runPropertiesTemplate = getRunPropertiesSource(paragraphNode);

  for (let index = paragraphNode.childNodes.length - 1; index >= 0; index -= 1) {
    const child = paragraphNode.childNodes[index];

    if (child.nodeType !== 1) {
      continue;
    }

    const localName = getLocalName(child);

    if (localName === "r" || localName === "fld" || localName === "br" || localName === "tab") {
      paragraphNode.removeChild(child);
    }
  }

  const parts = String(desiredText).split("\n");

  if (parts.length === 1 && parts[0] === "") {
    insertBeforeTailNode(paragraphNode, createRunNode(paragraphNode.ownerDocument, "", runPropertiesTemplate));
    return;
  }

  parts.forEach((part, index) => {
    if (index > 0) {
      const breakNode = paragraphNode.ownerDocument.createElementNS(XML_NAMESPACES.drawing, "a:br");
      insertBeforeTailNode(paragraphNode, breakNode);
    }

    if (part !== "" || parts.length === 1) {
      insertBeforeTailNode(paragraphNode, createRunNode(paragraphNode.ownerDocument, part, runPropertiesTemplate));
    }
  });
}

function replaceParagraphPlaceholders(paragraphNode, valuesByKey) {
  const segments = getTextSegments(paragraphNode);

  if (segments.length === 0) {
    return;
  }

  const originalText = segments.map((segment) => segment.value).join("");
  const replacedText = replacePlaceholdersInText(originalText, valuesByKey, {
    preserveMissingColumns: true
  });

  if (replacedText === originalText) {
    return;
  }

  const individuallyReplacedText = segments
    .map((segment) =>
      segment.type === "text"
        ? replacePlaceholdersInText(segment.value, valuesByKey, { preserveMissingColumns: true })
        : segment.value
    )
    .join("");

  if (individuallyReplacedText === replacedText) {
    for (const segment of segments) {
      if (segment.type !== "text") {
        continue;
      }

      const updatedValue = replacePlaceholdersInText(segment.value, valuesByKey, {
        preserveMissingColumns: true
      });
      setTextContentWithSpace(segment.node, updatedValue);
    }

    return;
  }

  rebuildParagraphWithText(paragraphNode, replacedText);
}

function replacePlaceholdersInSlideXml(slideXml, valuesByKey) {
  const document = parseXml(slideXml);
  const paragraphs = findDescendants(document, "p");

  for (const paragraph of paragraphs) {
    replaceParagraphPlaceholders(paragraph, valuesByKey);
  }

  return serializeXml(document);
}

function stripUnsupportedRelationships(slideRelsXml) {
  const document = parseXml(slideRelsXml);
  const relationships = findDescendants(document, "Relationship");

  for (const relationship of relationships) {
    const relationshipType = getAttributeByLocalName(relationship, "Type");

    if (!FORBIDDEN_CLONED_RELATIONSHIP_TYPES.has(relationshipType)) {
      continue;
    }

    relationship.parentNode.removeChild(relationship);
  }

  return serializeXml(document);
}

async function cloneSlideForRecord(zip, sourceSlidePath, valuesByKey, slideNumber) {
  const newSlidePath = `ppt/slides/slide${slideNumber}.xml`;
  const sourceSlideXml = await readZipText(zip, sourceSlidePath);
  const generatedSlideXml = replacePlaceholdersInSlideXml(sourceSlideXml, valuesByKey);
  zip.file(newSlidePath, generatedSlideXml);

  const sourceRelsPath = getSlideRelsPath(sourceSlidePath);
  const sourceRelsFile = zip.file(sourceRelsPath);

  if (sourceRelsFile) {
    const targetRelsPath = getSlideRelsPath(newSlidePath);
    const relsXml = await sourceRelsFile.async("string");
    zip.file(targetRelsPath, stripUnsupportedRelationships(relsXml));
  }

  return {
    path: newSlidePath
  };
}

async function updateContentTypes(zip, slidePaths) {
  const contentTypesXml = await readZipText(zip, "[Content_Types].xml");
  const document = parseXml(contentTypesXml);
  const overrides = findDescendants(document, "Override");
  const existingPartNames = new Set(overrides.map((node) => getAttributeByLocalName(node, "PartName")));

  for (const slidePath of slidePaths) {
    const partName = `/${slidePath}`;

    if (existingPartNames.has(partName)) {
      continue;
    }

    const overrideNode = document.createElementNS(XML_NAMESPACES.contentTypes, "Override");
    overrideNode.setAttribute("PartName", partName);
    overrideNode.setAttribute("ContentType", CONTENT_TYPE_SLIDE);
    document.documentElement.appendChild(overrideNode);
  }

  zip.file("[Content_Types].xml", serializeXml(document));
}

async function rewritePresentation(zip, slidePaths) {
  const presentationXml = await readZipText(zip, "ppt/presentation.xml");
  const presentationRelsXml = await readZipText(zip, "ppt/_rels/presentation.xml.rels");

  const presentationDocument = parseXml(presentationXml);
  const relationshipsDocument = parseXml(presentationRelsXml);

  const relationshipsRoot = relationshipsDocument.documentElement;
  const relationshipNodes = findDescendants(relationshipsDocument, "Relationship");
  let nextRelationshipIndex = 1;

  for (const relationship of relationshipNodes) {
    const relationshipId = getAttributeByLocalName(relationship, "Id");
    const relationshipType = getAttributeByLocalName(relationship, "Type");

    if (relationshipType === SLIDE_RELATIONSHIP_TYPE) {
      relationship.parentNode.removeChild(relationship);
      continue;
    }

    const numericPart = Number(String(relationshipId ?? "").replace(/^rId/i, ""));
    if (Number.isFinite(numericPart)) {
      nextRelationshipIndex = Math.max(nextRelationshipIndex, numericPart + 1);
    }
  }

  const slideRelationshipIds = [];

  for (const slidePath of slidePaths) {
    const relationshipId = `rId${nextRelationshipIndex}`;
    nextRelationshipIndex += 1;

    const relationshipNode = relationshipsDocument.createElementNS(
      XML_NAMESPACES.packageRelationships,
      "Relationship"
    );
    relationshipNode.setAttribute("Id", relationshipId);
    relationshipNode.setAttribute("Type", SLIDE_RELATIONSHIP_TYPE);
    relationshipNode.setAttribute("Target", path.posix.relative("ppt", slidePath));
    relationshipsRoot.appendChild(relationshipNode);
    slideRelationshipIds.push(relationshipId);
  }

  const slideIdList = findDescendants(presentationDocument, "sldIdLst")[0];
  ensure(slideIdList, "No se encontro la lista de diapositivas dentro del PPTX.", 400);

  let nextSlideId = 256;

  for (let index = slideIdList.childNodes.length - 1; index >= 0; index -= 1) {
    const child = slideIdList.childNodes[index];

    if (!isElement(child, "sldId")) {
      continue;
    }

    const currentId = Number(child.getAttribute("id"));
    if (Number.isFinite(currentId)) {
      nextSlideId = Math.max(nextSlideId, currentId + 1);
    }

    slideIdList.removeChild(child);
  }

  slideRelationshipIds.forEach((relationshipId) => {
    const slideIdNode = presentationDocument.createElementNS(
      "http://schemas.openxmlformats.org/presentationml/2006/main",
      "p:sldId"
    );
    slideIdNode.setAttribute("id", String(nextSlideId));
    slideIdNode.setAttributeNS(XML_NAMESPACES.officeRelationships, "r:id", relationshipId);
    slideIdList.appendChild(slideIdNode);
    nextSlideId += 1;
  });

  zip.file("ppt/presentation.xml", serializeXml(presentationDocument));
  zip.file("ppt/_rels/presentation.xml.rels", serializeXml(relationshipsDocument));
}

export async function generatePptxFromTemplate({ templateBuffer, records }) {
  ensure(Array.isArray(records) && records.length > 0, "No hay registros para generar la salida.");

  const zip = await loadPptx(templateBuffer);
  const orderedSlides = await getOrderedSlides(zip);
  const slides = [];

  for (const slide of orderedSlides) {
    const slideXml = await readZipText(zip, slide.path);
    const analysis = analyzeSlideXml(slideXml);

    slides.push({
      ...slide,
      hasPlaceholders: analysis.placeholders.length > 0
    });
  }

  const generationPlan = buildGenerationPlan(slides);
  const hasTemplateSlides = generationPlan.some((group) => group.type === "template");

  if (!hasTemplateSlides) {
    throw new AppError("El template no contiene placeholders. No se puede generar una version procesada.");
  }

  const finalSlides = [];
  let nextSlideNumber = getNextSlideNumber(zip);

  for (const group of generationPlan) {
    if (group.type === "static") {
      finalSlides.push(...group.slides.map((slide) => slide.path));
      continue;
    }

    for (const record of records) {
      for (const slide of group.slides) {
        const clonedSlide = await cloneSlideForRecord(zip, slide.path, record.values, nextSlideNumber);
        nextSlideNumber += 1;
        finalSlides.push(clonedSlide.path);
      }
    }
  }

  await updateContentTypes(zip, finalSlides);
  await rewritePresentation(zip, finalSlides);

  return zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: {
      level: 6
    }
  });
}
