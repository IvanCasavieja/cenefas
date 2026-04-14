import { DOMParser, XMLSerializer } from "@xmldom/xmldom";

export const XML_NAMESPACES = {
  drawing: "http://schemas.openxmlformats.org/drawingml/2006/main",
  officeRelationships: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  packageRelationships: "http://schemas.openxmlformats.org/package/2006/relationships",
  contentTypes: "http://schemas.openxmlformats.org/package/2006/content-types",
  xml: "http://www.w3.org/XML/1998/namespace"
};

const parser = new DOMParser();
const serializer = new XMLSerializer();

export function parseXml(xmlText) {
  return parser.parseFromString(xmlText, "application/xml");
}

export function serializeXml(document) {
  return serializer.serializeToString(document);
}

export function getLocalName(node) {
  return node?.localName ?? node?.nodeName?.split(":").pop() ?? "";
}

export function isElement(node, localName) {
  return node?.nodeType === 1 && getLocalName(node) === localName;
}

export function getChildElements(node, localName = null) {
  const children = [];

  for (let index = 0; index < node.childNodes.length; index += 1) {
    const child = node.childNodes[index];

    if (child.nodeType !== 1) {
      continue;
    }

    if (!localName || getLocalName(child) === localName) {
      children.push(child);
    }
  }

  return children;
}

export function findDescendants(node, localName, accumulator = []) {
  if (!node?.childNodes) {
    return accumulator;
  }

  for (let index = 0; index < node.childNodes.length; index += 1) {
    const child = node.childNodes[index];

    if (child.nodeType !== 1) {
      continue;
    }

    if (getLocalName(child) === localName) {
      accumulator.push(child);
    }

    findDescendants(child, localName, accumulator);
  }

  return accumulator;
}

export function getAttributeByLocalName(node, localName) {
  if (!node?.attributes) {
    return null;
  }

  for (let index = 0; index < node.attributes.length; index += 1) {
    const attribute = node.attributes[index];
    const attributeLocalName = attribute.localName ?? attribute.name?.split(":").pop();

    if (attributeLocalName === localName) {
      return attribute.value;
    }
  }

  return null;
}

export function setTextContentWithSpace(textElement, value) {
  while (textElement.firstChild) {
    textElement.removeChild(textElement.firstChild);
  }

  const textValue = String(value ?? "");
  const preserveSpace = /^\s|\s$|\s{2,}/.test(textValue);

  if (preserveSpace) {
    textElement.setAttributeNS(XML_NAMESPACES.xml, "xml:space", "preserve");
  } else if (textElement.hasAttribute("xml:space")) {
    textElement.removeAttribute("xml:space");
  }

  textElement.appendChild(textElement.ownerDocument.createTextNode(textValue));
}

