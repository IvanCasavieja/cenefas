const PLACEHOLDER_REGEX = /\{\{\s*([^{}]+?)\s*\}\}/g;

function stripAccents(value) {
  return value.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

export function normalizeFieldName(value = "") {
  return stripAccents(String(value ?? ""))
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

export function extractPlaceholders(text = "") {
  const placeholders = [];
  const seen = new Set();

  for (const match of String(text).matchAll(PLACEHOLDER_REGEX)) {
    const label = String(match[1] ?? "").trim();
    const key = normalizeFieldName(label);

    if (!key || seen.has(key)) {
      continue;
    }

    seen.add(key);
    placeholders.push({
      key,
      label,
      token: match[0]
    });
  }

  return placeholders;
}

export function replacePlaceholdersInText(text, valuesByKey, options = {}) {
  const preserveMissingColumns = options.preserveMissingColumns ?? true;

  return String(text ?? "").replace(PLACEHOLDER_REGEX, (match, rawKey) => {
    const key = normalizeFieldName(rawKey);

    if (!key) {
      return match;
    }

    if (!Object.prototype.hasOwnProperty.call(valuesByKey, key)) {
      return preserveMissingColumns ? match : "";
    }

    return String(valuesByKey[key] ?? "");
  });
}

export function buildMatchSummary(placeholders, columns) {
  const columnsByKey = new Map(columns.map((column) => [column.key, column]));
  const placeholderKeys = new Set(placeholders.map((placeholder) => placeholder.key));

  const matchedFields = placeholders
    .filter((placeholder) => columnsByKey.has(placeholder.key))
    .map((placeholder) => ({
      placeholder: placeholder.label,
      column: columnsByKey.get(placeholder.key).label
    }));

  const missingPlaceholders = placeholders
    .filter((placeholder) => !columnsByKey.has(placeholder.key))
    .map((placeholder) => placeholder.label);

  const unusedColumns = columns
    .filter((column) => !placeholderKeys.has(column.key))
    .map((column) => column.label);

  return {
    matchedFields,
    missingPlaceholders,
    unusedColumns
  };
}

