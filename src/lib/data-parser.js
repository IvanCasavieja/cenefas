import path from "node:path";

import { parse as parseCsv } from "csv-parse/sync";
import ExcelJS from "exceljs";

import { AppError, ensure } from "./errors.js";
import { normalizeFieldName } from "./placeholder-utils.js";

const SUPPORTED_EXTENSIONS = new Set([".xlsx", ".csv"]);

function formatExcelCell(cell) {
  return String(cell?.text ?? "").trim();
}

function parseCsvMatrix(buffer) {
  return parseCsv(buffer.toString("utf8"), {
    bom: true,
    relax_column_count: true,
    skip_empty_lines: false
  }).map((row) => row.map((value) => String(value ?? "")));
}

async function parseXlsxMatrix(buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const worksheet = workbook.worksheets[0];
  ensure(worksheet, "El archivo de datos no contiene hojas legibles.");

  const columnCount = worksheet.columnCount;
  const matrix = [];

  worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const cells = [];

    for (let columnIndex = 1; columnIndex <= columnCount; columnIndex += 1) {
      cells.push(formatExcelCell(row.getCell(columnIndex)));
    }

    matrix[rowNumber - 1] = cells;
  });

  return {
    sheetName: worksheet.name,
    matrix
  };
}

async function readMatrix(buffer, extension) {
  try {
    if (extension === ".csv") {
      return {
        sheetName: "CSV",
        matrix: parseCsvMatrix(buffer)
      };
    }

    return await parseXlsxMatrix(buffer);
  } catch (error) {
    throw new AppError("No se pudo leer el archivo de datos. Verifica que no este corrupto.", 400, {
      cause: error.message
    });
  }
}

export async function parseDataFile(buffer, originalName) {
  const extension = path.extname(originalName).toLowerCase();

  if (!SUPPORTED_EXTENSIONS.has(extension)) {
    throw new AppError("El archivo de datos debe ser .xlsx o .csv.");
  }

  const { sheetName, matrix } = await readMatrix(buffer, extension);
  const headerRow = Array.isArray(matrix[0]) ? matrix[0] : [];
  const columns = headerRow
    .map((label, index) => ({
      index,
      label: String(label ?? "").trim()
    }))
    .filter((column) => column.label)
    .map((column) => ({
      ...column,
      key: normalizeFieldName(column.label)
    }));

  ensure(columns.length > 0, "El archivo de datos no tiene encabezados validos en la primera fila.");

  const duplicatedKeys = new Set();
  const seenKeys = new Set();

  for (const column of columns) {
    if (!column.key) {
      throw new AppError(`La columna "${column.label}" no se puede convertir en un identificador usable.`);
    }

    if (seenKeys.has(column.key)) {
      duplicatedKeys.add(column.label);
      continue;
    }

    seenKeys.add(column.key);
  }

  if (duplicatedKeys.size > 0) {
    throw new AppError(
      `Hay columnas duplicadas o equivalentes al normalizar nombres: ${Array.from(duplicatedKeys).join(", ")}.`
    );
  }

  const records = [];

  for (let rowIndex = 1; rowIndex < matrix.length; rowIndex += 1) {
    const row = Array.isArray(matrix[rowIndex]) ? matrix[rowIndex] : [];
    const values = {};
    let hasContent = false;

    for (const column of columns) {
      const cellValue = String(row[column.index] ?? "").trim();
      values[column.key] = cellValue;

      if (cellValue !== "") {
        hasContent = true;
      }
    }

    if (!hasContent) {
      continue;
    }

    records.push({
      rowNumber: rowIndex + 1,
      values
    });
  }

  ensure(records.length > 0, "El archivo de datos no contiene registros con contenido.");

  return {
    fileType: extension.slice(1),
    sheetName,
    columns,
    records,
    rowCount: records.length
  };
}
