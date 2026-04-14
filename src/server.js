import path from "node:path";

import express from "express";
import multer from "multer";

import { parseDataFile } from "./lib/data-parser.js";
import { AppError } from "./lib/errors.js";
import { generateOutput } from "./lib/output-service.js";
import { buildMatchSummary } from "./lib/placeholder-utils.js";
import { analyzePptxTemplate } from "./lib/pptx-reader.js";
import { buildDownloadName, createJob, initStorage, readJob } from "./lib/storage.js";

const PORT = Number(process.env.PORT || 3000);
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 25 * 1024 * 1024
  }
});
const LOCAL_ORIGIN_REGEX = /^https?:\/\/(?:localhost|127\.0\.0\.1)(?::\d+)?$/i;

function asyncHandler(handler) {
  return (request, response, next) => {
    Promise.resolve(handler(request, response, next)).catch(next);
  };
}

function ensureTemplateFile(file) {
  if (!file) {
    throw new AppError("Tenes que subir un archivo base .pptx.");
  }

  if (!file.originalname.toLowerCase().endsWith(".pptx")) {
    throw new AppError("El template debe ser un archivo .pptx.");
  }
}

function ensureDataFile(file) {
  if (!file) {
    throw new AppError("Tenes que subir un archivo de datos .xlsx o .csv.");
  }

  const lowerName = file.originalname.toLowerCase();

  if (!lowerName.endsWith(".xlsx") && !lowerName.endsWith(".csv")) {
    throw new AppError("El archivo de datos debe ser .xlsx o .csv.");
  }
}

function buildWarnings({ templateAnalysis, matchSummary }) {
  const warnings = [];

  if (templateAnalysis.placeholders.length === 0) {
    warnings.push("El PPTX no contiene placeholders detectables con el formato {{CAMPO}}.");
  }

  if (matchSummary.missingPlaceholders.length > 0) {
    warnings.push(
      `Hay placeholders sin columna equivalente: ${matchSummary.missingPlaceholders.join(", ")}. Se conservaran visibles en la salida para que puedas revisarlos.`
    );
  }

  if (matchSummary.unusedColumns.length > 0) {
    warnings.push(
      `Hay columnas cargadas que no se usan en el template: ${matchSummary.unusedColumns.join(", ")}.`
    );
  }

  return warnings;
}

await initStorage();

const app = express();
const ROOT_DIRECTORY = process.cwd();

app.use((request, response, next) => {
  const origin = request.headers.origin;

  if (origin && LOCAL_ORIGIN_REGEX.test(origin)) {
    response.setHeader("Access-Control-Allow-Origin", origin);
    response.setHeader("Vary", "Origin");
    response.setHeader("Access-Control-Allow-Methods", "GET,POST,OPTIONS");
    response.setHeader("Access-Control-Allow-Headers", "Content-Type");
  }

  if (request.method === "OPTIONS") {
    response.status(204).end();
    return;
  }

  next();
});

app.use(express.json());

function sendRootIndex(response) {
  response.sendFile(path.resolve(ROOT_DIRECTORY, "index.html"));
}

app.get("/", (_request, response) => {
  sendRootIndex(response);
});

app.get("/index.html", (_request, response) => {
  sendRootIndex(response);
});

app.get("/app.js", (_request, response) => {
  response.type("application/javascript");
  response.sendFile(path.resolve(ROOT_DIRECTORY, "app.js"));
});

app.get("/styles.css", (_request, response) => {
  response.type("text/css");
  response.sendFile(path.resolve(ROOT_DIRECTORY, "styles.css"));
});

app.get("/api/health", (_request, response) => {
  response.json({
    ok: true
  });
});

app.post(
  "/api/analyze",
  upload.fields([
    { name: "template", maxCount: 1 },
    { name: "data", maxCount: 1 }
  ]),
  asyncHandler(async (request, response) => {
    const templateFile = request.files?.template?.[0];
    const dataFile = request.files?.data?.[0];

    ensureTemplateFile(templateFile);
    ensureDataFile(dataFile);

    const dataset = await parseDataFile(dataFile.buffer, dataFile.originalname);
    const templateAnalysis = await analyzePptxTemplate(templateFile.buffer);
    const matchSummary = buildMatchSummary(templateAnalysis.placeholders, dataset.columns);
    const warnings = buildWarnings({
      templateAnalysis,
      matchSummary
    });

    const job = await createJob({
      templateBuffer: templateFile.buffer,
      templateFileName: templateFile.originalname,
      dataBuffer: dataFile.buffer,
      dataFileName: dataFile.originalname,
      analysis: {
        recordCount: dataset.rowCount,
        columns: dataset.columns.map((column) => column.label),
        placeholders: templateAnalysis.placeholders.map((placeholder) => placeholder.label),
        templateSlides: templateAnalysis.slides
          .filter((slide) => slide.placeholders.length > 0)
          .map((slide) => slide.index)
      }
    });

    response.json({
      jobId: job.jobId,
      templateFileName: templateFile.originalname,
      dataFileName: dataFile.originalname,
      recordCount: dataset.rowCount,
      slideCount: templateAnalysis.slideCount,
      columns: dataset.columns.map((column) => column.label),
      placeholders: templateAnalysis.placeholders.map((placeholder) => placeholder.label),
      matchedFields: matchSummary.matchedFields,
      missingPlaceholders: matchSummary.missingPlaceholders,
      unusedColumns: matchSummary.unusedColumns,
      warnings,
      canGenerate: templateAnalysis.placeholders.length > 0,
      templateSlides: templateAnalysis.slides
        .filter((slide) => slide.placeholders.length > 0)
        .map((slide) => ({
          index: slide.index,
          previewText: slide.previewText,
          placeholders: slide.placeholders.map((placeholder) => placeholder.label)
        })),
      generationMode:
        "Se duplican las slides template contiguas que contienen placeholders, una vez por cada fila del archivo de datos."
    });
  })
);

app.post(
  "/api/generate",
  asyncHandler(async (request, response) => {
    const { jobId, format = "pptx" } = request.body ?? {};

    if (!jobId) {
      throw new AppError("Falta el identificador del analisis previo.");
    }

    const job = await readJob(jobId);
    const dataset = await parseDataFile(job.dataBuffer, job.dataFileName);
    const output = await generateOutput({
      format,
      templateBuffer: job.templateBuffer,
      records: dataset.records
    });

    const fileName = buildDownloadName(job.templateFileName);
    response.setHeader("Content-Type", output.mimeType);
    response.setHeader("Content-Disposition", `attachment; filename="${fileName}"`);
    response.send(output.buffer);
  })
);

app.use("/api/*", (_request, response) => {
  response.status(404).json({
    message: "Ruta API no encontrada."
  });
});

app.use((error, _request, response, _next) => {
  if (error?.type === "entity.parse.failed") {
    response.status(400).json({
      message: "El cuerpo JSON de la solicitud no es valido."
    });
    return;
  }

  if (error instanceof multer.MulterError) {
    const message =
      error.code === "LIMIT_FILE_SIZE"
        ? "Uno de los archivos supera el limite de 25 MB."
        : "No se pudieron procesar los archivos subidos.";

    response.status(400).json({
      message
    });
    return;
  }

  if (error instanceof AppError) {
    response.status(error.statusCode).json({
      message: error.message,
      details: error.details
    });
    return;
  }

  console.error(error);
  response.status(500).json({
    message: "Ocurrio un error interno procesando la solicitud."
  });
});

app.listen(PORT, () => {
  console.log(`Servidor listo en http://localhost:${PORT}`);
});
