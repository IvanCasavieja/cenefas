import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

import { AppError, ensure } from "./errors.js";

const JOBS_DIRECTORY = path.resolve(process.cwd(), "storage", "jobs");
const JOB_MAX_AGE_MS = 24 * 60 * 60 * 1000;

function sanitizeFileName(fileName, fallback = "archivo") {
  const baseName = String(fileName ?? fallback)
    .replace(/[^\w.\-]+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "");

  return baseName || fallback;
}

function getJobDirectory(jobId) {
  ensure(/^[a-f0-9-]+$/i.test(jobId), "El identificador del proceso no es válido.", 400);
  return path.join(JOBS_DIRECTORY, jobId);
}

async function cleanupOldJobs() {
  const entries = await fs.readdir(JOBS_DIRECTORY, {
    withFileTypes: true
  }).catch(() => []);

  const now = Date.now();

  for (const entry of entries) {
    if (!entry.isDirectory()) {
      continue;
    }

    const jobPath = path.join(JOBS_DIRECTORY, entry.name);
    const stats = await fs.stat(jobPath).catch(() => null);

    if (!stats) {
      continue;
    }

    if (now - stats.mtimeMs > JOB_MAX_AGE_MS) {
      await fs.rm(jobPath, {
        recursive: true,
        force: true
      });
    }
  }
}

export async function initStorage() {
  await fs.mkdir(JOBS_DIRECTORY, {
    recursive: true
  });
  await cleanupOldJobs();
}

export async function createJob({ templateBuffer, templateFileName, dataBuffer, dataFileName, analysis }) {
  const jobId = crypto.randomUUID();
  const jobDirectory = getJobDirectory(jobId);
  const templateTargetName = sanitizeFileName(templateFileName, "template.pptx");
  const dataTargetName = sanitizeFileName(dataFileName, "datos");

  await fs.mkdir(jobDirectory, {
    recursive: true
  });

  await fs.writeFile(path.join(jobDirectory, "template.bin"), templateBuffer);
  await fs.writeFile(path.join(jobDirectory, "data.bin"), dataBuffer);
  await fs.writeFile(
    path.join(jobDirectory, "meta.json"),
    JSON.stringify(
      {
        createdAt: new Date().toISOString(),
        templateFileName,
        dataFileName,
        storedTemplateName: templateTargetName,
        storedDataName: dataTargetName,
        analysis
      },
      null,
      2
    )
  );

  return {
    jobId
  };
}

export async function readJob(jobId) {
  const jobDirectory = getJobDirectory(jobId);
  const metaPath = path.join(jobDirectory, "meta.json");
  const templatePath = path.join(jobDirectory, "template.bin");
  const dataPath = path.join(jobDirectory, "data.bin");

  try {
    const [metaRaw, templateBuffer, dataBuffer] = await Promise.all([
      fs.readFile(metaPath, "utf8"),
      fs.readFile(templatePath),
      fs.readFile(dataPath)
    ]);

    return {
      ...JSON.parse(metaRaw),
      jobId,
      templateBuffer,
      dataBuffer
    };
  } catch (error) {
    throw new AppError("No se encontró el análisis previo. Volvé a subir los archivos.", 404, {
      cause: error.message
    });
  }
}

export function buildDownloadName(templateFileName) {
  const parsed = path.parse(templateFileName ?? "cenefas-template.pptx");
  return sanitizeFileName(`${parsed.name || "cenefas"}-procesado.pptx`, "cenefas-procesado.pptx");
}
