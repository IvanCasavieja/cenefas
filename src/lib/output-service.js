import { AppError } from "./errors.js";
import { generatePptxFromTemplate } from "./pptx-generator.js";

export async function generateOutput({ format = "pptx", templateBuffer, records }) {
  if (format !== "pptx") {
    throw new AppError("La exportación solicitada todavía no está implementada. Por ahora solo se genera PPTX.");
  }

  return {
    buffer: await generatePptxFromTemplate({
      templateBuffer,
      records
    }),
    mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    extension: "pptx"
  };
}

