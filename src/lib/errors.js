export class AppError extends Error {
  constructor(message, statusCode = 400, details = {}) {
    super(message);
    this.name = "AppError";
    this.statusCode = statusCode;
    this.details = details;
  }
}

export function ensure(condition, message, statusCode = 400, details = {}) {
  if (!condition) {
    throw new AppError(message, statusCode, details);
  }
}

