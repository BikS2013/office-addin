export class AppBaseError extends Error {
  constructor(
    message: string,
    public readonly retryable: boolean,
    public readonly userMessage: string,
    public readonly cause?: Error
  ) {
    super(message);
    this.name = this.constructor.name;
  }
}

export class ConfigurationError extends AppBaseError {
  constructor(
    public readonly fieldPath: string,
    message: string,
    cause?: Error
  ) {
    super(message, false, `Configuration error at '${fieldPath}': ${message}`, cause);
  }
}

export class ConfigFetchError extends AppBaseError {
  constructor(
    message: string,
    public readonly statusCode?: number,
    cause?: Error
  ) {
    const retryable = !statusCode || statusCode >= 500;
    const userMsg = statusCode
      ? statusCode >= 500
        ? `Configuration server error (HTTP ${statusCode}). Try reloading later.`
        : `Configuration URL returned HTTP ${statusCode}. Verify the URL is correct.`
      : "Cannot reach configuration URL. Check your network connection and try again.";
    super(message, retryable, userMsg, cause);
  }
}

export class ApiExecutionError extends AppBaseError {
  constructor(
    message: string,
    public readonly apiName: string,
    public readonly errorType: "network" | "cors" | "timeout" | "http" | "parse" | "field_missing",
    retryable: boolean,
    cause?: Error
  ) {
    super(message, retryable, message, cause);
  }
}

export class OfficeApiError extends AppBaseError {
  constructor(message: string, cause?: Error) {
    super(message, true, message, cause);
  }
}

export class HistoryStorageError extends AppBaseError {
  constructor(message: string, cause?: Error) {
    super(message, false, message, cause);
  }
}
