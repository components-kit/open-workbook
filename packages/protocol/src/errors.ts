export type ErrorSeverity = "info" | "warning" | "error" | "fatal";

export const ExcelErrorCode = {
  AddinDisconnected: "ADDIN_DISCONNECTED",
  WorkbookNotFound: "WORKBOOK_NOT_FOUND",
  SheetNotFound: "SHEET_NOT_FOUND",
  RangeInvalid: "RANGE_INVALID",
  CapabilityUnavailable: "CAPABILITY_UNAVAILABLE",
  PermissionDenied: "PERMISSION_DENIED",
  ConfirmationRequired: "CONFIRMATION_REQUIRED",
  ExternalChangeDetected: "EXTERNAL_CHANGE_DETECTED",
  TemplateMismatch: "TEMPLATE_MISMATCH",
  FormulaValidationFailed: "FORMULA_VALIDATION_FAILED",
  StyleValidationFailed: "STYLE_VALIDATION_FAILED",
  BackupUnavailable: "BACKUP_UNAVAILABLE",
  LockConflict: "LOCK_CONFLICT",
  NotFound: "NOT_FOUND",
  InvalidArgument: "INVALID_ARGUMENT",
  PayloadTooLarge: "PAYLOAD_TOO_LARGE",
  OperationFailed: "OPERATION_FAILED",
  TransactionCancelled: "TRANSACTION_CANCELLED",
  Timeout: "TIMEOUT"
} as const;

export type ExcelErrorCode = (typeof ExcelErrorCode)[keyof typeof ExcelErrorCode];

export interface ExcelRuntimeError {
  code: ExcelErrorCode;
  message: string;
  severity: ErrorSeverity;
  details?: Record<string, unknown>;
  retryable: boolean;
}

export function runtimeError(
  code: ExcelErrorCode,
  message: string,
  options: Partial<Omit<ExcelRuntimeError, "code" | "message">> = {}
): ExcelRuntimeError {
  const error: ExcelRuntimeError = {
    code,
    message,
    severity: options.severity ?? "error",
    retryable: options.retryable ?? false
  };

  if (options.details !== undefined) {
    error.details = options.details;
  }

  return error;
}
