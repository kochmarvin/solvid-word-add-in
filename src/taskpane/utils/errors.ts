/**
 * Structured error types for EditPlan processing
 */

export class ValidationError extends Error {
  constructor(
    message: string,
    public details?: Record<string, unknown>
  ) {
    super(message);
    this.name = "ValidationError";
  }
}

export class AnchorNotFoundError extends Error {
  constructor(
    message: string,
    public anchor: string
  ) {
    super(message);
    this.name = "AnchorNotFoundError";
  }
}

export class ExecutionError extends Error {
  constructor(
    message: string,
    public originalError?: unknown
  ) {
    super(message);
    this.name = "ExecutionError";
  }
}

export interface ErrorResult {
  ok: false;
  error_type: "validation" | "anchor_not_found" | "execution_failed";
  message: string;
  details?: Record<string, unknown>;
}

export interface SuccessResult {
  ok: true;
  message: string;
}

export type ExecutionResult = SuccessResult | ErrorResult;

