export type ApiErrorCode =
  | "INVALID_SCHEMA"
  | "SLIDE_NOT_FOUND"
  | "SHAPE_NOT_FOUND"
  | "SHAPE_AMBIGUOUS"
  | "NOT_APPLICABLE"
  | "UNSUPPORTED"
  | "FAILED";

export interface ShapeCandidate {
  id: string;
  name: string;
  text: string;
}

/** An error a model can act on: what went wrong, where, and what it could have meant. */
export class ApiError extends Error {
  constructor(
    public readonly code: ApiErrorCode,
    message: string,
    public readonly path?: string,
    public readonly candidates?: ShapeCandidate[]
  ) {
    super(message);
    this.name = "ApiError";
  }

  public toJSON() {
    return {
      code: this.code,
      message: this.message,
      ...(this.path ? { path: this.path } : {}),
      ...(this.candidates ? { candidates: this.candidates } : {})
    };
  }
}
