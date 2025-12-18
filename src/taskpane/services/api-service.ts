/**
 * API service for communicating with backend to generate EditPlan
 */
import { EditPlanResponse, EditPlan } from "../types/edit-plan";
import { ValidationError } from "../utils/errors";
import { validateEditPlan } from "../utils/validation";

// API base URL - can be configured via window variable or defaults to localhost:8000
// In browser environment, process.env is not available, so we use window or default
const getApiBaseUrl = (): string => {
  // Check window for API_BASE_URL (can be set via script tag or webpack)
  if (typeof window !== "undefined" && (window as any).API_BASE_URL) {
    return (window as any).API_BASE_URL;
  }
  // Fallback to default development URL
  return "http://localhost:8000";
};

const API_BASE_URL = getApiBaseUrl();

export interface ApiError {
  message: string;
  status?: number;
  details?: unknown;
}

export interface EditPlanApiResponse {
  ok: true;
  response: string;
  editPlan: EditPlan;
}

export interface EditPlanApiError {
  ok: false;
  error: ApiError;
}

export type EditPlanResult = EditPlanApiResponse | EditPlanApiError;

/**
 * Calls the backend API to generate an EditPlan from a user prompt
 */
export async function generateEditPlan(prompt: string): Promise<EditPlanResult> {
  try {
    const response = await fetch(`${API_BASE_URL}/api/generate-edit-plan`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ prompt }),
    });

    if (!response.ok) {
      const errorText = await response.text();
      return {
        ok: false,
        error: {
          message: `API request failed with status ${response.status}`,
          status: response.status,
          details: errorText,
        },
      };
    }

    const data: unknown = await response.json();

    // Validate the response structure
    if (!data || typeof data !== "object") {
      return {
        ok: false,
        error: {
          message: "Invalid API response: expected an object",
        },
      };
    }

    const responseData = data as Record<string, unknown>;

    // Check for edit_plan field
    if (!responseData.edit_plan) {
      return {
        ok: false,
        error: {
          message: "Invalid API response: missing edit_plan field",
        },
      };
    }

    // Validate the EditPlan schema
    let editPlan: EditPlan;
    try {
      editPlan = validateEditPlan(responseData.edit_plan);
    } catch (error) {
      if (error instanceof ValidationError) {
        return {
          ok: false,
          error: {
            message: `EditPlan validation failed: ${error.message}`,
            details: error.details,
          },
        };
      }
      throw error;
    }

    // Extract response text
    const responseText = typeof responseData.response === "string"
      ? responseData.response
      : "Edit plan generated successfully";

    return {
      ok: true,
      response: responseText,
      editPlan,
    };
  } catch (error) {
    // Handle network errors, JSON parsing errors, etc.
    return {
      ok: false,
      error: {
        message: error instanceof Error ? error.message : "Unknown error occurred",
        details: error,
      },
    };
  }
}

