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
 * Reads document structure (headings and content) for context awareness
 */
async function getDocumentContext(): Promise<{ 
  headings: Array<{ text: string; level: number }>;
  content_summary: string;
  has_content: boolean;
}> {
  if (typeof Word === "undefined") {
    return { headings: [], content_summary: "", has_content: false };
  }

  try {
    return await Word.run(async (context) => {
      const body = context.document.body;
      const paragraphs = body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      const headings: Array<{ text: string; level: number }> = [];
      const contentPieces: string[] = [];
      let hasContent = false;

      for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        para.load("style,text");
        await context.sync();

        const style = para.style;
        const text = para.text.trim();
        
        if (style === "Heading 1" || style === "Heading 2" || style === "Heading 3") {
          const level = style === "Heading 1" ? 1 : style === "Heading 2" ? 2 : 3;
          headings.push({
            text,
            level,
          });
        } else if (text.length > 0) {
          // Collect paragraph content (first 200 chars of each paragraph for context)
          contentPieces.push(text.substring(0, 200));
          hasContent = true;
        }
      }

      // Create a summary of document content (first few paragraphs)
      const contentSummary = contentPieces.slice(0, 5).join(" ").substring(0, 1000);

      return { 
        headings,
        content_summary: contentSummary,
        has_content: hasContent || headings.length > 0
      };
    });
  } catch (error) {
    console.error("Error reading document context:", error);
    return { headings: [], content_summary: "", has_content: false };
  }
}

/**
 * Calls the backend API to generate an EditPlan from a user prompt
 * @param prompt - The user's prompt/request
 * @param conversationHistory - Optional conversation history for context
 */
export async function generateEditPlan(
  prompt: string,
  conversationHistory?: Array<{ role: "user" | "ai"; content: string }>
): Promise<EditPlanResult> {
  try {
    // Get document context (headings) for content awareness
    const documentContext = await getDocumentContext();

    const response = await fetch(`${API_BASE_URL}/api/generate-edit-plan`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ 
        prompt,
        conversation_history: conversationHistory || [],
        document_context: documentContext
      }),
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

