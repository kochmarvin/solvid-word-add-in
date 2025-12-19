/**
 * API service for communicating with backend to generate EditPlan
 */
import { EditPlanResponse, EditPlan } from "../types/edit-plan";
import { SemanticDocument, SemanticEditResponse, DocumentBlock } from "../types/semantic-edit";
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
  semanticEditPlan?: { ops: Array<{ action: string; target_block_id: string; content: string; reason: string }> };
}

export interface EditPlanApiError {
  ok: false;
  error: ApiError;
}

export type EditPlanResult = EditPlanApiResponse | EditPlanApiError;

/**
 * Extracts keywords from user prompt for intelligent content search
 */
function extractKeywords(prompt: string): string[] {
  // Remove common stop words and extract meaningful keywords
  const stopWords = new Set([
    "the", "a", "an", "and", "or", "but", "in", "on", "at", "to", "for", "of", "with", "by",
    "is", "are", "was", "were", "be", "been", "being", "have", "has", "had", "do", "does", "did",
    "will", "would", "should", "could", "may", "might", "must", "can", "this", "that", "these", "those",
    "i", "you", "he", "she", "it", "we", "they", "his", "her", "its", "our", "their",
    "add", "insert", "write", "create", "make", "change", "fix", "update", "more", "about", "information"
  ]);
  
  const words = prompt.toLowerCase()
    .replace(/[^\w\s]/g, " ")
    .split(/\s+/)
    .filter(word => word.length > 2 && !stopWords.has(word));
  
  // Remove duplicates
  const uniqueWords: string[] = [];
  const seen = new Set<string>();
  for (const word of words) {
    if (!seen.has(word)) {
      seen.add(word);
      uniqueWords.push(word);
    }
  }
  return uniqueWords;
}

/**
 * Extracts semantic document structure with sections and blocks with stable IDs
 * Exported for use in execution engine
 */
export async function getSemanticDocument(): Promise<SemanticDocument> {
  if (typeof Word === "undefined") {
    return { sections: [], blocks: {} };
  }

  try {
    return await Word.run(async (context) => {
      const body = context.document.body;
      const paragraphs = body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      const sections: Array<{ id: string; title: string; level: number; blocks: string[] }> = [];
      const blocks: Record<string, DocumentBlock> = {};
      
      let blockCounter = 1;
      let sectionCounter = 1;
      let currentSectionId: string | null = null;
      let currentSectionBlocks: string[] = [];

      for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        para.load("style,text");
        await context.sync();

        const style = para.style;
        const text = para.text.trim();
        
        if (!text) continue;

        const blockId = `b${blockCounter++}`;

        if (style === "Heading 1" || style === "Heading 2" || style === "Heading 3") {
          // Save previous section if it exists
          if (currentSectionId && currentSectionBlocks.length > 0) {
            sections.push({
              id: currentSectionId,
              title: sections.find(s => s.id === currentSectionId)?.title || "",
              level: sections.find(s => s.id === currentSectionId)?.level || 1,
              blocks: currentSectionBlocks
            });
          }

          // Create new section
          const level = style === "Heading 1" ? 1 : style === "Heading 2" ? 2 : 3;
          const sectionId = `s${sectionCounter++}`;
          
          blocks[blockId] = {
            id: blockId,
            type: "heading",
            text,
            level
          };

          sections.push({
            id: sectionId,
            title: text,
            level,
            blocks: [blockId]
          });

          currentSectionId = sectionId;
          currentSectionBlocks = [blockId];
        } else {
          // Regular paragraph
          blocks[blockId] = {
            id: blockId,
            type: "paragraph",
            text
          };

          if (currentSectionId) {
            currentSectionBlocks.push(blockId);
          } else {
            // No section yet, create a default one
            const sectionId = `s${sectionCounter++}`;
            sections.push({
              id: sectionId,
              title: "Introduction",
              level: 1,
              blocks: [blockId]
            });
            currentSectionId = sectionId;
            currentSectionBlocks = [blockId];
          }
        }
      }

      // Save last section
      if (currentSectionId && currentSectionBlocks.length > 0) {
        const existingSection = sections.find(s => s.id === currentSectionId);
        if (existingSection) {
          existingSection.blocks = currentSectionBlocks;
        }
      }

      return {
        sections: sections.map(s => ({
          id: s.id,
          title: s.title,
          level: s.level,
          blocks: s.blocks
        })),
        blocks
      };
    });
  } catch (error) {
    console.error("Error extracting semantic document:", error);
    return { sections: [], blocks: {} };
  }
}

/**
 * Reads document structure and relevant content based on user prompt
 */
async function getDocumentContext(prompt: string): Promise<{ 
  headings: Array<{ text: string; level: number }>;
  heading_hierarchy?: string;
  relevant_content: Array<{ heading: string; level: number; paragraphs: string[] }>;
  content_summary: string;
  has_content: boolean;
}> {
  if (typeof Word === "undefined") {
    return { headings: [], heading_hierarchy: "", relevant_content: [], content_summary: "", has_content: false };
  }

  try {
    return await Word.run(async (context) => {
      const body = context.document.body;
      const paragraphs = body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      const headings: Array<{ text: string; level: number; index: number }> = [];
      const relevantContent: Array<{ heading: string; level: number; paragraphs: string[]; heading_index: number }> = [];
      const allParagraphs: string[] = [];
      let hasContent = false;
      let currentHeading: string | null = null;
      let currentHeadingLevel = 0;
      let currentHeadingIndex = -1;
      let currentSectionParagraphs: string[] = [];

      // Extract keywords from prompt for intelligent matching
      const keywords = extractKeywords(prompt);
      const promptLower = prompt.toLowerCase();
      
      // Check if this is a reordering/restructuring request
      const isReorderingRequest = /reorder|reorganize|restructure|rewrite.*order|rearrange|reorder.*text/i.test(prompt);

      for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        para.load("style,text");
        await context.sync();

        const style = para.style;
        const text = para.text.trim();
        
        if (style === "Heading 1" || style === "Heading 2" || style === "Heading 3") {
          // Save previous section if it was relevant
          if (currentHeading && currentSectionParagraphs.length > 0) {
            const sectionText = (currentHeading + " " + currentSectionParagraphs.join(" ")).toLowerCase();
            const isRelevant = keywords.some(keyword => sectionText.includes(keyword)) ||
                             promptLower.includes(currentHeading.toLowerCase().substring(0, 10));
            
            if (isRelevant) {
              relevantContent.push({
                heading: currentHeading,
                level: currentHeadingLevel,
                paragraphs: currentSectionParagraphs.slice(0, 3), // Max 3 paragraphs per section
                heading_index: currentHeadingIndex
              });
            }
          }
          
          const level = style === "Heading 1" ? 1 : style === "Heading 2" ? 2 : 3;
          headings.push({
            text,
            level,
            index: headings.length
          });
          
          currentHeading = text;
          currentHeadingLevel = level;
          currentHeadingIndex = headings.length - 1;
          currentSectionParagraphs = [];
        } else if (text.length > 0) {
          allParagraphs.push(text);
          currentSectionParagraphs.push(text.substring(0, 300)); // First 300 chars per paragraph
          hasContent = true;
        }
      }

      // Check last section
      if (currentHeading && currentSectionParagraphs.length > 0) {
        const sectionText = (currentHeading + " " + currentSectionParagraphs.join(" ")).toLowerCase();
        const isRelevant = keywords.some(keyword => sectionText.includes(keyword)) ||
                         promptLower.includes(currentHeading.toLowerCase().substring(0, 10));
        
        if (isRelevant) {
          relevantContent.push({
            heading: currentHeading,
            level: currentHeadingLevel,
            paragraphs: currentSectionParagraphs.slice(0, 3),
            heading_index: currentHeadingIndex
          });
        }
      }

      // Create a summary from first few paragraphs if no relevant content found
      const contentSummary = allParagraphs.length > 0 
        ? allParagraphs.slice(0, 3).join(" ").substring(0, 500)
        : "";
      
      // For reordering requests, include all paragraphs in relevant_content so AI can reorder them
      if (isReorderingRequest && allParagraphs.length > 0) {
        // Add all paragraphs as a single "section" for reordering
        relevantContent.push({
          heading: currentHeading || "Document Content",
          level: currentHeadingLevel || 1,
          paragraphs: allParagraphs, // Include ALL paragraphs for reordering
          heading_index: currentHeadingIndex >= 0 ? currentHeadingIndex : 0
        });
      }

      // Build hierarchical structure representation
      const headingHierarchy: string[] = [];
      const parentStack: Array<{ text: string; level: number }> = [];
      
      for (const heading of headings) {
        // Pop parents that are at same or higher level
        while (parentStack.length > 0 && parentStack[parentStack.length - 1].level >= heading.level) {
          parentStack.pop();
        }
        
        // Build hierarchy path
        const path = parentStack.length > 0 
          ? `${parentStack.map(p => p.text).join(" > ")} > ${heading.text}`
          : heading.text;
        
        headingHierarchy.push(`  ${"  ".repeat(heading.level - 1)}[H${heading.level}] ${heading.text}`);
        
        // Add to parent stack
        parentStack.push({ text: heading.text, level: heading.level });
      }

      return { 
        headings: headings.map(h => ({ text: h.text, level: h.level })), // Remove index for API
        heading_hierarchy: headingHierarchy.join("\n"), // Hierarchical representation
        relevant_content: relevantContent,
        content_summary: contentSummary,
        has_content: hasContent || headings.length > 0
      };
    });
  } catch (error) {
    console.error("Error reading document context:", error);
    return { headings: [], heading_hierarchy: "", relevant_content: [], content_summary: "", has_content: false };
  }
}

/**
 * Calls the backend API to generate an EditPlan from a user prompt
 * @param prompt - The user's prompt/request
 * @param conversationHistory - Optional conversation history for context
 * @param selectedRange - Optional selected text range to target
 */
export async function generateEditPlan(
  prompt: string,
  conversationHistory?: Array<{ role: "user" | "ai"; content: string }>,
  selectedRange?: { tag: string; text: string } | null
): Promise<EditPlanResult> {
  try {
    // Get semantic document structure (sections and blocks with IDs)
    const semanticDocument = await getSemanticDocument();
    
    // Also get legacy document context for backward compatibility
    const documentContext = await getDocumentContext(prompt);

    const response = await fetch(`${API_BASE_URL}/api/generate-edit-plan`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ 
        prompt,
        conversation_history: conversationHistory || [],
        document_context: documentContext,
        semantic_document: semanticDocument,
        selected_range: selectedRange ? {
          tag: selectedRange.tag,
          text: selectedRange.text
        } : null
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

    // Check if this is a semantic edit plan (has "ops" field) or legacy edit plan (has "edit_plan" field)
    if (responseData.ops && Array.isArray(responseData.ops)) {
      // This is a semantic edit plan
      const semanticEditPlan = responseData as { response: string; ops: Array<{ action: string; target_block_id: string; content: string; reason: string }> };
      
      // Validate that ops array is not empty
      if (semanticEditPlan.ops.length === 0) {
        return {
          ok: false,
          error: {
            message: "Semantic edit plan must have at least one operation",
          },
        };
      }
      
      const response: EditPlanApiResponse = {
        ok: true,
        response: semanticEditPlan.response || "Semantic edit plan generated",
        editPlan: {
          version: "1.0",
          actions: [] // Semantic plans use ops, not actions - empty array is OK for semantic plans
        } as EditPlan,
        semanticEditPlan: {
          ops: semanticEditPlan.ops
        }
      };
      
      console.log("API Service - Returning semantic edit plan with ops count:", semanticEditPlan.ops.length);
      return response;
    }

    // Check for edit_plan field (legacy format)
    if (!responseData.edit_plan) {
      return {
        ok: false,
        error: {
          message: "Invalid API response: missing edit_plan or ops field",
        },
      };
    }

    // Validate the EditPlan schema (legacy format)
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

