/**
 * Semantic document editing types
 * Based on semantic document model with sections and blocks
 */

export interface DocumentBlock {
  id: string; // Stable block ID (e.g., "b1", "b2")
  type: "paragraph" | "heading";
  text: string;
  level?: number; // For headings
}

export interface DocumentSection {
  id: string; // Stable section ID (e.g., "s1", "s2")
  title: string;
  level: number;
  blocks: string[]; // Array of block IDs
}

export interface SemanticDocument {
  sections: DocumentSection[];
  blocks: Record<string, DocumentBlock>; // Map of block_id -> block
}

export interface SemanticOperation {
  action: "insert_after" | "insert_before" | "replace";
  target_block_id: string;
  content: string;
  reason: string;
}

export interface SemanticEditPlan {
  ops: SemanticOperation[];
}

export interface SemanticEditResponse {
  response: string;
  edit_plan: SemanticEditPlan;
}

