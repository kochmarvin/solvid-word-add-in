/**
 * Type definitions for EditPlan schema
 * Supports paragraphs, headings, and color formatting
 */

export interface BlockStyle {
  color?: string;
}

export interface ParagraphBlock {
  type: "paragraph";
  text: string;
}

export interface HeadingBlock {
  type: "heading";
  level: 1 | 2 | 3;
  text: string;
  style?: BlockStyle;
}

export type Block = ParagraphBlock | HeadingBlock;

export interface ReplaceSectionAction {
  type: "replace_section";
  anchor: string;
  blocks: Block[];
}

export interface UpdateHeadingStyleAction {
  type: "update_heading_style";
  target: "all";
  style: BlockStyle;
}

export interface CorrectTextAction {
  type: "correct_text";
  search_text: string;
  replacement_text: string;
  case_sensitive?: boolean;
}

export interface InsertTextAction {
  type: "insert_text";
  anchor: string;
  location: "start" | "end" | "after_heading";
  heading_text?: string; // Required when location is "after_heading"
  blocks: Block[];
}

export type EditAction = ReplaceSectionAction | UpdateHeadingStyleAction | CorrectTextAction | InsertTextAction;

export interface EditPlan {
  version: "1.0";
  actions: EditAction[];
}

export interface EditPlanResponse {
  response: string;
  edit_plan: EditPlan;
}

