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

export type EditAction = ReplaceSectionAction | UpdateHeadingStyleAction;

export interface EditPlan {
  version: "1.0";
  actions: EditAction[];
}

export interface EditPlanResponse {
  response: string;
  edit_plan: EditPlan;
}

