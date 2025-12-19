/**
 * Strict validation for EditPlan schema
 */
import {
  EditPlan,
  EditAction,
  Block,
  ParagraphBlock,
  HeadingBlock,
  CorrectTextAction,
  InsertTextAction,
} from "../types/edit-plan";
import { ValidationError } from "./errors";

const MAX_ACTIONS = 50;
const MAX_BLOCKS_PER_ACTION = 100;
const MAX_CHARACTERS_PER_TEXT = 100000;

/**
 * Validates CSS color format (hex, rgb, rgba, or named colors)
 */
function isValidColor(color: string): boolean {
  // Hex color
  if (/^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$/.test(color)) {
    return true;
  }
  // RGB/RGBA
  if (/^rgba?\(\s*\d+\s*,\s*\d+\s*,\s*\d+\s*(,\s*[\d.]+\s*)?\)$/.test(color)) {
    return true;
  }
  // Named colors (basic check - could be expanded)
  const namedColors = [
    "black", "white", "red", "green", "blue", "yellow", "cyan", "magenta",
    "gray", "grey", "orange", "purple", "pink", "brown", "navy", "teal"
  ];
  return namedColors.includes(color.toLowerCase());
}

/**
 * Validates a paragraph block
 */
function validateParagraphBlock(block: ParagraphBlock, index: number): void {
  if (!block.text || typeof block.text !== "string") {
    throw new ValidationError(
      `Paragraph block at index ${index} must have a text property (string)`
    );
  }
  if (block.text.length > MAX_CHARACTERS_PER_TEXT) {
    throw new ValidationError(
      `Paragraph block at index ${index} exceeds maximum text length (${MAX_CHARACTERS_PER_TEXT})`
    );
  }
  if (block.text.includes("\n")) {
    throw new ValidationError(
      `Paragraph block at index ${index} contains newline characters. Paragraphs must be single blocks.`
    );
  }
  if (block.style?.color && !isValidColor(block.style.color)) {
    throw new ValidationError(
      `Paragraph block at index ${index} has invalid color format: ${block.style.color}`
    );
  }
  if (block.style?.alignment && !["left", "center", "right", "justify"].includes(block.style.alignment)) {
    throw new ValidationError(
      `Paragraph block at index ${index} has invalid alignment: ${block.style.alignment}`
    );
  }
  if (block.style?.bold !== undefined && typeof block.style.bold !== "boolean") {
    throw new ValidationError(`Paragraph block at index ${index} bold must be a boolean`);
  }
}

/**
 * Validates a heading block
 */
function validateHeadingBlock(block: HeadingBlock, index: number): void {
  if (!block.text || typeof block.text !== "string") {
    throw new ValidationError(
      `Heading block at index ${index} must have a text property (string)`
    );
  }
  if (block.level < 1 || block.level > 3 || !Number.isInteger(block.level)) {
    throw new ValidationError(
      `Heading block at index ${index} must have a level between 1 and 3`
    );
  }
  if (block.style?.color && !isValidColor(block.style.color)) {
    throw new ValidationError(
      `Heading block at index ${index} has invalid color format: ${block.style.color}`
    );
  }
  if (block.style?.alignment && !["left", "center", "right", "justify"].includes(block.style.alignment)) {
    throw new ValidationError(
      `Heading block at index ${index} has invalid alignment: ${block.style.alignment}`
    );
  }
  if (block.style?.bold !== undefined && typeof block.style.bold !== "boolean") {
    throw new ValidationError(`Heading block at index ${index} bold must be a boolean`);
  }
}

/**
 * Validates a block
 */
function validateBlock(block: Block, index: number): void {
  if (block.type === "paragraph") {
    validateParagraphBlock(block, index);
  } else if (block.type === "heading") {
    validateHeadingBlock(block, index);
  } else {
    throw new ValidationError(
      `Block at index ${index} has invalid type: ${(block as Block).type}. Only "paragraph" and "heading" are supported.`
    );
  }
}

/**
 * Validates a replace_section action
 */
function validateReplaceSectionAction(action: Extract<EditAction, { type: "replace_section" }>): void {
  if (!action.anchor || typeof action.anchor !== "string" || action.anchor.trim() === "") {
    throw new ValidationError("replace_section action must have a non-empty anchor string");
  }
  if (!Array.isArray(action.blocks)) {
    throw new ValidationError("replace_section action must have a blocks array");
  }
  if (action.blocks.length === 0) {
    throw new ValidationError("replace_section action must have at least one block");
  }
  if (action.blocks.length > MAX_BLOCKS_PER_ACTION) {
    throw new ValidationError(
      `replace_section action exceeds maximum blocks per action (${MAX_BLOCKS_PER_ACTION})`
    );
  }
  action.blocks.forEach((block, index) => {
    validateBlock(block, index);
  });
}

/**
 * Validates an update_heading_style action
 */
function validateUpdateHeadingStyleAction(action: Extract<EditAction, { type: "update_heading_style" }>): void {
  if (!["all", "specific"].includes(action.target)) {
    throw new ValidationError('update_heading_style action target must be "all" or "specific"');
  }
  if (action.target === "specific") {
    if (!action.heading_text || typeof action.heading_text !== "string" || action.heading_text.trim() === "") {
      throw new ValidationError('update_heading_style action with target "specific" must have a non-empty heading_text');
    }
  }
  if (!action.style || typeof action.style !== "object") {
    throw new ValidationError("update_heading_style action must have a style object");
  }
  if (action.style.color && !isValidColor(action.style.color)) {
    throw new ValidationError(`update_heading_style action has invalid color format: ${action.style.color}`);
  }
  if (action.style.alignment && !["left", "center", "right", "justify"].includes(action.style.alignment)) {
    throw new ValidationError(`update_heading_style action has invalid alignment: ${action.style.alignment}`);
  }
  if (action.style.bold !== undefined && typeof action.style.bold !== "boolean") {
    throw new ValidationError("update_heading_style action bold must be a boolean");
  }
}

/**
 * Validates an update_text_format action
 */
function validateUpdateTextFormatAction(action: Extract<EditAction, { type: "update_text_format" }>): void {
  if (!["all", "headings", "paragraphs"].includes(action.target)) {
    throw new ValidationError('update_text_format action target must be "all", "headings", or "paragraphs"');
  }
  if (!action.style || typeof action.style !== "object") {
    throw new ValidationError("update_text_format action must have a style object");
  }
  if (action.style.color && !isValidColor(action.style.color)) {
    throw new ValidationError(`update_text_format action has invalid color format: ${action.style.color}`);
  }
  if (action.style.alignment && !["left", "center", "right", "justify"].includes(action.style.alignment)) {
    throw new ValidationError(`update_text_format action has invalid alignment: ${action.style.alignment}`);
  }
  if (action.style.bold !== undefined && typeof action.style.bold !== "boolean") {
    throw new ValidationError("update_text_format action bold must be a boolean");
  }
}

/**
 * Validates an insert_text action
 */
function validateInsertTextAction(action: InsertTextAction): void {
  if (!action.anchor || typeof action.anchor !== "string" || action.anchor.trim() === "") {
    throw new ValidationError("insert_text action must have a non-empty anchor string");
  }
  if (action.location !== "start" && action.location !== "end" && action.location !== "after_heading" && action.location !== "at_position") {
    throw new ValidationError('insert_text action location must be "start", "end", "after_heading", or "at_position"');
  }
  if (action.location === "after_heading") {
    if (!action.heading_text || typeof action.heading_text !== "string" || action.heading_text.trim() === "") {
      throw new ValidationError("insert_text action must have heading_text when location is 'after_heading'");
    }
  }
  if (action.location === "at_position") {
    if (typeof action.position !== "number") {
      throw new ValidationError("insert_text action must have position (number) when location is 'at_position'");
    }
  }
  if (!Array.isArray(action.blocks)) {
    throw new ValidationError("insert_text action must have a blocks array");
  }
  if (action.blocks.length === 0) {
    throw new ValidationError("insert_text action must have at least one block");
  }
  if (action.blocks.length > MAX_BLOCKS_PER_ACTION) {
    throw new ValidationError(
      `insert_text action exceeds maximum blocks per action (${MAX_BLOCKS_PER_ACTION})`
    );
  }
  action.blocks.forEach((block, index) => {
    validateBlock(block, index);
  });
}

/**
 * Validates a correct_text action
 */
function validateCorrectTextAction(action: CorrectTextAction): void {
  if (!action.search_text || typeof action.search_text !== "string") {
    throw new ValidationError("correct_text action must have a search_text string");
  }
  if (action.search_text.length === 0) {
    throw new ValidationError("correct_text action search_text cannot be empty");
  }
  if (action.search_text.length > MAX_CHARACTERS_PER_TEXT) {
    throw new ValidationError(
      `correct_text action search_text exceeds maximum length (${MAX_CHARACTERS_PER_TEXT})`
    );
  }
  if (action.replacement_text === undefined || typeof action.replacement_text !== "string") {
    throw new ValidationError("correct_text action must have a replacement_text string");
  }
  if (action.replacement_text.length > MAX_CHARACTERS_PER_TEXT) {
    throw new ValidationError(
      `correct_text action replacement_text exceeds maximum length (${MAX_CHARACTERS_PER_TEXT})`
    );
  }
  if (action.case_sensitive !== undefined && typeof action.case_sensitive !== "boolean") {
    throw new ValidationError("correct_text action case_sensitive must be a boolean if provided");
  }
}

/**
 * Validates an action
 */
function validateAction(action: EditAction, index: number): void {
  if (action.type === "replace_section") {
    validateReplaceSectionAction(action);
  } else if (action.type === "update_heading_style") {
    validateUpdateHeadingStyleAction(action);
  } else if (action.type === "update_text_format") {
    validateUpdateTextFormatAction(action);
  } else if (action.type === "correct_text") {
    validateCorrectTextAction(action);
  } else if (action.type === "insert_text") {
    validateInsertTextAction(action);
  } else {
    throw new ValidationError(
      `Action at index ${index} has invalid type: ${(action as EditAction).type}. Only "replace_section", "update_heading_style", "update_text_format", "correct_text", and "insert_text" are supported.`
    );
  }
}

/**
 * Validates an EditPlan
 */
export function validateEditPlan(editPlan: unknown): EditPlan {
  if (!editPlan || typeof editPlan !== "object") {
    throw new ValidationError("EditPlan must be an object");
  }

  const plan = editPlan as Record<string, unknown>;

  // Validate version
  if (plan.version !== "1.0") {
    throw new ValidationError(`EditPlan version must be "1.0", got: ${plan.version}`);
  }

  // Validate actions
  if (!Array.isArray(plan.actions)) {
    throw new ValidationError("EditPlan must have an actions array");
  }
  
  // Allow empty actions array - this might be a semantic plan placeholder
  // The actual validation will be done at the API service level
  // Semantic plans use 'ops' instead of 'actions', so empty actions is valid for them
  if (plan.actions.length === 0) {
    // Return early - allow empty actions (semantic plans will be validated separately)
    // Create a proper EditPlan object with validated properties
    return {
      version: plan.version as "1.0",
      actions: plan.actions as EditAction[]
    };
  }
  if (plan.actions.length > MAX_ACTIONS) {
    throw new ValidationError(
      `EditPlan exceeds maximum actions (${MAX_ACTIONS})`
    );
  }

  // Validate each action
  plan.actions.forEach((action, index) => {
    validateAction(action as EditAction, index);
  });

  // Check for unexpected fields (fail-closed)
  const allowedFields = ["version", "actions"];
  const unexpectedFields = Object.keys(plan).filter(key => !allowedFields.includes(key));
  if (unexpectedFields.length > 0) {
    throw new ValidationError(
      `EditPlan contains unexpected fields: ${unexpectedFields.join(", ")}`
    );
  }

  return editPlan as EditPlan;
}

