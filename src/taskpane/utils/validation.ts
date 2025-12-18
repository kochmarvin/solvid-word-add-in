/**
 * Strict validation for EditPlan schema
 */
import {
  EditPlan,
  EditAction,
  Block,
  ParagraphBlock,
  HeadingBlock,
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
  if (action.target !== "all") {
    throw new ValidationError('update_heading_style action target must be "all"');
  }
  if (!action.style || typeof action.style !== "object") {
    throw new ValidationError("update_heading_style action must have a style object");
  }
  if (action.style.color && !isValidColor(action.style.color)) {
    throw new ValidationError(`update_heading_style action has invalid color format: ${action.style.color}`);
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
  } else {
    throw new ValidationError(
      `Action at index ${index} has invalid type: ${(action as EditAction).type}. Only "replace_section" and "update_heading_style" are supported.`
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
  if (plan.actions.length === 0) {
    throw new ValidationError("EditPlan must have at least one action");
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

