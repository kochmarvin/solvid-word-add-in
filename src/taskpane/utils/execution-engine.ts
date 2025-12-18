/**
 * Execution engine for EditPlan actions
 */
/* global Word */

import {
  EditPlan,
  ReplaceSectionAction,
  UpdateHeadingStyleAction,
  Block,
} from "../types/edit-plan";
import { ExecutionError, ExecutionResult, AnchorNotFoundError } from "./errors";
import {
  insertParagraphBlock,
  insertHeadingBlock,
  applyHeadingColor,
  findSectionEnd,
  deleteRangeContent,
} from "./word-api";

/**
 * Multi-strategy anchor resolution for EditPlan actions
 */


export interface AnchorResolutionResult {
  found: boolean;
  range?: Word.Range;
  error?: string;
}

/**
 * Strategy 1: Content Controls
 * Searches for a content control with matching tag
 */
async function resolveContentControl(anchor: string): Promise<AnchorResolutionResult> {
  try {
    return await Word.run(async (context) => {
      const contentControls = context.document.contentControls;
      contentControls.load("items");
      await context.sync();

      for (let i = 0; i < contentControls.items.length; i++) {
        const cc = contentControls.items[i];
        cc.load("tag");
        await context.sync();

        if (cc.tag === anchor) {
          const range = cc.getRange();
          range.load("text");
          await context.sync();
          return { found: true, range };
        }
      }

      return { found: false };
    });
  } catch (error) {
    return {
      found: false,
      error: `Error resolving content control: ${error}`,
    };
  }
}

/**
 * Strategy 2: Bookmarks
 * Searches for a bookmark with matching name
 */
async function resolveBookmark(anchor: string): Promise<AnchorResolutionResult> {
  try {
    return await Word.run(async (context) => {
      const bookmarks = context.document.bookmarks;
      bookmarks.load("items");
      await context.sync();

      for (let i = 0; i < bookmarks.items.length; i++) {
        const bookmark = bookmarks.items[i];
        bookmark.load("name");
        await context.sync();

        if (bookmark.name === anchor) {
          bookmark.load("range");
          await context.sync();
          const range = bookmark.range;
          range.load("text");
          await context.sync();
          return { found: true, range };
        }
      }

      return { found: false };
    });
  } catch (error) {
    return {
      found: false,
      error: `Error resolving bookmark: ${error}`,
    };
  }
}

/**
 * Strategy 3: Heading Text Matching
 * Matches heading paragraph text to anchor (case-insensitive, partial match)
 */
async function resolveHeadingText(anchor: string): Promise<AnchorResolutionResult> {
  try {
    return await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      const anchorLower = anchor.toLowerCase().trim();

      for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        para.load("style,text");
        await context.sync();

        const style = para.style;
        const isHeading = style === "Heading 1" || style === "Heading 2" || style === "Heading 3";

        if (isHeading) {
          const paraText = para.text.toLowerCase().trim();
          // Check for exact match or if anchor is contained in heading text
          if (paraText === anchorLower || paraText.includes(anchorLower) || anchorLower.includes(paraText)) {
            const range = para.getRange();
            range.load("text");
            await context.sync();
            return { found: true, range };
          }
        }
      }

      return { found: false };
    });
  } catch (error) {
    return {
      found: false,
      error: `Error resolving heading text: ${error}`,
    };
  }
}

/**
 * Strategy 4: Default "main" anchor
 * Returns current selection or end of document
 */
async function resolveDefaultAnchor(): Promise<AnchorResolutionResult> {
  return await Word.run(async (context) => {
    try {
      // Try to get current selection
      const range = context.document.getSelection();
      range.load("text");
      await context.sync();
      return { found: true, range };
    } catch {
      // If no selection, use body range
      const range = context.document.body.getRange();
      range.load("text");
      await context.sync();
      return { found: true, range };
    }
  });
}



/**
 * Resolves an anchor using multiple strategies in order
 */
export async function resolveAnchor(anchor: string): Promise<Word.Range> {
  // Check if Word API is available
  if (typeof Word === "undefined") {
    throw new AnchorNotFoundError(
      "Word API is not available. Please ensure the add-in is running in Word.",
      anchor
    );
  }

  // Strategy 1: Content Controls
  let result = await resolveContentControl(anchor);
  if (result.found && result.range) {
    return result.range;
  }

  // Strategy 2: Bookmarks
  result = await resolveBookmark(anchor);
  if (result.found && result.range) {
    return result.range;
  }

  // Strategy 3: Heading Text
  result = await resolveHeadingText(anchor);
  if (result.found && result.range) {
    return result.range;
  }

  // Strategy 4: Default for "main" anchor
  if (anchor === "main") {
    result = await resolveDefaultAnchor();
    if (result.found && result.range) {
      return result.range;
    }
  }

  // All strategies failed
  throw new AnchorNotFoundError(
    `Could not resolve anchor "${anchor}". Tried content controls, bookmarks, heading text, and default location.`,
    anchor
  );
}


/**
 * Executes a replace_section action
 */
async function executeReplaceSection(action: ReplaceSectionAction): Promise<void> {
  try {
    // Check if Word API is available
    if (typeof Word === "undefined") {
      throw new ExecutionError("Word API is not available. Please ensure the add-in is running in Word.");
    }

    // For "main" anchor, replace entire document body content
    // For other anchors, we'll also use the same approach for simplicity
    await Word.run(async (context) => {
      const body = context.document.body;
      
      // Clear the body content
      body.clear();
      await context.sync();

      // Get the body range for insertion
      const bodyRange = body.getRange();
      let insertLocation: Word.InsertLocation.before | Word.InsertLocation.after = Word.InsertLocation.before;
      let currentRange = bodyRange;

      // Insert blocks sequentially
      for (const block of action.blocks) {
        if (block.type === "paragraph") {
          const para = currentRange.insertParagraph(block.text, insertLocation);
          para.style = "Normal";
          currentRange = para.getRange();
          insertLocation = Word.InsertLocation.after;
        } else if (block.type === "heading") {
          const para = currentRange.insertParagraph(block.text, insertLocation);
          const headingStyle = `Heading ${block.level}`;
          para.style = headingStyle;

          if (block.style?.color) {
            const colorHex = block.style.color.startsWith("#")
              ? block.style.color.substring(1)
              : block.style.color;
            para.font.color = colorHex;
          }

          currentRange = para.getRange();
          insertLocation = Word.InsertLocation.after;
        }
      }

      await context.sync();
    });
  } catch (error) {
    if (error instanceof Error) {
      throw new ExecutionError(`Failed to execute replace_section action: ${error.message}`, error);
    }
    throw new ExecutionError(`Failed to execute replace_section action: ${String(error)}`, error);
  }
}

/**
 * Executes an update_heading_style action
 */
async function executeUpdateHeadingStyle(action: UpdateHeadingStyleAction): Promise<void> {
  try {
    // Check if Word API is available
    if (typeof Word === "undefined") {
      throw new ExecutionError("Word API is not available. Please ensure the add-in is running in Word.");
    }

    if (action.target !== "all") {
      throw new ExecutionError(`Unsupported target for update_heading_style: ${action.target}`);
    }

    if (!action.style.color) {
      // No color to apply
      return;
    }

    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      const headingParagraphs: Word.Paragraph[] = [];

      // Find all heading paragraphs
      for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        para.load("style");
        await context.sync();

        const style = para.style;
        if (style === "Heading 1" || style === "Heading 2" || style === "Heading 3") {
          headingParagraphs.push(para);
        }
      }

      // Apply color to all headings
      if (headingParagraphs.length > 0) {
        const colorHex = action.style.color.startsWith("#")
          ? action.style.color.substring(1)
          : action.style.color;

        headingParagraphs.forEach((heading) => {
          heading.font.color = colorHex;
        });

        await context.sync();
      }
    });
  } catch (error) {
    if (error instanceof Error) {
      throw new ExecutionError(`Failed to execute update_heading_style action: ${error.message}`, error);
    }
    throw new ExecutionError(`Failed to execute update_heading_style action: ${String(error)}`, error);
  }
}

/**
 * Executes an EditPlan
 */
export async function executeEditPlan(editPlan: EditPlan): Promise<ExecutionResult> {
  try {
    for (const action of editPlan.actions) {
      if (action.type === "replace_section") {
        await executeReplaceSection(action);
      } else if (action.type === "update_heading_style") {
        await executeUpdateHeadingStyle(action);
      } else {
        throw new ExecutionError(`Unknown action type: ${(action as { type: string }).type}`);
      }
    }

    return {
      ok: true,
      message: `Successfully executed ${editPlan.actions.length} action(s)`,
    };
  } catch (error) {
    if (error instanceof ExecutionError) {
      return {
        ok: false,
        error_type: "execution_failed",
        message: error.message,
        details: {
          originalError: error.originalError,
        },
      };
    }

    return {
      ok: false,
      error_type: "execution_failed",
      message: `Execution failed: ${error instanceof Error ? error.message : String(error)}`,
      details: {
        error,
      },
    };
  }
}

