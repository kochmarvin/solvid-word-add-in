/**
 * Execution engine for EditPlan actions
 */
/* global Word */

import {
  EditPlan,
  ReplaceSectionAction,
  UpdateHeadingStyleAction,
  CorrectTextAction,
  InsertTextAction,
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
 * Executes an insert_text action
 * Inserts blocks at a specific location without deleting existing content
 */
async function executeInsertText(action: InsertTextAction): Promise<void> {
  try {
    // Check if Word API is available
    if (typeof Word === "undefined") {
      throw new ExecutionError("Word API is not available. Please ensure the add-in is running in Word.");
    }

    await Word.run(async (context) => {
      const body = context.document.body;
      
      // Get the insertion point based on location
      let insertRange: Word.Range;
      
      if (action.location === "after_heading") {
        // Find the heading and insert after it
        if (!action.heading_text) {
          throw new ExecutionError("heading_text is required when location is 'after_heading'");
        }
        
        const paragraphs = body.paragraphs;
        paragraphs.load("items");
        await context.sync();
        
        let headingFound = false;
        const searchText = action.heading_text.toLowerCase().trim();
        let bestMatch: Word.Paragraph | null = null;
        let bestMatchScore = 0;
        
        for (let i = 0; i < paragraphs.items.length; i++) {
          const para = paragraphs.items[i];
          para.load("style,text");
          await context.sync();
          
          const style = para.style;
          const isHeading = style === "Heading 1" || style === "Heading 2" || style === "Heading 3";
          
          if (isHeading) {
            const paraText = para.text.toLowerCase().trim();
            
            // Exact match - highest priority
            if (paraText === searchText) {
              insertRange = para.getRange();
              headingFound = true;
              break;
            }
            
            // Partial matching - check if search text is contained in heading or vice versa
            // Score based on how much of the search text matches
            let matchScore = 0;
            if (paraText.includes(searchText)) {
              // Search text is contained in heading - score based on length ratio
              matchScore = searchText.length / paraText.length;
            } else if (searchText.includes(paraText)) {
              // Heading is contained in search text - lower score
              matchScore = paraText.length / searchText.length * 0.5;
            } else {
              // Check for word-level matches
              const searchWords = searchText.split(/\s+/);
              const headingWords = paraText.split(/\s+/);
              let wordMatches = 0;
              for (const word of searchWords) {
                if (word.length > 2 && headingWords.some(hw => hw.includes(word) || word.includes(hw))) {
                  wordMatches++;
                }
              }
              if (wordMatches > 0) {
                matchScore = wordMatches / searchWords.length * 0.3;
              }
            }
            
            // Keep track of best match
            if (matchScore > bestMatchScore) {
              bestMatchScore = matchScore;
              bestMatch = para;
            }
          }
        }
        
        // Use best match if we found one (even if not exact)
        if (!headingFound && bestMatch && bestMatchScore > 0.1) {
          insertRange = bestMatch.getRange();
          headingFound = true;
        }
        
        if (!headingFound) {
          throw new ExecutionError(
            `Heading not found: "${action.heading_text}". Please check the heading text and try again.`
          );
        }
      } else if (action.location === "start") {
        insertRange = body.getRange(Word.RangeLocation.start);
      } else {
        insertRange = body.getRange(Word.RangeLocation.end);
      }

      // Insert blocks sequentially
      let currentRange = insertRange;
      let insertLocation: Word.InsertLocation.before | Word.InsertLocation.after = 
        action.location === "start" ? Word.InsertLocation.before : Word.InsertLocation.after;

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
      throw new ExecutionError(`Failed to execute insert_text action: ${error.message}`, error);
    }
    throw new ExecutionError(`Failed to execute insert_text action: ${String(error)}`, error);
  }
}

/**
 * Executes a correct_text action
 * Finds and replaces text in the document
 */
async function executeCorrectText(action: CorrectTextAction): Promise<void> {
  try {
    // Check if Word API is available
    if (typeof Word === "undefined") {
      throw new ExecutionError("Word API is not available. Please ensure the add-in is running in Word.");
    }

    await Word.run(async (context) => {
      const body = context.document.body;
      const searchText = action.search_text;
      const replacementText = action.replacement_text;
      const caseSensitive = action.case_sensitive ?? false;

      // Search for the text in the document
      const searchResults = body.search(searchText, {
        matchCase: caseSensitive,
        matchWholeWord: false,
        matchWildcards: false,
      });

      searchResults.load("items");
      await context.sync();

      if (searchResults.items.length === 0) {
        throw new ExecutionError(
          `Text not found: "${searchText}". Please check the text and try again.`
        );
      }

      // Replace all occurrences
      for (let i = 0; i < searchResults.items.length; i++) {
        const range = searchResults.items[i];
        range.insertText(replacementText, Word.InsertLocation.replace);
      }

      await context.sync();
    });
  } catch (error) {
    if (error instanceof Error) {
      throw new ExecutionError(`Failed to execute correct_text action: ${error.message}`, error);
    }
    throw new ExecutionError(`Failed to execute correct_text action: ${String(error)}`, error);
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
      } else if (action.type === "correct_text") {
        await executeCorrectText(action);
      } else if (action.type === "insert_text") {
        await executeInsertText(action);
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

