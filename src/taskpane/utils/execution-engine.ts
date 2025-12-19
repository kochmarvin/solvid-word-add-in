/**
 * Execution engine for EditPlan actions
 */
/* global Word */

import {
  EditPlan,
  ReplaceSectionAction,
  UpdateHeadingStyleAction,
  UpdateTextFormatAction,
  CorrectTextAction,
  InsertTextAction,
  Block,
  BlockStyle,
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

    await Word.run(async (context) => {
      // For "selected" anchor, resolve via Content Control
      // For "main" anchor or other anchors, use standard resolution
      let targetRange: Word.Range;
      
      if (action.anchor === "selected") {
        // Find the Content Control with solvid-selected- tag
        const contentControls = context.document.contentControls;
        contentControls.load("items");
        await context.sync();
        
        let foundCC: Word.ContentControl | null = null;
        for (let i = 0; i < contentControls.items.length; i++) {
          const cc = contentControls.items[i];
          cc.load("tag");
          await context.sync();
          
          if (cc.tag && cc.tag.startsWith("solvid-selected-")) {
            foundCC = cc;
            break;
          }
        }
        
        if (!foundCC) {
          throw new ExecutionError("Selected content not found. Please select text and try again.");
        }
        
        targetRange = foundCC.getRange();
        targetRange.load("text");
        await context.sync();
      } else {
        // Use standard anchor resolution for other anchors
        targetRange = await resolveAnchor(action.anchor);
      }

      // For "main" anchor, clear entire body; for others, clear just the target range
      if (action.anchor === "main") {
        const body = context.document.body;
        body.clear();
        await context.sync();
        targetRange = body.getRange();
      } else {
        // Clear the target range content
        targetRange.clear();
        await context.sync();
      }

      // Insert new blocks
      let insertLocation: Word.InsertLocation.before | Word.InsertLocation.after = Word.InsertLocation.before;
      let currentRange = targetRange;

      // Insert blocks sequentially
      for (const block of action.blocks) {
        if (block.type === "paragraph") {
          const para = currentRange.insertParagraph(block.text, insertLocation);
          para.style = "Normal";
          applyBlockFormatting(para, block.style);
          currentRange = para.getRange();
          insertLocation = Word.InsertLocation.after;
        } else if (block.type === "heading") {
          const para = currentRange.insertParagraph(block.text, insertLocation);
          const headingStyle = `Heading ${block.level}`;
          para.style = headingStyle;
          applyBlockFormatting(para, block.style);
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
      } else if (action.location === "at_position") {
        // Insert at a specific position
        if (typeof action.position !== "number") {
          throw new ExecutionError("position is required when location is 'at_position'");
        }
        
        const bodyRange = body.getRange();
        bodyRange.load("start");
        await context.sync();
        
        // Create a range at the specified position
        const positionRange = bodyRange.getRange();
        // We need to find the paragraph that contains this position
        const paragraphs = body.paragraphs;
        paragraphs.load("items");
        await context.sync();
        
        let foundPara: Word.Paragraph | null = null;
        for (let i = 0; i < paragraphs.items.length; i++) {
          const para = paragraphs.items[i];
          const paraRange = para.getRange();
          paraRange.load("start,end");
          await context.sync();
          
          if (paraRange.start <= action.position && paraRange.end >= action.position) {
            foundPara = para;
            break;
          }
        }
        
        if (foundPara) {
          const paraRange = foundPara.getRange();
          paraRange.load("start");
          await context.sync();
          
          // Create a range at the exact position
          insertRange = paraRange.getRange();
          insertRange.start = action.position;
          insertRange.end = action.position;
        } else {
          // If position not found in any paragraph, use body range
          insertRange = bodyRange;
        }
      } else if (action.location === "start") {
        insertRange = body.getRange(Word.RangeLocation.start);
      } else {
        insertRange = body.getRange(Word.RangeLocation.end);
      }

      // Insert blocks sequentially
      let currentRange = insertRange;
      let insertLocation: Word.InsertLocation.before | Word.InsertLocation.after = 
        (action.location === "start" || action.location === "at_position") ? Word.InsertLocation.before : Word.InsertLocation.after;

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
          applyBlockFormatting(para, block.style);
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
 * Helper function to apply formatting to a paragraph
 * Must be called within a Word.run context
 */
function applyBlockFormatting(para: Word.Paragraph, style?: BlockStyle): void {
  if (!style) return;

  // Apply color
  if (style.color) {
    const colorHex = style.color.startsWith("#")
      ? style.color.substring(1)
      : style.color;
    para.font.color = colorHex;
  }

  // Apply bold
  if (style.bold !== undefined) {
    para.font.bold = style.bold;
  }

  // Apply alignment
  if (style.alignment) {
    const alignmentMap: Record<string, Word.Alignment> = {
      left: Word.Alignment.left,
      center: Word.Alignment.centered,
      right: Word.Alignment.right,
      justify: Word.Alignment.justified,
    };
    para.alignment = alignmentMap[style.alignment] || Word.Alignment.left;
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

    if (action.target === "specific" && !action.heading_text) {
      throw new ExecutionError("heading_text is required when target is 'specific'");
    }

    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      const headingParagraphs: Word.Paragraph[] = [];

      if (action.target === "all") {
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
      } else {
        // Find specific heading by text matching
        const searchText = action.heading_text!.toLowerCase().trim();
        let bestMatch: Word.Paragraph | null = null;
        let bestMatchScore = 0;

        for (let i = 0; i < paragraphs.items.length; i++) {
          const para = paragraphs.items[i];
          para.load("style,text");
          await context.sync();

          const style = para.style;
          if (style === "Heading 1" || style === "Heading 2" || style === "Heading 3") {
            const paraText = para.text.toLowerCase().trim();

            // Exact match - highest priority
            if (paraText === searchText) {
              bestMatch = para;
              bestMatchScore = 1.0;
              break;
            }

            // Partial matching
            let matchScore = 0;
            if (paraText.includes(searchText)) {
              matchScore = searchText.length / paraText.length;
            } else if (searchText.includes(paraText)) {
              matchScore = paraText.length / searchText.length * 0.5;
            } else {
              // Word-level matching
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

            if (matchScore > bestMatchScore) {
              bestMatchScore = matchScore;
              bestMatch = para;
            }
          }
        }

        if (bestMatch && bestMatchScore > 0.1) {
          headingParagraphs.push(bestMatch);
        } else {
          throw new ExecutionError(
            `Heading not found: "${action.heading_text}". Please check the heading text and try again.`
          );
        }
      }

      // Apply formatting to target headings
      if (headingParagraphs.length > 0) {
        headingParagraphs.forEach((heading) => {
          applyBlockFormatting(heading, action.style);
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
 * Executes an update_text_format action
 */
async function executeUpdateTextFormat(action: UpdateTextFormatAction): Promise<void> {
  try {
    // Check if Word API is available
    if (typeof Word === "undefined") {
      throw new ExecutionError("Word API is not available. Please ensure the add-in is running in Word.");
    }

    await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      const targetParagraphs: Word.Paragraph[] = [];

      // Find target paragraphs based on action.target
      for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        para.load("style");
        await context.sync();

        const style = para.style;
        const isHeading = style === "Heading 1" || style === "Heading 2" || style === "Heading 3";
        const isParagraph = style === "Normal" || (!isHeading && style !== "");

        if (action.target === "all") {
          targetParagraphs.push(para);
        } else if (action.target === "headings" && isHeading) {
          targetParagraphs.push(para);
        } else if (action.target === "paragraphs" && isParagraph) {
          targetParagraphs.push(para);
        }
      }

      // Apply formatting to target paragraphs
      if (targetParagraphs.length > 0) {
        targetParagraphs.forEach((para) => {
          applyBlockFormatting(para, action.style);
        });

        await context.sync();
      }
    });
  } catch (error) {
    if (error instanceof Error) {
      throw new ExecutionError(`Failed to execute update_text_format action: ${error.message}`, error);
    }
    throw new ExecutionError(`Failed to execute update_text_format action: ${String(error)}`, error);
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
      } else if (action.type === "update_text_format") {
        await executeUpdateTextFormat(action);
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

