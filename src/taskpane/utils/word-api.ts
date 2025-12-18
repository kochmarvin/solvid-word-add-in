/**
 * Helper functions for Word API operations
 */
/* global Word */

import { ParagraphBlock, HeadingBlock } from "../types/edit-plan";

/**
 * Converts CSS color to Word color format
 * Word expects hex colors without the # prefix
 */
function convertColorToWordFormat(color: string): string {
  // Remove # if present
  if (color.startsWith("#")) {
    return color.substring(1);
  }
  // For named colors, return as-is (Word may handle some)
  return color;
}

/**
 * Inserts a paragraph block at the specified location
 */
export async function insertParagraphBlock(
  block: ParagraphBlock,
  location: Word.InsertLocation
): Promise<void> {
  await Word.run(async (context) => {
    const body = context.document.body;
    const paragraph = body.insertParagraph(block.text, location as Word.InsertLocation.start | Word.InsertLocation.end);
    paragraph.style = "Normal";
    await context.sync();
  });
}

/**
 * Inserts a heading block at the specified location
 */
export async function insertHeadingBlock(
  block: HeadingBlock,
  location: Word.InsertLocation
): Promise<void> {
  await Word.run(async (context) => {
    const body = context.document.body;
    const paragraph = body.insertParagraph(block.text, location as Word.InsertLocation.start | Word.InsertLocation.end);
    
    // Set heading style based on level
    const headingStyle = `Heading ${block.level}`;
    paragraph.style = headingStyle;
    
    // Apply color if specified
    if (block.style?.color) {
      const colorHex = convertColorToWordFormat(block.style.color);
      paragraph.font.color = colorHex;
    }
    
    await context.sync();
  });
}

/**
 * Applies color formatting to heading paragraphs
 */
export async function applyHeadingColor(
  headings: Word.Paragraph[],
  color: string
): Promise<void> {
  await Word.run(async (context) => {
    const colorHex = convertColorToWordFormat(color);
    
    headings.forEach((heading) => {
      heading.font.color = colorHex;
    });
    
    await context.sync();
  });
}

/**
 * Finds the end of a section starting from a given range
 * Section ends at: next anchor, next heading of same/higher level, or end of document
 */
export async function findSectionEnd(startRange: Word.Range): Promise<Word.Range> {
  return await Word.run(async (context) => {
    const body = context.document.body;
    const paragraphs = body.paragraphs;
    paragraphs.load("items");
    await context.sync();

    // Get the starting paragraph from the range
    startRange.load("paragraphs");
    await context.sync();
    
    if (startRange.paragraphs.items.length === 0) {
      // No paragraph found, use end of document
      return body.getRange();
    }
    
    const startPara = startRange.paragraphs.items[0];
    startPara.load("style,parentBody");
    await context.sync();

    // Find the index of the starting paragraph by comparing paragraph objects
    let startIndex = -1;
    for (let i = 0; i < paragraphs.items.length; i++) {
      const para = paragraphs.items[i];
      para.load("text");
      await context.sync();
      
      // Try to match by checking if the start range intersects with this paragraph
      const paraRange = para.getRange();
      paraRange.load("start");
      startRange.load("start");
      await context.sync();
      
      // Check if start range begins within this paragraph's range
      if (startRange.start >= paraRange.start) {
        paraRange.load("end");
        await context.sync();
        if (startRange.start <= paraRange.end) {
          startIndex = i;
          break;
        }
      }
    }

    // If we couldn't find the start, use end of document
    if (startIndex === -1) {
      return body.getRange();
    }

    // Get the heading level of the starting paragraph (if it's a heading)
    let startHeadingLevel = 0;
    if (startIndex >= 0 && startIndex < paragraphs.items.length) {
      const startPara = paragraphs.items[startIndex];
      startPara.load("style");
      await context.sync();
      
      const style = startPara.style;
      if (style === "Heading 1") startHeadingLevel = 1;
      else if (style === "Heading 2") startHeadingLevel = 2;
      else if (style === "Heading 3") startHeadingLevel = 3;
    }

    // Search forward for section end
    for (let i = startIndex + 1; i < paragraphs.items.length; i++) {
      const para = paragraphs.items[i];
      para.load("style");
      await context.sync();
      
      const style = para.style;
      
      // Check if this is a heading of same or higher level
      if (style === "Heading 1" || style === "Heading 2" || style === "Heading 3") {
        const level = style === "Heading 1" ? 1 : style === "Heading 2" ? 2 : 3;
        if (level <= startHeadingLevel || startHeadingLevel === 0) {
          // Found a heading that ends the section
          const range = para.getRange();
          range.load("text");
          await context.sync();
          return range;
        }
      }
    }

    // No heading found, use end of document
    return body.getRange();
  });
}

/**
 * Deletes content between two ranges
 */
export async function deleteRangeContent(startRange: Word.Range, endRange: Word.Range): Promise<void> {
  await Word.run(async (context) => {
    startRange.load("start");
    endRange.load("start");
    await context.sync();

    // Create a range from start to end
    const deleteRange = startRange.expandTo(endRange);
    deleteRange.delete();
    await context.sync();
  });
}

