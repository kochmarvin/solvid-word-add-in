/**
 * Execution engine for semantic document editing operations
 * Handles insert_after, insert_before, and replace operations based on block IDs
 */
import { SemanticOperation, SemanticEditPlan } from "../types/semantic-edit";
import { ExecutionError } from "./errors";

/**
 * Executes semantic edit operations
 */
export async function executeSemanticEditPlan(editPlan: SemanticEditPlan, documentStructure: { sections: Array<{ id: string; title: string; level: number; blocks: string[] }>; blocks: Record<string, { type: "paragraph" | "heading"; text: string; level?: number; id?: string }> }): Promise<void> {
  if (typeof Word === "undefined") {
    throw new ExecutionError("Word API is not available. Please ensure the add-in is running in Word.");
  }

  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      const paragraphs = body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      // Load all paragraph data at once for efficiency
      const paragraphData: Array<{ index: number; text: string; style: string }> = [];
      for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        para.load("style,text");
      }
      await context.sync();
      
      for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        paragraphData.push({
          index: i,
          text: para.text.trim(),
          style: para.style
        });
      }

      // Build a map of block_id to paragraph index using the document structure
      const blockIdToParagraphIndex: Record<string, number> = {};
      const usedParagraphIndices = new Set<number>();
      
      // Iterate through all blocks in order (from sections)
      for (const section of documentStructure.sections) {
        for (const blockId of section.blocks) {
          // Find the paragraph that matches this block
          const block = documentStructure.blocks[blockId];
          if (!block) {
            console.warn(`Block ${blockId} not found in document structure`);
            continue;
          }
          
          // Search for matching paragraph (try exact match first, then fuzzy)
          let matchedIndex = -1;
          
          // First pass: exact text match
          for (let i = 0; i < paragraphData.length; i++) {
            if (usedParagraphIndices.has(i)) continue;
            
            const para = paragraphData[i];
            const paraText = para.text;
            const blockText = block.text.trim();
            
            if (paraText === blockText) {
              // For headings, also check style
              if (block.type === "heading") {
                const expectedStyle = `Heading ${block.level}`;
                if (para.style === expectedStyle) {
                  matchedIndex = i;
                  break;
                }
              } else {
                // Regular paragraph - exact match
                matchedIndex = i;
                break;
              }
            }
          }
          
          // Second pass: fuzzy match if exact match failed
          if (matchedIndex === -1) {
            for (let i = 0; i < paragraphData.length; i++) {
              if (usedParagraphIndices.has(i)) continue;
              
              const para = paragraphData[i];
              const paraText = para.text.toLowerCase();
              const blockText = block.text.trim().toLowerCase();
              
              // Check if one contains the other (for partial matches)
              if (paraText === blockText || 
                  (blockText.length > 10 && paraText.includes(blockText)) ||
                  (blockText.length > 10 && blockText.includes(paraText))) {
                // For headings, check style
                if (block.type === "heading") {
                  const expectedStyle = `Heading ${block.level}`;
                  if (para.style === expectedStyle) {
                    matchedIndex = i;
                    break;
                  }
                } else {
                  // Regular paragraph - fuzzy match
                  matchedIndex = i;
                  break;
                }
              }
            }
          }
          
          if (matchedIndex !== -1) {
            blockIdToParagraphIndex[blockId] = matchedIndex;
            usedParagraphIndices.add(matchedIndex);
          } else {
            console.warn(`Could not find paragraph matching block ${blockId} with text: "${block.text.substring(0, 50)}..."`);
          }
        }
      }

      // Log the mapping for debugging
      console.log("Block ID to paragraph index mapping:", blockIdToParagraphIndex);
      console.log("Available block IDs in document structure:", Object.keys(documentStructure.blocks));

      // Execute operations in order
      for (const op of editPlan.ops) {
        await executeSemanticOperation(op, blockIdToParagraphIndex, paragraphs, documentStructure);
      }
    });
  } catch (error) {
    if (error instanceof Error) {
      throw new ExecutionError(`Failed to execute semantic edit plan: ${error.message}`, error);
    }
    throw new ExecutionError(`Failed to execute semantic edit plan: ${String(error)}`, error);
  }
}

/**
 * Executes a single semantic operation
 */
async function executeSemanticOperation(
  op: SemanticOperation,
  blockIdToParagraphIndex: Record<string, number>,
  paragraphs: Word.ParagraphCollection,
  documentStructure: { sections: Array<{ id: string; title: string; level: number; blocks: string[] }>; blocks: Record<string, { type: "paragraph" | "heading"; text: string; level?: number; id?: string }> }
): Promise<void> {
  if (typeof Word === "undefined") {
    throw new ExecutionError("Word API is not available.");
  }

  await Word.run(async (context) => {
    // Find the target block's paragraph
    const targetBlockId = op.target_block_id;
    
    // Check if block exists in document structure
    if (!documentStructure.blocks[targetBlockId]) {
      const availableBlockIds = Object.keys(documentStructure.blocks).join(", ");
      throw new ExecutionError(`Block ID ${targetBlockId} not found in document structure. Available block IDs: ${availableBlockIds || "none"}`);
    }
    
    if (blockIdToParagraphIndex[targetBlockId] === undefined) {
      const block = documentStructure.blocks[targetBlockId];
      const blockText = block ? block.text.substring(0, 50) : "unknown";
      throw new ExecutionError(`Block ID ${targetBlockId} not found in document. Block text: "${blockText}...". This may happen if the document was modified after the edit plan was generated.`);
    }

    const targetParaIndex = blockIdToParagraphIndex[targetBlockId];
    
    // Verify the paragraph index is valid
    if (targetParaIndex < 0 || targetParaIndex >= paragraphs.items.length) {
      throw new ExecutionError(`Invalid paragraph index ${targetParaIndex} for block ID ${targetBlockId}`);
    }
    
    const targetPara = paragraphs.items[targetParaIndex];
    const targetRange = targetPara.getRange();
    targetRange.load("text");
    await context.sync();

    if (op.action === "replace") {
      // Replace the target block's content
      targetRange.clear();
      await context.sync();
      targetRange.insertText(op.content, Word.InsertLocation.replace);
      await context.sync();
    } else if (op.action === "insert_after") {
      // Insert after the target block
      targetRange.insertText(op.content + "\n", Word.InsertLocation.after);
      await context.sync();
    } else if (op.action === "insert_before") {
      // Insert before the target block
      targetRange.insertText(op.content + "\n", Word.InsertLocation.before);
      await context.sync();
    }
  });
}

