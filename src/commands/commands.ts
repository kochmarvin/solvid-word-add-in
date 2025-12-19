/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office Word */

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

/**
 * Handles the context menu action for formatting selected text
 * Gets the selected text and opens the task pane
 * @param event
 */
/* global Office */

function formatSelectedText(event: Office.AddinCommands.Event) {
  // IMPORTANT: event.completed() must be called synchronously
  event.completed();
  
  // Check if Word API is available
  if (typeof Word === "undefined") {
    Office.addin.showAsTaskpane();
    return;
  }

  // Capture selection as Content Control
  Word.run(async (context) => {
    try {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      const selectedText = selection.text.trim();
      if (!selectedText) {
        Office.addin.showAsTaskpane();
        return;
      }

      // IMPORTANT: Store the selection text BEFORE any operations
      const selectionRange = selection.getRange();
      selectionRange.load("text");
      await context.sync();
      const selectionTextToPreserve = selectionRange.text;

      // First, check for existing solvid-selected Content Controls
      const contentControls = context.document.contentControls;
      contentControls.load("items");
      await context.sync();
      
      // Check if selection is inside an existing solvid-selected Content Control
      let existingCC: Word.ContentControl | null = null;
      const oldControls: Word.ContentControl[] = [];
      
      for (let i = 0; i < contentControls.items.length; i++) {
        const cc = contentControls.items[i];
        cc.load("tag");
        await context.sync();
        
        if (cc.tag && cc.tag.startsWith("solvid-selected-")) {
          const ccRange = cc.getRange();
          ccRange.load("text");
          await context.sync();
          
          // Check if current selection text matches or overlaps with this Content Control
          if (ccRange.text === selectedText || 
              ccRange.text.includes(selectedText) || 
              selectedText.includes(ccRange.text)) {
            // Selection is the same or overlaps - reuse this control
            existingCC = cc;
          } else {
            // Different selection - mark for deletion
            oldControls.push(cc);
          }
        }
      }
      
      // Hide old Content Controls (remove borders) instead of deleting
      // This preserves the text while removing the visual selection
      for (const oldCC of oldControls) {
        // Hide the Content Control border and mark as inactive
        oldCC.appearance = "Hidden";
        oldCC.tag = `solvid-inactive-${Date.now()}`;
        oldCC.load("tag,appearance");
      }
      await context.sync();

      // Create or update Content Control
      let cc: Word.ContentControl;
      const tag = `solvid-selected-${Date.now()}`;
      
      if (existingCC) {
        // Update existing control with new tag
        cc = existingCC;
        cc.tag = tag;
        cc.load("tag");
        await context.sync();
      } else {
        // Create a new Content Control to mark the selection
        cc = selection.insertContentControl();
        cc.tag = tag;
        cc.title = "Solvid Selection";
        cc.appearance = "BoundingBox";
        cc.load("id,tag");
        await context.sync();
      }

      // Store the tag and text in document settings
      Office.context.document.settings.set("selectedText", selectedText);
      Office.context.document.settings.set("selectedTag", tag);
      await Office.context.document.settings.saveAsync();

      // Also ping taskpane via localStorage
      localStorage.setItem("solvid:refreshSelection", String(Date.now()));

      // Ensure taskpane is visible
      Office.addin.showAsTaskpane();
    } catch (error) {
      // If error, just open taskpane
      Office.addin.showAsTaskpane();
    }
  }).catch(() => {
    Office.addin.showAsTaskpane();
  });
}

Office.onReady(() => {
  Office.actions.associate("formatSelectedText", formatSelectedText);
});



// Make function globally accessible (required for Office Add-in commands)
// @ts-ignore
window.formatSelectedText = formatSelectedText;

// Register functions when Office.js is ready
Office.onReady(() => {
  try {
    // Register the functions with Office.
    // These functions will be called when the corresponding commands are executed
    Office.actions.associate("action", action);
    Office.actions.associate("formatSelectedText", formatSelectedText);
  } catch (error) {
    // Log registration errors if possible
    if (typeof console !== "undefined") {
      console.error("Error registering actions:", error);
    }
  }
});
