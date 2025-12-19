/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office Word */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item?.notificationMessages.replaceAsync(
    "ActionPerformanceNotification",
    message
  );

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

/**
 * Handles the context menu action for formatting selected text
 * Gets the selected text and opens the task pane
 * @param event
 */
function formatSelectedText(event: Office.AddinCommands.Event) {
  // Check if Word API is available
  if (typeof Word === "undefined") {
    // Just open the task pane if Word API is not available
    Office.addin.showAsTaskpane();
    event.completed();
    return;
  }

  Word.run(async (context) => {
    try {
      // Get the selected text
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      const selectedText = selection.text.trim();

      if (selectedText) {
        // Store the selected text in Office.context.document.settings
        // This will be available when the task pane opens
        Office.context.document.settings.set("selectedText", selectedText);
        await Office.context.document.settings.saveAsync();
      }

      // Open the task pane
      Office.addin.showAsTaskpane();
    } catch (error) {
      // If there's an error, just open the task pane
      console.error("Error getting selected text:", error);
      Office.addin.showAsTaskpane();
    } finally {
      // Always complete the event
      event.completed();
    }
  });
}

// Register the functions with Office.
Office.actions.associate("action", action);
Office.actions.associate("formatSelectedText", formatSelectedText);
