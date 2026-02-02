/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item.notificationMessages.replaceAsync(
    "ActionPerformanceNotification",
    message
  );

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

function showAlert(event) {
  // Show alert box
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Hello! This is your custom button alert!",
    icon: "Icon.80x80",
    persistent: false
  };

  // Show notification
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);
  
  // Or use browser alert (simpler option)
  alert("Button clicked! This is your custom alert.");
  
  // Signal that the command is complete
  event.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);
Office.actions.associate("showAlert", showAlert);