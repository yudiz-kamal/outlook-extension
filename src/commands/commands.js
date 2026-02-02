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

async function showAlert(event) {
  try {
    // Get the current email item
    const item = Office.context.mailbox.item;

    // Extract email details
    const emailData = {
      subject: item.subject,
      from: item.from?.emailAddress || item.from?.displayName || "Unknown",
      to: item.to?.map(recipient => recipient.emailAddress || recipient.displayName) || [],
      cc: item.cc?.map(recipient => recipient.emailAddress || recipient.displayName) || [],
      dateTimeCreated: item.dateTimeCreated?.toISOString() || null,
      itemId: item.itemId || null
    };

    // Get email body
    item.body.getAsync(Office.CoercionType.Text, async (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        emailData.body = result.value;
      }

      // Replace with your actual API endpoint
      const apiEndpoint = "https://your-api-endpoint.com/api/email";

      // Make POST API call
      const response = await fetch(apiEndpoint, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          // Add any additional headers here (e.g., Authorization)
          // "Authorization": "Bearer YOUR_TOKEN"
        },
        body: JSON.stringify(emailData)
      });

      if (response.ok) {
        // Show success notification
        const successMessage = {
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message: "Email data sent successfully!",
          icon: "Icon.80x80",
          persistent: false
        };
        Office.context.mailbox.item.notificationMessages.replaceAsync("apiSuccess", successMessage);
      } else {
        throw new Error(`API call failed with status: ${response.status}`);
      }

      // Signal that the command is complete
      event.completed();
    });

  } catch (error) {
    console.error("Error calling API:", error);

    // Show error notification
    const errorMessage = {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message: "Failed to send email data. Please try again.",
      icon: "Icon.80x80",
      persistent: true
    };
    Office.context.mailbox.item.notificationMessages.replaceAsync("apiError", errorMessage);

    // Signal that the command is complete
    event.completed();
  }
}

// Register the function with Office.
Office.actions.associate("action", action);
Office.actions.associate("showAlert", showAlert);