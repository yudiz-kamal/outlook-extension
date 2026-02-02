/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
  console.log("Office.js is ready!");
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
    console.log("Button clicked! Calling API...");

    // Get the current email item
    const item = Office.context.mailbox.item;

    // Extract email details
    const emailData = {
      subject: item.subject || "No Subject",
      from: item.from?.emailAddress || item.from?.displayName || "Unknown",
      to: item.to ? item.to.map(recipient => recipient.emailAddress || recipient.displayName) : [],
      cc: item.cc ? item.cc.map(recipient => recipient.emailAddress || recipient.displayName) : [],
      dateTimeCreated: item.dateTimeCreated ? item.dateTimeCreated.toISOString() : null,
      itemId: item.itemId || null
    };

    console.log("Email data extracted:", emailData);

    // Get email body (async operation)
    item.body.getAsync(Office.CoercionType.Text, async (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        emailData.body = result.value.substring(0, 500); // Limit body length for testing
        console.log("Email body extracted");
      } else {
        console.warn("Could not get email body");
        emailData.body = "Body not available";
      }

      try {
        // Using JSONPlaceholder fake API for testing
        const apiEndpoint = "https://jsonplaceholder.typicode.com/posts";

        console.log("Calling API:", apiEndpoint);

        // Make POST API call
        const response = await fetch(apiEndpoint, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            title: emailData.subject,
            body: JSON.stringify(emailData),
            userId: 1
          })
        });

        console.log("API Response status:", response.status);

        if (response.ok) {
          const responseData = await response.json();
          console.log("API Response data:", responseData);

          // Show success notification
          const successMessage = {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: `✓ API called successfully! Response ID: ${responseData.id}`,
            icon: "Icon.80x80",
            persistent: true
          };
          Office.context.mailbox.item.notificationMessages.replaceAsync("apiSuccess", successMessage);

          // Also show browser alert for testing
          alert(`API Success!\n\nEmail Subject: ${emailData.subject}\nAPI Response ID: ${responseData.id}`);
        } else {
          throw new Error(`API call failed with status: ${response.status}`);
        }

      } catch (error) {
        console.error("Error calling API:", error);

        // Show error notification
        const errorMessage = {
          type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
          message: `✗ API call failed: ${error.message}`,
          icon: "Icon.80x80",
          persistent: true
        };
        Office.context.mailbox.item.notificationMessages.replaceAsync("apiError", errorMessage);

        // Also show browser alert for testing
        alert(`API Error!\n\n${error.message}`);
      }

      // Signal that the command is complete
      event.completed();
    });

  } catch (error) {
    console.error("Error in showAlert function:", error);

    // Show error notification
    const errorMessage = {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message: `Error: ${error.message}`,
      icon: "Icon.80x80",
      persistent: true
    };
    Office.context.mailbox.item.notificationMessages.replaceAsync("generalError", errorMessage);

    // Signal that the command is complete
    event.completed();
  }
}

// Register the functions with Office.
Office.actions.associate("action", action);
Office.actions.associate("showAlert", showAlert);