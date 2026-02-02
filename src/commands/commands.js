/* global Office */

Office.onReady(() => {
  console.log("Office.js is ready!");
});

async function showAlert(event) {
  try {
    console.log("Button clicked! Calling GET request to jsonplaceholder...");

    const apiEndpoint = "https://jsonplaceholder.typicode.com/posts";
    const response = await fetch(apiEndpoint);
    console.log("GET Response status:", response.status);

    if (!response.ok) {
      throw new Error(`API returned status: ${response.status}`);
    }

    const responseData = await response.json();
    console.log("GET Response payload:", responseData);

    const successMessage = {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: `Fetched ${responseData.length} posts from jsonplaceholder.`,
      icon: "Icon.80x80",
      persistent: true
    };
    Office.context.mailbox.item.notificationMessages.replaceAsync("apiSuccess", successMessage);

  } catch (error) {
    console.error("Error calling GET API:", error);
    const errorMessage = {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message: `GET request failed: ${error.message}`,
      icon: "Icon.16x16",
      persistent: true
    };
    Office.context.mailbox.item.notificationMessages.replaceAsync("apiError", errorMessage);
  } finally {
    event.completed();
  }
}

Office.actions.associate("showAlert", showAlert);
