/* global Office */

Office.onReady(() => {
  console.log("Office.js is ready!");
});

async function showAlert(event) {
  try {
    console.log("Button clicked! Calling POST to fakestoreapi...");
    console.log("emailData:");
    console.dir(Office.context.mailbox.item, { depth: Infinity });
    
    const apiEndpoint = "https://fakestoreapi.com/products";
    const response = await fetch(apiEndpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        title: "New Product",
        price: 29.99
      }),
    });
    console.log("POST Response status:", response.status);

    if (!response.ok) {
      throw new Error(`API returned status: ${response.status}`);
    }

    const responseData = await response.json();
    console.log("POST Response payload:", responseData);

    const successMessage = {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: `Created product ID ${responseData.id} at price $${responseData.price}.`,
      icon: "Icon.80x80",
      persistent: true,
    };
    Office.context.mailbox.item.notificationMessages.replaceAsync("apiSuccess", successMessage);

  } catch (error) {
    console.error("Error calling POST API:", error);
    const errorMessage = {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message: `POST request failed: ${error.message}`,
      icon: "Icon.16x16",
      persistent: true
    };
    Office.context.mailbox.item.notificationMessages.replaceAsync("apiError", errorMessage);
  } finally {
    event.completed();
  }
}

Office.actions.associate("showAlert", showAlert);
