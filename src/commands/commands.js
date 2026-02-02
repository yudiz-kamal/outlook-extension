/* global Office */

Office.onReady(() => {
  console.log("Office.js is ready!");
});

async function showAlert(event) {
  try {
    console.log("Button clicked! Calling API...");

    const item = Office.context.mailbox.item;

    const emailData = {
      subject: item.subject || "No Subject",
      from: item.from?.emailAddress || item.from?.displayName || "Unknown",
      dateTimeCreated: item.dateTimeCreated ? item.dateTimeCreated.toISOString() : null
    };

    console.log("Email data:", emailData);

    // Call fake API
    const apiEndpoint = "https://jsonplaceholder.typicode.com/posts";

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
      console.log("API Response:", responseData);

      alert(`✓ API Success!\n\nEmail: ${emailData.subject}\nResponse ID: ${responseData.id}`);

      const successMessage = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: `API called successfully! ID: ${responseData.id}`,
        icon: "Icon.80x80",
        persistent: true
      };
      Office.context.mailbox.item.notificationMessages.replaceAsync("apiSuccess", successMessage);
    } else {
      throw new Error(`API returned status: ${response.status}`);
    }

    event.completed();

  } catch (error) {
    console.error("Error:", error);
    alert(`✗ Error!\n\n${error.message}`);
    event.completed();
  }
}

Office.actions.associate("showAlert", showAlert);