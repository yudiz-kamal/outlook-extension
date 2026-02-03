/* global Office */

Office.onReady(() => {
  console.log("Office.js is ready!");
});

function simplifyRecipient(recipient) {
  if (!recipient) {
    return null;
  }
  return {
    displayName: recipient.displayName,
    emailAddress: recipient.emailAddress,
    recipientType: recipient.recipientType,
  };
}

function simplifyRecipientList(list) {
  if (!Array.isArray(list)) {
    return [];
  }
  return list.map(simplifyRecipient);
}

function getBodyText() {
  const item = Office.context.mailbox.item;
  return new Promise((resolve, reject) => {
    if (!item?.body?.getAsync) {
      resolve("");
      return;
    }

    item.body.getAsync(Office.CoercionType.Text, (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        resolve(asyncResult.value);
      } else {
        reject(asyncResult.error);
      }
    });
  });
}

async function getItemSnapshot() {
  const item = Office.context.mailbox.item;
  if (!item) {
    return null;
  }

  const snapshot = {
    subject: item.subject,
    itemClass: item.itemClass,
    itemId: item.itemId,
    dateTimeCreated: item.dateTimeCreated,
    dateTimeModified: item.dateTimeModified,
    from: simplifyRecipient(item.from),
    to: simplifyRecipientList(item.to),
    cc: simplifyRecipientList(item.cc),
    bcc: simplifyRecipientList(item.bcc),
  };
  
  console.table('snapshot', snapshot)
  try {
    snapshot.body = await getBodyText();
  } catch (error) {
    console.warn("Unable to read item body as text:", error);
    snapshot.body = "";
  }

  return snapshot;
}

async function showAlert(event) {
  try {
    const emailData = await getItemSnapshot();
    console.table('emailData', emailData)
    console.log("Button clicked! Calling POST to fakestoreapi...");
    

    const apiEndpoint = "https://fakestoreapi.com/products";
    const response = await fetch(apiEndpoint, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        title: "New Product",
        price: 29.99,
        emailData,
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
