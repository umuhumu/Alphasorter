// Utility function to validate email addresses
function isValidEmail(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

// Utility function to retrieve recipients with display name (if available)
function fetchRecipientsWithDisplayName(field) {
  return new Promise((resolve) => field.getAsync((result) => resolve(result.value || [])));
}

// Main action function
function action(event) {
  const item = Office.context.mailbox.item;

  // Ensure that the item is a message
  if (item.itemType !== Office.MailboxEnums.ItemType.Message) {
    notifyUser("This action can only be performed on email messages.", "ErrorMessage", true);
    event.completed();
    return;
  }

  // Notify user that the process has started
  //notifyUser("", "InformationalMessage", false);

  // Fetch recipients asynchronously (To, Cc, and Bcc fields)
  const toPromise = fetchRecipientsWithDisplayName(item.to);
  const ccPromise = fetchRecipientsWithDisplayName(item.cc);
  const bccPromise = fetchRecipientsWithDisplayName(item.bcc);

  Promise.all([toPromise, ccPromise, bccPromise])
    .then(([toRecipients, ccRecipients, bccRecipients]) => {
      // Process each recipient field using the display name if available
      const processedTo = processFieldRecipients(toRecipients);
      const processedCC = processFieldRecipients(ccRecipients);
      const processedBCC = processFieldRecipients(bccRecipients);

      // Update the recipients after processing (sorting, removing duplicates)
      return Promise.all([
        updateRecipients(item.to, processedTo.processed),
        updateRecipients(item.cc, processedCC.processed),
        updateRecipients(item.bcc, processedBCC.processed),
      ]).then((results) => {
        const allSuccess = results.every((success) => success);
        const totalRecipients = processedTo.processed.length + processedCC.processed.length + processedBCC.processed.length;
        const totalDuplicatesRemoved = processedTo.deduplicationCount + processedCC.deduplicationCount + processedBCC.deduplicationCount;

        // Prepare a summary message
        let summaryMessage = `Alphasorter completed: ${totalRecipients} recipients sorted, ${totalDuplicatesRemoved} duplicates removed.`;
        
        // Notify user of the result
        if (allSuccess) {
          notifyUser(summaryMessage, "InformationalMessage", false);
        } else {
          notifyUser("An error occurred while updating recipients.", "ErrorMessage", true);
        }
      });
    })
    .catch((error) => {
      console.error("Error processing recipients:", error);
      notifyUser("An error occurred while processing recipients.", "ErrorMessage", true);
    })
    .finally(() => {
      event.completed();
    });
}

// Process recipient field (sort and deduplicate)
function processFieldRecipients(recipients) {
  // Create an array of recipient objects with both displayName and emailAddress
  const recipientDetails = recipients.map((r) => ({
    emailAddress: r.emailAddress,
    displayName: r.displayName || r.emailAddress, // Fallback to email if no display name
  }));

  // Remove duplicates based on email address
  const uniqueRecipients = Array.from(new Map(recipientDetails.map(r => [r.emailAddress, r])).values());

  return {
    processed: uniqueRecipients.sort((a, b) => a.displayName.localeCompare(b.displayName)), // Sort by displayName
    deduplicationCount: recipients.length - uniqueRecipients.length, // Count duplicates
  };
}

// Update recipients in the message
function updateRecipients(field, recipients) {
  return new Promise((resolve) =>
    field.setAsync(recipients.map((r) => ({ emailAddress: r.emailAddress, displayName: r.displayName })), (result) =>
      resolve(result.status === Office.AsyncResultStatus.Succeeded)
    )
  );
}

// Notify the user with a message
function notifyUser(message, type, persistent) {
  // Assign icon based on the message type
  let icon;
  
  switch (type) {
    case "ErrorMessage":
      icon = "error"; // Built-in error icon
      break;
    case "WarningMessage":
      icon = "warning"; // Built-in warning icon
      break;
    case "InformationalMessage":
      icon = "info"; // Built-in informational icon
      break;
    default:
      icon = "info"; // Default to informational icon
      break;
  }

  // Set notification
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", {
    type: Office.MailboxEnums.ItemNotificationMessageType[type] || Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: message,
    icon: icon, // Use built-in Office icons
    persistent: persistent,
  });
}

// Register the action
Office.actions.associate("action", action);
