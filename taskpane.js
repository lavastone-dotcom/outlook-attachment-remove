// taskpane.js

Office.onReady(() => {
  // Ensure the DOM is fully loaded
  document.addEventListener("DOMContentLoaded", () => {
    const removeButton = document.getElementById("remove-attachments");
    const subjectLabel = document.getElementById("item-subject");

    if (!removeButton) {
      console.error("Element with ID 'remove-attachments' not found in the DOM.");
      return;
    }

    if (!subjectLabel) {
      console.error("Element with ID 'item-subject' not found in the DOM.");
      return;
    }

    // Set up click handler for the remove attachments button
    removeButton.onclick = async () => {
      const item = Office.context.mailbox.item;

      if (item && item.internetMessageId) {
        try {
          await deleteAttachmentsByGraph(item.internetMessageId);
          console.log("Attachment removal triggered.");
        } catch (error) {
          console.error("Error removing attachments:", error);
        }
      } else {
        console.error("No internetMessageId found. Cannot proceed with attachment removal.");
      }
    };

    // Display the subject of the current item
    const item = Office.context.mailbox.item;
    if (item && item.subject) {
      subjectLabel.innerHTML = `<b>Subject:</b> ${item.subject}`;
    } else {
      subjectLabel.innerHTML = "<b>Subject:</b> (No subject available)";
    }
  });
});

// Placeholder function for deleting attachments using Microsoft Graph API
async function deleteAttachmentsByGraph(internetMessageId) {
  // Implement your logic to authenticate and call the Microsoft Graph API
  // to delete attachments based on the internetMessageId
  // This is a placeholder and should be replaced with actual implementation
  console.log(`deleteAttachmentsByGraph called with ID: ${internetMessageId}`);
  // Example:
  // await fetch(`https://graph.microsoft.com/v1.0/me/messages/${internetMessageId}/attachments`, {
  //   method: 'DELETE',
  //   headers: {
  //     'Authorization': `Bearer ${accessToken}`
  //   }
  // });
}
