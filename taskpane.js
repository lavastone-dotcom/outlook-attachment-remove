Office.onReady(() => {
  console.log("Office.js is ready");
});

async function clearAttachments(event) {
  try {
    const itemId = Office.context.mailbox.item.itemId;
    const accessToken = await Office.auth.getAccessToken({ allowSignInPrompt: true });

    console.log("Access token acquired.");

    // Decode the itemId to REST format if needed
    let restId = Office.context.mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0);

    // Get attachments
    let attachmentsResponse = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${restId}/attachments`, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: "application/json"
      }
    });

    if (!attachmentsResponse.ok) {
      throw new Error("Failed to fetch attachments.");
    }

    let attachmentsData = await attachmentsResponse.json();

    const attachments = attachmentsData.value;

    if (attachments.length === 0) {
      console.log("No attachments found.");
      event.completed();
      return;
    }

    // Delete each attachment
    for (const attachment of attachments) {
      let deleteResponse = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${restId}/attachments/${attachment.id}`, {
        method: "DELETE",
        headers: {
          Authorization: `Bearer ${accessToken}`
        }
      });

      if (!deleteResponse.ok) {
        console.error(`Failed to delete attachment ${attachment.id}`);
      } else {
        console.log(`Attachment ${attachment.id} deleted.`);
      }
    }

    console.log("All attachments removed.");
    event.completed();
  } catch (error) {
    console.error("Error removing attachments: ", error);
    event.completed();
  }
}
