Office.onReady(() => {
  console.log("Office.js is ready");
});

async function clearAttachments() {
  try {
    const itemId = Office.context.mailbox.item.itemId;
    const accessToken = await Office.auth.getAccessToken({ allowSignInPrompt: true });

    console.log("Access token acquired.");

    // Correct ID format handling
    let restId = Office.context.mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0);

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
      return;
    }

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
  } catch (error) {
    console.error("Error removing attachments: ", error);
  }
}
