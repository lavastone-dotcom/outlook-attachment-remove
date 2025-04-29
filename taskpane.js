Office.onReady(() => {
  console.log("Office.js is ready");
});

async function clearAttachments() {
  const item = Office.context.mailbox.item;

  if (item.attachments.length === 0) {
    console.log("No attachments found.");
    return;
  }

  const attachmentIds = item.attachments.map(att => att.id);

  for (const id of attachmentIds) {
    await new Promise((resolve, reject) => {
      Office.context.mailbox.item.removeAttachmentAsync(id, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log(`Attachment ${id} removed.`);
          resolve();
        } else {
          console.error(`Failed to remove attachment ${id}`, result.error);
          reject(result.error);
        }
      });
    });
  }

  console.log("All attachments removed.");
}
