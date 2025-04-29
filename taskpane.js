
Office.onReady(() => {
  console.log("Office.js is ready.");
});

function removeAttachments() {
  Office.context.mailbox.item.attachments.forEach(att => {
    Office.context.mailbox.item.removeAttachmentAsync(att.id, result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Removed attachment: " + att.name);
      } else {
        console.error("Failed to remove attachment: " + result.error.message);
      }
    });
  });
}
