async function clearAttachments() {
    const item = Office.context.mailbox.item;

    if (!item || typeof item.removeAttachmentAsync !== 'function') {
        console.error("removeAttachmentAsync is not available in this context.");
        alert("⚠️ Cannot remove attachments. Open a single email in full view.");
        return;
    }

    const attachmentIds = item.attachments.map(att => att.id);

    for (const id of attachmentIds) {
        await new Promise((resolve, reject) => {
            item.removeAttachmentAsync(id, (asyncResult) => {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    console.log(`Attachment ${id} removed successfully.`);
                    resolve();
                } else {
                    console.error(`Failed to remove attachment ${id}`, asyncResult.error);
                    reject(asyncResult.error);
                }
            });
        });
    }

    alert("✅ Attachments removed!");
}
