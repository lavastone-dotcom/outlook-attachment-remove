// taskpane.js
Office.onReady(() => {
    document.getElementById("remove-attachments").onclick = async () => {
        const item = Office.context.mailbox.item;
        if (item && item.internetMessageId) {
            await deleteAttachmentsByGraph(item.internetMessageId);
            console.log("Attachment removal triggered.");
        } else {
            console.error("No internetMessageId found.");
        }
    };
});
