Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        removeAttachments();
    }
});

async function removeAttachments() {
    try {
        const item = Office.context.mailbox.item;
        const internetMessageId = item.internetMessageId;

        if (!internetMessageId) {
            console.error("No InternetMessageId found.");
            return;
        }

        await deleteAttachmentsByGraph(internetMessageId);
    } catch (error) {
        console.error("Error removing attachments:", error);
    }
}
