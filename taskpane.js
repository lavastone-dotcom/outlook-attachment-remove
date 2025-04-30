Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("remove-attachments").onclick = async () => {
            try {
                const item = Office.context.mailbox.item;

                if (!item.internetMessageId) {
                    console.error("This message doesn't support 'internetMessageId'.");
                    return;
                }

                const internetMessageId = item.internetMessageId;

                // Call Graph API function to delete attachments
                await deleteAttachmentsByGraph(internetMessageId);
                console.log("Attachment removal triggered.");
            } catch (error) {
                console.error("Error removing attachments:", error);
            }
        };
    }
});
