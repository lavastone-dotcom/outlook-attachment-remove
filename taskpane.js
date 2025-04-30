Office.onReady(() => {
    if (Office.context.mailbox.item) {
        document.getElementById("remove-attachments").onclick = async () => {
            try {
                const item = Office.context.mailbox.item;

                // Office.context.mailbox.item.internetMessageId is not supported in some clients.
                // We log it if it's undefined for debugging.
                if (!item.internetMessageId) {
                    console.error("No internetMessageId found. Ensure this runs in supported Outlook clients.");
                    alert("InternetMessageId not available. Try running in Outlook Web App.");
                    return;
                }

                console.log("InternetMessageId:", item.internetMessageId);

                await deleteAttachmentsByGraph(item.internetMessageId);
                alert("Attachments removed successfully.");
            } catch (error) {
                console.error("Error removing attachments:", error);
                alert("Error occurred while removing attachments. See console for details.");
            }
        };
    }
});
