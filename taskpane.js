document.addEventListener("DOMContentLoaded", function () {
    const btn = document.getElementById("removeBtn");
    if (btn) {
        btn.addEventListener("click", onRemoveClick);
    }
});

async function onRemoveClick() {
    try {
        const item = Office.context.mailbox.item;

        // Get the internetMessageId of the selected email
        item.internetMessageId.getAsync(async (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const internetMessageId = result.value;
                console.log("Message ID:", internetMessageId);

                // Call the function from graph.js to delete attachments
                await deleteAttachmentsByGraph(internetMessageId);

                alert("Attachments removed successfully.");
            } else {
                console.error("Error retrieving internetMessageId:", result.error);
                alert("Failed to retrieve message ID.");
            }
        });

    } catch (err) {
        console.error("Unexpected error:", err);
        alert("An unexpected error occurred. Check console.");
    }
}
