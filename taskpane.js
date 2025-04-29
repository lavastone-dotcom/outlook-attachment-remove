// Wait for Office.js to be ready
Office.onReady(function(info) {
    if (info.host === Office.HostType.Outlook) {
        console.log("Office.js ready in Outlook!");

        // Bind the button click to the function
        document.getElementById("removeButton").onclick = function() {
            removeAttachment();
        };
    }
});

// Function to remove an attachment
function removeAttachment() {
    var item = Office.context.mailbox.item;

    // Make sure there are attachments
    if (!item.attachments || item.attachments.length === 0) {
        console.error("No attachments found on this email.");
        return;
    }

    // Example: Remove the first attachment
    var attachmentId = item.attachments[0].id;

    item.removeAttachmentAsync(attachmentId, function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log("Attachment removed successfully!");
        } else {
            console.error("Failed to remove attachment: " + asyncResult.error.message);
        }
    });
}
