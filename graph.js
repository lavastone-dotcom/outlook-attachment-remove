function deleteAttachment(attachmentId) {
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, function (result) {
        if (result.status === "succeeded") {
            const token = result.value;
            const itemId = Office.context.mailbox.item.itemId;
            const encodedItemId = Office.context.mailbox.convertToRestId(itemId, Office.MailboxEnums.RestVersion.v2_0);

            fetch(`https://graph.microsoft.com/v1.0/me/messages/${encodedItemId}/attachments/${attachmentId}`, {
                method: "DELETE",
                headers: {
                    "Authorization": `Bearer ${token}`
                }
            }).then(response => {
                if (response.ok) {
                    alert("Attachment deleted successfully.");
                    location.reload();
                } else {
                    response.json().then(data => {
                        alert("Failed to delete: " + (data.error?.message || response.statusText));
                    });
                }
            }).catch(error => {
                alert("Error: " + error.message);
            });

        } else {
            alert("Could not get token: " + result.error.message);
        }
    });
}
