// graph.js
async function getGraphToken() {
    return await Office.auth.getAccessToken({ allowSignInPrompt: true });
}

async function deleteAttachmentsByGraph(internetMessageId) {
    const token = await getGraphToken();
    const encodedMessageId = encodeURIComponent(internetMessageId);

    // 1. Find the message by InternetMessageId
    const response = await fetch(`https://graph.microsoft.com/v1.0/me/messages?$filter=internetMessageId eq '${internetMessageId}'`, {
        headers: {
            Authorization: `Bearer ${token}`
        }
    });
    const data = await response.json();

    if (data.value.length === 0) {
        console.error("Message not found in Graph.");
        return;
    }

    const messageId = data.value[0].id;

    // 2. Get attachments
    const attachmentsResponse = await fetch(`https://graph.microsoft.com/v1.0/me/messages/${messageId}/attachments`, {
        headers: {
            Authorization: `Bearer ${token}`
        }
    });
    const attachmentsData = await attachmentsResponse.json();

    // 3. Delete each attachment
    for (const attachment of attachmentsData.value) {
        await fetch(`https://graph.microsoft.com/v1.0/me/messages/${messageId}/attachments/${attachment.id}`, {
            method: "DELETE",
            headers: {
                Authorization: `Bearer ${token}`
            }
        });
    }

    console.log("Attachments deleted successfully.");
}
