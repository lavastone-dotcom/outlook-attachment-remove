async function deleteAttachmentsByGraph(internetMessageId) {
    const token = await getGraphToken();
    const encodedMessageId = encodeURIComponent(internetMessageId);
    const graphEndpoint = `https://graph.microsoft.com/v1.0/me/messages?$filter=internetMessageId eq '${internetMessageId}'`;

    const response = await fetch(graphEndpoint, {
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

    const attachmentsEndpoint = `https://graph.microsoft.com/v1.0/me/messages/${messageId}/attachments`;
    const attachmentsResponse = await fetch(attachmentsEndpoint, {
        headers: {
            Authorization: `Bearer ${token}`
        }
    });

    const attachmentsData = await attachmentsResponse.json();
    for (const attachment of attachmentsData.value) {
        const deleteUrl = `https://graph.microsoft.com/v1.0/me/messages/${messageId}/attachments/${attachment.id}`;
        await fetch(deleteUrl, {
            method: "DELETE",
            headers: {
                Authorization: `Bearer ${token}`
            }
        });
    }
}
