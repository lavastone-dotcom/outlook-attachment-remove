async function deleteAttachmentsByGraph(internetMessageId) {
    const token = await getGraphToken(); // Assume this retrieves a valid Graph token

    const encodedFilter = encodeURIComponent(`internetMessageId eq '${internetMessageId}'`);
    const graphEndpoint = `https://graph.microsoft.com/v1.0/me/messages?$filter=${encodedFilter}`;

    const response = await fetch(graphEndpoint, {
        headers: {
            Authorization: `Bearer ${token}`
        }
    });

    const data = await response.json();
    if (!data.value || data.value.length === 0) {
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
        const delResponse = await fetch(deleteUrl, {
            method: "DELETE",
            headers: {
                Authorization: `Bearer ${token}`
            }
        });

        if (delResponse.status === 204) {
            console.log(`Deleted attachment: ${attachment.name}`);
        } else {
            console.warn(`Failed to delete attachment: ${attachment.name}`);
        }
    }
}
