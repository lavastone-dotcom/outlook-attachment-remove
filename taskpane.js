// taskpane.js

Office.onReady(() => {
    // Ensure the DOM is fully loaded
    document.addEventListener("DOMContentLoaded", () => {
        const removeBtn = document.getElementById("remove-attachments");
        if (removeBtn) {
            removeBtn.onclick = async () => {
                const item = Office.context.mailbox.item;
                if (item && item.internetMessageId) {
                    try {
                        await deleteAttachmentsByGraph(item.internetMessageId);
                        console.log("Attachment removal triggered.");
                    } catch (err) {
                        console.error("Error removing attachments:", err);
                    }
                } else {
                    console.error("No internetMessageId found.");
                }
            };
        }
    });
});

/**
 * Deletes all attachments from a message using Microsoft Graph API.
 * @param {string} internetMessageId - The Internet Message ID of the email.
 */
async function deleteAttachmentsByGraph(internetMessageId) {
    const accessToken = await getAccessToken();
    if (!accessToken) {
        throw new Error("Access token is missing.");
    }

    const encodedId = encodeURIComponent(internetMessageId);
    const graphEndpoint = `https://graph.microsoft.com/v1.0/me/messages?$filter=internetMessageId eq '${encodedId}'`;

    // Fetch the message ID from Graph
    const messageResponse = await fetch(graphEndpoint, {
        headers: {
            Authorization: `Bearer ${accessToken}`
        }
    });

    if (!messageResponse.ok) {
        throw new Error(`Failed to retrieve message: ${messageResponse.statusText}`);
    }

    const messageData = await messageResponse.json();
    const message = messageData.value[0];
    if (!message || !message.id) {
        throw new Error("Message not found.");
    }

    // Retrieve attachments
    const attachmentsEndpoint = `https://graph.microsoft.com/v1.0/me/messages/${message.id}/attachments`;
    const attachmentsResponse = await fetch(attachmentsEndpoint, {
        headers: {
            Authorization: `Bearer ${accessToken}`
        }
    });

    if (!attachmentsResponse.ok) {
        throw new Error(`Failed to retrieve attachments: ${attachmentsResponse.statusText}`);
    }

    const attachmentsData = await attachmentsResponse.json();
    const attachments = attachmentsData.value;

    // Delete each attachment
    for (const attachment of attachments) {
        const deleteEndpoint = `https://graph.microsoft.com/v1.0/me/messages/${message.id}/attachments/${attachment.id}`;
        const deleteResponse = await fetch(deleteEndpoint, {
            method: "DELETE",
            headers: {
                Authorization: `Bearer ${accessToken}`
            }
        });

        if (!deleteResponse.ok) {
            console.warn(`Failed to delete attachment ${attachment.id}: ${deleteResponse.statusText}`);
        } else {
            console.log(`Deleted attachment: ${attachment.name}`);
        }
    }
}

/**
 * Retrieves an access token for Microsoft Graph API.
 * @returns {Promise<string>} The access token.
 */
async function getAccessToken() {
    return new Promise((resolve, reject) => {
        Office.context.auth.getAccessTokenAsync({ allowSignInPrompt: true }, (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value);
            } else {
                reject(result.error);
            }
        });
    });
}
