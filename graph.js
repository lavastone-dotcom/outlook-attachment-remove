function deleteAttachment(attachmentId) {
    Office.context.mailbox.getCallbackTokenAsync({ isRest: false }, function (result) {
        if (result.status === "succeeded") {
            const token = result.value;
            const itemId = Office.context.mailbox.item.itemId;

            const request = `
                <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
                               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">
                  <soap:Body>
                    <DeleteAttachment xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">
                      <AttachmentIds>
                        <t:AttachmentId Id="${attachmentId}" />
                      </AttachmentIds>
                    </DeleteAttachment>
                  </soap:Body>
                </soap:Envelope>`;

            Office.context.mailbox.makeEwsRequestAsync(request, function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    showMessage("Attachment deleted successfully.", "success");
                    setTimeout(() => location.reload(), 1000);
                } else {
                    showMessage("Failed to delete attachment: " + asyncResult.error.message, "error");
                }
            });

        } else {
            showMessage("Could not get token: " + result.error.message, "error");
        }
    });
}
