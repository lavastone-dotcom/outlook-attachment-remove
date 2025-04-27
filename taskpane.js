Office.onReady(info => {
    if (info.host === Office.HostType.Outlook) {
        const attachments = Office.context.mailbox.item.attachments;
        const listDiv = document.getElementById("attachment-list");

        if (attachments.length === 0) {
            listDiv.innerHTML = "<p>No attachments found.</p>";
        } else {
            attachments.forEach(att => {
                const el = document.createElement("div");
                el.innerHTML = `<input type='checkbox' value='${att.id}' /> ${att.name}`;
                listDiv.appendChild(el);
            });
        }
    }
});

function deleteSelected() {
    const selected = Array.from(document.querySelectorAll("input[type=checkbox]:checked")).map(e => e.value);

    if (selected.length === 0) {
        showMessage("Please select at least one attachment to delete.", "error");
        return;
    }

    if (!confirm(`Are you sure you want to delete ${selected.length} attachment(s)?`)) {
        return;
    }

    selected.forEach(id => deleteAttachment(id));
}

function showMessage(message, type) {
    const messageArea = document.getElementById("message-area");
    messageArea.style.color = (type === "error") ? "red" : "green";
    messageArea.innerText = message;
}
