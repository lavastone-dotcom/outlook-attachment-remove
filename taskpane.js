Office.onReady(info => {
    if (info.host === Office.HostType.Outlook) {
        Office.context.mailbox.item.attachments.forEach(att => {
            const el = document.createElement("div");
            el.innerHTML = `<input type='checkbox' value='${att.id}' /> ${att.name}`;
            document.getElementById("attachment-list").appendChild(el);
        });
    }
});

function deleteSelected() {
    const selected = Array.from(document.querySelectorAll("input[type=checkbox]:checked")).map(e => e.value);
    selected.forEach(id => deleteAttachment(id));
}
