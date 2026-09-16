Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) {
        insertSignature();
    }
});

function insertSignature() {
    // Define your company HTML signature layout here
    var signatureHtml = `
        <br/><br/>
        <div style="font-family: Arial, sans-serif; font-size: 13px; color: #333;">
            <strong>Best regards,</strong><br/>
            <span style="color: #0056b3; font-size: 14px;"><strong>Company Support Team</strong></span><br/>
            <span>Contoso Ltd. | <a href="https://example.com">www.example.com</a></span>
        </div>
    `;

    // Insert text at current cursor position (top of reply, above quoted thread)
    Office.context.mailbox.item.body.setSelectedDataAsync(
        signatureHtml,
        { coercionType: Office.CoercionType.Html },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Signature insertion failed: " + asyncResult.error.message);
            } else {
                console.log("Signature inserted successfully.");
            }
        }
    );
}
