Office.onReady();

function onNewMessageComposeHandler(event) {
    insertSignature(event);
}

function onNewReplyComposeHandler(event) {
    insertSignature(event);
}

function insertSignature(event) {
    const signatureHtml = `<br><br>
<div style="font-family: Arial, Helvetica, sans-serif; font-size: 10pt; line-height: 1.4; color: #000000;">
  Best,<br><br>
  <strong>Amin Alzouabi</strong><br><br>
  –<br>
  <strong>Ripple Collective</strong><br>
  Multidisciplinary Design Ecosystem<br>
  <a href="https://www.ripplecollective.ae" style="color: #0066cc; text-decoration: none;">www.ripplecollective.ae</a><br><br>
  <span style="font-size: 9pt; color: #555555;">[Ripple Collective is an ecosystem of projects, including Slash, Fabrica, Fount, Central, and Base, all based in Abu Dhabi.]</span>
</div>`;

    Office.context.mailbox.item.body.setSignatureAsync(
        signatureHtml,
        { coercionType: Office.CoercionType.Html },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error(asyncResult.error.message);
            }
            event.completed();
        }
    );
}
