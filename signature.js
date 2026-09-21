Office.onReady();

function onNewMessageComposeHandler(event) {
    insertSignature(event);
}

function onNewReplyComposeHandler(event) {
    insertSignature(event);
}

// One entry per brand, keyed by email domain.
const BRANDS = {
    "slash.ae": {
        name: "Slash",
        tagline: "Strategic Design Studio",
        url: "www.slash.ae"
    }
    // Copy the entry above for each remaining brand, e.g.:
    // "fabrica.ae": { name: "Fabrica", tagline: "TAGLINE", url: "www.fabrica.ae" },
    // "fount.ae":   { name: "Fount",   tagline: "TAGLINE", url: "www.fount.ae" },
    // "central.ae": { name: "Central", tagline: "TAGLINE", url: "www.central.ae" },
    // "base.ae":    { name: "Base",    tagline: "TAGLINE", url: "www.base.ae" }
};

const FONT = "font-family:Helvetica,Arial,sans-serif;font-size:10pt;";
const RIPPLE_LINK = `<a href="https://www.ripplecollective.ae" style="${FONT}">www.ripplecollective.ae</a>`;

function getSignatureHtml(name, email) {
    const domain = (email.split("@")[1] || "").toLowerCase();
    const brand = BRANDS[domain];

    if (brand) {
        return `<br><br>
<div style="${FONT}line-height:normal;color:#000000;">
  Best,<br><br>
  <strong>${name}</strong><br><br>
  –<br>
  <strong>${brand.name}</strong><br>
  ${brand.tagline}<br>
  <a href="https://${brand.url}" style="${FONT}">${brand.url}</a><br><br>
  ${brand.name} is part of Ripple Collective, a Multidisciplinary Design Ecosystem, based in Abu Dhabi.<br>
  ${RIPPLE_LINK}
</div>`;
    }

    // Default: Ripple Collective
    return `<br><br>
<div style="${FONT}line-height:normal;color:#000000;">
  Best,<br><br>
  <strong>${name}</strong><br><br>
  –<br>
  <strong>Ripple Collective</strong><br>
  Multidisciplinary Design Ecosystem<br>
  ${RIPPLE_LINK}<br><br>
  Ripple Collective is an ecosystem of projects, including Slash, Fabrica, Fount, Central, and Base, all based in Abu Dhabi.
</div>`;
}

function insertSignature(event) {
    const profile = Office.context.mailbox.userProfile;
    const name = escapeHtml(profile.displayName || "");
    const signatureHtml = getSignatureHtml(name, profile.emailAddress || "");

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

function escapeHtml(text) {
    return text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
Office.actions.associate("onNewReplyComposeHandler", onNewReplyComposeHandler);
