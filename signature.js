Office.onReady();

function onNewMessageComposeHandler(event) {
    insertSignature(event);
}

function onNewReplyComposeHandler(event) {
    insertSignature(event);
}

// One entry per sub-brand, keyed by the sender's email domain.
const BRANDS = {
    "slash.ae": {
        name: "Slash",
        tagline: "Strategic Design Studio",
        url: "www.slash.ae"
    },
    "fabrica.ae": {
        name: "Fabrica",
        tagline: "Interdisciplinary Making Platform",
        url: "www.fabrica.ae"
    },
    "fountconcept.ae": {
        name: "Fount",
        tagline: "Curated Concept Store",
        url: "www.fountconcept.ae"
    },
    "central.ae": {
        name: "Central",
        tagline: "Logistics Platform",
        url: "www.central.ae"
    },
    "basearchitecture.ae": {
        name: "Base",
        tagline: "Collaborative Architecture Studio",
        url: "www.basearchitecture.ae"
    }
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
