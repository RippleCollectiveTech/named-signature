Office.onReady();

function onNewMessageComposeHandler(event) {
    insertSignature(event);
}

function onNewReplyComposeHandler(event) {
    insertSignature(event);
}

// One entry per sub-brand, keyed by the sender's email domain.
// "site" is the website address without "www." (added in code below).
const BRANDS = {
    "slash.ae": {
        name: "Slash",
        tagline: "Strategic Design Studio",
        site: "slash.ae"
    },
    "fabrica.ae": {
        name: "Fabrica",
        tagline: "Interdisciplinary Making Platform",
        site: "fabrica.ae"
    },
    "fountconcept.ae": {
        name: "Fount",
        tagline: "Curated Concept Store",
        site: "fountconcept.ae"
    },
    "central.ae": {
        name: "Central",
        tagline: "Logistics Platform",
        site: "central.ae"
    },
    "basearchitecture.ae": {
        name: "Base",
        tagline: "Collaborative Architecture Studio",
        site: "basearchitecture.ae"
    }
};

const FONT = "font-family:Helvetica,Arial,sans-serif;font-size:10pt;";

function makeLink(site) {
    const host = "www." + site;
    return '<a href="https://' + host + '" style="' + FONT + '">' + host + "</a>";
}

const RIPPLE_LINK = makeLink("ripplecollective.ae");

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
  ${makeLink(brand.site)}<br><br>
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