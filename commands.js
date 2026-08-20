/* UPG signature add-in - event-based activation */
var SIGNATURES = {
  "kbanks@codeplayground.io": "<table cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"font-family:Arial,Helvetica,sans-serif;background:#ffffff;\"> <tr> <td style=\"vertical-align:middle;\"><a href=\"https://www.upg.energy\" style=\"text-decoration:none;\"><img src=\"https://kalvinbanks.github.io/upg-signature-addin/img/kbanks.jpg\" width=\"310\" height=\"197\" alt=\"United Power Group\" style=\"display:block;border:0;\"></a></td> <td style=\"vertical-align:middle;padding:8px 0 8px 22px;border-left:3px solid #76b82a;\"> <div style=\"font-size:21px;font-weight:bold;color:#231f20;letter-spacing:.2px;\">Kalvin Banks</div> <div style=\"margin-top:8px;font-size:14px;\"> <a href=\"mailto:kbanks@codeplayground.io\" style=\"color:#231f20;text-decoration:none;\">kbanks@codeplayground.io</a> &nbsp;&nbsp;\u2022&nbsp;&nbsp;<a href=\"https://www.upg.energy\" style=\"color:#1d5632;font-weight:bold;text-decoration:none;\">www.upg.energy</a> </div> <div style=\"margin-top:9px;font-size:11px;color:#6d6e71;\">601 S. Railroad ST, Lewisville, TX 75057</div> </td> </tr> </table>"
};

function pickSignature() {
  var email = "";
  try {
    var p = Office.context.mailbox.userProfile;
    if (p && p.emailAddress) { email = String(p.emailAddress).toLowerCase(); }
  } catch (e) { email = ""; }

  if (SIGNATURES[email]) { return SIGNATURES[email]; }

  var keys = [];
  for (var k in SIGNATURES) { if (SIGNATURES.hasOwnProperty(k)) { keys.push(k); } }
  if (keys.length === 1) { return SIGNATURES[keys[0]]; }
  return null;
}

/* Fallback for Outlook on the web, where the sample uses body.setAsync
   rather than setSignatureAsync. */
function writeViaBody(html, event) {
  try {
    Office.context.mailbox.item.body.setAsync(
      "<br/><br/>" + html,
      { coercionType: "html", asyncContext: event },
      function (r) { r.asyncContext.completed(); }
    );
  } catch (e) {
    event.completed();
  }
}

function applySignature(event) {
  var html = null;
  try { html = pickSignature(); } catch (e) { html = null; }
  if (!html) { event.completed(); return; }

  var item = Office.context.mailbox.item;
  if (item && item.body && typeof item.body.setSignatureAsync === "function") {
    try {
      item.body.setSignatureAsync(
        html,
        { coercionType: Office.CoercionType.Html },
        function (result) {
          if (result && result.status === Office.AsyncResultStatus.Failed) {
            writeViaBody(html, event);
            return;
          }
          event.completed();
        }
      );
      return;
    } catch (e) { /* fall through to body.setAsync */ }
  }
  writeViaBody(html, event);
}

Office.actions.associate("applySignature", applySignature);

/* REQUIRED for event-based activation on Outlook on the web and new Outlook.
   Without this the runtime never signals readiness and launch events never fire. */
Office.onReady();
