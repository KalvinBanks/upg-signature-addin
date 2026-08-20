/* UPG signature add-in - event-based activation (diagnostic build) */
var SIGNATURES = {
  "kbanks@codeplayground.io": "<table cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"font-family:Arial,Helvetica,sans-serif;background:#ffffff;\"> <tr> <td style=\"vertical-align:middle;\"><a href=\"https://www.upg.energy\" style=\"text-decoration:none;\"><img src=\"https://kalvinbanks.github.io/upg-signature-addin/img/kbanks.jpg\" width=\"310\" height=\"197\" alt=\"United Power Group\" style=\"display:block;border:0;\"></a></td> <td style=\"vertical-align:middle;padding:8px 0 8px 22px;border-left:3px solid #76b82a;\"> <div style=\"font-size:21px;font-weight:bold;color:#231f20;letter-spacing:.2px;\">Kalvin Banks</div> <div style=\"margin-top:8px;font-size:14px;\"> <a href=\"mailto:kbanks@codeplayground.io\" style=\"color:#231f20;text-decoration:none;\">kbanks@codeplayground.io</a> &nbsp;&nbsp;\u2022&nbsp;&nbsp;<a href=\"https://www.upg.energy\" style=\"color:#1d5632;font-weight:bold;text-decoration:none;\">www.upg.energy</a> </div> <div style=\"margin-top:9px;font-size:11px;color:#6d6e71;\">601 S. Railroad ST, Lewisville, TX 75057</div> </td> </tr> </table>"
};

function pickSignature() {
  var email = "";
  try {
    var p = Office.context.mailbox.userProfile;
    if (p && p.emailAddress) { email = String(p.emailAddress).toLowerCase(); }
  } catch (e) { email = ""; }

  if (SIGNATURES[email]) { return { html: SIGNATURES[email], why: "exact:" + email }; }

  // Fallback: if only one signature is configured, use it regardless of address.
  var keys = [];
  for (var k in SIGNATURES) { if (SIGNATURES.hasOwnProperty(k)) { keys.push(k); } }
  if (keys.length === 1) { return { html: SIGNATURES[keys[0]], why: "sole-entry (profile was '" + email + "')" }; }

  return { html: null, why: "no match for '" + email + "' among " + keys.length + " entries" };
}

function applySignature(event) {
  var chosen;
  try {
    chosen = pickSignature();
  } catch (e) {
    chosen = { html: null, why: "pick threw: " + (e && e.message) };
  }

  var html = chosen.html;
  if (!html) {
    // Visible diagnostic so we can tell "handler never ran" from "handler ran but found nothing".
    html = '<div style="font-family:Arial,sans-serif;font-size:12px;color:#b00">' +
           'UPG add-in ran but did not apply a signature &mdash; ' + chosen.why + '</div>';
  }

  try {
    Office.context.mailbox.item.body.setSignatureAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      function (result) {
        if (result && result.status === Office.AsyncResultStatus.Failed) {
          try {
            Office.context.mailbox.item.body.setSignatureAsync(
              "UPG add-in: setSignatureAsync failed - " + (result.error && result.error.message),
              { coercionType: Office.CoercionType.Text },
              function () { event.completed(); }
            );
            return;
          } catch (e2) { /* fall through */ }
        }
        event.completed();
      }
    );
  } catch (e) {
    event.completed();
  }
}

Office.actions.associate("applySignature", applySignature);
