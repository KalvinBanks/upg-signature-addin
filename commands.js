/* UPG signature add-in - event-based activation */
var SIGNATURES = {
  "kbanks@codeplayground.io": "<table cellpadding=\"0\" cellspacing=\"0\" border=\"0\" style=\"font-family:Arial,Helvetica,sans-serif;background:#ffffff;\"> <tr> <td style=\"vertical-align:middle;\"><a href=\"https://www.upg.energy\" style=\"text-decoration:none;\"><img src=\"https://kalvinbanks.github.io/upg-signature-addin/img/kbanks.jpg\" width=\"310\" height=\"197\" alt=\"United Power Group\" style=\"display:block;border:0;\"></a></td> <td style=\"vertical-align:middle;padding:8px 0 8px 22px;border-left:3px solid #76b82a;\"> <div style=\"font-size:21px;font-weight:bold;color:#231f20;letter-spacing:.2px;\">Kalvin Banks</div> <div style=\"margin-top:8px;font-size:14px;\"> <a href=\"mailto:kbanks@codeplayground.io\" style=\"color:#231f20;text-decoration:none;\">kbanks@codeplayground.io</a> &nbsp;&nbsp;\u2022&nbsp;&nbsp;<a href=\"https://www.upg.energy\" style=\"color:#1d5632;font-weight:bold;text-decoration:none;\">www.upg.energy</a> </div> <div style=\"margin-top:9px;font-size:11px;color:#6d6e71;\">601 S. Railroad ST, Lewisville, TX 75057</div> </td> </tr> </table>"
};

function applySignature(event) {
  try {
    var profile = Office.context.mailbox.userProfile;
    var email = (profile && profile.emailAddress ? profile.emailAddress : "").toLowerCase();
    var html = SIGNATURES[email];
    if (!html) { event.completed(); return; }
    Office.context.mailbox.item.body.setSignatureAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      function () { event.completed(); }
    );
  } catch (e) {
    event.completed();
  }
}

Office.actions.associate("applySignature", applySignature);
