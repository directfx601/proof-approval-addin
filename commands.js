/*
 * directFX Solutions — Proof Approval Template add-in
 *
 * Inserts the company's standard proof-request / approval wording into
 * the body of whatever email the user is composing. Static wording only —
 * no per-job fields. Staff still fill in job specifics manually before
 * sending.
 *
 * To change the wording company-wide, edit PROOF_TEMPLATE_HTML below and
 * redeploy (re-upload this file to the hosting location). No changes to
 * the manifest are needed for a wording-only edit.
 */

Office.onReady(() => {
  // Required by Office.js; no additional startup logic needed.
});

const PROOF_TEMPLATE_HTML =
  "<p><strong>Proof Attached &mdash; Approval Required</strong></p>" +
  "<p>Please review the attached proof carefully, including all text, layout, and color before responding.</p>" +
  "<p>To approve this proof for production, reply to this email with exactly:</p>" +
  "<p style=\"margin-left:20px;\"><strong>APPROVED &mdash; [Version #] &mdash; no changes</strong></p>" +
  "<p>If you need changes, reply with exactly:</p>" +
  "<p style=\"margin-left:20px;\"><strong>CHANGES REQUESTED &mdash; [Version #]</strong> and list your requested changes below.</p>" +
  "<p>Please respond within 2 business days to hold your production date.</p>" +
  "<p style=\"font-size:11px;color:#555;\"><em>Approval of this proof authorizes production and confirms the accuracy of all text, " +
  "layout, and color shown. directFX Solutions is not responsible for errors present on an approved proof. Digital/PDF proofs do " +
  "not guarantee an exact color match &mdash; a physical contract proof is available on request for color-critical jobs.</em></p>" +
  "<p>&nbsp;</p>";

/**
 * Ribbon button handler: inserts the standard proof-approval wording
 * at the current cursor position in the compose body.
 * @param {Office.AddinCommands.Event} event
 */
function insertProofTemplate(event) {
  Office.context.mailbox.item.body.setSelectedDataAsync(
    PROOF_TEMPLATE_HTML,
    { coercionType: Office.CoercionType.Html },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        // Fail silently in production; surface in console for troubleshooting.
        console.error("Failed to insert proof template: " + asyncResult.error.message);
      }
      event.completed();
    }
  );
}

// Register the function so the manifest's ExecuteFunction action can find it.
Office.actions.associate("insertProofTemplate", insertProofTemplate);
