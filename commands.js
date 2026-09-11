/* global Office */
const CRM_BCC = "sync-639acf3e3a5d76cfe1a236bc199dec9d85b73184@inbound.crmbook.dk";
Office.onReady();
function addCrmBcc(event) {
  const item = Office.context.mailbox.item;
  item.bcc.getAsync((current) => {
    if (current.status !== Office.AsyncResultStatus.Succeeded) return finish(event, "Unable to read the Bcc field.", "error");
    if (current.value.some((p) => p.emailAddress.toLowerCase() === CRM_BCC.toLowerCase())) return finish(event, "CRM Book is already in Bcc.", "informationalMessage");
    item.bcc.addAsync([CRM_BCC], (added) => finish(event, added.status === Office.AsyncResultStatus.Succeeded ? "CRM Book added to Bcc." : "Could not add CRM Book to Bcc.", added.status === Office.AsyncResultStatus.Succeeded ? "informationalMessage" : "error"));
  });
}
function finish(event, message, type) { Office.context.mailbox.item.notificationMessages.replaceAsync("crm-bcc-status", {type, message, persistent:false}, () => event.completed()); }
