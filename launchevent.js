/* PULSE Smart Alerts (Pilot) - OnMessageSend handler
   Note: This MVP blocks send when simple keywords are detected.
   Next step (optional): replace local keyword check with HTTP call to Power Automate.
*/

/* global Office */

Office.onReady(() => {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
});

function onMessageSendHandler(event) {
  const item = Office.context.mailbox.item;

  item.subject.getAsync((subRes) => {
    if (subRes.status !== Office.AsyncResultStatus.Succeeded) {
      event.completed({ allowEvent: true });
      return;
    }

    item.body.getAsync(Office.CoercionType.Text, (bodyRes) => {
      if (bodyRes.status !== Office.AsyncResultStatus.Succeeded) {
        event.completed({ allowEvent: true });
        return;
      }

      const subject = (subRes.value || "").toLowerCase();
      const body = (bodyRes.value || "").toLowerCase();
      const text = `${subject}
${body}`;

      // MVP keyword set (can be externalized later)
      const keywords = ["price", "unit price", "quotation", "cost", "报价", "价格", "单价", "总价"]; 
      const hit = keywords.find(k => text.includes(k));

      if (hit) {
        event.completed({
          allowEvent: false,
          errorMessage: `PULSE Agent Alert: detected sensitive keyword "${hit}". Please review before sending.`
        });
        return;
      }

      event.completed({ allowEvent: true });
    });
  });
}
