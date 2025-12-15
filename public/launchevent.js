// This function triggers when the user hits 'Send'
function onMessageSendHandler(event) {
  const item = Office.context.mailbox.item;

  // 1. Process the Subject Line
  item.subject.getAsync((result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      let subject = result.value;
      // Regex \b ensures we match the whole word (not "LTE" inside "FILTER")
      const newSubject = subject
        .replace(/\bLTE\b/g, "LTE®")
        .replace(/\bLTEM\b/g, "LTEM®")
        .replace(/\bFlash Cigar\b/g, "Flash Cigar®")
        .replace(/\bDilse\b/g, "Dilse™");

      item.subject.setAsync(newSubject);
    }
  });

  // 2. Process the Email Body
  item.body.getAsync(Office.CoercionType.Html, (result) => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      let body = result.value;
      const newBody = body
        .replace(/\bLTE\b/g, "LTE&reg;")
        .replace(/\bLTEM\b/g, "LTEM&reg;")
        .replace(/\bFlash Cigar\b/g, "Flash Cigar&reg;")
        .replace(/\bDilse\b/g, "Dilse&trade;");

      item.body.setAsync(newBody, { coercionType: Office.CoercionType.Html }, (asyncResult) => {
        // 3. Signal to Outlook that it's okay to finish sending
        event.completed({ allowEvent: true });
      });
    }
  });
}

// Register the function
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
