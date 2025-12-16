// SEDL Trademark Enforcer – Smart Alerts OnMessageSend handler

// This function is called by Outlook when the user clicks Send.
function onMessageSendHandler(event) {
  try {
    const item = Office.context.mailbox.item;

    if (!item || !item.body) {
      // If we can't access body, never block sending
      event.completed({ allowEvent: true });
      return;
    }

    // Read the body as HTML
    item.body.getAsync(
      "html",
      { asyncContext: event },
      function (asyncResult) {
        const evt = asyncResult.asyncContext;

        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          // On any error, allow send
          evt.completed({ allowEvent: true });
          return;
        }

        let body = asyncResult.value || "";

        // Map of plain text ? trademarked form
        const trademarkMap = {
          "LTE": "LTE\u00AE",
          "LTEM": "LTEM\u00AE",
          "SCP": "SCP\u00AE",
          "Flash Cigar": "Flash Cigar\u00AE",
          "Dilse": "Dilse\u00AE"
        };

        Object.keys(trademarkMap).forEach(function (plain) {
          const marked = trademarkMap[plain];

          // Escape regex specials in the key
          const escapedPlain = plain.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

          // Replace whole word/phrase, but only when it's NOT already followed by ®
          const regex = new RegExp("\\b" + escapedPlain + "\\b(?!\\u00AE)", "g");

          body = body.replace(regex, marked);
        });

        // Write the updated HTML back
        item.body.setAsync(
          body,
          { asyncContext: evt, coercionType: "html" },
          function () {
            // Whatever happens, we must call completed and allow the send
            evt.completed({ allowEvent: true });
          }
        );
      }
    );
  } catch (e) {
    // Never block sending because of an exception
    event.completed({ allowEvent: true });
  }
}

// Register handler name for Smart Alerts runtime
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
