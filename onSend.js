// onSend.js
// Fires automatically at the moment Send is clicked (reply, forward, or new).
// Strips any day marker from the subject so it never goes out externally.
// Desktop (new + classic), Mac, and web only - does NOT run on Outlook mobile.
// (Known, accepted gap - see conversation with Kate, Aug 2026.)

// Matches an optional chain of RE:/FW:/FWD: prefixes, then the marker itself.
// Keeps the RE:/FW: prefixes, drops the marker + trailing dash.
const MARKER_REGEX =
  /^((?:\s*(?:RE|FW|FWD)\s*:\s*)*)(?:🔴|🟠|🟡|🟢|🔵|🟣|⚪)\s*(?:Monday|Tuesday|Wednesday|Thursday|Friday|Next Week|Future)\s*[—-]\s*/iu;

function onMessageSendHandler(event) {
  Office.context.mailbox.item.subject.getAsync((subjectResult) => {
    if (subjectResult.status !== Office.AsyncResultStatus.Succeeded) {
      // If we can't even read the subject, don't block the send - just let it go.
      event.completed({ allowEvent: true });
      return;
    }

    const original = subjectResult.value || "";
    const stripped = original.replace(MARKER_REGEX, "$1");

    if (stripped === original) {
      // No marker present - nothing to do.
      event.completed({ allowEvent: true });
      return;
    }

    Office.context.mailbox.item.subject.setAsync(stripped, (setResult) => {
      // Whether the strip succeeded or not, never block Harvey from sending.
      // (SendMode is set to SoftBlock in the manifest as a second layer of safety.)
      event.completed({ allowEvent: true });
    });
  });
}

// Register the handler so Office can find it by the FunctionName
// declared in the manifest's LaunchEvent element.
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
