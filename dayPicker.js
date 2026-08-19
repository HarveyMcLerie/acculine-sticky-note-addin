// dayPicker.js
// Runs inside the "Set Day" task pane on the message read surface.
// Lets Kate/Harvey tag the currently-open email with a coloured day marker.
// The actual subject rewrite happens server-side (n8n -> Microsoft Graph),
// because Office.js cannot edit the subject of an already-received message
// directly in read mode - only in compose mode.

// ---- CONFIG: update this if the n8n webhook path changes ----
const WEBHOOK_URL = "https://acculine135.app.n8n.cloud/webhook/day-marker";

// Keep this list in sync with onSend.js and the n8n workflow's strip logic.
const DAYS = [
  { key: "monday", label: "Monday", dot: "🔴" },
  { key: "tuesday", label: "Tuesday", dot: "🟠" },
  { key: "wednesday", label: "Wednesday", dot: "🟡" },
  { key: "thursday", label: "Thursday", dot: "🟢" },
  { key: "friday", label: "Friday", dot: "🔵" },
  { key: "nextweek", label: "Next Week", dot: "🟣" },
  { key: "future", label: "Future", dot: "⚪" },
];

Office.onReady(() => {
  const container = document.getElementById("buttons");

  DAYS.forEach((day) => {
    const btn = document.createElement("button");
    btn.className = "day-btn";
    btn.innerHTML = `<span class="dot">${day.dot}</span> ${day.label}`;
    btn.addEventListener("click", () => applyMarker(day));
    container.appendChild(btn);
  });

  document.getElementById("clearBtn").addEventListener("click", () => applyMarker(null));
});

function setStatus(msg) {
  document.getElementById("status").textContent = msg;
}

function applyMarker(day) {
  setStatus("Updating...");

  const item = Office.context.mailbox.item;
  const restId = Office.context.mailbox.convertToRestId(
    item.itemId,
    Office.MailboxEnums.RestVersion.v2_0
  );

  item.subject.getAsync((subjectResult) => {
    if (subjectResult.status !== Office.AsyncResultStatus.Succeeded) {
      setStatus("Couldn't read the subject. Try again.");
      return;
    }

    const payload = {
      restId: restId,
      restUrl: Office.context.mailbox.restUrl, // tells n8n which mailbox/tenant endpoint to hit
      currentSubject: subjectResult.value,
      day: day ? day.key : null, // null = clear
    };

    fetch(WEBHOOK_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    })
      .then((res) => {
        if (!res.ok) throw new Error("Webhook returned " + res.status);
        return res.json();
      })
      .then(() => {
        setStatus(day ? `Marked as ${day.label}.` : "Marker cleared.");
      })
      .catch((err) => {
        console.error(err);
        setStatus("Something went wrong updating the email. Please try again.");
      });
  });
}
