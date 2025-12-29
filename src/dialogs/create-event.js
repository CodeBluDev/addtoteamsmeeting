/* global Office */

function toInputValue(date) {
  const pad = (value) => String(value).padStart(2, "0");
  return (
    `${date.getFullYear()}-` +
    `${pad(date.getMonth() + 1)}-` +
    `${pad(date.getDate())}T` +
    `${pad(date.getHours())}:` +
    `${pad(date.getMinutes())}`
  );
}

function initializeDefaults() {
  const now = new Date();
  const later = new Date(now.getTime() + 30 * 60000);
  document.getElementById("start").value = toInputValue(now);
  document.getElementById("end").value = toInputValue(later);
}

function handleParentMessage(message) {
  let data;
  try {
    data = JSON.parse(message);
  } catch (error) {
    return;
  }

  if (data.subject) {
    document.getElementById("subject").value = data.subject;
  }

  if (data.teamsLink) {
    document.getElementById("teamsLink").textContent = data.teamsLink;
  }
}

function sendCreate() {
  const subject = document.getElementById("subject").value;
  const start = document.getElementById("start").value;
  const end = document.getElementById("end").value;

  Office.context.ui.messageParent(
    JSON.stringify({
      action: "create",
      subject,
      start,
      end
    })
  );
}

function sendCancel() {
  Office.context.ui.messageParent(JSON.stringify({ action: "cancel" }));
}

Office.onReady(() => {
  const status = document.getElementById("jsStatus");
  if (status) {
    status.textContent = "JS loaded";
  }
  // eslint-disable-next-line no-console
  console.log("[Dialog] JS loaded");

  if (!Office.context || !Office.context.ui || !Office.context.ui.addHandlerAsync) {
    if (status) {
      status.textContent = "Office.js not available (open in Outlook dialog)";
    }
    return;
  }

  initializeDefaults();
  document.getElementById("createBtn").addEventListener("click", sendCreate);
  document.getElementById("cancelBtn").addEventListener("click", sendCancel);

  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived,
    (arg) => handleParentMessage(arg.message)
  );
});
