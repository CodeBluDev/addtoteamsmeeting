/******/ (function() { // webpackBootstrap
/*!*************************************!*\
  !*** ./src/dialogs/create-event.js ***!
  \*************************************/
/* global Office */

function toInputValue(date) {
  var pad = function pad(value) {
    return String(value).padStart(2, "0");
  };
  return "".concat(date.getFullYear(), "-") + "".concat(pad(date.getMonth() + 1), "-") + "".concat(pad(date.getDate()), "T") + "".concat(pad(date.getHours()), ":") + "".concat(pad(date.getMinutes()));
}
function initializeDefaults() {
  var now = new Date();
  var later = new Date(now.getTime() + 30 * 60000);
  document.getElementById("start").value = toInputValue(now);
  document.getElementById("end").value = toInputValue(later);
}
function handleParentMessage(message) {
  var data;
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
  var subject = document.getElementById("subject").value;
  var start = document.getElementById("start").value;
  var end = document.getElementById("end").value;
  Office.context.ui.messageParent(JSON.stringify({
    action: "create",
    subject: subject,
    start: start,
    end: end
  }));
}
function sendCancel() {
  Office.context.ui.messageParent(JSON.stringify({
    action: "cancel"
  }));
}
Office.onReady(function () {
  var status = document.getElementById("jsStatus");
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
  Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, function (arg) {
    return handleParentMessage(arg.message);
  });
});
/******/ })()
;
//# sourceMappingURL=createEventDialog.js.map