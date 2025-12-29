/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

const BUILD_TAG = "v1.7.1";
const BUILD_MARKER = "2024-09-18T14:35Z";
const EWS_MESSAGES_NS = "http://schemas.microsoft.com/exchange/services/2006/messages";
const EWS_TYPES_NS = "http://schemas.microsoft.com/exchange/services/2006/types";
const DEBUG_LOGS = true;
const NOTIFICATION_ICON_URL = "https://mvteamsmeetinglink.netlify.app/assets/codeblu-teams-16.png?v=1.7.1";

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function addTeamsLinkToLocation(event) {
  const item = Office.context.mailbox.item;
  logDebug("Command invoked", { itemId: item.itemId, itemType: item.itemType });

  // Read the message body as HTML
  item.body.getAsync(Office.CoercionType.Html, (bodyResult) => {
    logDebug("Body getAsync result", { status: bodyResult.status });
    if (bodyResult.status !== Office.AsyncResultStatus.Succeeded) {
      item.notificationMessages.replaceAsync("error", {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: "Unable to read message body."
      });
      event.completed();
      return;
    }

    const bodyHtml = bodyResult.value;

    // Extract the Teams meeting link
    const teamsRegex = /https:\/\/teams\.microsoft\.com\/l\/meetup-join\/[^\s"<]+/i;
    const match = bodyHtml.match(teamsRegex);
    logDebug("Teams link match", { found: Boolean(match) });

    if (!match) {
      notifyInfo(item, "No Microsoft Teams meeting link found in this invite.");
      event.completed();
      return;
    }

    const teamsLink = match[0];
    logDebug("Teams link extracted", { teamsLink });
    findCalendarItemByTeamsLink(teamsLink, (findError, calendarItem) => {
      logDebug("Find by link", { error: Boolean(findError), found: Boolean(calendarItem) });
      if (findError) {
        notifyError(item, "Unable to search calendar items.");
        event.completed();
        return;
      }

      if (calendarItem) {
        updateCalendarItemLocation(calendarItem, teamsLink, (updateError) => {
          if (updateError) {
            notifyError(item, "Unable to update the calendar location.");
          } else {
            notifySuccess(item);
          }
          event.completed();
        });
        return;
      }

      getMessageTimeRange(item, (timeError, timeRange) => {
        logDebug("Message time range", { error: Boolean(timeError), hasRange: Boolean(timeRange) });
        if (timeError || !timeRange) {
          notifyInfo(item, "No matching calendar event found.");
          event.completed();
          return;
        }

        findCalendarItemByTimeRange(timeRange, (timeFindError, timeCalendarItem) => {
          logDebug("Find by time", {
            error: Boolean(timeFindError),
            found: Boolean(timeCalendarItem)
          });
          if (timeFindError) {
            notifyError(item, "Unable to search calendar items.");
            event.completed();
            return;
          }

          if (!timeCalendarItem) {
            notifyInfo(item, "No matching calendar event found.");
            event.completed();
            return;
          }

          updateCalendarItemLocation(timeCalendarItem, teamsLink, (updateError) => {
            if (updateError) {
              notifyError(item, "Unable to update the calendar location.");
            } else {
              notifySuccess(item);
            }
            event.completed();
          });
        });
      });
    });
  });
}

function notifySuccess(item) {
  item.notificationMessages.replaceAsync("success", {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    icon: NOTIFICATION_ICON_URL,
    message: `Teams meeting link added to Location. (${BUILD_TAG} | ${BUILD_MARKER})`
  });
}

function notifyError(item, message) {
  item.notificationMessages.replaceAsync("error", {
    type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
    message
  });
}

function notifyInfo(item, message) {
  item.notificationMessages.replaceAsync("info", {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    icon: NOTIFICATION_ICON_URL,
    message
  });
}

function getMessageTimeRange(item, callback) {
  if (!item.start || !item.end || !item.start.getAsync || !item.end.getAsync) {
    callback(null, null);
    return;
  }

  item.start.getAsync((startResult) => {
    if (startResult.status !== Office.AsyncResultStatus.Succeeded) {
      callback(startResult.error, null);
      return;
    }

    item.end.getAsync((endResult) => {
      if (endResult.status !== Office.AsyncResultStatus.Succeeded) {
        callback(endResult.error, null);
        return;
      }

      const start = new Date(startResult.value);
      const end = new Date(endResult.value);
      if (Number.isNaN(start.valueOf()) || Number.isNaN(end.valueOf())) {
        callback(null, null);
        return;
      }

      callback(null, { start, end });
    });
  });
}

function findCalendarItemByTeamsLink(teamsLink, callback) {
  const escapedLink = escapeXml(teamsLink);
  logDebug("EWS FindItem by link request prepared");
  const request = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="${EWS_TYPES_NS}"
               xmlns:m="${EWS_MESSAGES_NS}">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013" />
  </soap:Header>
  <soap:Body>
    <m:FindItem Traversal="Shallow">
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
      </m:ItemShape>
      <m:Restriction>
        <t:Contains ContainmentMode="Substring" ContainmentComparison="IgnoreCase">
          <t:FieldURI FieldURI="item:Body" />
          <t:Constant Value="${escapedLink}" />
        </t:Contains>
      </m:Restriction>
      <m:ParentFolderIds>
        <t:DistinguishedFolderId Id="calendar" />
      </m:ParentFolderIds>
    </m:FindItem>
  </soap:Body>
</soap:Envelope>`;

  Office.context.mailbox.makeEwsRequestAsync(request, (result) => {
    logDebug("EWS FindItem by link response", { status: result.status });
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      callback(result.error, null);
      return;
    }

    const calendarItem = parseFirstCalendarItem(result.value);
    callback(null, calendarItem);
  });
}

function findCalendarItemByTimeRange(timeRange, callback) {
  const startIso = timeRange.start.toISOString();
  const endIso = timeRange.end.toISOString();
  logDebug("EWS FindItem by time request prepared", { startIso, endIso });
  const request = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="${EWS_TYPES_NS}"
               xmlns:m="${EWS_MESSAGES_NS}">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013" />
  </soap:Header>
  <soap:Body>
    <m:FindItem Traversal="Shallow">
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="calendar:Start" />
          <t:FieldURI FieldURI="calendar:End" />
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:CalendarView StartDate="${startIso}" EndDate="${endIso}" />
      <m:ParentFolderIds>
        <t:DistinguishedFolderId Id="calendar" />
      </m:ParentFolderIds>
    </m:FindItem>
  </soap:Body>
</soap:Envelope>`;

  Office.context.mailbox.makeEwsRequestAsync(request, (result) => {
    logDebug("EWS FindItem by time response", { status: result.status });
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      callback(result.error, null);
      return;
    }

    const calendarItem = parseCalendarItemByTime(result.value, timeRange);
    callback(null, calendarItem);
  });
}

function updateCalendarItemLocation(calendarItem, location, callback) {
  const escapedLocation = escapeXml(location);
  logDebug("EWS UpdateItem request prepared", { itemId: calendarItem.id });
  const request = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="${EWS_TYPES_NS}"
               xmlns:m="${EWS_MESSAGES_NS}">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013" />
  </soap:Header>
  <soap:Body>
    <m:UpdateItem ConflictResolution="AlwaysOverwrite" SendMeetingInvitationsOrCancellations="SendToNone">
      <m:ItemChanges>
        <t:ItemChange>
          <t:ItemId Id="${calendarItem.id}" ChangeKey="${calendarItem.changeKey}" />
          <t:Updates>
            <t:SetItemField>
              <t:FieldURI FieldURI="calendar:Location" />
              <t:CalendarItem>
                <t:Location>${escapedLocation}</t:Location>
              </t:CalendarItem>
            </t:SetItemField>
          </t:Updates>
        </t:ItemChange>
      </m:ItemChanges>
    </m:UpdateItem>
  </soap:Body>
</soap:Envelope>`;

  Office.context.mailbox.makeEwsRequestAsync(request, (result) => {
    logDebug("EWS UpdateItem response", { status: result.status });
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      callback(result.error);
      return;
    }

    const xmlDoc = parseXml(result.value);
    if (!xmlDoc || !isEwsResponseSuccess(xmlDoc)) {
      callback(new Error("EWS UpdateItem failed."));
      return;
    }

    callback(null);
  });
}

function parseFirstCalendarItem(responseText) {
  const xmlDoc = parseXml(responseText);
  if (!xmlDoc) {
    return null;
  }

  const itemId = xmlDoc.getElementsByTagNameNS(EWS_TYPES_NS, "ItemId")[0];
  if (!itemId) {
    return null;
  }

  return {
    id: itemId.getAttribute("Id"),
    changeKey: itemId.getAttribute("ChangeKey")
  };
}

function parseCalendarItemByTime(responseText, timeRange) {
  const xmlDoc = parseXml(responseText);
  if (!xmlDoc) {
    return null;
  }

  const calendarItems = xmlDoc.getElementsByTagNameNS(EWS_TYPES_NS, "CalendarItem");
  for (let i = 0; i < calendarItems.length; i += 1) {
    const calendarItem = calendarItems[i];
    const startNode = calendarItem.getElementsByTagNameNS(EWS_TYPES_NS, "Start")[0];
    const endNode = calendarItem.getElementsByTagNameNS(EWS_TYPES_NS, "End")[0];
    const itemId = calendarItem.getElementsByTagNameNS(EWS_TYPES_NS, "ItemId")[0];

    if (!startNode || !endNode || !itemId) {
      continue;
    }

    const start = new Date(startNode.textContent);
    const end = new Date(endNode.textContent);
    if (Number.isNaN(start.valueOf()) || Number.isNaN(end.valueOf())) {
      continue;
    }

    if (isSameTimeRange(start, end, timeRange)) {
      return {
        id: itemId.getAttribute("Id"),
        changeKey: itemId.getAttribute("ChangeKey")
      };
    }
  }

  return null;
}

function isSameTimeRange(start, end, timeRange) {
  const toleranceMs = 60000;
  return (
    Math.abs(start - timeRange.start) <= toleranceMs &&
    Math.abs(end - timeRange.end) <= toleranceMs
  );
}

function parseXml(xmlString) {
  if (!xmlString) {
    return null;
  }

  return new DOMParser().parseFromString(xmlString, "text/xml");
}

function isEwsResponseSuccess(xmlDoc) {
  const responseCode = xmlDoc.getElementsByTagNameNS(EWS_MESSAGES_NS, "ResponseCode")[0];
  return responseCode && responseCode.textContent === "NoError";
}

function escapeXml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function logDebug(message, data) {
  if (!DEBUG_LOGS) {
    return;
  }

  if (data) {
    // eslint-disable-next-line no-console
    console.log(`[AddTeamsLink] ${message}`, data);
  } else {
    // eslint-disable-next-line no-console
    console.log(`[AddTeamsLink] ${message}`);
  }
}

// Register the function with Office.
Office.actions.associate("addTeamsLinkToLocation", addTeamsLinkToLocation);
