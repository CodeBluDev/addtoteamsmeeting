/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, OfficeRuntime */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

const BUILD_TAG = "v1.7.1";
const BUILD_MARKER = "2024-09-18T14:35Z";
const EWS_MESSAGES_NS = "http://schemas.microsoft.com/exchange/services/2006/messages";
const EWS_TYPES_NS = "http://schemas.microsoft.com/exchange/services/2006/types";
const DEBUG_LOGS = true;
const NOTIFICATION_ICON_ID = "Icon.16x16";
const DIALOG_URL = "https://mvteamsmeetinglink.netlify.app/create-event.html?v=1.7.1";
const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
const GRAPH_SEARCH_DAYS = 90;

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
    const teamsLink = extractTeamsLink(bodyHtml);
    logDebug("Teams link match", { found: Boolean(teamsLink) });

    if (!teamsLink) {
      notifyInfo(item, "No Microsoft Teams meeting link found in this invite.");
      event.completed();
      return;
    }

    logDebug("Teams link extracted", { teamsLink });
    const meetingId = extractMeetingId(bodyHtml, teamsLink);
    logDebug("Meeting id extracted", { meetingId });

    findCalendarEventByGraph(teamsLink, meetingId)
      .then((eventId) => {
        if (!eventId) {
          notifyInfo(item, "No matching calendar event found. Opening event dialog.");
          openCreateEventDialog(item, teamsLink);
          event.completed();
          return null;
        }

        return updateCalendarEventLocationGraph(eventId, teamsLink)
          .then(() => {
            notifySuccess(item);
            event.completed();
            return null;
          })
          .catch((updateError) => {
            logDebug("Graph update failed", { message: updateError.message });
            notifyError(item, "Unable to update the calendar location.");
            event.completed();
            return null;
          });
      })
      .catch((error) => {
        logDebug("Graph search failed", { message: error.message });
        notifyInfo(item, "Opening event dialog (calendar search blocked).");
        openCreateEventDialog(item, teamsLink);
        event.completed();
      });
  });
}

function notifySuccess(item) {
  item.notificationMessages.replaceAsync("success", {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    icon: NOTIFICATION_ICON_ID,
    persistent: false,
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
    icon: NOTIFICATION_ICON_ID,
    persistent: false,
    message
  });
}

function openCreateEventDialog(item, teamsLink) {
  const subject = item.subject || "Teams meeting";

  Office.context.ui.displayDialogAsync(
    DIALOG_URL,
    { height: 55, width: 35, displayInIframe: true },
    (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        notifyError(item, "Unable to open the event dialog.");
        return;
      }

      const dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        let data;
        try {
          data = JSON.parse(arg.message);
        } catch (parseError) {
          notifyError(item, "Invalid response from dialog.");
          dialog.close();
          return;
        }

        if (data.action === "cancel") {
          dialog.close();
          return;
        }

        if (data.action !== "create") {
          notifyError(item, "Unknown dialog response.");
          dialog.close();
          return;
        }

        const start = new Date(data.start);
        const end = new Date(data.end);
        if (Number.isNaN(start.valueOf()) || Number.isNaN(end.valueOf())) {
          notifyError(item, "Invalid date/time from dialog.");
          dialog.close();
          return;
        }

        Office.context.mailbox.displayNewAppointmentForm({
          subject: data.subject || subject,
          location: teamsLink,
          start,
          end
        });

        dialog.close();
        notifySuccess(item);
      });

      dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
        notifyInfo(item, "Event dialog closed.");
      });

      dialog.messageChild(
        JSON.stringify({
          subject,
          teamsLink
        })
      );
    }
  );
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
    logDebug("EWS FindItem by link response", {
      status: result.status,
      error: result.error ? { name: result.error.name, message: result.error.message } : null
    });
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
    logDebug("EWS FindItem by time response", {
      status: result.status,
      error: result.error ? { name: result.error.name, message: result.error.message } : null
    });
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
    logDebug("EWS UpdateItem response", {
      status: result.status,
      error: result.error ? { name: result.error.name, message: result.error.message } : null
    });
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

function extractMeetingId(bodyHtml, teamsLink) {
  const combined = `${bodyHtml} ${teamsLink || ""}`;
  const decoded = decodeHtmlEntities(decodeLink(combined));
  const match = decoded.match(/19:meeting_[^"'\\s<>]+/i);
  if (match) {
    return match[0];
  }

  const encodedMatch = combined.match(/19%3Ameeting_[^"'\\s<>%]+/i);
  if (encodedMatch) {
    return decodeLink(encodedMatch[0]);
  }

  return null;
}

async function findCalendarEventByGraph(teamsLink, meetingId) {
  const token = await getGraphAccessToken();
  const start = new Date();
  const end = new Date(start.getTime() + GRAPH_SEARCH_DAYS * 24 * 60 * 60 * 1000);
  let url = `${GRAPH_BASE_URL}/me/calendarView?startDateTime=${encodeURIComponent(
    start.toISOString()
  )}&endDateTime=${encodeURIComponent(end.toISOString())}` +
    "&$select=id,subject,body,location,onlineMeetingUrl,start,end";

  while (url) {
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Prefer: 'outlook.body-content-type="text"'
      }
    });

    if (!response.ok) {
      const text = await response.text();
      throw new Error(`Graph calendarView failed: ${response.status} ${text}`);
    }

    const data = await response.json();
    const items = data.value || [];

    for (let i = 0; i < items.length; i += 1) {
      if (eventMatchesTeams(items[i], teamsLink, meetingId)) {
        return items[i].id;
      }
    }

    url = data["@odata.nextLink"] || null;
  }

  return null;
}

async function updateCalendarEventLocationGraph(eventId, teamsLink) {
  const token = await getGraphAccessToken();
  const response = await fetch(`${GRAPH_BASE_URL}/me/events/${eventId}`, {
    method: "PATCH",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      location: {
        displayName: teamsLink
      }
    })
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Graph update failed: ${response.status} ${text}`);
  }
}

async function getGraphAccessToken() {
  if (!OfficeRuntime || !OfficeRuntime.auth || !OfficeRuntime.auth.getAccessToken) {
    throw new Error("OfficeRuntime auth is not available.");
  }

  const token = await OfficeRuntime.auth.getAccessToken({
    allowSignInPrompt: true,
    allowConsentPrompt: true,
    forMSGraphAccess: true
  });
  logDebug("Graph token acquired");
  return token;
}

function eventMatchesTeams(event, teamsLink, meetingId) {
  const bodyText = event.body && event.body.content ? event.body.content : "";
  const onlineUrl = event.onlineMeetingUrl || "";

  if (meetingId) {
    if (bodyText.includes(meetingId) || onlineUrl.includes(meetingId)) {
      return true;
    }
  }

  if (teamsLink) {
    if (bodyText.includes(teamsLink) || onlineUrl.includes(teamsLink)) {
      return true;
    }
  }

  return false;
}

function extractTeamsLink(bodyHtml) {
  const teamsRegex = /https:\/\/teams\.microsoft\.com\/l\/meetup-join\/[^\s"<]+/i;
  const safeLinksRegex = /https:\/\/[^\/]+\.safelinks\.protection\.outlook\.com\/[^\s"<]+/i;
  const akaTeamsRegex = /https:\/\/aka\.ms\/[^\s"<]*teams[^\s"<]*/i;
  const directLink = findTeamsLinkInText(bodyHtml);
  if (directLink) {
    return directLink;
  }

  const urls = bodyHtml.match(/https?:\/\/[^\s"'<>]+/gi) || [];
  for (let i = 0; i < urls.length; i += 1) {
    const rawUrl = urls[i];
    const cleanedUrl = rawUrl.replace(/&amp;/g, "&");

    if (teamsRegex.test(cleanedUrl)) {
      return decodeLink(cleanedUrl);
    }

    if (safeLinksRegex.test(cleanedUrl)) {
      const extracted = extractSafeLinkTarget(cleanedUrl);
      if (extracted) {
        const decoded = decodeLink(extracted);
        if (teamsRegex.test(decoded)) {
          return decoded;
        }
      }
    }

    if (akaTeamsRegex.test(cleanedUrl)) {
      return cleanedUrl;
    }
  }

  return null;
}

function findTeamsLinkInText(text) {
  const candidates = [];
  const rawMatches = text.match(/https:\/\/teams\.microsoft\.com\/l\/meetup-join\/[^\s"'<>]+/gi);
  if (rawMatches) {
    candidates.push(...rawMatches);
  }

  const decodedHtml = decodeHtmlEntities(text);
  const decodedMatches = decodedHtml.match(/https:\/\/teams\.microsoft\.com\/l\/meetup-join\/[^\s"'<>]+/gi);
  if (decodedMatches) {
    candidates.push(...decodedMatches);
  }

  for (let i = 0; i < candidates.length; i += 1) {
    const cleaned = decodeLink(candidates[i]);
    if (cleaned) {
      return cleaned;
    }
  }

  return null;
}

function decodeLink(url) {
  const cleaned = url.replace(/&amp;/g, "&");
  if (/%[0-9A-Fa-f]{2}/.test(cleaned)) {
    try {
      return decodeURIComponent(cleaned);
    } catch (error) {
      return cleaned;
    }
  }
  return cleaned;
}

function decodeHtmlEntities(text) {
  return text
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, "\"")
    .replace(/&#39;/g, "'");
}

function extractSafeLinkTarget(safeLinkUrl) {
  try {
    const url = new URL(safeLinkUrl);
    const target = url.searchParams.get("url");
    if (!target) {
      return null;
    }
    return decodeURIComponent(target);
  } catch (error) {
    return null;
  }
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
