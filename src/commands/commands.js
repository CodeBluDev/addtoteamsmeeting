/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, OfficeRuntime */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

const BUILD_TAG = "v1.8.17";
const BUILD_MARKER = "2026-01-01T16:09Z";
const DEFAULT_BASE_URL = "https://codebludev.github.io/addtoteamsmeeting";
const CACHE_BUSTER = "1.8.17";
const EWS_MESSAGES_NS = "http://schemas.microsoft.com/exchange/services/2006/messages";
const EWS_TYPES_NS = "http://schemas.microsoft.com/exchange/services/2006/types";
const DEBUG_LOGS = true;
const NOTIFICATION_ICON_ID = "Icon.16x16";
const DIALOG_URL = getDialogUrl("create-event.html");
const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
const GRAPH_SEARCH_DAYS = 365;
const AAD_CLIENT_ID = "226fcb0c-fa77-48bb-a20e-70a75ce176fd";
const AAD_AUTHORITY = "https://login.microsoftonline.com/organizations";
const GRAPH_SCOPES = ["https://graph.microsoft.com/Calendars.ReadWrite"];
const AUTH_DIALOG_URL = getDialogUrl("auth.html");
let cachedGraphToken = null;
let cachedGraphTokenExpiresAt = 0;
const GRAPH_TOKEN_STORAGE_KEY = "addTeamsGraphToken";

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

    const meetingId = extractMeetingId(bodyHtml, teamsLink);
    logDebug("Teams meeting id extracted", { meetingId });

    (async () => {
      try {
        const eventId = await findCalendarEventByGraph(teamsLink, meetingId);
        if (!eventId) {
          notifyInfo(item, "No matching calendar event found. Opening event dialog.");
          openCreateEventDialog(item, teamsLink);
          return;
        }

        await updateCalendarEventLocationGraph(eventId, teamsLink);
        notifySuccess(item);
      } catch (error) {
        logDebug("Graph update failed", { message: error.message });
        notifyError(item, "Unable to update the calendar location via Microsoft Graph.");
      } finally {
        event.completed();
      }
    })();
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

function getUtcWindowStart() {
  const now = new Date();
  now.setDate(now.getDate() - 7);
  return now.toISOString();
}

function getUtcWindowEnd() {
  const now = new Date();
  now.setDate(now.getDate() + 90);
  return now.toISOString();
}

function openCreateEventDialog(item, teamsLink) {
  const baseSubject = item.subject || "Teams meeting";
  const subject = prependBuildTag(baseSubject);

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

        const subjectWithTag = prependBuildTag(data.subject || baseSubject);
        Office.context.mailbox.displayNewAppointmentForm({
          subject: subjectWithTag,
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

function prependBuildTag(subject) {
  const prefix = `[${BUILD_TAG}] `;
  if (!subject) {
    return prefix.trim();
  }
  if (subject.startsWith(prefix)) {
    return subject;
  }
  return `${prefix}${subject}`;
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
      <m:CalendarView StartDate="${getUtcWindowStart()}" EndDate="${getUtcWindowEnd()}" />
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
    logEwsResult("FindItem by link", result);
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      callback(result.error, null);
      return;
    }

    const calendarItem = parseFirstCalendarItem(result.value);
    callback(null, calendarItem);
  });
}

function runEwsHealthCheck(callback) {
  logDebug("EWS health check request prepared");
  const request = `<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
               xmlns:t="${EWS_TYPES_NS}"
               xmlns:m="${EWS_MESSAGES_NS}">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013" />
  </soap:Header>
  <soap:Body>
    <m:GetFolder>
      <m:FolderShape>
        <t:BaseShape>IdOnly</t:BaseShape>
      </m:FolderShape>
      <m:FolderIds>
        <t:DistinguishedFolderId Id="calendar" />
      </m:FolderIds>
    </m:GetFolder>
  </soap:Body>
</soap:Envelope>`;

  Office.context.mailbox.makeEwsRequestAsync(request, (result) => {
    logEwsResult("health check", result);

    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      callback(false, formatEwsError(result.error));
      return;
    }

    const xmlDoc = parseXml(result.value);
    if (!xmlDoc || !isEwsResponseSuccess(xmlDoc)) {
      callback(false, "EWS GetFolder failed.");
      return;
    }

    callback(true, "OK");
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
    logEwsResult("FindItem by time", result);
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
    logEwsResult("UpdateItem", result);
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
  if (teamsLink) {
    const linkMatch = teamsLink.match(/19:meeting_[^/?"'\\s<>]+/i);
    if (linkMatch) {
      return linkMatch[0];
    }
  }

  const combined = `${bodyHtml} ${teamsLink || ""}`;
  const decoded = decodeHtmlEntities(decodeLink(combined));
  const match = decoded.match(/19:meeting_[^/?"'\\s<>]+/i);
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
    logDebug("Graph calendarView request", {
      method: "GET",
      url,
      headers: {
        Authorization: exposeAuthHeader(token),
        Prefer: 'outlook.body-content-type="text"'
      }
    });
    const response = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        Prefer: 'outlook.body-content-type="text"'
      }
    });

    await logGraphResponse("calendarView", response);
    if (!response.ok) {
      const text = await response.text();
      throw new Error(`Graph calendarView failed: ${response.status} ${text}`);
    }

    const data = await response.json();
    const items = data.value || [];
    logDebug("Graph calendarView items", {
      count: items.length,
      sample: items.slice(0, 5).map((item) => summarizeGraphItem(item, teamsLink, meetingId))
    });

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
  const body = {
    location: {
      displayName: teamsLink
    }
  };
  logDebug("Graph update request", {
    method: "PATCH",
    url: `${GRAPH_BASE_URL}/me/events/${eventId}`,
    headers: {
      Authorization: exposeAuthHeader(token),
      "Content-Type": "application/json"
    },
    body
  });
  const response = await fetch(`${GRAPH_BASE_URL}/me/events/${eventId}`, {
    method: "PATCH",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(body)
  });

  await logGraphResponse("updateEvent", response);
  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Graph update failed: ${response.status} ${text}`);
  }
}

async function getGraphAccessToken() {
  if (cachedGraphToken && Date.now() < cachedGraphTokenExpiresAt) {
    return cachedGraphToken;
  }

  const stored = readStoredGraphToken();
  if (stored) {
    cachedGraphToken = stored.token;
    cachedGraphTokenExpiresAt = stored.expiresAt;
    logDebug("Graph token loaded from storage");
    return stored.token;
  }

  if (OfficeRuntime && OfficeRuntime.auth && OfficeRuntime.auth.getAccessToken) {
    try {
      const token = await OfficeRuntime.auth.getAccessToken({
        allowSignInPrompt: true,
        allowConsentPrompt: true,
        forMSGraphAccess: true
      });
      logDebug("Graph token acquired (OfficeRuntime)");
      cacheGraphToken(token, 50);
      return token;
    } catch (error) {
      logDebug("OfficeRuntime auth failed", { message: error.message });
    }
  }

  const token = await getGraphAccessTokenViaDialog();
  logDebug("Graph token acquired (dialog)");
  cacheGraphToken(token, 45);
  storeGraphToken(token);
  return token;
}

function eventMatchesTeams(event, teamsLink, meetingId) {
  const bodyText = normalizeText(event.body && event.body.content ? event.body.content : "");
  const onlineUrl = normalizeText(event.onlineMeetingUrl || "");
  const locationText = normalizeText(
    event.location && event.location.displayName ? event.location.displayName : ""
  );
  const normalizedMeetingId = normalizeText(meetingId || "");
  const normalizedTeamsLink = normalizeText(teamsLink || "");

  if (normalizedMeetingId) {
    if (
      bodyText.includes(normalizedMeetingId) ||
      onlineUrl.includes(normalizedMeetingId) ||
      locationText.includes(normalizedMeetingId)
    ) {
      return true;
    }
  }

  if (normalizedTeamsLink) {
    if (
      bodyText.includes(normalizedTeamsLink) ||
      onlineUrl.includes(normalizedTeamsLink) ||
      locationText.includes(normalizedTeamsLink)
    ) {
      return true;
    }
  }

  return false;
}

function cacheGraphToken(token, minutes) {
  cachedGraphToken = token;
  cachedGraphTokenExpiresAt = Date.now() + minutes * 60 * 1000;
}

function storeGraphToken(token) {
  const expiresAt = decodeJwtExpiresAt(token);
  if (!expiresAt) {
    return;
  }
  try {
    localStorage.setItem(
      GRAPH_TOKEN_STORAGE_KEY,
      JSON.stringify({ token, expiresAt })
    );
  } catch (error) {
    // Ignore storage errors.
  }
}

function readStoredGraphToken() {
  try {
    const raw = localStorage.getItem(GRAPH_TOKEN_STORAGE_KEY);
    if (!raw) {
      return null;
    }
    const parsed = JSON.parse(raw);
    if (!parsed || !parsed.token || !parsed.expiresAt) {
      return null;
    }
    if (Date.now() >= parsed.expiresAt - 60 * 1000) {
      localStorage.removeItem(GRAPH_TOKEN_STORAGE_KEY);
      return null;
    }
    return parsed;
  } catch (error) {
    return null;
  }
}

function decodeJwtExpiresAt(token) {
  try {
    const parts = token.split(".");
    if (parts.length < 2) {
      return null;
    }
    const payload = JSON.parse(atob(parts[1].replace(/-/g, "+").replace(/_/g, "/")));
    if (!payload || !payload.exp) {
      return null;
    }
    return payload.exp * 1000;
  } catch (error) {
    return null;
  }
}

function getGraphAccessTokenViaDialog() {
  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      AUTH_DIALOG_URL,
      { height: 60, width: 40, displayInIframe: false },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          const message = result.error && result.error.message ? result.error.message : "Unknown dialog error.";
          logDebug("Auth dialog open failed", { code: result.error && result.error.code, message });
          reject(new Error(`Unable to open auth dialog: ${message}`));
          return;
        }

        const dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
          let data;
          try {
            data = JSON.parse(arg.message);
          } catch (parseError) {
            dialog.close();
            reject(new Error("Invalid auth dialog response."));
            return;
          }

          if (data.type === "log") {
            logDebug("Auth dialog log", {
              level: data.level,
              message: data.message,
              data: data.data || null
            });
            return;
          }

          if (data.type === "ready") {
            dialog.messageChild(
              JSON.stringify({
                clientId: AAD_CLIENT_ID,
                authority: AAD_AUTHORITY,
                scopes: GRAPH_SCOPES
              })
            );
            return;
          }

          if (data.type === "token" && data.accessToken) {
            dialog.close();
            resolve(data.accessToken);
            return;
          }

          dialog.close();
          reject(new Error(data.message || "Auth failed."));
        });

        dialog.addEventHandler(Office.EventType.DialogEventReceived, (arg) => {
          logDebug("Auth dialog event", { error: arg && arg.error ? arg.error : null });
          reject(new Error("Auth dialog closed."));
        });
      }
    );
  });
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

function normalizeText(text) {
  if (!text) {
    return "";
  }
  const decoded = decodeHtmlEntities(String(text));
  const cleaned = decodeLink(decoded);
  return cleaned.toLowerCase();
}

function summarizeGraphItem(item, teamsLink, meetingId) {
  const bodyText = normalizeText(item.body && item.body.content ? item.body.content : "");
  const onlineUrl = normalizeText(item.onlineMeetingUrl || "");
  const locationText = normalizeText(
    item.location && item.location.displayName ? item.location.displayName : ""
  );
  const normalizedMeetingId = normalizeText(meetingId || "");
  const normalizedTeamsLink = normalizeText(teamsLink || "");
  return {
    id: item.id || null,
    subject: item.subject || null,
    start: item.start && item.start.dateTime ? item.start.dateTime : null,
    onlineMeetingUrl: item.onlineMeetingUrl || null,
    location: item.location && item.location.displayName ? item.location.displayName : null,
    hasMeetingId: Boolean(
      normalizedMeetingId &&
        (bodyText.includes(normalizedMeetingId) ||
          onlineUrl.includes(normalizedMeetingId) ||
          locationText.includes(normalizedMeetingId))
    ),
    hasTeamsLink: Boolean(
      normalizedTeamsLink &&
        (bodyText.includes(normalizedTeamsLink) ||
          onlineUrl.includes(normalizedTeamsLink) ||
          locationText.includes(normalizedTeamsLink))
    )
  };
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

function getBaseUrl() {
  if (typeof window !== "undefined" && window.location && window.location.origin) {
    return window.location.origin;
  }
  return DEFAULT_BASE_URL;
}

function getDialogUrl(path) {
  const baseUrl = getBaseUrl();
  return `${baseUrl}/${path}?v=${CACHE_BUSTER}`;
}

function formatEwsError(error) {
  if (!error) {
    return "Unknown error";
  }

  const name = error.name || "EWS error";
  const message = error.message || "No message";
  return `${name}: ${message}`;
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

async function logGraphResponse(label, response) {
  try {
    const text = await response.clone().text();
    logDebug("Graph response", {
      label,
      status: response.status,
      ok: response.ok,
      url: response.url,
      text
    });
  } catch (error) {
    logDebug("Graph response read failed", { label, message: error.message });
  }
}

function logEwsResult(label, result) {
  const errorString = result.error ? formatEwsError(result.error) : null;
  const error = result.error
    ? {
        name: result.error.name,
        message: result.error.message,
        code: result.error.code,
        debugInfo: result.error.debugInfo || null
      }
    : null;
  const value = typeof result.value === "string" ? result.value : "";
  logDebug(`EWS ${label} response`, {
    status: result.status,
    errorString,
    errorJson: result.error ? JSON.stringify(result.error) : null,
    error,
    valueLength: value.length,
    valuePreview: value ? value.slice(0, 2000) : null
  });
}

function exposeAuthHeader(token) {
  if (!token) {
    return "Bearer [missing]";
  }
  return `Bearer ${token}`;
}

// Register the function with Office.
Office.actions.associate("addTeamsLinkToLocation", addTeamsLinkToLocation);
