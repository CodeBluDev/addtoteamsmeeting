/******/ (function() { // webpackBootstrap
/*!**********************************!*\
  !*** ./src/commands/commands.js ***!
  \**********************************/
function _regenerator() { /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/babel/babel/blob/main/packages/babel-helpers/LICENSE */ var e, t, r = "function" == typeof Symbol ? Symbol : {}, n = r.iterator || "@@iterator", o = r.toStringTag || "@@toStringTag"; function i(r, n, o, i) { var c = n && n.prototype instanceof Generator ? n : Generator, u = Object.create(c.prototype); return _regeneratorDefine2(u, "_invoke", function (r, n, o) { var i, c, u, f = 0, p = o || [], y = !1, G = { p: 0, n: 0, v: e, a: d, f: d.bind(e, 4), d: function d(t, r) { return i = t, c = 0, u = e, G.n = r, a; } }; function d(r, n) { for (c = r, u = n, t = 0; !y && f && !o && t < p.length; t++) { var o, i = p[t], d = G.p, l = i[2]; r > 3 ? (o = l === n) && (u = i[(c = i[4]) ? 5 : (c = 3, 3)], i[4] = i[5] = e) : i[0] <= d && ((o = r < 2 && d < i[1]) ? (c = 0, G.v = n, G.n = i[1]) : d < l && (o = r < 3 || i[0] > n || n > l) && (i[4] = r, i[5] = n, G.n = l, c = 0)); } if (o || r > 1) return a; throw y = !0, n; } return function (o, p, l) { if (f > 1) throw TypeError("Generator is already running"); for (y && 1 === p && d(p, l), c = p, u = l; (t = c < 2 ? e : u) || !y;) { i || (c ? c < 3 ? (c > 1 && (G.n = -1), d(c, u)) : G.n = u : G.v = u); try { if (f = 2, i) { if (c || (o = "next"), t = i[o]) { if (!(t = t.call(i, u))) throw TypeError("iterator result is not an object"); if (!t.done) return t; u = t.value, c < 2 && (c = 0); } else 1 === c && (t = i.return) && t.call(i), c < 2 && (u = TypeError("The iterator does not provide a '" + o + "' method"), c = 1); i = e; } else if ((t = (y = G.n < 0) ? u : r.call(n, G)) !== a) break; } catch (t) { i = e, c = 1, u = t; } finally { f = 1; } } return { value: t, done: y }; }; }(r, o, i), !0), u; } var a = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} t = Object.getPrototypeOf; var c = [][n] ? t(t([][n]())) : (_regeneratorDefine2(t = {}, n, function () { return this; }), t), u = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(c); function f(e) { return Object.setPrototypeOf ? Object.setPrototypeOf(e, GeneratorFunctionPrototype) : (e.__proto__ = GeneratorFunctionPrototype, _regeneratorDefine2(e, o, "GeneratorFunction")), e.prototype = Object.create(u), e; } return GeneratorFunction.prototype = GeneratorFunctionPrototype, _regeneratorDefine2(u, "constructor", GeneratorFunctionPrototype), _regeneratorDefine2(GeneratorFunctionPrototype, "constructor", GeneratorFunction), GeneratorFunction.displayName = "GeneratorFunction", _regeneratorDefine2(GeneratorFunctionPrototype, o, "GeneratorFunction"), _regeneratorDefine2(u), _regeneratorDefine2(u, o, "Generator"), _regeneratorDefine2(u, n, function () { return this; }), _regeneratorDefine2(u, "toString", function () { return "[object Generator]"; }), (_regenerator = function _regenerator() { return { w: i, m: f }; })(); }
function _regeneratorDefine2(e, r, n, t) { var i = Object.defineProperty; try { i({}, "", {}); } catch (e) { i = 0; } _regeneratorDefine2 = function _regeneratorDefine(e, r, n, t) { function o(r, n) { _regeneratorDefine2(e, r, function (e) { return this._invoke(r, n, e); }); } r ? i ? i(e, r, { value: n, enumerable: !t, configurable: !t, writable: !t }) : e[r] = n : (o("next", 0), o("throw", 1), o("return", 2)); }, _regeneratorDefine2(e, r, n, t); }
function _toConsumableArray(r) { return _arrayWithoutHoles(r) || _iterableToArray(r) || _unsupportedIterableToArray(r) || _nonIterableSpread(); }
function _nonIterableSpread() { throw new TypeError("Invalid attempt to spread non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function _unsupportedIterableToArray(r, a) { if (r) { if ("string" == typeof r) return _arrayLikeToArray(r, a); var t = {}.toString.call(r).slice(8, -1); return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray(r, a) : void 0; } }
function _iterableToArray(r) { if ("undefined" != typeof Symbol && null != r[Symbol.iterator] || null != r["@@iterator"]) return Array.from(r); }
function _arrayWithoutHoles(r) { if (Array.isArray(r)) return _arrayLikeToArray(r); }
function _arrayLikeToArray(r, a) { (null == a || a > r.length) && (a = r.length); for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e]; return n; }
function asyncGeneratorStep(n, t, e, r, o, a, c) { try { var i = n[a](c), u = i.value; } catch (n) { return void e(n); } i.done ? t(u) : Promise.resolve(u).then(r, o); }
function _asyncToGenerator(n) { return function () { var t = this, e = arguments; return new Promise(function (r, o) { var a = n.apply(t, e); function _next(n) { asyncGeneratorStep(a, r, o, _next, _throw, "next", n); } function _throw(n) { asyncGeneratorStep(a, r, o, _next, _throw, "throw", n); } _next(void 0); }); }; }
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office, OfficeRuntime */

Office.onReady(function () {
  // If needed, Office.js is ready to be called.
});
var BUILD_TAG = "v1.8.4";
var BUILD_MARKER = "2026-01-01T14:28Z";
var EWS_MESSAGES_NS = "http://schemas.microsoft.com/exchange/services/2006/messages";
var EWS_TYPES_NS = "http://schemas.microsoft.com/exchange/services/2006/types";
var DEBUG_LOGS = true;
var NOTIFICATION_ICON_ID = "Icon.16x16";
var DIALOG_URL = "https://codebludev.github.io/addtoteamsmeeting/create-event.html?v=1.8.4";
var GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
var GRAPH_SEARCH_DAYS = 90;
var AAD_CLIENT_ID = "226fcb0c-fa77-48bb-a20e-70a75ce176fd";
var AAD_AUTHORITY = "https://login.microsoftonline.com/organizations";
var GRAPH_SCOPES = ["https://graph.microsoft.com/Calendars.ReadWrite"];
var AUTH_DIALOG_URL = "https://codebludev.github.io/addtoteamsmeeting/auth.html?v=1.8.4";
var cachedGraphToken = null;
var cachedGraphTokenExpiresAt = 0;

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function addTeamsLinkToLocation(event) {
  var item = Office.context.mailbox.item;
  logDebug("Command invoked", {
    itemId: item.itemId,
    itemType: item.itemType
  });

  // Read the message body as HTML
  item.body.getAsync(Office.CoercionType.Html, function (bodyResult) {
    logDebug("Body getAsync result", {
      status: bodyResult.status
    });
    if (bodyResult.status !== Office.AsyncResultStatus.Succeeded) {
      item.notificationMessages.replaceAsync("error", {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: "Unable to read message body."
      });
      event.completed();
      return;
    }
    var bodyHtml = bodyResult.value;

    // Extract the Teams meeting link
    var teamsLink = extractTeamsLink(bodyHtml);
    logDebug("Teams link match", {
      found: Boolean(teamsLink)
    });
    if (!teamsLink) {
      notifyInfo(item, "No Microsoft Teams meeting link found in this invite.");
      event.completed();
      return;
    }
    logDebug("Teams link extracted", {
      teamsLink: teamsLink
    });
    runEwsHealthCheck(function (ewsOk, ewsMessage) {
      if (!ewsOk) {
        notifyError(item, "EWS health check failed: ".concat(ewsMessage));
        event.completed();
        return;
      }
      findCalendarItemByTeamsLink(teamsLink, function (findError, calendarItem) {
        logDebug("Find by link", {
          error: Boolean(findError),
          found: Boolean(calendarItem)
        });
        if (findError) {
          notifyError(item, "EWS error: ".concat(formatEwsError(findError)));
          notifyInfo(item, "Opening event dialog (calendar search blocked).");
          openCreateEventDialog(item, teamsLink);
          event.completed();
          return;
        }
        if (calendarItem) {
          updateCalendarItemLocation(calendarItem, teamsLink, function (updateError) {
            if (updateError) {
              notifyError(item, "Unable to update the calendar location.");
            } else {
              notifySuccess(item);
            }
            event.completed();
          });
          return;
        }
        getMessageTimeRange(item, function (timeError, timeRange) {
          logDebug("Message time range", {
            error: Boolean(timeError),
            hasRange: Boolean(timeRange)
          });
          if (timeError || !timeRange) {
            notifyInfo(item, "No matching calendar event found.");
            event.completed();
            return;
          }
          findCalendarItemByTimeRange(timeRange, function (timeFindError, timeCalendarItem) {
            logDebug("Find by time", {
              error: Boolean(timeFindError),
              found: Boolean(timeCalendarItem)
            });
            if (timeFindError) {
              notifyError(item, "EWS error: ".concat(formatEwsError(timeFindError)));
              notifyInfo(item, "Opening event dialog (calendar search blocked).");
              openCreateEventDialog(item, teamsLink);
              event.completed();
              return;
            }
            if (!timeCalendarItem) {
              notifyInfo(item, "No matching calendar event found. Opening event dialog.");
              openCreateEventDialog(item, teamsLink);
              event.completed();
              return;
            }
            updateCalendarItemLocation(timeCalendarItem, teamsLink, function (updateError) {
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
  });
}
function notifySuccess(item) {
  item.notificationMessages.replaceAsync("success", {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    icon: NOTIFICATION_ICON_ID,
    persistent: false,
    message: "Teams meeting link added to Location. (".concat(BUILD_TAG, " | ").concat(BUILD_MARKER, ")")
  });
}
function notifyError(item, message) {
  item.notificationMessages.replaceAsync("error", {
    type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
    message: message
  });
}
function notifyInfo(item, message) {
  item.notificationMessages.replaceAsync("info", {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    icon: NOTIFICATION_ICON_ID,
    persistent: false,
    message: message
  });
}
function getUtcWindowStart() {
  var now = new Date();
  now.setDate(now.getDate() - 7);
  return now.toISOString();
}
function getUtcWindowEnd() {
  var now = new Date();
  now.setDate(now.getDate() + 90);
  return now.toISOString();
}
function openCreateEventDialog(item, teamsLink) {
  var baseSubject = item.subject || "Teams meeting";
  var subject = prependBuildTag(baseSubject);
  Office.context.ui.displayDialogAsync(DIALOG_URL, {
    height: 55,
    width: 35,
    displayInIframe: true
  }, function (result) {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      notifyError(item, "Unable to open the event dialog.");
      return;
    }
    var dialog = result.value;
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
      var data;
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
      var start = new Date(data.start);
      var end = new Date(data.end);
      if (Number.isNaN(start.valueOf()) || Number.isNaN(end.valueOf())) {
        notifyError(item, "Invalid date/time from dialog.");
        dialog.close();
        return;
      }
      var subjectWithTag = prependBuildTag(data.subject || baseSubject);
      Office.context.mailbox.displayNewAppointmentForm({
        subject: subjectWithTag,
        location: teamsLink,
        start: start,
        end: end
      });
      dialog.close();
      notifySuccess(item);
    });
    dialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
      notifyInfo(item, "Event dialog closed.");
    });
    dialog.messageChild(JSON.stringify({
      subject: subject,
      teamsLink: teamsLink
    }));
  });
}
function prependBuildTag(subject) {
  var prefix = "[".concat(BUILD_TAG, "] ");
  if (!subject) {
    return prefix.trim();
  }
  if (subject.startsWith(prefix)) {
    return subject;
  }
  return "".concat(prefix).concat(subject);
}
function getMessageTimeRange(item, callback) {
  if (!item.start || !item.end || !item.start.getAsync || !item.end.getAsync) {
    callback(null, null);
    return;
  }
  item.start.getAsync(function (startResult) {
    if (startResult.status !== Office.AsyncResultStatus.Succeeded) {
      callback(startResult.error, null);
      return;
    }
    item.end.getAsync(function (endResult) {
      if (endResult.status !== Office.AsyncResultStatus.Succeeded) {
        callback(endResult.error, null);
        return;
      }
      var start = new Date(startResult.value);
      var end = new Date(endResult.value);
      if (Number.isNaN(start.valueOf()) || Number.isNaN(end.valueOf())) {
        callback(null, null);
        return;
      }
      callback(null, {
        start: start,
        end: end
      });
    });
  });
}
function findCalendarItemByTeamsLink(teamsLink, callback) {
  var escapedLink = escapeXml(teamsLink);
  logDebug("EWS FindItem by link request prepared");
  var request = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\"\n               xmlns:t=\"".concat(EWS_TYPES_NS, "\"\n               xmlns:m=\"").concat(EWS_MESSAGES_NS, "\">\n  <soap:Header>\n    <t:RequestServerVersion Version=\"Exchange2013\" />\n  </soap:Header>\n  <soap:Body>\n    <m:FindItem Traversal=\"Shallow\">\n      <m:ItemShape>\n        <t:BaseShape>IdOnly</t:BaseShape>\n      </m:ItemShape>\n      <m:CalendarView StartDate=\"").concat(getUtcWindowStart(), "\" EndDate=\"").concat(getUtcWindowEnd(), "\" />\n      <m:Restriction>\n        <t:Contains ContainmentMode=\"Substring\" ContainmentComparison=\"IgnoreCase\">\n          <t:FieldURI FieldURI=\"item:Body\" />\n          <t:Constant Value=\"").concat(escapedLink, "\" />\n        </t:Contains>\n      </m:Restriction>\n      <m:ParentFolderIds>\n        <t:DistinguishedFolderId Id=\"calendar\" />\n      </m:ParentFolderIds>\n    </m:FindItem>\n  </soap:Body>\n</soap:Envelope>");
  Office.context.mailbox.makeEwsRequestAsync(request, function (result) {
    logDebug("EWS FindItem by link response", {
      status: result.status,
      error: result.error ? {
        name: result.error.name,
        message: result.error.message
      } : null
    });
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      callback(result.error, null);
      return;
    }
    var calendarItem = parseFirstCalendarItem(result.value);
    callback(null, calendarItem);
  });
}
function runEwsHealthCheck(callback) {
  var request = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\"\n               xmlns:t=\"".concat(EWS_TYPES_NS, "\"\n               xmlns:m=\"").concat(EWS_MESSAGES_NS, "\">\n  <soap:Header>\n    <t:RequestServerVersion Version=\"Exchange2013\" />\n  </soap:Header>\n  <soap:Body>\n    <m:GetFolder>\n      <m:FolderShape>\n        <t:BaseShape>IdOnly</t:BaseShape>\n      </m:FolderShape>\n      <m:FolderIds>\n        <t:DistinguishedFolderId Id=\"calendar\" />\n      </m:FolderIds>\n    </m:GetFolder>\n  </soap:Body>\n</soap:Envelope>");
  Office.context.mailbox.makeEwsRequestAsync(request, function (result) {
    logDebug("EWS health check response", {
      status: result.status,
      error: result.error ? {
        name: result.error.name,
        message: result.error.message
      } : null
    });
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      callback(false, formatEwsError(result.error));
      return;
    }
    var xmlDoc = parseXml(result.value);
    if (!xmlDoc || !isEwsResponseSuccess(xmlDoc)) {
      callback(false, "EWS GetFolder failed.");
      return;
    }
    callback(true, "OK");
  });
}
function findCalendarItemByTimeRange(timeRange, callback) {
  var startIso = timeRange.start.toISOString();
  var endIso = timeRange.end.toISOString();
  logDebug("EWS FindItem by time request prepared", {
    startIso: startIso,
    endIso: endIso
  });
  var request = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\"\n               xmlns:t=\"".concat(EWS_TYPES_NS, "\"\n               xmlns:m=\"").concat(EWS_MESSAGES_NS, "\">\n  <soap:Header>\n    <t:RequestServerVersion Version=\"Exchange2013\" />\n  </soap:Header>\n  <soap:Body>\n    <m:FindItem Traversal=\"Shallow\">\n      <m:ItemShape>\n        <t:BaseShape>IdOnly</t:BaseShape>\n        <t:AdditionalProperties>\n          <t:FieldURI FieldURI=\"calendar:Start\" />\n          <t:FieldURI FieldURI=\"calendar:End\" />\n        </t:AdditionalProperties>\n      </m:ItemShape>\n      <m:CalendarView StartDate=\"").concat(startIso, "\" EndDate=\"").concat(endIso, "\" />\n      <m:ParentFolderIds>\n        <t:DistinguishedFolderId Id=\"calendar\" />\n      </m:ParentFolderIds>\n    </m:FindItem>\n  </soap:Body>\n</soap:Envelope>");
  Office.context.mailbox.makeEwsRequestAsync(request, function (result) {
    logDebug("EWS FindItem by time response", {
      status: result.status,
      error: result.error ? {
        name: result.error.name,
        message: result.error.message
      } : null
    });
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      callback(result.error, null);
      return;
    }
    var calendarItem = parseCalendarItemByTime(result.value, timeRange);
    callback(null, calendarItem);
  });
}
function updateCalendarItemLocation(calendarItem, location, callback) {
  var escapedLocation = escapeXml(location);
  logDebug("EWS UpdateItem request prepared", {
    itemId: calendarItem.id
  });
  var request = "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\"\n               xmlns:t=\"".concat(EWS_TYPES_NS, "\"\n               xmlns:m=\"").concat(EWS_MESSAGES_NS, "\">\n  <soap:Header>\n    <t:RequestServerVersion Version=\"Exchange2013\" />\n  </soap:Header>\n  <soap:Body>\n    <m:UpdateItem ConflictResolution=\"AlwaysOverwrite\" SendMeetingInvitationsOrCancellations=\"SendToNone\">\n      <m:ItemChanges>\n        <t:ItemChange>\n          <t:ItemId Id=\"").concat(calendarItem.id, "\" ChangeKey=\"").concat(calendarItem.changeKey, "\" />\n          <t:Updates>\n            <t:SetItemField>\n              <t:FieldURI FieldURI=\"calendar:Location\" />\n              <t:CalendarItem>\n                <t:Location>").concat(escapedLocation, "</t:Location>\n              </t:CalendarItem>\n            </t:SetItemField>\n          </t:Updates>\n        </t:ItemChange>\n      </m:ItemChanges>\n    </m:UpdateItem>\n  </soap:Body>\n</soap:Envelope>");
  Office.context.mailbox.makeEwsRequestAsync(request, function (result) {
    logDebug("EWS UpdateItem response", {
      status: result.status,
      error: result.error ? {
        name: result.error.name,
        message: result.error.message
      } : null
    });
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      callback(result.error);
      return;
    }
    var xmlDoc = parseXml(result.value);
    if (!xmlDoc || !isEwsResponseSuccess(xmlDoc)) {
      callback(new Error("EWS UpdateItem failed."));
      return;
    }
    callback(null);
  });
}
function parseFirstCalendarItem(responseText) {
  var xmlDoc = parseXml(responseText);
  if (!xmlDoc) {
    return null;
  }
  var itemId = xmlDoc.getElementsByTagNameNS(EWS_TYPES_NS, "ItemId")[0];
  if (!itemId) {
    return null;
  }
  return {
    id: itemId.getAttribute("Id"),
    changeKey: itemId.getAttribute("ChangeKey")
  };
}
function parseCalendarItemByTime(responseText, timeRange) {
  var xmlDoc = parseXml(responseText);
  if (!xmlDoc) {
    return null;
  }
  var calendarItems = xmlDoc.getElementsByTagNameNS(EWS_TYPES_NS, "CalendarItem");
  for (var i = 0; i < calendarItems.length; i += 1) {
    var calendarItem = calendarItems[i];
    var startNode = calendarItem.getElementsByTagNameNS(EWS_TYPES_NS, "Start")[0];
    var endNode = calendarItem.getElementsByTagNameNS(EWS_TYPES_NS, "End")[0];
    var itemId = calendarItem.getElementsByTagNameNS(EWS_TYPES_NS, "ItemId")[0];
    if (!startNode || !endNode || !itemId) {
      continue;
    }
    var start = new Date(startNode.textContent);
    var end = new Date(endNode.textContent);
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
  var toleranceMs = 60000;
  return Math.abs(start - timeRange.start) <= toleranceMs && Math.abs(end - timeRange.end) <= toleranceMs;
}
function parseXml(xmlString) {
  if (!xmlString) {
    return null;
  }
  return new DOMParser().parseFromString(xmlString, "text/xml");
}
function isEwsResponseSuccess(xmlDoc) {
  var responseCode = xmlDoc.getElementsByTagNameNS(EWS_MESSAGES_NS, "ResponseCode")[0];
  return responseCode && responseCode.textContent === "NoError";
}
function escapeXml(value) {
  return String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}
function extractMeetingId(bodyHtml, teamsLink) {
  if (teamsLink) {
    var linkMatch = teamsLink.match(/19:meeting_[^/?"'\\s<>]+/i);
    if (linkMatch) {
      return linkMatch[0];
    }
  }
  var combined = "".concat(bodyHtml, " ").concat(teamsLink || "");
  var decoded = decodeHtmlEntities(decodeLink(combined));
  var match = decoded.match(/19:meeting_[^/?"'\\s<>]+/i);
  if (match) {
    return match[0];
  }
  var encodedMatch = combined.match(/19%3Ameeting_[^"'\\s<>%]+/i);
  if (encodedMatch) {
    return decodeLink(encodedMatch[0]);
  }
  return null;
}
function findCalendarEventByGraph(_x, _x2) {
  return _findCalendarEventByGraph.apply(this, arguments);
}
function _findCalendarEventByGraph() {
  _findCalendarEventByGraph = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee(teamsLink, meetingId) {
    var token, start, end, url, response, text, data, items, i;
    return _regenerator().w(function (_context) {
      while (1) switch (_context.n) {
        case 0:
          _context.n = 1;
          return getGraphAccessToken();
        case 1:
          token = _context.v;
          start = new Date();
          end = new Date(start.getTime() + GRAPH_SEARCH_DAYS * 24 * 60 * 60 * 1000);
          url = "".concat(GRAPH_BASE_URL, "/me/calendarView?startDateTime=").concat(encodeURIComponent(start.toISOString()), "&endDateTime=").concat(encodeURIComponent(end.toISOString())) + "&$select=id,subject,body,location,onlineMeetingUrl,start,end";
        case 2:
          if (!url) {
            _context.n = 11;
            break;
          }
          logDebug("Graph calendarView request", {
            method: "GET",
            url: url,
            headers: {
              Authorization: exposeAuthHeader(token),
              Prefer: 'outlook.body-content-type="text"'
            }
          });
          _context.n = 3;
          return fetch(url, {
            headers: {
              Authorization: "Bearer ".concat(token),
              Prefer: 'outlook.body-content-type="text"'
            }
          });
        case 3:
          response = _context.v;
          _context.n = 4;
          return logGraphResponse("calendarView", response);
        case 4:
          if (response.ok) {
            _context.n = 6;
            break;
          }
          _context.n = 5;
          return response.text();
        case 5:
          text = _context.v;
          throw new Error("Graph calendarView failed: ".concat(response.status, " ").concat(text));
        case 6:
          _context.n = 7;
          return response.json();
        case 7:
          data = _context.v;
          items = data.value || [];
          i = 0;
        case 8:
          if (!(i < items.length)) {
            _context.n = 10;
            break;
          }
          if (!eventMatchesTeams(items[i], teamsLink, meetingId)) {
            _context.n = 9;
            break;
          }
          return _context.a(2, items[i].id);
        case 9:
          i += 1;
          _context.n = 8;
          break;
        case 10:
          url = data["@odata.nextLink"] || null;
          _context.n = 2;
          break;
        case 11:
          return _context.a(2, null);
      }
    }, _callee);
  }));
  return _findCalendarEventByGraph.apply(this, arguments);
}
function updateCalendarEventLocationGraph(_x3, _x4) {
  return _updateCalendarEventLocationGraph.apply(this, arguments);
}
function _updateCalendarEventLocationGraph() {
  _updateCalendarEventLocationGraph = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee2(eventId, teamsLink) {
    var token, body, response, text;
    return _regenerator().w(function (_context2) {
      while (1) switch (_context2.n) {
        case 0:
          _context2.n = 1;
          return getGraphAccessToken();
        case 1:
          token = _context2.v;
          body = {
            location: {
              displayName: teamsLink
            }
          };
          logDebug("Graph update request", {
            method: "PATCH",
            url: "".concat(GRAPH_BASE_URL, "/me/events/").concat(eventId),
            headers: {
              Authorization: exposeAuthHeader(token),
              "Content-Type": "application/json"
            },
            body: body
          });
          _context2.n = 2;
          return fetch("".concat(GRAPH_BASE_URL, "/me/events/").concat(eventId), {
            method: "PATCH",
            headers: {
              Authorization: "Bearer ".concat(token),
              "Content-Type": "application/json"
            },
            body: JSON.stringify(body)
          });
        case 2:
          response = _context2.v;
          _context2.n = 3;
          return logGraphResponse("updateEvent", response);
        case 3:
          if (response.ok) {
            _context2.n = 5;
            break;
          }
          _context2.n = 4;
          return response.text();
        case 4:
          text = _context2.v;
          throw new Error("Graph update failed: ".concat(response.status, " ").concat(text));
        case 5:
          return _context2.a(2);
      }
    }, _callee2);
  }));
  return _updateCalendarEventLocationGraph.apply(this, arguments);
}
function getGraphAccessToken() {
  return _getGraphAccessToken.apply(this, arguments);
}
function _getGraphAccessToken() {
  _getGraphAccessToken = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee3() {
    var _token, token, _t;
    return _regenerator().w(function (_context3) {
      while (1) switch (_context3.p = _context3.n) {
        case 0:
          if (!(cachedGraphToken && Date.now() < cachedGraphTokenExpiresAt)) {
            _context3.n = 1;
            break;
          }
          return _context3.a(2, cachedGraphToken);
        case 1:
          if (!(OfficeRuntime && OfficeRuntime.auth && OfficeRuntime.auth.getAccessToken)) {
            _context3.n = 5;
            break;
          }
          _context3.p = 2;
          _context3.n = 3;
          return OfficeRuntime.auth.getAccessToken({
            allowSignInPrompt: true,
            allowConsentPrompt: true,
            forMSGraphAccess: true
          });
        case 3:
          _token = _context3.v;
          logDebug("Graph token acquired (OfficeRuntime)");
          cacheGraphToken(_token, 50);
          return _context3.a(2, _token);
        case 4:
          _context3.p = 4;
          _t = _context3.v;
          logDebug("OfficeRuntime auth failed", {
            message: _t.message
          });
        case 5:
          _context3.n = 6;
          return getGraphAccessTokenViaDialog();
        case 6:
          token = _context3.v;
          logDebug("Graph token acquired (dialog)");
          cacheGraphToken(token, 45);
          return _context3.a(2, token);
      }
    }, _callee3, null, [[2, 4]]);
  }));
  return _getGraphAccessToken.apply(this, arguments);
}
function eventMatchesTeams(event, teamsLink, meetingId) {
  var bodyText = event.body && event.body.content ? event.body.content : "";
  var onlineUrl = event.onlineMeetingUrl || "";
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
function cacheGraphToken(token, minutes) {
  cachedGraphToken = token;
  cachedGraphTokenExpiresAt = Date.now() + minutes * 60 * 1000;
}
function getGraphAccessTokenViaDialog() {
  return new Promise(function (resolve, reject) {
    Office.context.ui.displayDialogAsync(AUTH_DIALOG_URL, {
      height: 60,
      width: 40,
      displayInIframe: true
    }, function (result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        var message = result.error && result.error.message ? result.error.message : "Unknown dialog error.";
        logDebug("Auth dialog open failed", {
          code: result.error && result.error.code,
          message: message
        });
        reject(new Error("Unable to open auth dialog: ".concat(message)));
        return;
      }
      var dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
        var data;
        try {
          data = JSON.parse(arg.message);
        } catch (parseError) {
          dialog.close();
          reject(new Error("Invalid auth dialog response."));
          return;
        }
        if (data.type === "ready") {
          dialog.messageChild(JSON.stringify({
            clientId: AAD_CLIENT_ID,
            authority: AAD_AUTHORITY,
            scopes: GRAPH_SCOPES
          }));
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
      dialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
        reject(new Error("Auth dialog closed."));
      });
    });
  });
}
function extractTeamsLink(bodyHtml) {
  var teamsRegex = /https:\/\/teams\.microsoft\.com\/l\/meetup-join\/[^\s"<]+/i;
  var safeLinksRegex = /https:\/\/[^\/]+\.safelinks\.protection\.outlook\.com\/[^\s"<]+/i;
  var akaTeamsRegex = /https:\/\/aka\.ms\/[^\s"<]*teams[^\s"<]*/i;
  var directLink = findTeamsLinkInText(bodyHtml);
  if (directLink) {
    return directLink;
  }
  var urls = bodyHtml.match(/https?:\/\/[^\s"'<>]+/gi) || [];
  for (var i = 0; i < urls.length; i += 1) {
    var rawUrl = urls[i];
    var cleanedUrl = rawUrl.replace(/&amp;/g, "&");
    if (teamsRegex.test(cleanedUrl)) {
      return decodeLink(cleanedUrl);
    }
    if (safeLinksRegex.test(cleanedUrl)) {
      var extracted = extractSafeLinkTarget(cleanedUrl);
      if (extracted) {
        var decoded = decodeLink(extracted);
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
  var candidates = [];
  var rawMatches = text.match(/https:\/\/teams\.microsoft\.com\/l\/meetup-join\/[^\s"'<>]+/gi);
  if (rawMatches) {
    candidates.push.apply(candidates, _toConsumableArray(rawMatches));
  }
  var decodedHtml = decodeHtmlEntities(text);
  var decodedMatches = decodedHtml.match(/https:\/\/teams\.microsoft\.com\/l\/meetup-join\/[^\s"'<>]+/gi);
  if (decodedMatches) {
    candidates.push.apply(candidates, _toConsumableArray(decodedMatches));
  }
  for (var i = 0; i < candidates.length; i += 1) {
    var cleaned = decodeLink(candidates[i]);
    if (cleaned) {
      return cleaned;
    }
  }
  return null;
}
function decodeLink(url) {
  var cleaned = url.replace(/&amp;/g, "&");
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
  return text.replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, "\"").replace(/&#39;/g, "'");
}
function extractSafeLinkTarget(safeLinkUrl) {
  try {
    var url = new URL(safeLinkUrl);
    var target = url.searchParams.get("url");
    if (!target) {
      return null;
    }
    return decodeURIComponent(target);
  } catch (error) {
    return null;
  }
}
function formatEwsError(error) {
  if (!error) {
    return "Unknown error";
  }
  var name = error.name || "EWS error";
  var message = error.message || "No message";
  return "".concat(name, ": ").concat(message);
}
function logDebug(message, data) {
  if (!DEBUG_LOGS) {
    return;
  }
  if (data) {
    // eslint-disable-next-line no-console
    console.log("[AddTeamsLink] ".concat(message), data);
  } else {
    // eslint-disable-next-line no-console
    console.log("[AddTeamsLink] ".concat(message));
  }
}
function logGraphResponse(_x5, _x6) {
  return _logGraphResponse.apply(this, arguments);
}
function _logGraphResponse() {
  _logGraphResponse = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee4(label, response) {
    var text, _t2;
    return _regenerator().w(function (_context4) {
      while (1) switch (_context4.p = _context4.n) {
        case 0:
          _context4.p = 0;
          _context4.n = 1;
          return response.clone().text();
        case 1:
          text = _context4.v;
          logDebug("Graph response", {
            label: label,
            status: response.status,
            ok: response.ok,
            url: response.url,
            text: text
          });
          _context4.n = 3;
          break;
        case 2:
          _context4.p = 2;
          _t2 = _context4.v;
          logDebug("Graph response read failed", {
            label: label,
            message: _t2.message
          });
        case 3:
          return _context4.a(2);
      }
    }, _callee4, null, [[0, 2]]);
  }));
  return _logGraphResponse.apply(this, arguments);
}
function exposeAuthHeader(token) {
  if (!token) {
    return "Bearer [missing]";
  }
  return "Bearer ".concat(token);
}

// Register the function with Office.
Office.actions.associate("addTeamsLinkToLocation", addTeamsLinkToLocation);
/******/ })()
;
//# sourceMappingURL=commands.js.map