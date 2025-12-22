/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function addTeamsLinkToLocation(event) {
  const item = Office.context.mailbox.item;

  // Read the meeting body as HTML
  item.body.getAsync(Office.CoercionType.Html, (bodyResult) => {
    if (bodyResult.status !== Office.AsyncResultStatus.Succeeded) {
      item.notificationMessages.replaceAsync("error", {
        type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
        message: "Unable to read meeting body."
      });
      event.completed();
      return;
    }

    const bodyHtml = bodyResult.value;

    // Extract the Teams meeting link
    const teamsRegex = /https:\/\/teams\.microsoft\.com\/l\/meetup-join\/[^\s"<]+/i;
    const match = bodyHtml.match(teamsRegex);

    if (!match) {
      item.notificationMessages.replaceAsync("noLink", {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "No Microsoft Teams meeting link found in this invite."
      });
      event.completed();
      return;
    }

    const teamsLink = match[0];
    const newLocation = `Microsoft Teams Meeting\n${teamsLink}`;

    // Write the Teams link into the Location field
    item.location.setAsync(newLocation, () => {
      item.notificationMessages.replaceAsync("success", {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "Teams meeting link added to Location."
      });
      event.completed();
    });
  });
}

// Register the function with Office.
Office.actions.associate("addTeamsLinkToLocation", addTeamsLinkToLocation);
