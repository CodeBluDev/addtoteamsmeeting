/* global Office, msal */

function sendMessage(type, payload) {
  Office.context.ui.messageParent(JSON.stringify({ type, ...payload }));
}

function receiveConfig(message) {
  try {
    return JSON.parse(message);
  } catch (error) {
    return null;
  }
}

Office.onReady(() => {
  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived,
    async (arg) => {
      const config = receiveConfig(arg.message);
      if (!config || !config.clientId) {
        sendMessage("error", { message: "Missing auth configuration." });
        return;
      }

      const msalConfig = {
        auth: {
          clientId: config.clientId,
          authority: config.authority || "https://login.microsoftonline.com/organizations",
          redirectUri: window.location.origin + window.location.pathname
        },
        cache: {
          cacheLocation: "sessionStorage"
        }
      };

      const app = new msal.PublicClientApplication(msalConfig);
      const scopes = config.scopes || ["https://graph.microsoft.com/Calendars.ReadWrite"];

      try {
        const loginResult = await app.loginPopup({ scopes });
        const account = loginResult.account;
        const tokenResult = await app.acquireTokenSilent({
          scopes,
          account
        });
        sendMessage("token", { accessToken: tokenResult.accessToken });
      } catch (error) {
        try {
          const fallback = await app.acquireTokenPopup({ scopes });
          sendMessage("token", { accessToken: fallback.accessToken });
        } catch (popupError) {
          sendMessage("error", { message: popupError.message || "Auth failed." });
        }
      }
    }
  );
});
