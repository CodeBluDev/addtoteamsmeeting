/* global Office, msal */

function sendMessage(type, payload) {
  Office.context.ui.messageParent(JSON.stringify({ type, ...payload }));
}

function setStatus(message) {
  const el = document.getElementById("authStatus");
  if (el) {
    el.textContent = message;
  }
}

function receiveConfig(message) {
  try {
    return JSON.parse(message);
  } catch (error) {
    return null;
  }
}

Office.onReady(() => {
  setStatus("Auth dialog ready");
  // eslint-disable-next-line no-console
  console.log("[AuthDialog] ready");

  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived,
    async (arg) => {
      const config = receiveConfig(arg.message);
      if (!config || !config.clientId) {
        setStatus("Missing auth config.");
        sendMessage("error", { message: "Missing auth configuration." });
        return;
      }

      setStatus("Starting sign-in…");
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
        const loginResult = await app.loginPopup({
          scopes,
          prompt: "login"
        });
        const account = loginResult.account;
        const tokenResult = await app.acquireTokenSilent({
          scopes,
          account
        });
        setStatus("Token acquired.");
        sendMessage("token", { accessToken: tokenResult.accessToken });
      } catch (error) {
        setStatus(`Auth error: ${error.message || "login failed"}`);
        try {
          const fallback = await app.acquireTokenPopup({
            scopes,
            prompt: "login"
          });
          setStatus("Token acquired.");
          sendMessage("token", { accessToken: fallback.accessToken });
        } catch (popupError) {
          setStatus(`Auth error: ${popupError.message || "login failed"}`);
          sendMessage("error", { message: popupError.message || "Auth failed." });
        }
      }
    }
  );

  sendMessage("ready", {});
});
