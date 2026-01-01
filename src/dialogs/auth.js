/* global Office, msal */

const STORAGE_KEY = "addTeamsAuthConfig";
let authInProgress = false;
let currentConfig = null;

function sendMessage(type, payload) {
  Office.context.ui.messageParent(JSON.stringify({ type, ...payload }));
}

function sendLog(level, message, data) {
  sendMessage("log", { level, message, data });
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

function saveConfig(config) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(config));
  } catch (error) {
    // Ignore storage errors in constrained browsers.
  }
}

function loadConfig() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return null;
    }
    return JSON.parse(raw);
  } catch (error) {
    return null;
  }
}

async function startAuth(config) {
  if (!config || !config.clientId) {
    return;
  }
  if (authInProgress) {
    sendLog("info", "Auth already in progress, skipping");
    return;
  }
  authInProgress = true;
  setStatus("Starting sign-in…");
  sendLog("info", "Auth config received", {
    clientId: config.clientId,
    authority: config.authority,
    scopes: config.scopes
  });
  const app = createMsalApp(config);
  const scopes = config.scopes || ["https://graph.microsoft.com/Calendars.ReadWrite"];

  try {
    sendLog("info", "Starting loginPopup");
    const loginResult = await app.loginPopup({
      scopes,
      prompt: "login"
    });
    const loginAccount = loginResult.account;
    sendLog("info", "loginPopup succeeded", {
      username: loginAccount && loginAccount.username ? loginAccount.username : null
    });
    const tokenResult = await app.acquireTokenSilent({
      scopes,
      account: loginAccount
    });
    setStatus("Token acquired.");
    sendMessage("token", { accessToken: tokenResult.accessToken });
    sendLog("info", "Token acquired via acquireTokenSilent");
  } catch (error) {
    setStatus(`Auth error: ${error.message || "login failed"}`);
    sendLog("error", "loginPopup failed", {
      name: error.name,
      message: error.message,
      errorCode: error.errorCode
    });
    try {
      sendLog("info", "Starting acquireTokenPopup fallback");
      const fallback = await app.acquireTokenPopup({
        scopes,
        prompt: "login"
      });
      setStatus("Token acquired.");
      sendMessage("token", { accessToken: fallback.accessToken });
      sendLog("info", "Token acquired via acquireTokenPopup");
    } catch (popupError) {
      setStatus(`Auth error: ${popupError.message || "login failed"}`);
      sendMessage("error", { message: popupError.message || "Auth failed." });
      sendLog("error", "acquireTokenPopup failed", {
        name: popupError.name,
        message: popupError.message,
        errorCode: popupError.errorCode
      });
    }
  } finally {
    authInProgress = false;
  }
}

function createMsalApp(config) {
  const msalConfig = {
    auth: {
      clientId: config.clientId,
      authority: config.authority || "https://login.microsoftonline.com/organizations",
      redirectUri: window.location.origin + window.location.pathname
    },
    cache: {
      cacheLocation: "localStorage",
      storeAuthStateInCookie: true
    }
  };

  return new msal.PublicClientApplication(msalConfig);
}

async function trySilentAuth(config) {
  const app = createMsalApp(config);
  const scopes = config.scopes || ["https://graph.microsoft.com/Calendars.ReadWrite"];
  const account = app.getAllAccounts()[0];
  if (!account) {
    return false;
  }

  try {
    sendLog("info", "Account found, acquiring token silently");
    const tokenResult = await app.acquireTokenSilent({
      scopes,
      account
    });
    setStatus("Token acquired.");
    sendMessage("token", { accessToken: tokenResult.accessToken });
    sendLog("info", "Token acquired via acquireTokenSilent");
    return true;
  } catch (error) {
    sendLog("error", "Silent token failed", {
      name: error.name,
      message: error.message,
      errorCode: error.errorCode
    });
    return false;
  }
}

function showStartButton() {
  const button = document.getElementById("authStartButton");
  if (button) {
    button.style.display = "inline-block";
  }
}

Office.onReady(() => {
  setStatus("Auth dialog ready");
  // eslint-disable-next-line no-console
  console.log("[AuthDialog] ready");
  sendLog("info", "Auth dialog ready", { origin: window.location.origin });

  window.addEventListener("error", (event) => {
    sendLog("error", "Auth dialog error", {
      message: event.message,
      filename: event.filename,
      lineno: event.lineno,
      colno: event.colno
    });
  });

  window.addEventListener("unhandledrejection", (event) => {
    const reason = event.reason || {};
    sendLog("error", "Auth dialog unhandled rejection", {
      message: reason.message || String(reason)
    });
  });

  Office.context.ui.addHandlerAsync(
    Office.EventType.DialogParentMessageReceived,
    async (arg) => {
      const config = receiveConfig(arg.message);
      if (!config || !config.clientId) {
        setStatus("Missing auth config.");
        sendMessage("error", { message: "Missing auth configuration." });
        sendLog("error", "Missing auth configuration.");
        return;
      }

      saveConfig(config);
      currentConfig = config;
      const silentOk = await trySilentAuth(config);
      if (!silentOk) {
        setStatus("Ready to sign in.");
        showStartButton();
      }
    }
  );

  sendMessage("ready", {});

  const storedConfig = loadConfig();
  if (storedConfig && storedConfig.clientId) {
    currentConfig = storedConfig;
    trySilentAuth(storedConfig).then((silentOk) => {
      if (!silentOk) {
        setStatus("Ready to sign in.");
        showStartButton();
      }
    });
  }

  const button = document.getElementById("authStartButton");
  if (button) {
    button.addEventListener("click", () => {
      sendLog("info", "User clicked sign-in button");
      showStartButton();
      startAuth(currentConfig);
    });
  }
});
