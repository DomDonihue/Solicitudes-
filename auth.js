const msalConfig = {
  auth: {
    clientId: CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    redirectUri: CONFIG.redirectUri
  },
  cache: { cacheLocation: "sessionStorage", storeAuthStateInCookie: false }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

// Scopes separados por recurso
const graphRequest   = { scopes: ["User.Read", "Mail.Send"] };
const spRequest      = { scopes: [`https://mdonihue.sharepoint.com/AllSites.Write`] };

let currentUser = null;

async function initAuth() {
  await msalInstance.initialize();
  const response = await msalInstance.handleRedirectPromise();
  if (response) msalInstance.setActiveAccount(response.account);
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    msalInstance.setActiveAccount(accounts[0]);
    return true;
  }
  return false;
}

async function login() {
  // Pide ambos scopes en el mismo login para consentimiento único
  await msalInstance.loginRedirect({
    scopes: [...graphRequest.scopes, ...spRequest.scopes],
    prompt: "select_account"
  });
}

async function logout() {
  await msalInstance.logoutRedirect();
}

async function getGraphToken() {
  const account = msalInstance.getActiveAccount();
  if (!account) throw new Error("No active account");
  const res = await msalInstance.acquireTokenSilent({ ...graphRequest, account });
  return res.accessToken;
}

async function getSharePointToken() {
  const account = msalInstance.getActiveAccount();
  if (!account) throw new Error("No active account");
  try {
    const res = await msalInstance.acquireTokenSilent({ ...spRequest, account });
    return res.accessToken;
  } catch {
    const res = await msalInstance.acquireTokenPopup({ ...spRequest, account });
    return res.accessToken;
  }
}

async function getCurrentUser() {
  if (currentUser) return currentUser;
  const token = await getGraphToken();
  const res = await fetch("https://graph.microsoft.com/v1.0/me", {
    headers: { Authorization: `Bearer ${token}` }
  });
  currentUser = await res.json();
  return currentUser;
}
