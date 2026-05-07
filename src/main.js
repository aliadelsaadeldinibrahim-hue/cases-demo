import * as msal from "@azure/msal-browser";

const msalConfig = {
  auth: {
    clientId: "PASTE_CLIENT_ID_HERE",
    authority: "https://login.microsoftonline.com/PASTE_TENANT_ID_HERE",
    redirectUri: window.location.origin,
  }
};

const dynamicsUrl = "https://PASTE_YOUR_ORG.crm4.dynamics.com";
const viewId = "PASTE_VIEW_ID_HERE";
const scopes = [`${dynamicsUrl}/user_impersonation`];

const msalInstance = new msal.PublicClientApplication(msalConfig);
await msalInstance.initialize();
await msalInstance.handleRedirectPromise();

const loginBtn = document.getElementById("login-btn");
const message = document.getElementById("message");
const casesDiv = document.getElementById("cases");

loginBtn.addEventListener("click", async () => {
  try {
    await msalInstance.loginPopup({ scopes });
    loadCases();
  } catch (e) {
    message.textContent = "Login failed: " + e.message;
  }
});

async function loadCases() {
  message.textContent = "Loading cases...";
  loginBtn.style.display = "none";

  const account = msalInstance.getAllAccounts()[0];
  const tokenRes = await msalInstance.acquireTokenSilent({ scopes, account });

  const res = await fetch(
    `${dynamicsUrl}/api/data/v9.2/incidents?savedQuery=${viewId}&$select=title,ticketnumber,statecode,createdon`,
    {
      headers: {
        Authorization: `Bearer ${tokenRes.accessToken}`,
        Accept: "application/json",
        "OData-MaxVersion": "4.0",
        "OData-Version": "4.0"
      }
    }
  );

  const data = await res.json();
  message.textContent = `${data.value.length} cases loaded`;

  casesDiv.innerHTML = data.value.map(c => `
    <div class="case">
      <strong>${c.ticketnumber}</strong> — ${c.title}<br/>
      <small class="${c.statecode === 0 ? 'status-active' : 'status-resolved'}">
        ${c.statecode === 0 ? "Active" : "Resolved"}
      </small>
      &nbsp;|&nbsp;
      <small>Created: ${new Date(c.createdon).toLocaleDateString()}</small>
    </div>
  `).join("");
}