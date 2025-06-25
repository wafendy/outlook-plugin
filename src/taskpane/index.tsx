import { createNestablePublicClientApplication, IPublicClientApplication } from "@azure/msal-browser";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";

/* global document, Office, module, require, HTMLElement */

const title = "SmartCompose Task Pane Add-in";

const rootElement: HTMLElement | null = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

let pca: IPublicClientApplication = undefined;
Office.onReady(async (info) => {
  if (info.host) {
    // Initialize the public client application
    pca = await createNestablePublicClientApplication({
      auth: {
        clientId: "91093319-ecda-4b06-9973-be6d731cae8d",
        authority: "https://login.microsoftonline.com/common",
      },
    });

    run();
  }
});
/* Render application after Office initializes */
Office.onReady(() => {
  console.log("======> Office.onReady");
  root?.render(
    <FluentProvider theme={webLightTheme}>
      <App title={title} />
    </FluentProvider>
  );
});

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root?.render(NextApp);
  });
}

async function run() {
  // Specify minimum scopes needed for the access token.
  const tokenRequest = {
    scopes: ["Files.Read", "User.Read", "openid", "profile"],
  };
  let accessToken = null;

  // 1: Call acquireTokenSilent.
  try {
    console.log("Trying to acquire token silently...");
    const userAccount = await pca.acquireTokenSilent(tokenRequest);
    console.log("Acquired token silently.");
    accessToken = userAccount.accessToken;
  } catch (error) {
    console.log(`Unable to acquire token silently: ${error}`);
  }

  // 2: Call acquireTokenPopup.
  if (accessToken === null) {
    // Acquire token silent failure. Send an interactive request via popup.
    try {
      console.log("Trying to acquire token interactively...");
      const userAccount = await pca.acquireTokenPopup(tokenRequest);
      console.log("Acquired token interactively.");
      accessToken = userAccount.accessToken;
    } catch (popupError) {
      // Acquire token interactive failure.
      console.log(`Unable to acquire token interactively: ${popupError}`);
    }
  }

  // 3: Log error if token still null.
  // Log error if both silent and popup requests failed.
  if (accessToken === null) {
    console.error(`Unable to acquire access token.`);
    return;
  }

  // 4: Find out who am i
  var headers = new Headers();
  var bearer = "Bearer " + accessToken;
  headers.append("Authorization", bearer);
  var options = {
    method: "GET",
    headers: headers,
  };
  var graphEndpoint = "https://graph.microsoft.com/v1.0/me";

  fetch(graphEndpoint, options)
    .then((resp) => resp.json())
    .then((data) => {
      console.log("User information:", data);
      console.log("Display Name:", data.displayName);
      console.log("Email:", data.email);
      console.log("Job Title:", data.jobTitle);
    })
    .catch((error) => {
      console.error("Error fetching user data:", error);
    });

  // 5: Call the Microsoft Graph API.
  // Call the Microsoft Graph API with the access token.
  const response = await fetch(`https://graph.microsoft.com/v1.0/me/drive/root/children?$select=name&$top=10`, {
    headers: { Authorization: accessToken },
  });

  if (response.ok) {
    // Write file names to the console.
    const data = await response.json();
    const names = data.value.map((item) => item.name);

    // Be sure the taskpane.html has an element with Id = item-subject.
    const label = document.getElementById("item-subject");

    // Write file names to task pane and the console.
    const nameText = names.join(", ");
    if (label) label.textContent = nameText;
    console.log(nameText);
  } else {
    const errorText = await response.text();
    console.error("Microsoft Graph call failed - error text: " + errorText);
  }
}
