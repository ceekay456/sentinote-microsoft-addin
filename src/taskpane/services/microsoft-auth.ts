/* global Office */

import { environment } from "./environment";

// Azure AD app registration details
const MS_CLIENT_ID = environment.microsoft.clientId;
const MS_TENANT_ID = environment.microsoft.tenantId;
const MS_REDIRECT_URI = environment.microsoft.redirectUri;
const MS_SCOPE = environment.microsoft.scope;

// Cache the token in memory
let cachedToken: string | null = null;
let tokenExpiry = 0;

function generateCodeVerifier(): string {
  const array = new Uint8Array(32);
  crypto.getRandomValues(array);
  return Array.from(array, (b) => b.toString(16).padStart(2, "0")).join("");
}

async function generateCodeChallenge(verifier: string): Promise<string> {
  const encoder = new TextEncoder();
  const data = encoder.encode(verifier);
  const digest = await crypto.subtle.digest("SHA-256", data);
  return btoa(String.fromCharCode(...new Uint8Array(digest)))
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/, "");
}

/**
 * Get a Microsoft Graph API access token.
 * Uses cached token if valid, otherwise opens a dialog for auth.
 * Uses login_hint to skip the account picker when possible.
 */
export async function getMicrosoftToken(): Promise<string> {
  // Return cached token if still valid (with 5 min buffer)
  if (cachedToken && Date.now() < tokenExpiry - 5 * 60 * 1000) {
    return cachedToken;
  }

  const codeVerifier = generateCodeVerifier();
  const codeChallenge = await generateCodeChallenge(codeVerifier);

  const state = btoa(
    JSON.stringify({
      clientId: MS_CLIENT_ID,
      tenantId: MS_TENANT_ID,
      codeVerifier,
    })
  );

  const authUrl =
    `https://login.microsoftonline.com/${MS_TENANT_ID}/oauth2/v2.0/authorize` +
    `?client_id=${encodeURIComponent(MS_CLIENT_ID)}` +
    `&response_type=code` +
    `&redirect_uri=${encodeURIComponent(MS_REDIRECT_URI)}` +
    `&scope=${encodeURIComponent(MS_SCOPE)}` +
    `&code_challenge=${encodeURIComponent(codeChallenge)}` +
    `&code_challenge_method=S256` +
    `&state=${encodeURIComponent(state)}`;

  return new Promise<string>((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      authUrl,
      { height: 60, width: 40, promptBeforeOpen: false },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          reject(new Error(`Failed to open sign-in dialog: ${asyncResult.error.message}`));
          return;
        }

        const dialog = asyncResult.value;

        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          (arg: { message: string }) => {
            dialog.close();
            try {
              const message = JSON.parse(arg.message);
              if (message.status === "success" && message.accessToken) {
                cachedToken = message.accessToken;
                tokenExpiry = Date.now() + (message.expiresIn || 3600) * 1000;
                resolve(message.accessToken);
              } else {
                reject(new Error(message.errorDescription || "Microsoft authentication failed"));
              }
            } catch {
              reject(new Error("Invalid response from Microsoft sign-in"));
            }
          }
        );

        dialog.addEventHandler(
          Office.EventType.DialogEventReceived,
          (arg: { error: number }) => {
            if (arg.error === 12006) {
              reject(new Error("Sign-in dialog was closed"));
            }
          }
        );
      }
    );
  });
}
