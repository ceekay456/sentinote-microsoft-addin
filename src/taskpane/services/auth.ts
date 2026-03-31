import { environment } from "./environment";

/* global Office */

export interface AuthUser {
  email: string;
  tenantId: string;
  sub: string;
  groups: string[];
  isAdmin: boolean;
}

export interface AuthTokens {
  idToken: string;
  accessToken: string;
  refreshToken: string;
  expiresAt: number;
}

type AuthChangeCallback = (authenticated: boolean) => void;

let currentTokens: AuthTokens | null = null;
let currentUser: AuthUser | null = null;
const listeners: AuthChangeCallback[] = [];

function decodeJwtPayload(token: string): Record<string, unknown> {
  const base64 = token.split(".")[1].replace(/-/g, "+").replace(/_/g, "/");
  const json = atob(base64);
  return JSON.parse(json);
}

function extractUser(idToken: string): AuthUser {
  const claims = decodeJwtPayload(idToken);
  const groups = (claims["cognito:groups"] as string[]) || [];
  return {
    email: claims.email as string,
    tenantId: claims["custom:tenantId"] as string,
    sub: claims.sub as string,
    groups,
    isAdmin: groups.includes("admin"),
  };
}

function notifyListeners() {
  const authenticated = currentTokens !== null;
  listeners.forEach((cb) => cb(authenticated));
}

export function onAuthChange(callback: AuthChangeCallback): () => void {
  listeners.push(callback);
  return () => {
    const index = listeners.indexOf(callback);
    if (index >= 0) listeners.splice(index, 1);
  };
}

export function isAuthenticated(): boolean {
  if (!currentTokens) return false;
  if (Date.now() >= currentTokens.expiresAt) {
    currentTokens = null;
    currentUser = null;
    return false;
  }
  return true;
}

export function getIdToken(): string | null {
  if (!isAuthenticated()) return null;
  return currentTokens!.idToken;
}

export function getUser(): AuthUser | null {
  if (!isAuthenticated()) return null;
  return currentUser;
}

export function signIn(): Promise<void> {
  return new Promise((resolve, reject) => {
    const { domain, clientId, redirectUri } = environment.cognito;

    // Store config in localStorage so callback.html can read it
    localStorage.setItem(
      "rc_auth_config",
      JSON.stringify({ domain, clientId })
    );

    const loginUrl =
      `https://${domain}/login?` +
      `response_type=code&` +
      `client_id=${encodeURIComponent(clientId)}&` +
      `redirect_uri=${encodeURIComponent(redirectUri)}&` +
      `scope=openid+email+profile`;

    Office.context.ui.displayDialogAsync(
      loginUrl,
      { height: 60, width: 40, promptBeforeOpen: false },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(new Error(result.error.message));
          return;
        }

        const dialog = result.value;

        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          (arg: { message: string } | { error: number }) => {
            dialog.close();

            if ("error" in arg) {
              reject(new Error(`Dialog error: ${arg.error}`));
              return;
            }

            const message = JSON.parse(arg.message);
            if (message.status === "error") {
              reject(new Error(message.errorDescription || message.error));
              return;
            }

            currentTokens = {
              idToken: message.idToken,
              accessToken: message.accessToken,
              refreshToken: message.refreshToken,
              expiresAt: Date.now() + (message.expiresIn || 3600) * 1000,
            };
            currentUser = extractUser(message.idToken);
            notifyListeners();
            resolve();
          }
        );

        dialog.addEventHandler(
          Office.EventType.DialogEventReceived,
          (arg: { error: number }) => {
            // Dialog was closed by user (12006) or navigated away
            if (arg.error === 12006) {
              reject(new Error("Sign-in dialog was closed"));
            }
          }
        );
      }
    );
  });
}

export function signOut(): void {
  currentTokens = null;
  currentUser = null;
  localStorage.removeItem("rc_auth_config");
  notifyListeners();
}
