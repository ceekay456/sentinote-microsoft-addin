import { environment } from "./environment";
import { getIdToken } from "./auth";

/**
 * Make an authenticated API call (uses the Cognito ID token from the session).
 */
async function apiCall<T>(path: string, options: { method?: string; params?: Record<string, string>; body?: unknown } = {}): Promise<T> {
  const { method = "GET", params, body } = options;
  const token = getIdToken();

  let url = `${environment.apiUrl}${path}`;
  if (params) {
    url += `?${new URLSearchParams(params).toString()}`;
  }

  const headers: Record<string, string> = {
    "Content-Type": "application/json",
  };
  if (token) {
    headers.Authorization = token;
  }

  const response = await fetch(url, {
    method,
    headers,
    body: body ? JSON.stringify(body) : undefined,
  });

  if (!response.ok) {
    const errorBody = await response.text();
    throw new Error(errorBody || response.statusText);
  }

  return response.json();
}

// Session bootstrap (unauthenticated)

export interface AddinSession {
  checkoutId: string;
  fileKey: string;
  fileName: string;
  studyId: string;
  tenantId: string;
  filePath: string;
  checkedOutByEmail: string;
  checkedOutAt: string;
  oneDrivePath: string;
  cognitoIdToken: string;
  msGraphToken: string | null;
}

export async function getAddinSession(checkoutId: string): Promise<AddinSession> {
  return apiCall<AddinSession>("/files/checkout/addin-session", {
    params: { checkoutId },
  });
}

// Check-in (authenticated — uses token from session)

export async function requestCheckin(key: string): Promise<{ uploadUrl: string; oneDrivePath: string }> {
  return apiCall<{ uploadUrl: string; oneDrivePath: string }>("/files/checkin", {
    method: "POST",
    body: { key },
  });
}

export async function completeCheckin(key: string, versionId?: string): Promise<void> {
  await apiCall<{ message: string }>("/files/checkin/complete", {
    method: "POST",
    body: { key, versionId },
  });
}

export async function cancelCheckout(key: string): Promise<void> {
  await apiCall<{ message: string }>("/files/checkout/cancel", {
    method: "POST",
    body: { key },
  });
}
