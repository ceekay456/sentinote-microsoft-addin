import { environment } from "./environment";

let currentIdToken: string | null = null;

/**
 * Set the Cognito ID token (provided by the backend session endpoint).
 */
export function setIdToken(token: string): void {
  currentIdToken = token;
}

/**
 * Get the current Cognito ID token for API calls.
 */
export function getIdToken(): string | null {
  return currentIdToken;
}

/**
 * Get the API base URL.
 */
export function getApiUrl(): string {
  return environment.apiUrl;
}
