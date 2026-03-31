import { environment } from "./environment";
import { getIdToken, signOut } from "./auth";

interface ApiOptions {
  method?: string;
  params?: Record<string, string>;
  body?: unknown;
}

export class ApiError extends Error {
  constructor(
    public status: number,
    message: string
  ) {
    super(message);
    this.name = "ApiError";
  }
}

export async function apiCall<T>(path: string, options: ApiOptions = {}): Promise<T> {
  const { method = "GET", params, body } = options;

  const token = getIdToken();
  if (!token) {
    signOut();
    throw new ApiError(401, "Not authenticated");
  }

  let url = `${environment.apiUrl}${path}`;
  if (params) {
    const searchParams = new URLSearchParams(params);
    url += `?${searchParams.toString()}`;
  }

  const headers: Record<string, string> = {
    Authorization: token,
    "Content-Type": "application/json",
  };

  const response = await fetch(url, {
    method,
    headers,
    body: body ? JSON.stringify(body) : undefined,
  });

  if (response.status === 401) {
    signOut();
    throw new ApiError(401, "Session expired");
  }

  if (!response.ok) {
    const errorBody = await response.text();
    throw new ApiError(response.status, errorBody || response.statusText);
  }

  return response.json();
}

// API types matching the backend responses

export interface StudySummary {
  studyId: string;
  name: string;
  protocolNumber: string;
  status: string;
  templateId: string;
  createdAt: string;
  createdBy: string;
}

export interface FileItem {
  key: string;
  name: string;
  size?: number;
  lastModified?: string;
  type: "file" | "folder";
  metadataType?: string;
}

export interface FilesResponse {
  files: FileItem[];
  folders: FileItem[];
  prefix: string;
  source: string;
  hasMore: boolean;
  nextContinuationToken?: string;
}

export interface PermissionsResponse {
  userId: string;
  email: string;
  tenantId: string;
  orgRole: string;
  studies: Array<{
    studyId: string;
    role: string;
    permissions: Record<string, boolean>;
  }>;
}

// API methods

export async function fetchStudies(): Promise<StudySummary[]> {
  const result = await apiCall<{ studies: StudySummary[] }>("/studies");
  return result.studies;
}

export async function fetchFiles(prefix: string, source = "storage"): Promise<FilesResponse> {
  return apiCall<FilesResponse>("/files", {
    params: { prefix, source, maxKeys: "200" },
  });
}

export async function fetchPermissions(): Promise<PermissionsResponse> {
  return apiCall<PermissionsResponse>("/permissions/me");
}
