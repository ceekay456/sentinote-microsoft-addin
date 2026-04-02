const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

/**
 * Upload a file to the user's OneDrive via Microsoft Graph API.
 * Uses simple upload for files ≤ 4MB, chunked upload session for larger files.
 * Returns the OneDrive item metadata (including webUrl).
 */
export async function uploadToOneDrive(
  graphToken: string,
  path: string,
  fileData: ArrayBuffer
): Promise<{ webUrl: string }> {
  const encodedPath = encodeURIComponent(path).replace(/%2F/g, "/");

  if (fileData.byteLength <= 4 * 1024 * 1024) {
    const response = await fetch(
      `${GRAPH_BASE}/me/drive/root:/${encodedPath}:/content`,
      {
        method: "PUT",
        headers: {
          Authorization: `Bearer ${graphToken}`,
          "Content-Type":
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        },
        body: fileData,
      }
    );
    if (!response.ok) {
      const error = await response.text();
      throw new Error(`OneDrive upload failed (${response.status}): ${error}`);
    }
    return response.json();
  }

  return uploadLargeFile(graphToken, encodedPath, fileData);
}

async function uploadLargeFile(
  graphToken: string,
  encodedPath: string,
  fileData: ArrayBuffer
): Promise<{ webUrl: string }> {
  // Create upload session
  const sessionResponse = await fetch(
    `${GRAPH_BASE}/me/drive/root:/${encodedPath}:/createUploadSession`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${graphToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        item: { "@microsoft.graph.conflictBehavior": "replace" },
      }),
    }
  );

  if (!sessionResponse.ok) {
    const error = await sessionResponse.text();
    throw new Error(`Upload session failed (${sessionResponse.status}): ${error}`);
  }

  const session = await sessionResponse.json();
  const uploadUrl = session.uploadUrl;

  // Upload in 5MB chunks
  const chunkSize = 5 * 1024 * 1024;
  const totalSize = fileData.byteLength;
  let result: { webUrl: string } | undefined;

  for (let offset = 0; offset < totalSize; offset += chunkSize) {
    const end = Math.min(offset + chunkSize, totalSize);
    const chunk = fileData.slice(offset, end);

    const chunkResponse = await fetch(uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Length": chunk.byteLength.toString(),
        "Content-Range": `bytes ${offset}-${end - 1}/${totalSize}`,
      },
      body: chunk,
    });

    if (!chunkResponse.ok && chunkResponse.status !== 202) {
      const error = await chunkResponse.text();
      throw new Error(`Chunk upload failed (${chunkResponse.status}): ${error}`);
    }

    if (chunkResponse.status === 200 || chunkResponse.status === 201) {
      result = await chunkResponse.json();
    }
  }

  if (!result) throw new Error("Upload completed without response");
  return result;
}

/**
 * Download a file from the user's OneDrive.
 */
export async function downloadFromOneDrive(
  graphToken: string,
  path: string
): Promise<ArrayBuffer> {
  const encodedPath = encodeURIComponent(path).replace(/%2F/g, "/");

  const response = await fetch(
    `${GRAPH_BASE}/me/drive/root:/${encodedPath}:/content`,
    { headers: { Authorization: `Bearer ${graphToken}` } }
  );

  if (!response.ok) {
    const error = await response.text();
    throw new Error(`OneDrive download failed (${response.status}): ${error}`);
  }

  return response.arrayBuffer();
}

/**
 * Delete a file from the user's OneDrive.
 */
export async function deleteFromOneDrive(
  graphToken: string,
  path: string
): Promise<void> {
  const encodedPath = encodeURIComponent(path).replace(/%2F/g, "/");

  const response = await fetch(
    `${GRAPH_BASE}/me/drive/root:/${encodedPath}`,
    {
      method: "DELETE",
      headers: { Authorization: `Bearer ${graphToken}` },
    }
  );

  if (!response.ok && response.status !== 404) {
    const error = await response.text();
    throw new Error(`OneDrive delete failed (${response.status}): ${error}`);
  }
}
