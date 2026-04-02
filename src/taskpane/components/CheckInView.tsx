import { useState } from "react";
import {
  makeStyles,
  tokens,
  Title3,
  Body1,
  Body2,
  Caption1,
  Button,
  Spinner,
  MessageBar,
  MessageBarBody,
} from "@fluentui/react-components";
import {
  Document24Regular,
  ArrowUpload20Regular,
  Dismiss20Regular,
} from "@fluentui/react-icons";
import { requestCheckin, completeCheckin, cancelCheckout } from "../services/api";
import { getMicrosoftToken } from "../services/microsoft-auth";
import { deleteFromOneDrive } from "../services/onedrive";

/* global Office */

/**
 * Get the current document content directly from Word Online.
 * This captures all edits, bypassing OneDrive caching.
 */
function getDocumentContent(): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(
      Office.FileType.Compressed,
      { sliceSize: 4 * 1024 * 1024 },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error("Failed to get document content"));
          return;
        }

        const file = result.value;
        const sliceCount = file.sliceCount;
        const chunks: Uint8Array[] = [];
        let slicesRead = 0;

        const readSlice = (index: number) => {
          file.getSliceAsync(index, (sliceResult) => {
            if (sliceResult.status !== Office.AsyncResultStatus.Succeeded) {
              file.closeAsync();
              reject(new Error(`Failed to read slice ${index}`));
              return;
            }

            chunks.push(new Uint8Array(sliceResult.value.data));
            slicesRead++;

            if (slicesRead === sliceCount) {
              file.closeAsync();
              // Combine all chunks into a single ArrayBuffer
              const totalLength = chunks.reduce((sum, c) => sum + c.length, 0);
              const combined = new Uint8Array(totalLength);
              let offset = 0;
              for (const chunk of chunks) {
                combined.set(chunk, offset);
                offset += chunk.length;
              }
              resolve(combined.buffer);
            } else {
              readSlice(index + 1);
            }
          });
        };

        readSlice(0);
      }
    );
  });
}

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    height: "100%",
    overflow: "hidden",
  },
  header: {
    padding: "16px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  content: {
    flex: 1,
    padding: "16px",
    display: "flex",
    flexDirection: "column",
    gap: "16px",
  },
  fileInfo: {
    display: "flex",
    gap: "12px",
    alignItems: "flex-start",
  },
  fileIcon: {
    flexShrink: 0,
    color: tokens.colorBrandForeground1,
    marginTop: "2px",
  },
  fileDetails: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
    minWidth: 0,
  },
  fileName: {
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  meta: {
    display: "flex",
    flexDirection: "column",
    gap: "4px",
    padding: "12px",
    backgroundColor: tokens.colorNeutralBackground2,
    borderRadius: "4px",
  },
  actions: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    marginTop: "auto",
  },
});

interface CheckInViewProps {
  fileKey: string;
  fileName: string;
  studyId: string;
  filePath: string;
  checkedOutAt: string;
  oneDrivePath: string;
  msGraphToken: string | null;
  onComplete: () => void;
}

export function CheckInView({
  fileKey,
  fileName,
  filePath,
  checkedOutAt,
  oneDrivePath,
  msGraphToken,
  onComplete,
}: CheckInViewProps) {
  const styles = useStyles();
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);

  const getGraphToken = async (): Promise<string> => {
    // Use stored token if available, fall back to interactive auth
    if (msGraphToken) return msGraphToken;
    return getMicrosoftToken();
  };

  const handleCheckin = async () => {
    setBusy(true);
    setError(null);
    try {
      const { uploadUrl } = await requestCheckin(fileKey);

      // Get the document content directly from Word Online (includes all edits)
      const fileData = await getDocumentContent();

      // Upload to S3
      const s3Response = await fetch(uploadUrl, {
        method: "PUT",
        headers: {
          "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        },
        body: fileData,
      });
      if (!s3Response.ok) throw new Error("Failed to upload to storage");
      const versionId = s3Response.headers.get("x-amz-version-id") || undefined;

      await completeCheckin(fileKey, versionId);

      // Clean up OneDrive (best effort)
      try {
        const msToken = await getGraphToken();
        await deleteFromOneDrive(msToken, oneDrivePath);
      } catch { /* ok */ }

      setSuccess("Document saved to ResearchCloud successfully.");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Check-in failed");
    } finally {
      setBusy(false);
    }
  };

  const handleCancel = async () => {
    setBusy(true);
    setError(null);
    try {
      await cancelCheckout(fileKey);

      // Clean up OneDrive (best effort)
      try {
        const msToken = await getGraphToken();
        await deleteFromOneDrive(msToken, oneDrivePath);
      } catch { /* ok */ }

      setSuccess("Checkout cancelled. Changes discarded.");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Cancel failed");
    } finally {
      setBusy(false);
    }
  };

  // Extract the folder path from filePath (remove the filename)
  const folderPath = filePath.includes("/")
    ? filePath.substring(0, filePath.lastIndexOf("/"))
    : "";

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <Title3>ResearchCloud</Title3>
      </div>

      <div className={styles.content}>
        <div className={styles.fileInfo}>
          <Document24Regular className={styles.fileIcon} />
          <div className={styles.fileDetails}>
            <Body1 className={styles.fileName}><b>{fileName}</b></Body1>
            {folderPath && <Caption1>{folderPath}</Caption1>}
          </div>
        </div>

        <div className={styles.meta}>
          <Caption1>
            Checked out: {new Date(checkedOutAt).toLocaleString()}
          </Caption1>
        </div>

        {success && (
          <MessageBar intent="success">
            <MessageBarBody>{success}</MessageBarBody>
          </MessageBar>
        )}

        {error && (
          <MessageBar intent="error">
            <MessageBarBody>{error}</MessageBarBody>
          </MessageBar>
        )}

        {!success && (
          <div className={styles.actions}>
            <Button
              appearance="primary"
              icon={busy ? <Spinner size="tiny" /> : <ArrowUpload20Regular />}
              disabled={busy}
              onClick={handleCheckin}
            >
              Save to ResearchCloud
            </Button>
            <Button
              appearance="secondary"
              icon={<Dismiss20Regular />}
              disabled={busy}
              onClick={handleCancel}
            >
              Cancel and Release
            </Button>
          </div>
        )}
      </div>
    </div>
  );
}
