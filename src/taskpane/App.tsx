import { useState, useEffect } from "react";
import { makeStyles, tokens, Spinner } from "@fluentui/react-components";
import { setIdToken } from "./services/auth";
import { getAddinSession, AddinSession } from "./services/api";
import { CheckInView } from "./components/CheckInView";
import { NotResearchCloud } from "./components/NotResearchCloud";
import { ReadOnlyView } from "./components/ReadOnlyView";

/* global Office */

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    overflow: "hidden",
    backgroundColor: tokens.colorNeutralBackground1,
  },
  loading: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    height: "100%",
    gap: "12px",
  },
});

/**
 * Check if the document is a VerityOne preview (read-only, no checkout).
 * Preview files are uploaded to VerityOne/.preview/{uuid}_{filename}
 */
function isPreviewDocument(source: string): boolean {
  try {
    const decoded = decodeURIComponent(source);
    return decoded.includes('.preview/') || decoded.includes('.preview%2F');
  } catch {
    return false;
  }
}

/**
 * Extract the checkout ID from the OneDrive document URL or filename.
 * Our convention: OneDrive path is VerityOne/{studyId}/{checkoutId}_{filename}
 * The document title in Word Online is "{checkoutId}_{filename}"
 */
function extractCheckoutId(documentUrl: string): string | null {
  try {
    // Try to find the checkout ID pattern in the URL
    // OneDrive URLs contain the filename with our {checkoutId}_ prefix
    const decoded = decodeURIComponent(documentUrl);

    // Match our checkout ID pattern: a base36 timestamp + random chars before an underscore
    // Pattern: {base36 ~10 chars}{random ~6 chars}_{original filename}
    const match = decoded.match(/([a-z0-9]{14,20})_[^/]+\.docx/i);
    if (match) {
      return match[1];
    }
  } catch {
    // ignore
  }
  return null;
}

export function App() {
  const styles = useStyles();
  const [state, setState] = useState<"loading" | "checkin" | "preview" | "not-found">("loading");
  const [session, setSession] = useState<AddinSession | null>(null);

  useEffect(() => {
    detectCheckout();
  }, []);

  const detectCheckout = async () => {
    try {
      // Collect all possible sources for the checkout ID
      const sources: string[] = [];

      // Browser document title (shows the filename in Word Online)
      try { sources.push(document.title); } catch { /* ignore */ }

      // Office document URL
      try { sources.push(Office.context.document.url || ""); } catch { /* ignore */ }

      // Office file properties
      try {
        const props = await new Promise<{ url: string }>((resolve, reject) => {
          Office.context.document.getFilePropertiesAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              resolve(result.value as unknown as { url: string });
            } else {
              reject(new Error("Failed"));
            }
          });
        });
        sources.push(props.url || "");
      } catch { /* ignore */ }

      // Browser location
      try { sources.push(window.location.href); } catch { /* ignore */ }

      // Check if this is a preview (read-only) document
      for (const source of sources) {
        if (source && isPreviewDocument(source)) {
          setState("preview");
          return;
        }
      }

      // Try to extract checkout ID from any source
      let checkoutId: string | null = null;
      for (const source of sources) {
        if (source) {
          checkoutId = extractCheckoutId(source);
          if (checkoutId) break;
        }
      }

      if (!checkoutId) {
        setState("not-found");
        return;
      }

      // Fetch the session from the backend
      const addinSession = await getAddinSession(checkoutId);

      // Set the Cognito token for authenticated API calls
      setIdToken(addinSession.cognitoIdToken);

      setSession(addinSession);
      setState("checkin");
    } catch {
      setState("not-found");
    }
  };

  return (
    <div className={styles.root}>
      {state === "loading" && (
        <div className={styles.loading}>
          <Spinner label="Detecting document..." />
        </div>
      )}

      {state === "checkin" && session && (
        <CheckInView
          fileKey={session.fileKey}
          fileName={session.fileName}
          studyId={session.studyId}
          filePath={session.filePath}
          checkedOutAt={session.checkedOutAt}
          oneDrivePath={session.oneDrivePath}
          msGraphToken={session.msGraphToken}
          onComplete={() => setState("not-found")}
        />
      )}

      {state === "preview" && <ReadOnlyView />}

      {state === "not-found" && <NotResearchCloud />}
    </div>
  );
}
