import { useState, useEffect, useCallback } from "react";
import {
  makeStyles,
  tokens,
  Title3,
  Body1,
  Caption1,
  Button,
  Spinner,
  MessageBar,
  MessageBarBody,
} from "@fluentui/react-components";
import {
  ArrowLeft24Regular,
  Folder24Regular,
  Document24Regular,
  DocumentArrowRight24Regular,
} from "@fluentui/react-icons";
import { fetchFiles, FileItem, FilesResponse } from "../services/api";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    height: "100%",
    overflow: "hidden",
  },
  header: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "12px 16px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
  },
  breadcrumbs: {
    padding: "8px 16px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
    display: "flex",
    flexWrap: "wrap",
    gap: "4px",
    alignItems: "center",
  },
  breadcrumbSeparator: {
    color: tokens.colorNeutralForeground3,
  },
  content: {
    flex: 1,
    overflow: "auto",
    padding: "4px 0",
  },
  row: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "8px 16px",
    cursor: "pointer",
    ":hover": {
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
  },
  rowIcon: {
    flexShrink: 0,
    color: tokens.colorNeutralForeground2,
  },
  rowInfo: {
    flex: 1,
    minWidth: 0,
  },
  fileName: {
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  fileMeta: {
    display: "flex",
    gap: "12px",
  },
  docxIcon: {
    color: tokens.colorBrandForeground1,
  },
  centerMessage: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    height: "100%",
    gap: "12px",
    color: tokens.colorNeutralForeground3,
  },
  emptyFolder: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    height: "100%",
    gap: "8px",
    color: tokens.colorNeutralForeground3,
  },
});

interface FileBrowserProps {
  studyId: string;
  studyName: string;
  onBack: () => void;
}

function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function isDocx(name: string): boolean {
  return name.toLowerCase().endsWith(".docx");
}

export function FileBrowser({ studyId, studyName, onBack }: FileBrowserProps) {
  const styles = useStyles();
  const [currentPrefix, setCurrentPrefix] = useState(`studies/${studyId}/`);
  const [folders, setFolders] = useState<FileItem[]>([]);
  const [files, setFiles] = useState<FileItem[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);

  const loadFiles = useCallback(async (prefix: string) => {
    setLoading(true);
    setError(null);
    try {
      const result: FilesResponse = await fetchFiles(prefix);
      setFolders(result.folders || []);
      setFiles(result.files || []);
      setCurrentPrefix(prefix);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to load files");
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    loadFiles(currentPrefix);
  }, []);

  const navigateToFolder = (folder: FileItem) => {
    loadFiles(folder.key);
  };

  // Build breadcrumbs from the current prefix
  const breadcrumbs = (() => {
    const basePrefix = `studies/${studyId}/`;
    const relativePath = currentPrefix.replace(basePrefix, "");
    const parts = relativePath.split("/").filter(Boolean);
    const crumbs = [{ label: studyName, prefix: basePrefix }];
    let accumulated = basePrefix;
    for (const part of parts) {
      accumulated += `${part}/`;
      crumbs.push({ label: part, prefix: accumulated });
    }
    return crumbs;
  })();

  const allItems = [...folders, ...files];

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <Button
          icon={<ArrowLeft24Regular />}
          appearance="subtle"
          onClick={onBack}
          title="Back to studies"
        />
        <Title3>{studyName}</Title3>
      </div>

      {breadcrumbs.length > 1 && (
        <div className={styles.breadcrumbs}>
          {breadcrumbs.map((crumb, i) => (
            <span key={crumb.prefix}>
              {i > 0 && <span className={styles.breadcrumbSeparator}> / </span>}
              {i < breadcrumbs.length - 1 ? (
                <Button
                  appearance="subtle"
                  size="small"
                  onClick={() => loadFiles(crumb.prefix)}
                >
                  {crumb.label}
                </Button>
              ) : (
                <Caption1><b>{crumb.label}</b></Caption1>
              )}
            </span>
          ))}
        </div>
      )}

      <div className={styles.content}>
        {loading && (
          <div className={styles.centerMessage}>
            <Spinner label="Loading..." />
          </div>
        )}

        {error && (
          <MessageBar intent="error">
            <MessageBarBody>{error}</MessageBarBody>
          </MessageBar>
        )}

        {!loading && !error && allItems.length === 0 && (
          <div className={styles.emptyFolder}>
            <Folder24Regular />
            <Body1>This folder is empty</Body1>
          </div>
        )}

        {!loading &&
          folders.map((folder) => (
            <div
              key={folder.key}
              className={styles.row}
              onClick={() => navigateToFolder(folder)}
            >
              <Folder24Regular className={styles.rowIcon} />
              <div className={styles.rowInfo}>
                <Body1 className={styles.fileName}>{folder.name}</Body1>
              </div>
            </div>
          ))}

        {!loading &&
          files.map((file) => (
            <div key={file.key} className={styles.row}>
              {isDocx(file.name) ? (
                <DocumentArrowRight24Regular className={styles.docxIcon} />
              ) : (
                <Document24Regular className={styles.rowIcon} />
              )}
              <div className={styles.rowInfo}>
                <Body1 className={styles.fileName}>{file.name}</Body1>
                <div className={styles.fileMeta}>
                  {file.size != null && (
                    <Caption1>{formatFileSize(file.size)}</Caption1>
                  )}
                  {file.lastModified && (
                    <Caption1>
                      {new Date(file.lastModified).toLocaleDateString()}
                    </Caption1>
                  )}
                </div>
              </div>
            </div>
          ))}
      </div>
    </div>
  );
}
