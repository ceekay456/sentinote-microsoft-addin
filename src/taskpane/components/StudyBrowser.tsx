import { useState, useEffect } from "react";
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
  SignOut24Regular,
  Beaker24Regular,
  ChevronRight20Regular,
} from "@fluentui/react-icons";
import { fetchStudies, StudySummary } from "../services/api";
import { signOut, getUser } from "../services/auth";

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
    justifyContent: "space-between",
    padding: "12px 16px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
    flexShrink: 0,
  },
  headerLeft: {
    display: "flex",
    flexDirection: "column",
    gap: "2px",
    minWidth: 0,
  },
  content: {
    flex: 1,
    overflow: "auto",
  },
  row: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "10px 16px",
    cursor: "pointer",
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    ":hover": {
      backgroundColor: tokens.colorNeutralBackground1Hover,
    },
  },
  rowText: {
    flex: 1,
    minWidth: 0,
  },
  studyName: {
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  chevron: {
    flexShrink: 0,
    color: tokens.colorNeutralForeground3,
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
});

interface StudyBrowserProps {
  onStudySelect: (studyId: string, name: string) => void;
}

export function StudyBrowser({ onStudySelect }: StudyBrowserProps) {
  const styles = useStyles();
  const [studies, setStudies] = useState<StudySummary[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const user = getUser();

  useEffect(() => {
    loadStudies();
  }, []);

  const loadStudies = async () => {
    setLoading(true);
    setError(null);
    try {
      const result = await fetchStudies();
      setStudies(result.filter((s) => s.status === "active"));
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to load studies");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className={styles.container}>
      <div className={styles.header}>
        <div className={styles.headerLeft}>
          <Title3>Studies</Title3>
          {user && <Caption1>{user.email}</Caption1>}
        </div>
        <Button
          icon={<SignOut24Regular />}
          appearance="subtle"
          onClick={signOut}
          title="Sign out"
        />
      </div>

      <div className={styles.content}>
        {loading && (
          <div className={styles.centerMessage}>
            <Spinner label="Loading studies..." />
          </div>
        )}

        {error && (
          <MessageBar intent="error">
            <MessageBarBody>{error}</MessageBarBody>
          </MessageBar>
        )}

        {!loading && !error && studies.length === 0 && (
          <div className={styles.centerMessage}>
            <Beaker24Regular />
            <Body1>No studies available</Body1>
          </div>
        )}

        {studies.map((study) => (
          <div
            key={study.studyId}
            className={styles.row}
            onClick={() => onStudySelect(study.studyId, study.name)}
          >
            <div className={styles.rowText}>
              <Body2 className={styles.studyName}>{study.name}</Body2>
              {study.protocolNumber && (
                <Caption1>{study.protocolNumber}</Caption1>
              )}
            </div>
            <ChevronRight20Regular className={styles.chevron} />
          </div>
        ))}
      </div>
    </div>
  );
}
