import { useState, useEffect, useCallback } from "react";
import { makeStyles, tokens } from "@fluentui/react-components";
import { isAuthenticated, onAuthChange } from "./services/auth";
import { Login } from "./components/Login";
import { StudyBrowser } from "./components/StudyBrowser";
import { FileBrowser } from "./components/FileBrowser";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    overflow: "hidden",
    backgroundColor: tokens.colorNeutralBackground1,
  },
});

export function App() {
  const styles = useStyles();
  const [authenticated, setAuthenticated] = useState(isAuthenticated());
  const [selectedStudy, setSelectedStudy] = useState<{
    studyId: string;
    name: string;
  } | null>(null);

  useEffect(() => {
    return onAuthChange((auth) => {
      setAuthenticated(auth);
      if (!auth) setSelectedStudy(null);
    });
  }, []);

  const handleStudySelect = useCallback(
    (studyId: string, name: string) => {
      setSelectedStudy({ studyId, name });
    },
    []
  );

  const handleBack = useCallback(() => {
    setSelectedStudy(null);
  }, []);

  if (!authenticated) {
    return (
      <div className={styles.root}>
        <Login />
      </div>
    );
  }

  return (
    <div className={styles.root}>
      {selectedStudy ? (
        <FileBrowser
          studyId={selectedStudy.studyId}
          studyName={selectedStudy.name}
          onBack={handleBack}
        />
      ) : (
        <StudyBrowser onStudySelect={handleStudySelect} />
      )}
    </div>
  );
}
