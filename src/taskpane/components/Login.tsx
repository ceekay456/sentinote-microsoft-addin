import { useState } from "react";
import {
  makeStyles,
  tokens,
  Title3,
  Body1,
  Button,
  MessageBar,
  MessageBarBody,
  Spinner,
} from "@fluentui/react-components";
import { DocumentArrowRight24Regular } from "@fluentui/react-icons";
import { signIn } from "../services/auth";

const useStyles = makeStyles({
  container: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    height: "100%",
    padding: "24px",
    gap: "16px",
    textAlign: "center",
  },
  icon: {
    fontSize: "48px",
    color: tokens.colorBrandForeground1,
    marginBottom: "8px",
  },
  description: {
    color: tokens.colorNeutralForeground2,
    maxWidth: "280px",
  },
  error: {
    maxWidth: "280px",
  },
});

export function Login() {
  const styles = useStyles();
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);

  const handleSignIn = async () => {
    setLoading(true);
    setError(null);
    try {
      await signIn();
    } catch (err) {
      const message = err instanceof Error ? err.message : "Sign-in failed";
      if (!message.includes("dialog was closed")) {
        setError(message);
      }
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className={styles.container}>
      <DocumentArrowRight24Regular className={styles.icon} />
      <Title3>ResearchCloud</Title3>
      <Body1 className={styles.description}>
        Sign in to browse your studies and check out documents for editing in Word.
      </Body1>

      {error && (
        <MessageBar intent="error" className={styles.error}>
          <MessageBarBody>{error}</MessageBarBody>
        </MessageBar>
      )}

      {loading ? (
        <Spinner label="Signing in..." />
      ) : (
        <Button appearance="primary" size="large" onClick={handleSignIn}>
          Sign in
        </Button>
      )}
    </div>
  );
}
