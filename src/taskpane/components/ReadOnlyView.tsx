import {
  makeStyles,
  tokens,
  Body1,
  Body2,
  Title3,
} from "@fluentui/react-components";
import { Eye24Regular } from "@fluentui/react-icons";

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
    color: tokens.colorNeutralForeground3,
  },
});

export function ReadOnlyView() {
  const styles = useStyles();
  return (
    <div className={styles.container}>
      <Eye24Regular />
      <Title3>VerityOne</Title3>
      <Body1>
        This document is opened in view-only mode.
      </Body1>
      <Body2>
        To edit and save changes, use "Edit in Word" from the VerityOne web app.
      </Body2>
    </div>
  );
}
