import {
  makeStyles,
  tokens,
  Body1,
  Body2,
  Title3,
} from "@fluentui/react-components";
import { Info24Regular } from "@fluentui/react-icons";

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

export function NotResearchCloud() {
  const styles = useStyles();
  return (
    <div className={styles.container}>
      <Info24Regular />
      <Title3>ResearchCloud</Title3>
      <Body1>
        This document is not checked out from ResearchCloud.
      </Body1>
      <Body2>
        Open a document from the ResearchCloud web app to use this add-in.
      </Body2>
    </div>
  );
}
