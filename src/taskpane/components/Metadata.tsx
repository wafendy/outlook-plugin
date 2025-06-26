import { makeStyles } from "@fluentui/react-components";
import * as React from "react";
import { useGetVersionQuery } from "../../hooks/version";

const useStyles = makeStyles({
  metadata: {
    display: "flex",
    position: "absolute",
    right: 0,
    fontSize: "10px",
    color: "#686868",
    opacity: 0.5,
  },
});

const Metadata: React.FC<{}> = () => {
  const { data: versionData } = useGetVersionQuery();
  const styles = useStyles();

  return <div className={styles.metadata}>{versionData && JSON.stringify(versionData)}</div>;
};

export default Metadata;
