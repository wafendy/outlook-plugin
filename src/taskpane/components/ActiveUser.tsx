import { makeStyles } from "@fluentui/react-components";
import * as React from "react";

const useStyles = makeStyles({
  activeUser: {
    display: "flex",
    position: "absolute",
    left: "50%",
    transform: "translateX(-50%)",
    fontSize: "10px",
    color: "#686868",
    opacity: 0.5,
  },
});

const ActiveUser: React.FC<{}> = () => {
  const styles = useStyles();

  return (
    <div className={styles.activeUser} id="active-user">
      Show Active User Here
    </div>
  );
};

export default ActiveUser;
