import * as React from "react";
import Header from "./Header";
import TextInsertion from "./TextInsertion";
import { makeStyles } from "@fluentui/react-components";
import { insertText } from "../taskpane";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    fontFamily: "Open Sans, Helvetica"
  },
});

const App: React.FC<AppProps> = (props: AppProps) => {
  const styles = useStyles();

  console.log("======> App React");
  return (
    <div className={styles.root}>
      <Header logo="assets/logo-filled.png" title={props.title} />
      <TextInsertion insertText={insertText} />
    </div>
  );
};

export default App;
