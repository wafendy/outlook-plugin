import * as React from "react";
import Header from "./Header";
import TextInsertion from "./TextInsertion";
import { makeStyles } from "@fluentui/react-components";
import { insertText } from "../taskpane";
import Metadata from "./Metadata";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "100vh",
    fontFamily: "Open Sans, Helvetica",
  },
});

const queryClient = new QueryClient();

const App: React.FC<AppProps> = (props: AppProps) => {
  const styles = useStyles();

  return (
    <QueryClientProvider client={queryClient}>
      <div className={styles.root}>
        <Metadata />
        <Header logo="assets/logo-filled.png" title={props.title} />
        <TextInsertion insertText={insertText} />
      </div>
    </QueryClientProvider>
  );
};

export default App;
