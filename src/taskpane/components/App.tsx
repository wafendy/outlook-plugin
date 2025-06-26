import { makeStyles } from "@fluentui/react-components";
import { QueryClient, QueryClientProvider } from "@tanstack/react-query";
import * as React from "react";
import { insertText } from "../taskpane";
import ActiveUser from "./ActiveUser";
import Header from "./Header";
import Metadata from "./Metadata";
import TextInsertion from "./TextInsertion";

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
        <ActiveUser />
        <TextInsertion insertText={insertText} />
      </div>
    </QueryClientProvider>
  );
};

export default App;
