import * as React from "react";
import { useState } from "react";
import { Button, Field, Textarea, tokens, makeStyles } from "@fluentui/react-components";

/* global HTMLTextAreaElement */

interface TextInsertionProps {
  insertText: (text: string) => void;
  replaceText: (text: string) => void;
}

const useStyles = makeStyles({
  instructions: {
    fontWeight: tokens.fontWeightSemibold,
    marginTop: "20px",
    marginBottom: "10px",
  },
  textPromptAndInsertion: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
  },
  textAreaField: {
    marginLeft: "0px",
    marginTop: "30px",
    marginBottom: "20px",
    marginRight: "0px",
    maxWidth: "100%",
  },

  buttonAction: {
    width: "150px",
    marginLeft: "5px",
    marginRight: "5px",
  },
});

const TextInsertion: React.FC<TextInsertionProps> = (props: TextInsertionProps) => {
  const [text, setText] = useState<string>("Some text.");

  const handleTextInsertion = async () => {
    await props.insertText(text);
  };

  const handleTextReplacement = async () => {
    await props.replaceText(text);
  };

  const handleTextChange = async (event: React.ChangeEvent<HTMLTextAreaElement>) => {
    setText(event.target.value);
  };

  const styles = useStyles();

  return (
    <div className={styles.textPromptAndInsertion}>
      <Field className={styles.textAreaField} size="large" label="Customer Query.">
        <Textarea size="large" value={text} onChange={handleTextChange} />
      </Field>
      <Field className={styles.instructions}>Click the button to insert text.</Field>
      <div>
        <span>
          <Button
            className={styles.buttonAction}
            appearance="primary"
            disabled={false}
            size="large"
            onClick={handleTextReplacement}
          >
            Replace
          </Button>
        </span>
        <span>
          <Button
            className={styles.buttonAction}
            appearance="secondary"
            disabled={false}
            size="large"
            onClick={handleTextInsertion}
          >
            Insert
          </Button>
        </span>
      </div>
    </div>
  );
};

export default TextInsertion;
