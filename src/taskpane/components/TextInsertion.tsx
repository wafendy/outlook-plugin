import * as React from "react";
import { useState, useEffect } from "react";
import {
  Button,
  Field,
  Textarea,
  makeStyles,
  Dropdown,
  useId,
  Option,
  OptionOnSelectData,
  SelectionEvents,
  Spinner,
} from "@fluentui/react-components";
import { getEmailData } from "../taskpane";

/* global HTMLTextAreaElement */

interface TextInsertionProps {
  insertText: (text: string) => void;
}

const useStyles = makeStyles({
  textPromptAndInsertion: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    padding: "12px 24px 48px 24px",
  },
  textAreaField: {
    marginBottom: "24px",
    width: "100%",
    "& label": {
      margin: "12px 0",
      color: "#1C76D5",
      fontWeight: 600,
    },
  },
  buttonAction: {
    width: "100%",
    borderRadius: "8px",
    height: "40px",
  },
  subtitle: {
    fontSize: "12px",
    marginBottom: "16px",
    color: "#686868",
  },
});

type LoadingState = "initial" | "loading" | "loaded";

const TextInsertion: React.FC<TextInsertionProps> = (props: TextInsertionProps) => {
  const [queryText, setQueryText] = useState<string>("Some query text.");
  const [replyText, setReplyText] = useState<string>("");
  const [selectedValue, setSelectedValue] = useState<string[]>([]);
  const [selectedText, setSelectedText] = useState<string>("");

  const dropdownId = useId("dropdown-default");
  const options = [
    { value: "doc1", text: "HMG | Renting Out A Flat Or Bedroom | Terms and Conditions" },
    { value: "doc2", text: "FAQ | Childcare Leave" },
  ];

  const handleTextInsertion = async () => {
    props.insertText(replyText);
  };

  const handleSelectChange = (_: SelectionEvents, data: OptionOnSelectData) => {
    setSelectedValue(data.selectedOptions); // Update the state with the selected value
  };

  const autoSelectRelevantDocuments = (text: string) => {
    if (text.includes("rent")) {
      setSelectedValue([options[0].value]);
      setSelectedText(options[0].text);
    }

    if (text.includes("daughter")) {
      setSelectedValue([options[1].value]);
      setSelectedText(options[1].text);
    }
  };

  // Function that runs after an item is selected
  const handleItemSelected = (value: string[]) => {
    console.log("Selected item:", value);

    setLoadingState("loading");

    let replyText = "";

    if (value.includes("doc1")) {
      replyText = `Dear [Customer's Name],

We refer to your enquiry regarding the regulations for renting out your HDB unit as a Singaporean Permanent Resident (PR).

As a Singaporean PR, you are eligible to rent out your HDB flat, provided you meet certain conditions. Firstly, you must have fulfilled the Minimum Occupation Period (MOP) of five years. Additionally, you must seek approval from the Housing & Development Board (HDB) before renting out your flat. The approval process involves submitting an application with the necessary details, including the particulars of your tenants and the rental period.

Please note that the maximum rental period per application is 36 months for Singaporean or Malaysian tenants and 24 months for non-Malaysian non-citizen tenants. The rental period must not be less than six months. Furthermore, you must ensure that the tenants meet the eligibility criteria, which include being Singapore Citizens, Singapore Permanent Residents, or non-citizens legally residing in Singapore with valid passes.

For more detailed information on the terms and conditions, you may refer to the HDB's official guidelines at https://www.hdb.gov.sg/cs/infoweb/application-to-rent-out-bedrooms/terms-and-conditions.

Yours sincerely,
(Customer Service Officer)`;
    }

    if (value.includes("doc2")) {
      replyText = `Dear [Customer's Name],

We refer to your enquiry regarding your childcare leave entitlement as a Singaporean Permanent Resident with a 2-year-old daughter and an 11-year-old son.

As a Singaporean Permanent Resident, you are entitled to 2 days of childcare leave per year if you have a child below the age of 7 years, provided you have worked with your employer for at least 3 months. This entitlement is applicable to your 2-year-old daughter. The childcare leave is paid by your employer and is not claimable from the government. Unfortunately, there is no childcare leave entitlement for your 11-year-old son as the leave is only applicable for children below the age of 7 years.

For further details on childcare leave eligibility, you may refer to the Ministry of Manpower's guidelines here: https://mom.gov.sg/employment-practices/leave/childcare-leave/eligibility-and-entitlement.

Yours sincerely,
(Customer Service Officer)`;
    }

    setTimeout(() => {
      setLoadingState("loaded");
      setReplyText(replyText);
    }, 3000);
  };

  const styles = useStyles();

  const [loadingState, setLoadingState] = React.useState<LoadingState>("initial");

  const buttonIcon = loadingState === "loading" ? <Spinner size="tiny" /> : null;

  useEffect(() => {
    const loadText = async () => {
      try {
        const fetchedText = await getEmailData(); // Fetch the text asynchronously
        setQueryText(fetchedText); // Set the text in the state
        autoSelectRelevantDocuments(fetchedText);
      } catch (error) {
        console.error("Failed to fetch text", error);
      }
    };

    loadText(); // Invoke the async function
  }, []);

  return (
    <div className={styles.textPromptAndInsertion}>
      <Field className={styles.textAreaField} size="medium" label="Customer Query">
        <Textarea size="medium" rows={14} value={queryText} />
      </Field>

      <Field className={styles.textAreaField} size="medium" label="Reference Documents">
        <div className={styles.subtitle}>Select up to 5 documents for SmartCompose to use when generating a reply.</div>
        <Dropdown
          size="large"
          aria-labelledby={dropdownId}
          multiselect={true}
          placeholder="Select or search from a list"
          onOptionSelect={handleSelectChange}
          selectedOptions={selectedValue}
          value={selectedText}
          {...props}
        >
          {options.map((option) => (
            <Option key={option.value} value={option.value}>
              {option.text}
            </Option>
          ))}
        </Dropdown>

        <br />

        <Button
          className={styles.buttonAction}
          appearance="primary"
          disabled={false}
          size="medium"
          icon={buttonIcon}
          iconPosition="before"
          onClick={() => handleItemSelected(selectedValue)}
        >
          Generate reply
        </Button>
      </Field>

      {replyText && (
        <>
          <Field className={styles.textAreaField} size="medium" label="Generated Reply">
            <div className={styles.subtitle}>
              Content is generated by Generative AI, a type of artificial intelligence technology. Please verify the
              information before using it.
            </div>
            <Textarea size="medium" rows={14} value={replyText} />
          </Field>

          <Button
            className={styles.buttonAction}
            appearance="primary"
            disabled={false}
            size="medium"
            onClick={handleTextInsertion}
          >
            Insert to reply window
          </Button>
        </>
      )}
    </div>
  );
};

export default TextInsertion;
