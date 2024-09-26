import * as React from "react";
import { useState, useEffect } from "react";
import { Button, Field, Textarea, tokens, makeStyles } from "@fluentui/react-components";
import { getEmailData } from "../taskpane";

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
  radioButtonField: {},

  buttonAction: {
    width: "150px",
    marginLeft: "5px",
    marginRight: "5px",
  },
});

const TextInsertion: React.FC<TextInsertionProps> = (props: TextInsertionProps) => {
  const [queryText, setQueryText] = useState<string>("Some query text.");
  const [replyText, setReplyText] = useState<string>("");
  const [selectedValue, setSelectedValue] = useState<string>("");

  const handleTextInsertion = async () => {
    await props.insertText(replyText);
  };

  const handleTextReplacement = async () => {
    await props.replaceText(replyText);
  };

  const handleSelectChange = (event: React.ChangeEvent<HTMLSelectElement>) => {
    const value = event.target.value;
    setSelectedValue(value); // Update the state with the selected value
    handleItemSelected(value); // Call a custom function when item is selected
  };

  // Function that runs after an item is selected
  const handleItemSelected = (value: string) => {
    console.log("Selected item:", value);

    if (value === "doc1") {
      setReplyText(`Dear [Customer's Name],

We refer to your enquiry regarding the regulations for renting out your HDB unit as a Singaporean Permanent Resident (PR).

As a Singaporean PR, you are eligible to rent out your HDB flat, provided you meet certain conditions. Firstly, you must have fulfilled the Minimum Occupation Period (MOP) of five years. Additionally, you must seek approval from the Housing & Development Board (HDB) before renting out your flat. The approval process involves submitting an application with the necessary details, including the particulars of your tenants and the rental period.

Please note that the maximum rental period per application is 36 months for Singaporean or Malaysian tenants and 24 months for non-Malaysian non-citizen tenants. The rental period must not be less than six months. Furthermore, you must ensure that the tenants meet the eligibility criteria, which include being Singapore Citizens, Singapore Permanent Residents, or non-citizens legally residing in Singapore with valid passes.

For more detailed information on the terms and conditions, you may refer to the HDB's official guidelines at [https://www.hdb.gov.sg/cs/infoweb/application-to-rent-out-bedrooms/terms-and-conditions](https://www.hdb.gov.sg/cs/infoweb/application-to-rent-out-bedrooms/terms-and-conditions).

Yours sincerely,
(Customer Service Officer)`);
    }

    if (value == "doc2") {
      setReplyText(`Dear [Customer's Name],

We refer to your enquiry regarding your childcare leave entitlement as a Singaporean Permanent Resident with a 2-year-old daughter and an 11-year-old son.

As a Singaporean Permanent Resident, you are entitled to 2 days of childcare leave per year if you have a child below the age of 7 years, provided you have worked with your employer for at least 3 months. This entitlement is applicable to your 2-year-old daughter. The childcare leave is paid by your employer and is not claimable from the government. Unfortunately, there is no childcare leave entitlement for your 11-year-old son as the leave is only applicable for children below the age of 7 years.

For further details on childcare leave eligibility, you may refer to the Ministry of Manpower's guidelines here: [https://mom.gov.sg/employment-practices/leave/childcare-leave/eligibility-and-entitlement](https://mom.gov.sg/employment-practices/leave/childcare-leave/eligibility-and-entitlement).

Yours sincerely,
(Customer Service Officer)`);
    }

    // Add your custom logic here (e.g., update a form, make an API call, etc.)
  };

  const styles = useStyles();

  useEffect(() => {
    const loadText = async () => {
      try {
        const fetchedText = await getEmailData(); // Fetch the text asynchronously
        setQueryText(fetchedText); // Set the text in the state
        // setText("fetchedText"); // Set the text in the state
      } catch (error) {
        console.error("Failed to fetch text", error);
      }
    };

    loadText(); // Invoke the async function
  }, []);

  return (
    <div className={styles.textPromptAndInsertion}>
      <Field className={styles.textAreaField} size="large" label="Customer Query.">
        <Textarea size="large" rows={10} value={queryText} />
      </Field>
      <Field className={styles.radioButtonField} label="Reference Documents">
        <select value={selectedValue} onChange={handleSelectChange}>
          <option value="">Select a reference document</option>
          <option value="doc1">HMG | Renting Out A Flat Or Bedroom | Terms and Conditions</option>
          <option value="doc2">FAQ | Childcare Leave</option>
        </select>
      </Field>

      <Field className={styles.textAreaField} size="large" label="Generated reply">
        <Textarea size="large" rows={10} value={replyText} />
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
