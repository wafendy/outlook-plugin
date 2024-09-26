/* global Office console */
function cleanEmailText(text: string) {
  let lines = text.split("\n");

  let cleanedLines = lines.filter((line) => {
    return !(
      line.startsWith("From:") ||
      line.startsWith("Date:") ||
      line.startsWith("To:") ||
      line.startsWith("Subject:") ||
      line.trim().endsWith("wrote:")
    );
  });

  return cleanedLines.join("\n");
}

export function getEmailData(): Promise<string> {
  return new Promise((resolve, reject) => {
    try {
      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          const text = cleanEmailText(asyncResult.value).trim();
          console.log("text", text);
          resolve(text); // Resolve the promise with the result
        } else {
          reject(new Error("Failed to retrieve email body"));
        }
      });
    } catch (error) {
      console.log("Error: " + error);
      reject(error); // Reject the promise in case of an exception
    }
  });
}

export async function insertText(text: string) {
  // Write text to the cursor point in the compose surface.
  try {
    Office.context.mailbox.item?.body.setSelectedDataAsync(
      text,
      { coercionType: Office.CoercionType.Text },
      (asyncResult: Office.AsyncResult<void>) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          throw asyncResult.error.message;
        }
      }
    );
    console.log("Done Inserting");
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function replaceText(text: string) {
  // Write text to the cursor point in the compose surface.
  try {
    Office.context.mailbox.item?.body.setSelectedDataAsync(
      text,
      { coercionType: Office.CoercionType.Text },
      (asyncResult: Office.AsyncResult<void>) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          throw asyncResult.error.message;
        }
      }
    );
    console.log("Done Replacing");
  } catch (error) {
    console.log("Error: " + error);
  }
}
