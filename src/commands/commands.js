// /*
//  * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
//  * See LICENSE in the project root for license information.
//  */

// /* global global, Office, self, window */

// Office.onReady(() => {
//   // If needed, Office.js is ready to be called
// });

// /**
//  * Shows a notification when the add-in command is executed.
//  * @param event {Office.AddinCommands.Event}
//  */
// function action(event) {
//   const message = {
//     type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
//     message: "Performed action.",
//     icon: "Icon.80x80",
//     persistent: true,
//   };

//   // Show a notification message
//   Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

//   // Be sure to indicate when the add-in command function is complete
//   event.completed();
// }

// function getGlobal() {
//   return typeof self !== "undefined"
//     ? self
//     : typeof window !== "undefined"
//     ? window
//     : typeof global !== "undefined"
//     ? global
//     : undefined;
// }

// const g = getGlobal();

// // The add-in command functions need to be available in global scope
// g.action = action;

// function classifyEmail() {
//   // Get the current email item
//   Office.context.mailbox.item.getAsync(Office.CoercionType.Html, function (result) {
//       if (result.status === Office.AsyncResultStatus.Succeeded) {
//           // Get the email content in HTML format
//           var emailContent = result.value;

//           // Extract the text content from the HTML
//           var emailText = extractTextFromHtml(emailContent);

//           // Preprocess the email text using the email classification pipeline
//           preprocessAndClassifyEmail(emailText);
//       } else {
//           // Handle error if getting email content fails
//           console.error("Error getting email content:", result.error.message);
//       }
//   });
// }

//**********************************************************************************************************************************************

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    iconUrl: "https://localhost:3000/assets/icon-80.png", // Update with the correct icon URL
    persistent: true,
  };

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Call the classifyEmail function to start email classification
  classifyEmail();

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// The add-in command functions need to be available in global scope
g.action = action;

function classifyEmail() {
  // Get the current email item
  Office.context.mailbox.item.getAsync(Office.CoercionType.Html, function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      // Get the email content in HTML format
      var emailContent = result.value;

      // Extract the text content from the HTML
      var emailText = extractTextFromHtml(emailContent);

      // Preprocess the email text using the email classification pipeline
      preprocessAndClassifyEmail(emailText);
    } else {
      // Handle error if getting email content fails
      console.error("Error getting email content:", result.error.message);
    }
  });
}

function extractTextFromHtml(htmlContent) {
  // Use DOMParser to convert HTML string to a DOM object
  var parser = new DOMParser();
  var doc = parser.parseFromString(htmlContent, "text/html");

  // Extract the text from the DOM object
  return doc.body.innerText;
}

function preprocessAndClassifyEmail(emailText) {
  // Send the email text to the Flask server for preprocessing and classification
  const url = "your-flask-server-url/predict"; // Replace 'your-flask-server-url' with the actual URL of your Flask server

  fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ email_body: emailText }),
  })
    .then((response) => response.json())
    .then((result) => {
      // Handle the prediction result from the server here
      console.log("Prediction Result:", result);
      // Display the prediction result to the user using a notification or any other method
      showPredictionResult(result);
    })
    .catch((error) => {
      console.error("Error sending email to server:", error);
    });
}

function showPredictionResult(result) {
  // Show a notification or perform any other action to display the prediction result to the user
  const predictionMessage = result.prediction === "spam" ? "This email is classified as spam." : "This email is classified as ham.";
  const predictionNotification = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: predictionMessage,
    iconUrl: "https://localhost:3000/assets/icon-80.png", // Update with the correct icon URL
    persistent: true,
  };

  Office.context.mailbox.item.notificationMessages.replaceAsync("prediction", predictionNotification);
}




