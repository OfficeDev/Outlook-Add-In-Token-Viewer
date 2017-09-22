// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE in the project root for license information.

Office.initialize = function () {
};

function showMessage(message, icon, event) {
  Office.context.mailbox.item.notificationMessages.replaceAsync('msg', {
    type: 'informationalMessage',
    icon: icon,
    message: message,
    persistent: false
  }, function(result){
    event.completed();
  });
}

function reportError(errorMessage, event) {
  Office.context.mailbox.item.notificationMessages.replaceAsync('error', {
    type: "errorMessage",
    message: errorMessage
  }, function (result) {
    event.completed();
  })
}

function showDialog(data, event) {
  // Convert the JSON validation data to query params
  var query = $.param(data);
  var dialogUrl = "https://localhost:44359/add-in/dialog/dialog.html?" + query;
  Office.context.ui.displayDialogAsync(dialogUrl, {displayInIframe: true});
  event.completed();
}

function validateIdToken(event) {
  Office.context.mailbox.getUserIdentityTokenAsync(function(result) {
    if (result.status == "succeeded") {
      idToken = result.value;

      // Send token to validation service
      $.ajax({
        type: "POST",
        url: "/api/validateexchangetoken",
        data: JSON.stringify(idToken),
        contentType: "application/json; charset=utf-8"
      }).done(function (data) {
          // Display dialog with validation results
          showDialog(data, event);
      }).fail(function (error) {
          reportError("Error validating ID token: " + error.status, event);
      });
    } else {
      reportError("Error retrieving ID token: " + result.error.message, event);
    }
  });
}