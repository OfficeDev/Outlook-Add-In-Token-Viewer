// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE in the project root for license information.
$(document).ready(function(){

  // Check is valid
  var isValid = getParameterByName("IsValid");
  if (isValid === "true") {
    $("#status-valid").show();
  } else {
    $("#status-invalid").show();
  }

  // Check user ID
  var userId = getParameterByName("ComputedUserId");
  if (userId) {
    $("#userid").text(userId);
    $("#unique-id").show();
  }

  // Check signature result
  var sigResult = getParameterByName("SignatureResult");
  generateIcon(sigResult, $("#signature-result"));

  // Check audience result
  var audResult = getParameterByName("AudienceResult");
  generateIcon(audResult, $("#audience-result"));

  // Check lifetime result
  var lifeResult = getParameterByName("LifetimeResult");
  generateIcon(lifeResult, $("#lifetime-result"));

  // Check version result
  var verResult = getParameterByName("VersionResult");
  generateIcon(verResult, $("#version-result"));

  // Check for message
  var message = getParameterByName("Message");
  if (message) {
    $("#message").text(message);
    $("#validation-message").show();
  }
});

function generateIcon(status, parent) {
  var iconClass = "ms-Icon--Warning";
  var iconColor = "ms=fontColor-warning";

  if (status === "passed") {
    iconClass = "ms-Icon--CheckMark";
    iconColor = "ms-fontColor-success";
  } else if (status === "failed") {
    iconClass = "ms-Icon--Error";
    iconColor = "ms-fontColor-error";
  }

  parent.addClass(iconColor);

  $("<i>")
    .addClass("ms-Icon")
    .addClass(iconClass)
    .attr("title", status)
    .appendTo(parent);
}

function getParameterByName(name, url) {
  if (!url) {
    url = window.location.href;
  }

  name = name.replace(/[\[\]]/g, "\\$&");

  var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
  results = regex.exec(url);
  if (!results) return null;
  if (!results[2]) return '';

  return decodeURIComponent(results[2].replace(/\+/g, " "));
}