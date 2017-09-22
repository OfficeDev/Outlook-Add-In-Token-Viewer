// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See LICENSE in the project root for license information.
(function(){
  "use strict";

  var restCallbackSupported = false;
  var idToken = "";
  var ewsToken = "";
  var restToken = "";
  var ssoToken = "";

  // The Office initialize function must be run each time a new page is loaded
  Office.initialize = function(reason){
    $(document).ready(function(){
      var ToggleElements = document.querySelectorAll(".ms-Toggle");
      for(var i = 0; i < ToggleElements.length; i++) {
          new fabric["Toggle"](ToggleElements[i]);
      }

      $("#parse-id-token-toggle").click(function() {
        showIdToken($("#parse-id-token-toggle").is(":checked"));
      });

      $("#parse-ews-token-toggle").click(function() {
        showEwsToken($("#parse-ews-token-toggle").is(":checked"));
      });

      $("#parse-rest-token-toggle").click(function() {
        showRestToken($("#parse-rest-token-toggle").is(":checked"));
      });

      $("#parse-sso-token-toggle").click(function() {
        showSsoToken($("#parse-sso-token-toggle").is(":checked"));
      });

      if (Office.context.mailbox.restUrl !== undefined) {
        restCallbackSupported = true;
      }

      getTokens();
    });
  };

  // Displays the callback token for the current item
  function getTokens(){
    // Identity token
    Office.context.mailbox.getUserIdentityTokenAsync(function(result) {
      if (result.status == "succeeded") {
        idToken = result.value;
        showIdToken($("parse-id-token-toggle").is(":checked"));
      } else {
        reportError("id-token", result.error);
      }
    });

    // EWS token
    Office.context.mailbox.getCallbackTokenAsync(function(result) {
      if (result.status == "succeeded") {
        ewsToken = result.value;
        showEwsToken($("#parse-ews-token-toggle").is(":checked"));
      }
      else {
        reportError("ews-token", result.error);
      }
    });

    // REST token
    if (restCallbackSupported) {
      // Get the REST token
      Office.context.mailbox.getCallbackTokenAsync(
        {isRest: true}, function (result) {
          if (result.status == "succeeded") {
            restToken = result.value;
            showRestToken($("#parse-rest-token-toggle").is(":checked"));
          }
          else {
            reportError("rest-token", result.error);
          }
        }
      );
    } else {
      reportWarning("rest-token", "REST callback token not supported by client");
    }

    // Get SSO token
    if (Office.context.auth && Office.context.auth.getAccessTokenAsync !== undefined) {
      Office.context.auth.getAccessTokenAsync(function (result){
        if (result.status == "succeeded") {
          ssoToken = result.value;
          showSsoToken($("#parse-sso-token-toggle").is(":checked"))
        } else{
          reportError("sso-token", result.error);
        }
      });
    } else {
      reportWarning("sso-token", "SSO token is not supported by client");
    }
  }

  function reportError(target, errorMsg) {
    $("#" + target).text(JSON.stringify(errorMsg, null, 2));
    $("#" + target).parent().addClass("ms-bgColor-error");
    $("#" + target).parent().siblings(".ms-Toggle").hide();
  }

  function reportWarning(target, warningMsg) {
    $("#" + target).text(JSON.stringify(warningMsg, null, 2));
    $("#" + target).parent().addClass("ms-bgColor-warning");
    $("#" + target).parent().siblings(".ms-Toggle").hide();
  }

  function showIdToken(parseToken) {
    if (parseToken) {
      $("#id-token").text(JSON.stringify(jwt_decode(idToken), null, 2));
    } else {
      $("#id-token").text(idToken);
    }
  }
  
  function showEwsToken(parseToken) {
    if (parseToken) {
      $("#ews-token").text(JSON.stringify(jwt_decode(ewsToken), null, 2));
    } else {
      $("#ews-token").text(ewsToken);
    }
  }

  function showRestToken(parseToken) {
    if (parseToken) {
      $("#rest-token").text(JSON.stringify(jwt_decode(restToken), null, 2)).show();
    } else {
      $("#rest-token").text(restToken).show();
    }
  }

  function showSsoToken(parseToken) {
    if (parseToken) {
      $("#sso-token").text(JSON.stringify(jwt_decode(ssoToken), null, 2)).show();
    } else {
      $("#sso-token").text(ssoToken).show();
    }
  }
})();