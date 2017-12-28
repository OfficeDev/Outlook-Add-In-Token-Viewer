# Outlook Add-in: Token Viewer

This sample add-in demonstrates how to retrieve and parse the various tokens available to an Outlook add-in, including:

- The [Exchange user identity token](https://docs.microsoft.com/en-us/outlook/add-ins/inside-the-identity-token)
- The [callback tokens](https://dev.office.com/reference/add-ins/outlook/1.5/Office.context.mailbox?product=outlook) used for making EWS or REST calls
- The [single-sign-on token](https://docs.microsoft.com/en-us/outlook/add-ins/authenticate-a-user-with-an-sso-token)

## Key components

This sample includes two main parts, the add-in that retrieves and displays the tokens, and the back-end Web API that does validation of the Exchange user-identity token.

### Add-in

The add-in is contained in the [TokenValidationService/Add-in](TokenValidationService/Add-in) folder.

### Web API

The Web API is implemented in the **TokenValidationService** project.

## Configure the sample

### Register the add-in

Because this sample retrieves an SSO token, you must register the add-in in the [Application Registration Portal](https://apps.dev.microsoft.com/) to get an app ID and secret.

1. Register an app using the instructions at https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token#registering-your-add-in. You do not need to register a Web app unless you intend to modify this sample to call Microsoft Graph.
1. Open the add-in manifest [manifest-outlook-token-viewer.xml](TokenValidationService/Add-in/manifest-outlook-token-viewer.xml).
1. Replace all instances of `YOUR_APP_ID` in the manifest with the app ID generated in your app registration.
1. Update the `<Scopes>` element in the manifest to reflect the permissions you configured in the **Microsoft Graph Permissions** section of your app registration.
1. Open the [Web.config](TokenValidationService/Web.config) file and replace all instances of `YOUR_APP_ID` in the manifest with the app ID generated in your app registration.

### Provide user consent

Because you will sideload this add-in, you need to provide user consent to enable the SSO flow. Follow the instructions at https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token#providing-consent-when-sideloading-an-add-in to provide consent.

## Run the sample

### Sideload the add-in

Follow the instructions at https://docs.microsoft.com/outlook/add-ins/sideload-outlook-add-ins-for-testing to sideload [manifest-outlook-token-viewer.xml](TokenValidationService/Add-in/manifest-outlook-token-viewer.xml).

> **Note:** This step only needs to be done once *unless* you modify the manifest. If you modify the manifest, you need to remove the add-in, then sideload the updated manifest.

### Run the project

Open **TokenValidationService.sln** in Visual Studio and press **F5** to debug the project. Select a message in Outlook and use the add-in buttons to view the tokens or validate the identity token.

## Copyright

Copyright (c) Microsoft. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
