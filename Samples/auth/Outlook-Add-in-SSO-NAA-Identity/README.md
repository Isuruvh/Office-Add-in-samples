# Outlook SSO To-Do Add-in

This project is a personal Outlook add-in built around the Microsoft sample for SSO and Nested App Authentication (NAA). It allows a user to sign in through Microsoft Entra ID, request an access token, and send that token to a protected backend API to manage a personal to-do list tied to the signed-in user.

This project is intended as a practical base for ongoing experimentation and future work with Outlook add-ins, Microsoft identity, and secure API access patterns.

## Overview

The add-in is made of two main parts:

- SPA client: the Outlook task pane UI built with JavaScript/TypeScript
- API backend: a Node.js Express server that validates the access token and stores/retrieves user-specific tasks

The overall flow is:

1. User signs in to the Outlook add-in.
2. The client requests an Entra ID access token using MSAL.
3. The token is used to call the protected backend API.
4. The API validates the token and returns data for the current user only.
5. Tasks are stored and retrieved per user using the authenticated identity.

## Purpose

This project demonstrates how to:

- use Microsoft SSO in an Outlook add-in
- request delegated access tokens for a protected API
- validate tokens on the server side
- keep data scoped to the signed-in user
- build a practical to-do list add-in for Outlook

## Features

- Outlook task pane experience
- Microsoft Entra ID sign-in
- Access token acquisition using MSAL
- Secure API calls with access tokens
- User-scoped to-do list operations
- Create, read, and delete tasks
- Backend token validation with JWT checks

## Tech Stack

- Outlook add-ins
- JavaScript / TypeScript
- Node.js
- Express.js
- Microsoft Authentication Library (MSAL)
- Microsoft Entra ID
- LowDB for local data storage in the sample

## Project Structure

```text
Outlook-Add-in-SSO-NAA-Identity/
├── API/
│   ├── app.js
│   ├── package.json
│   ├── server-helpers/
│   ├── db/
│   └── bin/
├── SPA/
│   ├── package.json
│   ├── manifest.xml
│   ├── src/
│   └── webpack.config.js
├── images/
├── README.md
├── .gitignore
└── package-lock.json
```

## Prerequisites

Before running this project, make sure you have:

- Microsoft 365 account with Outlook access
- Outlook desktop or Outlook on the web for testing
- Node.js installed
- npm installed
- Azure / Microsoft Entra app registration access
- Permission to sideload Outlook add-ins in your environment

## Microsoft Entra Setup

You need to configure an Entra app registration for the sample.

### Required configuration

- Register an application in Microsoft Entra ID
- Expose an API and define delegated scopes such as:
  - `Todolist.Read`
  - `Todolist.ReadWrite`
- Add the app as a client application for itself if required
- Add redirect URIs for local development, such as:
  - `http://localhost:3000`
  - `http://localhost:3000/redirect`
  - `brk-multihub://localhost:3000`
- Ensure the app manifest has `requestedAccessTokenVersion` set to `2`

### Important configuration files

Update the application and tenant values in the project settings:

- `API/server-helpers/authConfig.js`
- `SPA/src/taskpane/msalconfig.ts`
- `SPA/src/taskpane/authConfig.ts`

Use the Application (client) ID and Directory (tenant) ID from your Entra app registration.

## Local Development Setup

### 1. Install API dependencies

```bash
cd API
npm install
```

### 2. Start the API server

```bash
npm start
```

The API runs locally and validates incoming access tokens and serves the to-do data endpoints.

### 3. Install SPA dependencies

Open a second terminal:

```bash
cd SPA
npm install
```

### 4. Start the add-in

```bash
npm start
```

This launches the Outlook add-in in development mode.

## Run in Outlook

1. Open Outlook.
2. Sign in with the same Microsoft 365 account used for the app registration.
3. Open an email or item in Outlook.
4. Launch the task pane for the add-in.
5. Add or retrieve your tasks from the UI.

## How the app works

### Client side

The Outlook add-in uses MSAL to obtain a token for the API resource. The token is requested for the resource and scopes needed by the backend.

### Server side

The Node.js API validates the token, checks audience and issuer claims, and then authorizes access to the to-do operations. Tasks are scoped to the signed-in user identity.

## Security Notes

This project is a sample and is meant for learning and development.

Important points:

- Do not expose tokens in logs or debug output.
- Do not pass ID tokens across the network to a backend resource.
- Only use access tokens for backend API validation.
- In production, use tighter CORS policies and secure deployment controls.

## Current Development Notes

This project is a working sample for personal extension and testing. The default behavior is based on the official Microsoft sample, but this repository is intended to serve as a custom development sandbox for future Outlook SSO experiments.

## Planned Improvements

Possible future enhancements include:

- better error handling and user feedback
- improved UI styling and responsiveness
- persistent storage with a real backend database
- adding filtering and task categories
- support for multiple Outlook scenarios beyond a basic to-do list
- additional secure identity validation and logging

## Troubleshooting

### Add-in does not start

- confirm Node.js and npm are installed
- run `npm install` in both `API` and `SPA`
- verify the app registration values are correct

### Token request fails

- check Entra app registration values
- confirm redirect URIs match your local development setup
- verify the required delegated permissions are granted

### API returns unauthorized

- confirm the access token is being sent correctly
- validate the token's audience and issuer
- check the app registration's resource and scope configuration

## References

- Microsoft Office add-ins documentation
- Microsoft Entra ID documentation
- Outlook add-in SSO and identity samples

## License

This project follows the licensing terms of the parent repository and applicable Microsoft sample usage guidelines. Check the repository root for license information.

## Notes for Future Work

This README is intentionally kept simple and practical so it can be reused as a personal project baseline. If you continue expanding the app, update this document as the solution evolves with architecture changes, security decisions, and feature improvements.


### Access to data

Controllers should check if the presented access token has the necessary permissions to access the data, depending on the type of permission. This is illustrated in [todolist.js](./API/controllers/todolist.js):

```JavaScript
exports.getTodos = (req, res, next) => {
    // Check that caller has the delegated todolist.read permission from the user.
    if (hasRequiredDelegatedPermissions(req.authInfo, authConfig.protectedRoutes.todolist.delegatedPermissions.read)) {

        try {
            const owner = req.authInfo['oid'];

            const todos = db.get('todos')
                .filter({ owner: owner })
                .value();

            res.status(200).send(todos);
        } catch (error) {
            next(error);
        }
    } else {
        next(new Error('User does not have the required permissions.'));
    }
}
```

When granting access to data based on scopes, be sure to follow [the principle of least privilege](https://docs.microsoft.com/azure/active-directory/develop/secure-least-privileged-access).

## Learn more about the Microsoft identity platform

- [Microsoft identity platform (Microsoft Entra ID for developers)](https://docs.microsoft.com/azure/active-directory/develop/)
- [Overview of Microsoft Authentication Library (MSAL)](https://docs.microsoft.com/azure/active-directory/develop/msal-overview)
- [Understanding Microsoft Entra application consent experiences](https://docs.microsoft.com/azure/active-directory/develop/application-consent-experience)
- [Understand user and admin consent](https://docs.microsoft.com/azure/active-directory/develop/howto-convert-app-to-be-multi-tenant#understand-user-and-admin-consent)
- [Microsoft identity platform and OpenID Connect protocol](https://docs.microsoft.com/azure/active-directory/develop/v2-protocols-oidc)
- [Microsoft identity platform ID Tokens](https://docs.microsoft.com/azure/active-directory/develop/id-tokens)

For more information about how OAuth 2.0 protocols work in this scenario and other scenarios, see [Authentication Scenarios for Microsoft Entra ID](https://docs.microsoft.com/azure/active-directory/develop/authentication-flows-app-scenarios).

## Questions and feedback

- Did you experience any problems with the sample? [Create an issue](https://github.com/OfficeDev/Office-Add-in-samples/issues/new/choose) and we'll help you out.
- We'd love to get your feedback about this sample. Go to our [Office samples survey](https://aka.ms/OfficeSamplesSurvey) to give feedback and suggest improvements.
- For general questions about developing Office Add-ins, go to [Microsoft Q&A](https://learn.microsoft.com/answers/topics/office-js-dev.html) using the office-js-dev tag.

## Copyright

Copyright (c) 2025 Microsoft Corporation. All rights reserved.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

<img src="https://pnptelemetry.azurewebsites.net/pnp-officeaddins/samples/outlook-add-in-sso-naa-identity" />