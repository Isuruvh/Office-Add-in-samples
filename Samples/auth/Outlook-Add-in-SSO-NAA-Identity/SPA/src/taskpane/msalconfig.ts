// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// This file provides the default MSAL configuration for the add-in project.

import { LogLevel, Configuration } from "@azure/msal-browser";
import { createLocalUrl } from "./util";

/* global console */

export const clientId = "bb1874a8-699c-400a-959d-25671fe84a1b";
export const msalConfig: Configuration = {
  auth: {
    clientId,
    redirectUri: createLocalUrl("auth.html"),
    postLogoutRedirectUri: createLocalUrl("auth.html"),
  },
  cache: {
    cacheLocation: "localStorage",
  },
  system: {
    loggerOptions: {
      logLevel: LogLevel.Verbose,
      loggerCallback: (level: LogLevel, message: string) => {
        switch (level) {
          case LogLevel.Error:
            console.error(message);
            return;
          case LogLevel.Info:
            console.info(message);
            return;
          case LogLevel.Verbose:
            console.debug(message);
            return;
          case LogLevel.Warning:
            console.warn(message);
            return;
        }
      },
      piiLoggingEnabled: true,
    },
  },
};

// Default scopes to use in the fallback dialog.
export const defaultScopes = ["user.read", "files.read"];
