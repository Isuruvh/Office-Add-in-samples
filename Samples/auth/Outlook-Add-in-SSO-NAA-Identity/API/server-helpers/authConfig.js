// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

const authConfig = {
    credentials: {
        clientID: "bb1874a8-699c-400a-959d-25671fe84a1b"
    },
    metadata: {
        authority: "login.microsoftonline.com",
    },
    protectedRoutes: {
        todolist: {
            endpoint: "/api/todolist",
            delegatedPermissions: {
                read: ["Todolist.Read", "Todolist.ReadWrite"],
                readWrite: ["Todolist.ReadWrite"]
            }
        }
    }
}

module.exports = authConfig;
