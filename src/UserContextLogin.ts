// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

import * as config from "config";
import * as winston from "winston";
import * as constants from "./constants";
import { IAppDataStore } from "./storage/AppDataStore";
import { IOAuth2Provider } from "./providers";
import { Request, Response } from "express";
const uuidv4 = require("uuid/v4");

// Manage user context
export class UserContextLogin
{
    private replyUrl: string;

    constructor(
        private appDataStore: IAppDataStore,
        private oauthProvider: IOAuth2Provider,
    ) {
        this.replyUrl = config.get("app.baseUri") + "/usercontext/callback";
    }

    // Handle login to establish user context
    public async handleLogin(req: Request, res: Response): Promise<void> {
        // Create the url to the AAD authorize endpoint
        let state = uuidv4();
        let extraParams = {
            redirect_uri: this.replyUrl,
            prompt: "login",
        };
        let authUrl = this.oauthProvider.getAuthorizationUrl(state, extraParams);

        // Store the expected OAuth state
        await this.appDataStore.setAppDataAsync(constants.AppDataKey.oauthState, state);

        // Redirect the user to the AAD authorize endpoint
        res.redirect(authUrl);
    }

    // Handle OAuth callback
    public async handleCallback(req: Request, res: Response): Promise<void> {
        try {
            const incomingState = req.query.state as string;
            const authCode = req.query.code;

            // Check that the state matches
            let storedState = await this.appDataStore.getAppDataAsync(constants.AppDataKey.oauthState);
            await this.appDataStore.setAppDataAsync(constants.AppDataKey.oauthState, null);
            if (storedState !== incomingState) {
                throw new Error("OAuth state does not match");
            }

            // Redeem auth code for a token
            let userToken = await this.oauthProvider.getAccessTokenAsync(authCode, this.replyUrl);
            this.appDataStore.setAppDataAsync(constants.AppDataKey.userToken, userToken);
        } catch (e) {
            winston.error(`Error logging in: ${e.message}`, e);
            res.render("oauth-callback-error", {
                providerName: "Azure AD",
            });
        }
    }
}
