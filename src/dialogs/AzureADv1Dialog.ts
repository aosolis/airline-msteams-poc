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

import * as builder from "botbuilder";
import * as config from "config";
import * as constants from "../constants";
import * as utils from "../utils";
import { IOAuth2Provider } from "../providers";
let uuidv4 = require("uuid/v4");

// Base identity dialog
export class AzureADv1Dialog extends builder.IntentDialog
{
    private authProvider: IOAuth2Provider;
    private providerDisplayName: string;
    private providerName: string;

    constructor() {
        super();
        this.providerName = constants.IdentityProvider.azureADv1;
    }

    // Register the dialog with the bot
    public register(bot: builder.UniversalBot, rootDialog: builder.IntentDialog): void {
        bot.dialog(constants.DialogId.AzureADv1, this);

        this.authProvider = bot.get(this.providerName) as IOAuth2Provider;
        this.providerDisplayName = this.authProvider.displayName;

        this.onBegin((session, args, next) => { this.onDialogBegin(session, args, next); });
        this.onDefault((session) => { this.onMessageReceived(session); });
    }

    // Handle start of dialog
    private async onDialogBegin(session: builder.Session, args: any, next: () => void): Promise<void> {
        switch (args.matched.input) {
            case "login":
                this.handleLogin(session);
                break;

            case "logout":
                this.handleLogout(session);
                break;
        }
        next();
    }

    // Handle message
    private async onMessageReceived(session: builder.Session): Promise<void> {
        let messageAsAny = session.message as any;
        if (messageAsAny.originalInvoke) {
            // This was originally an invoke message
            let event = messageAsAny.originalInvoke;
            if (event.name === "signin/verifyState") {
                await this.handleLoginCallback(session);
            } else {
                // Unrecognized invoke, exit the dialog
                session.endDialog();
            }
        } else {
            // See if we are waiting for a verification code and got one
            if (utils.isUserTokenPendingVerification(session, this.providerName)) {
                let verificationCode = utils.findVerificationCode(session.message.text);
                utils.validateVerificationCode(session, this.providerName, verificationCode);

                if (utils.getUserToken(session, this.providerName)) {
                    session.send(`Thank you for signing in.`);
                } else {
                    session.send(`Sorry, there was an error signing in to ${this.providerDisplayName}. Please try again.`);
                }
            } else {
                // Unrecognized input
                session.send("Sorry, I didn't understand.");
            }
        }

        session.endDialog();
    }

    // Handle user login callback
    private async handleLoginCallback(session: builder.Session): Promise<void> {
        let messageAsAny = session.message as any;
        let verificationCode = messageAsAny.originalInvoke.value.state;

        utils.validateVerificationCode(session, this.providerName, verificationCode);

        // End of auth flow: if the token is marked as validated, then the user is logged in

        if (utils.getUserToken(session, this.providerName)) {
            session.send(`Thank you for signing in.`);
        } else {
            session.send(`Sorry, there was an error signing in to ${this.providerDisplayName}. Please try again.`);
        }
        session.endDialog();
    }

    // Handle user login request
    private async handleLogin(session: builder.Session): Promise<void> {
        if (utils.getUserToken(session, this.providerName)) {
            // User is already logged in
            session.send(`You're already signed in.`);
            session.endDialog();
        } else {
            // Create the OAuth state, including a random anti-forgery state token
            let address = session.message.address;
            let state = JSON.stringify({
                security: uuidv4(),
                address: {
                    user: {
                        id: address.user.id,
                    },
                    conversation: {
                        id: address.conversation.id,
                    },
                },
            });
            utils.setOAuthState(session, this.providerName, state);

            // Create the authorization URL
            let authUrl = this.authProvider.getAuthorizationUrl(state);

            // Build the sign-in url
            let signinUrl = config.get("app.baseUri") + `/html/auth-start.html?authorizationUrl=${encodeURIComponent(authUrl)}`;

            // Send card with signin action
            let msg = new builder.Message(session)
                .addAttachment(new builder.HeroCard(session)
                    .text(`Click below to sign in to ${this.providerDisplayName}`)
                    .buttons([
                        new builder.CardAction(session)
                            .type("signin")
                            .value(signinUrl)
                            .title("Sign in"),
                    ]));
            session.send(msg);

            // The auth flow resumes when we handle the identity provider's OAuth callback in AuthBot.handleOAuthCallback()
        }
    }

    // Handle user logout request
    private async handleLogout(session: builder.Session): Promise<void> {
        if (!utils.getUserToken(session, this.providerName)) {
            session.send(`You're not currently signed in.`);
        } else {
            utils.setUserToken(session, this.providerName, null);
            session.send(`You're now signed out of ${this.providerDisplayName}.`);
        }
        session.endDialog();
    }
}
