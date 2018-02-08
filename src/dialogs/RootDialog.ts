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
import * as constants from "../constants";
import { AzureADv1Dialog } from "./AzureADv1Dialog";
import * as teams from "../TeamsApi";
import * as utils from "../utils";

// Root dialog provides choices in identity providers
export class RootDialog extends builder.IntentDialog
{
    constructor() {
        super();
    }

    // Register the dialog with the bot
    public register(bot: builder.UniversalBot): void {
        bot.dialog(constants.DialogId.Root, this);

        this.onBegin((session, args, next) => { this.onDialogBegin(session, args, next); });
        this.onDefault((session) => { this.onMessageReceived(session); });

        new AzureADv1Dialog().register(bot, this);
        this.matches(/azureADv1/i, constants.DialogId.AzureADv1);
        this.matches(/createTeam/i, (session) => { this.handleCreateTeam(session); });
        this.matches(/archiveTeam/i, (session) => { this.handleArchiveTeam(session); });
        this.matches(/deleteTeam/i, (session) => { this.handleDeleteTeam(session); });
    }

    // Handle resumption of dialog
    public dialogResumed<T>(session: builder.Session, result: builder.IDialogResult<T>): void {
        session.send("Ok, tell me what to do");
    }

    private async handleCreateTeam(session: builder.Session): Promise<void> {
        let userInfo = utils.getUserToken(session, constants.IdentityProviders.azureADv1);
        let teamsApi = new teams.TeamsApi(userInfo.accessToken);

        // Create the team
        let team: teams.Team;
        try {
            let teamSettings: teams.Team = {
                memberSettings: {
                    allowAddRemoveApps: false,
                    allowCreateUpdateChannels: false,
                    allowCreateUpdateRemoveConnectors: false,
                    allowCreateUpdateRemoveTabs: false,
                    allowDeleteChannels: false,
                },
                guestSettings: {
                    allowCreateUpdateChannels: false,
                    allowDeleteChannels: false,
                },
            };
            team = await teamsApi.createTeamAsync("Test Team", "Test Team Description", teamSettings);
            session.userData.teamId = team.id;
            session.send(`Created a new team, group id is ${team.id}`);
        } catch (e) {
            session.send(`Error creating team: ${e.message}`);
            return;
        }

        // Set up channels
        let channelsToAdd: teams.Channel[] = [
            {
                displayName: "Flight",
                description: "Aircraft and flight path",
            },
            {
                displayName: "Crew",
            },
        ];
        let channelsAddPromises = channelsToAdd.map(async channel => {
            try {
                await teamsApi.createChannelAsync(team.id, channel.displayName, channel.description);
            } catch (e) {
                session.send(`Error creating channel ${channel.displayName}: ${e.message}`);
            }
        });
        await Promise.all(channelsAddPromises);

        // Add team members
        let membersToAdd = [ "fff2cfa8-0eb6-4fdc-9902-fa0ba06219b3", "0f429da5-2cbf-4d95-bc2c-16a1bef3ed1c", "f431b248-8e59-4afa-a307-054a1f220f24" ];
        let memberAddPromises = membersToAdd.map(async memberId => {
            try {
                await teamsApi.addMemberToTeamAsync(team.id, memberId);
            } catch (e) {
                session.send(`Error adding member ${memberId}: ${e.message}`);
            }
        });
        await Promise.all(memberAddPromises);

        session.send(`Done setting up team, group id is ${team.id}`);
    }

    private async handleArchiveTeam(session: builder.Session): Promise<void> {
        let userInfo = utils.getUserToken(session, constants.IdentityProviders.azureADv1);
        let teamsApi = new teams.TeamsApi(userInfo.accessToken);

        try {
            // Remove all team members
            let teamId = session.userData.teamId;
            let teamMembers = await teamsApi.getMembersAsync(teamId);
            console.log(`Found ${teamMembers.length} members in the team`);

            let memberRemovePromises = teamMembers.map(async member => {
                try {
                    await teamsApi.removeMemberFromTeamAsync(teamId, member.id);
                } catch (e) {
                    session.send(`Error removing member ${member.id}: ${e.message}`);
                }
            });
            await Promise.all(memberRemovePromises);

            // Rename group
            await teamsApi.updateGroupAsync(teamId, {
                displayName: "[ARCHIVED] Test team",
            });

            session.send(`Archived the team with group id ${teamId}`);
        } catch (e) {
            session.send(`Error archiving team: ${e.message}`);
        }
    }

    private async handleDeleteTeam(session: builder.Session): Promise<void> {
        let userInfo = utils.getUserToken(session, constants.IdentityProviders.azureADv1);
        let teamsApi = new teams.TeamsApi(userInfo.accessToken);

        try {
            let teamId = session.userData.teamId;
            await teamsApi.deleteTeamAsync(teamId);
            delete session.userData.teamId;
            session.send(`Deleted the team with group id ${teamId}`);
        } catch (e) {
            session.send(`Error deleting team: ${e.message}`);
        }
    }

    // Handle start of dialog
    private async onDialogBegin(session: builder.Session, args: any, next: () => void): Promise<void> {
        session.dialogData.isFirstTurn = true;
        this.promptForIdentityProvider(session);
        next();
    }

    // Handle message
    private async onMessageReceived(session: builder.Session): Promise<void> {
        if (!session.dialogData.isFirstTurn) {
            // Unrecognized input
            session.send("I didn't understand that.");
            this.promptForIdentityProvider(session);
        } else {
            delete session.dialogData.isFirstTurn;
        }
    }

    // Prompt the user to pick an identity provider
    private promptForIdentityProvider(session: builder.Session): void {
        let msg = new builder.Message(session)
            .addAttachment(new builder.ThumbnailCard(session)
                .title("Select an identity provider")
                .buttons([
                    builder.CardAction.messageBack(session, "{}", "AzureAD (v1)")
                        .displayText("AzureAD (v1)")
                        .text("AzureADv1"),
                ]));
        session.send(msg);
    }
}
