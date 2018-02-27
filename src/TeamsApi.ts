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

import * as request from "request-promise";
import * as winston from "winston";

// =========================================================
// Teams Graph API
// =========================================================

const graphBaseUrl = "https://graph.microsoft.com/testTeamsTestEnv";

export interface DirectoryObject {
    id: string;
}

export interface Group {
    id?: string;
    displayName?: string;
    description?: string;
    members?: DirectoryObject[];
    owners?: DirectoryObject[];
    groupTypes?: string[];
    mailEnabled?: boolean;
    mailNickname?: string;
    securityEnabled?: boolean;
    visibility?: "private" | "public";
    creationOptions?: string[];
}

export interface Team {
    id?: string;
    memberSettings?: TeamMemberSettings;
    messagingSettings?: TeamMessagingSettings;
    funSettings?: TeamFunSettings;
    guestSettings?: TeamGuestSettings;
}

export interface Channel {
    id?: string;
    displayName?: string;
    description?: string;
}

export interface TeamMemberSettings {
    allowCreateUpdateChannels?: boolean;
    allowDeleteChannels?: boolean;
    allowAddRemoveApps?: boolean;
    allowCreateUpdateRemoveTabs?: boolean;
    allowCreateUpdateRemoveConnectors?: boolean;
}

export interface TeamMessagingSettings {
    allowUserEditMessages?: boolean;
    allowUserDeleteMessages?: boolean;
    allowOwnerDeleteMessages?: boolean;
    allowTeamMentions?: boolean;
    allowChannelMentions?: boolean;
}

export interface TeamFunSettings {
    allowGiphy?: boolean;
    giphyContentRating?: "moderate" | "strict";
    allowStickersAndMemes?: boolean;
    allowCustomMemes?: boolean;
}

export interface TeamGuestSettings {
    allowCreateUpdateChannels?: boolean;
    allowDeleteChannels?: boolean;
}

// Wrapper around the Microsoft Graph APIs for Teams
export abstract class TeamsApi {

    protected accessToken: string;
    protected expirationTime: number;

    // Refresh the access token
    protected abstract async refreshAccessTokenAsync(): Promise<void>;

    // Create a new team
    // Parameters:
    //   - displayName: team display name
    //   - description: team description
    //   - mailNickname: e-mail alias for the team (must be unique within the tenant)
    //   - teamSettings: team settings
    public async createTeamAsync(displayName: string, description: string, mailNickname: string, teamSettings: Team): Promise<Team>
    {
        await this.refreshAccessTokenAsync();

        // First create a modern group, which we will then convert to a team
        let newGroup = await this.createGroupAsync(displayName, description, mailNickname);
        winston.info(`Created new group ${newGroup.id}`);

        // The operation to create a team from a group can fail, particularly when the group is newly-created,
        // as knowledge about the newly-created group and its owners/members propagates through Azure AD.
        // To work around this, we retry the attempt several times, with a delay between each retry.

        let attemptCount = 0;
        let migrateTeamError: any;
        const maxAttempts = 5;                      // Max retries
        const retryWaitInMilliseconds = 10000;      // Delay between each retry (10 s)

        while (attemptCount < maxAttempts) {
            attemptCount++;

            try {
                return await this.createTeamFromGroupAsync(newGroup.id, teamSettings);
            } catch (e) {
                migrateTeamError = e;
                winston.warn(`Error converting group ${newGroup.id} to a team (attempt #${attemptCount}): ${e.message}`, e);

                if (e.statusCode === 404 || e.statusCode === 500) {
                    // Retry if status is 404 Not Found or 500 Internal Server Error
                    await new Promise((resolve, reject) => {
                        setTimeout(() => { resolve(); }, retryWaitInMilliseconds);
                    });
                } else if (e.statusCode === 409) {
                    // If status is 409 Conflict, a previous attempt succeeded behind the scenes
                    winston.info(`Treating conflict as success of a previous attempt`);
                    return await this.getTeamSettingsAsync(newGroup.id);
                } else {
                    break;
                }
            }
        }

        // Attempt to delete the group if conversion to a team failed
        if (migrateTeamError) {
            winston.error(`Error converting group ${newGroup.id} to a team, will delete it: ${migrateTeamError.message}`, migrateTeamError);
            try {
                await this.deleteGroupAsync(newGroup.id);
            } catch (e) {
                winston.error(`Failed to delete the group ${newGroup.id}: ${e.message}`, e);
            }
        }

        // If we get here there must have been an error
        throw migrateTeamError;
    }

    // Delete a team (group)
    // Parameters:
    //   - groupId: team (group) id
    public async deleteGroupAsync(groupId: string): Promise<void> {
        await this.refreshAccessTokenAsync();

        let options = {
            url: `${graphBaseUrl}/groups/${groupId}`,
            json: true,
            headers: {
                "Authorization": `Bearer ${this.accessToken}`,
            },
        };
        await request.delete(options);
    }

    // Add an owner to a team (group)
    // Parameters:
    //   - groupId: team (group) id
    //   - userObjectId: AAD object id of the owner to add
    public async addOwnerToGroupAsync(groupId: string, userObjectId: string): Promise<void> {
        await this.refreshAccessTokenAsync();

        let requestBody = {
            "@odata.id": `https://graph.microsoft.com/beta/directoryObjects/${userObjectId}`,
        };
        let options = {
            url: `${graphBaseUrl}/groups/${groupId}/owners/$ref`,
            body: requestBody,
            json: true,
            headers: {
                "Authorization": `Bearer ${this.accessToken}`,
            },
        };
        await request.post(options);
    }

    // Remove an owner from a team (group)
    // Parameters:
    //   - groupId: team (group) id
    //   - userObjectId: AAD object id of the owner to remove
    public async removeOwnerFromGroupAsync(groupId: string, userObjectId: string): Promise<void> {
        await this.refreshAccessTokenAsync();

        let options = {
            url: `${graphBaseUrl}/groups/${groupId}/owners/${userObjectId}/$ref`,
            json: true,
            headers: {
                "Authorization": `Bearer ${this.accessToken}`,
            },
        };
        await request.delete(options);
    }

    // Get the owners of a team (group)
    // Parameters:
    //   - groupId: team (group) id
    // Returns: a list of team owners, or an empty list if there are no owners
    public async getOwnersOfGroupAsync(groupId: string): Promise<DirectoryObject[]> {
        await this.refreshAccessTokenAsync();

        let options = {
            url: `${graphBaseUrl}/groups/${groupId}/owners`,
            json: true,
            headers: {
                "Authorization": `Bearer ${this.accessToken}`,
            },
        };
        let responseBody = await request.get(options);
        return responseBody.value || [];
    }

    // Add a member to a team (group)
    // Parameters:
    //   - groupId: team (group) id
    //   - userObjectId: AAD object id of the member to add
    public async addMemberToGroupAsync(groupId: string, userObjectId: string): Promise<void> {
        await this.refreshAccessTokenAsync();

        let requestBody = {
            "@odata.id": `https://graph.microsoft.com/beta/directoryObjects/${userObjectId}`,
        };
        let options = {
            url: `${graphBaseUrl}/groups/${groupId}/members/$ref`,
            body: requestBody,
            json: true,
            headers: {
                "Authorization": `Bearer ${this.accessToken}`,
            },
        };
        await request.post(options);
    }

    // Remove a member from a team (group)
    // Parameters:
    //   - groupId: team (group) id
    //   - userObjectId: AAD object id of the member to remove
    public async removeMemberFromGroupAsync(groupId: string, userObjectId: string): Promise<void> {
        await this.refreshAccessTokenAsync();

        let options = {
            url: `${graphBaseUrl}/groups/${groupId}/members/${userObjectId}/$ref`,
            json: true,
            headers: {
                "Authorization": `Bearer ${this.accessToken}`,
            },
        };
        await request.delete(options);
    }

    // Get the members of a team (group)
    // Parameters:
    //   - groupId: team (group) id
    // Returns: list of team members, or an empty list if the team has no members
    public async getMembersOfGroupAsync(groupId: string): Promise<DirectoryObject[]> {
        await this.refreshAccessTokenAsync();

        let options = {
            url: `${graphBaseUrl}/groups/${groupId}/members`,
            json: true,
            headers: {
                "Authorization": `Bearer ${this.accessToken}`,
            },
        };
        let responseBody = await request.get(options);
        return responseBody.value || [];
    }

    // Get group information
    // Parameters:
    //   - groupId: team (group) id
    public async getGroupAsync(groupId: string): Promise<Group> {
        await this.refreshAccessTokenAsync();

        let options = {
            url: `${graphBaseUrl}/groups/${groupId}`,
            json: true,
            headers: {
                "Authorization": `Bearer ${this.accessToken}`,
            },
        };
        return await request.get(options);
    }

    // Update group information
    // Parameters:
    //   - groupId: team (group) id
    //   - groupUpdates: new group information to update, populate only the properties that need to be updated
    public async updateGroupAsync(groupId: string, groupUpdates: Group): Promise<void> {
        await this.refreshAccessTokenAsync();

        let options = {
            url: `${graphBaseUrl}/groups/${groupId}`,
            body: groupUpdates,
            json: true,
            headers: {
                "Authorization": `Bearer ${this.accessToken}`,
            },
        };
        await request.patch(options);
    }

    // Create a new group
    // Parameters:
    //   - displayName: group display name
    //   - description: group description
    //   - mailNickname: e-mail alias for the group (must be unique within the tenant)
    private async createGroupAsync(displayName: string, description: string, mailNickname: string): Promise<Group>
    {
        let requestBody: Group = {
            displayName: displayName,
            description: description,
            mailEnabled: true,
            mailNickname: mailNickname,
            securityEnabled: false,
            visibility: "private",
            groupTypes: [ "unified" ],
        };

        let options = {
            url: `${graphBaseUrl}/groups`,
            body: requestBody,
            json: true,
            headers: {
                "Authorization": `Bearer ${this.accessToken}`,
            },
        };
        return await request.post(options);
    }

    // Create a team given an existing group
    // Parameters:
    //   - groupId: id of the group to convert into a team
    //   - teamSettings: team settings
    private async createTeamFromGroupAsync(groupId: string, teamSettings: Team): Promise<Team>
    {
        let options = {
            url: `${graphBaseUrl}/groups/${groupId}/team`,
            body: teamSettings,
            json: true,
            headers: {
                "Authorization": `Bearer ${this.accessToken}`,
            },
        };
        return await request.put(options);
    }

    // Get settings for an existing team
    // Parameters:
    //   - groupId: team (group) id
    private async getTeamSettingsAsync(groupId: string): Promise<Team>
    {
        let options = {
            url: `${graphBaseUrl}/groups/${groupId}/team`,
            json: true,
            headers: {
                "Authorization": `Bearer ${this.accessToken}`,
            },
        };
        return await request.get(options);
    }
}

// Teams API that uses delegated user context
export class UserContextTeamsApi extends TeamsApi {

    constructor(
        accessToken: string,
        expirationTime: number,
    )
    {
        super();
        this.accessToken = accessToken;
        this.expirationTime = expirationTime;
    }

    // Refresh the access token
    protected async refreshAccessTokenAsync(): Promise<void> {
        // This does not support refresh
    }
}

// Teams API that uses application context
export class AppContextTeamsApi extends TeamsApi {

    constructor(
        private tenantId: string,
        private appId: string,
        private appPassword: string,
    )
    {
        super();
    }

    // Refresh the access token
    protected async refreshAccessTokenAsync(): Promise<void> {
        // Check if the token requires refresh
        if (this.accessToken && (Date.now() < this.expirationTime)) {
            return;
        }

        // Get an access token using the client_credentials grant
        let accessTokenUrl = `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`;
        let params = {
            grant_type: "client_credentials",
            client_id: this.appId,
            client_secret: this.appPassword,
            scope: "https://graph.microsoft.com/.default",
        };

        let response = await request.post({ url: accessTokenUrl, form: params, json: true });
        this.accessToken = response.access_token;

        const expirationTimeBufferInSeconds = 60;       // Include a 1-minute buffer in the access token expiration time
        this.expirationTime = Date.now() + ((response.expires_in - expirationTimeBufferInSeconds) * 1000);
    }
}
