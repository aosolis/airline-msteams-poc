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

// Wrapper around the Teams Graph APIs
export abstract class TeamsApi {

    protected accessToken: string;
    protected expirationTime: number;

    // Refresh the access token
    protected abstract async refreshAccessTokenAsync(): Promise<void>;

    // Create a new team
    public async createTeamAsync(displayName: string, description: string, mailNickname: string, teamSettings: Team): Promise<Team>
    {
        await this.refreshAccessTokenAsync();

        let newGroup = await this.createGroupAsync(displayName, description, mailNickname);

        // The group may not be created yet, so retry up to 3 times, waiting 10 seconds between retries
        const maxAttempts = 3;
        const retryWaitInMilliseconds = 10000;

        let lastError: Error;
        let attemptCount = 0;
        while (attemptCount < maxAttempts) {
            attemptCount++;
            try {
                return await this.createTeamFromGroupAsync(newGroup.id, teamSettings);
            } catch (e) {
                lastError = e;
                // Allow retry if error is 404
                if (e.statusCode === 404) {
                    await new Promise((resolve, reject) => {
                        setTimeout(() => { resolve(); }, retryWaitInMilliseconds);
                    });
                } else {
                    break;
                }
            }
        }

        throw lastError;
    }

    // Delete a team (group)
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

    // Add a member to a team (group)
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

    // Create a new channel
    public async createChannelAsync(groupId: string, displayName: string, description?: string): Promise<Channel> {
        await this.refreshAccessTokenAsync();

        let requestBody: Channel = {
            displayName: displayName,
        };
        if (description) {
            requestBody.description = description;
        }

        let options = {
            url: `${graphBaseUrl}/groups/${groupId}/channels`,
            body: requestBody,
            json: true,
            headers: {
                "Authorization": `Bearer ${this.accessToken}`,
            },
        };
        return await request.post(options);
    }

    // Create a new group
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
}

// Teams API that uses delegated user permissions
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

// Teams API that uses application permissions
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
        if (this.accessToken && (Date.now() < this.expirationTime)) {
            return;
        }

        let accessTokenUrl = `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`;
        let params = {
            grant_type: "client_credentials",
            client_id: this.appId,
            client_secret: this.appPassword,
            scope: "https://graph.microsoft.com/.default",
        };

        let response = await request.post({ url: accessTokenUrl, form: params, json: true });
        this.accessToken = response.access_token;
        this.expirationTime = Date.now() + (response.expires_in * 1000) - (60 * 100);
    }
}
