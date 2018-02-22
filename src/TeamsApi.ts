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
export class TeamsApi {

    private accessToken: string;
    private expirationTime: number;

    constructor(
        private tenantId: string,
        private appId: string,
        private appPassword: string,
    )
    {
    }

    // Create a new team
    public async createTeamAsync(displayName: string, description: string, mailNickname: string, teamSettings: Team): Promise<Team>
    {
        await this.refreshAccessTokenAsync();

        let newGroup = await this.createGroupAsync(displayName, description, mailNickname);
        let newTeam: Team;

        // The group may not be created yet, so retry up to 3 times, waiting 10 seconds between retries
        let attemptCount = 0;
        while (attemptCount < 3) {
            attemptCount++;
            try {
                newTeam = await this.createTeamFromGroupAsync(newGroup.id, teamSettings);
                if (newTeam) {
                    break;
                }
            } catch (e) {
                // Check if error is 404; if so, retry
                if (true) {
                    await new Promise((resolve, reject) => {
                        setTimeout(() => { resolve(); }, 10000);
                    });
                }
            }
        }

        return newTeam;
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

    // Get an access token
    private async refreshAccessTokenAsync(): Promise<void> {
        // if (this.accessToken && (Date.now() < this.expirationTime)) {
        //     return;
        // }

        // let accessTokenUrl = `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`;
        // let params = {
        //     grant_type: "client_credentials",
        //     client_id: this.appId,
        //     client_secret: this.appPassword,
        //     scope: "https://graph.microsoft.com/.default",
        // };

        // let response = await request.post({ url: accessTokenUrl, form: params, json: true });
        // this.accessToken = response.access_token;
        // this.expirationTime = Date.now() + (response.expires_in * 1000) - (60 * 100);
        this.accessToken = "eyJ0eXAiOiJKV1QiLCJub25jZSI6IkFRQUJBQUFBQUFCSGg0a21TX2FLVDVYcmp6eFJBdEh6RmdQbDJDMTdIRV9YckpMQUo3RmdablpGOFkxR0o4eDBwaEpVU2Z1M1JWRXowWmJ5a2MxdUdXeVc5WGdYRGdVLVVJVFUzUE56SFA4YkFxcW1yd0MtSVNBQSIsImFsZyI6IlJTMjU2IiwieDV0IjoiU1NRZGhJMWNLdmhRRURTSnhFMmdHWXM0MFEwIiwia2lkIjoiU1NRZGhJMWNLdmhRRURTSnhFMmdHWXM0MFEwIn0.eyJhdWQiOiJodHRwczovL2dyYXBoLm1pY3Jvc29mdC5jb20iLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9hNWJiYjlkZi0wNmNjLTQ3ZjQtOGYyNC05ODFhMjAyNGI5NGMvIiwiaWF0IjoxNTE5MzI4NDk2LCJuYmYiOjE1MTkzMjg0OTYsImV4cCI6MTUxOTMzMjM5NiwiYWNyIjoiMSIsImFpbyI6IlkyTmdZRmlTS01uVnkzeCtjL1Y2WXpIVHJaTTNUMU5WVTI5eis3SFAyOTA2b3JPL1B4VUEiLCJhbXIiOlsicHdkIl0sImFwcF9kaXNwbGF5bmFtZSI6IkVtaXJhdGVzIiwiYXBwaWQiOiIxOWI5MjEzZS0yODM1LTRjNWMtYmRhZS03NzkzYjRmNDE3NzQiLCJhcHBpZGFjciI6IjEiLCJmYW1pbHlfbmFtZSI6IkFkbWluaXN0cmF0b3IiLCJnaXZlbl9uYW1lIjoiTU9EIiwiaXBhZGRyIjoiMTMxLjEwNy4xNTkuNzUiLCJuYW1lIjoiTU9EIEFkbWluaXN0cmF0b3IiLCJvaWQiOiIyN2Y5YzQ4Zi00YjM0LTQxYmMtYWRmYy1kMDY4OTZjNmY1ZWMiLCJwbGF0ZiI6IjMiLCJwdWlkIjoiMTAwM0JGRkRBNzk2MUY3RCIsInNjcCI6Ikdyb3VwLlJlYWQuQWxsIEdyb3VwLlJlYWRXcml0ZS5BbGwgb2ZmbGluZV9hY2Nlc3MgVXNlci5SZWFkIFVzZXIuUmVhZC5BbGwgVXNlci5SZWFkQmFzaWMuQWxsIFVzZXIuUmVhZFdyaXRlIiwic2lnbmluX3N0YXRlIjpbImttc2kiXSwic3ViIjoid1FSUHNWVjVJRUljVjA5MDRoWjB4eVMyYzlZVEl6X2kxcHM1eEcwVWVEVSIsInRpZCI6ImE1YmJiOWRmLTA2Y2MtNDdmNC04ZjI0LTk4MWEyMDI0Yjk0YyIsInVuaXF1ZV9uYW1lIjoiYWRtaW5ATTM2NXgxNDYxODgub25taWNyb3NvZnQuY29tIiwidXBuIjoiYWRtaW5ATTM2NXgxNDYxODgub25taWNyb3NvZnQuY29tIiwidXRpIjoiX1g5WWczUFAxRVdYV2FmY3hKQUhBQSIsInZlciI6IjEuMCIsIndpZHMiOlsiNjJlOTAzOTQtNjlmNS00MjM3LTkxOTAtMDEyMTc3MTQ1ZTEwIl19.bY7LLs6li9cuV2gpOsNKWb1b2A1GdGP643fL4Kp3KY70tsgFLeXYRwhklMrl1KctinAsKuPV47Q0KHLlF-77txsx1lI7MgnDJQp6nTo9JWAfYb90GLOY_wPvIPtX6YF_edYYygmSrN1YQKaUfTz0vy953Ur8dXtJQkS_tz1uAJ0uLZ-yn3jTLgobTQqd9tZvpGDN6akZ4JSgzniPEXEJiwt8T2ySBXgCpb0Uph5y6GpUdNnzQes89muQVAxGw9yw8UKOa5wlfdLgHvXxe_A8IHqLdmeWVldeAF7gmXG7vX7uGQ2zI9wHzdF0kKENSDqJJHUfL-Z0DotMBZXTn-ftaA";
    }
}
