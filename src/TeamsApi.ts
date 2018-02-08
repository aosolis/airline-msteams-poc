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

const graphBaseUrl = "https://graph.microsoft.com/testTeamsDevEnv";

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

    constructor(
        private accessToken: string,
    )
    {
    }

    // Create a new team
    public async createTeamAsync(displayName: string, description: string, teamSettings: Team): Promise<Team>
    {
        let newGroup = await this.createGroupAsync(displayName, description);
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
                        setTimeout(() => { resolve(); }, 1000);
                    });
                }
            }
        }

        return newTeam;
    }

    // Delete a team
    public async deleteTeamAsync(groupId: string): Promise<void> {
        let options = {
            url: `${graphBaseUrl}/groups/${groupId}`,
            json: true,
            headers: {
                "Authorization": `Bearer ${this.accessToken}`,
            },
        };
        await request.delete(options);
    }

    // Add a member to a team
    public async addMemberToTeamAsync(groupId: string, userObjectId: string): Promise<void> {
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

    // Remove a member from a team
    public async removeMemberFromTeamAsync(groupId: string, userObjectId: string): Promise<void> {
        let options = {
            url: `${graphBaseUrl}/groups/${groupId}/members/${userObjectId}/$ref`,
            json: true,
            headers: {
                "Authorization": `Bearer ${this.accessToken}`,
            },
        };
        await request.delete(options);
    }

    // Get the members of a team
    public async getMembersAsync(groupId: string): Promise<DirectoryObject[]> {
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

    // Update group information
    public async updateGroupAsync(groupId: string, groupUpdates: Group): Promise<void> {
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
    private async createGroupAsync(displayName: string, description: string): Promise<Group>
    {
        let requestBody = {
            displayName: displayName,
            description: description,
            mailEnabled: true,
            mailNickname: "teamAlias100",
            securityEnabled: false,
            visibility: "Private",
            groupTypes: [ "Unified" ],
        };
        let options = {
            url: `${graphBaseUrl}/groups`,
            body: requestBody,
            json: true,
            headers: {
                "Authorization": `Bearer ${this.accessToken}`,
            },
        };
        let responseBody = await request.post(options);
        return responseBody as Group;
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
        let responseBody = await request.put(options);
        return responseBody as Team;
    }
}
