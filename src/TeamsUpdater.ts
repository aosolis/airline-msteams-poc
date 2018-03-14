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

import * as _ from "lodash";
import * as moment from "moment";
import * as winston from "winston";
import * as config from "config";
import * as teams from "./TeamsApi";
import * as trips from "./trips/TripsApi";
import { GroupData, IAppDataStore } from "./storage/AppDataStore";

const daysInAdvanceToCreateTrips = 7;           // Create teams for trips departing X days in the future
const daysInPastToMonitorTrips = 7;             // Actively monitor future trips and trips that departed in the past Y days
const daysInPastToArchiveTrips = 14;            // Archive teams for trips that departed more than Z days ago
const archivedTag = "[ARCHIVED]";               // Tag prepended to team name when it is archived

const teamCreationDelayInSeconds = 7;          // Seconds to wait after creating a team, before adding members to it

// Default settings for new teams
const teamSettings: teams.Team = {
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

// Updates teams to be in sync with trips
export class TeamsUpdater
{
    private activeTeamOwnerId: string;
    private activeTeamOwnerUpn: string;
    private archivedTeamOwnerId: string;
    private archivedTeamOwnerUpn: string;

    constructor(
        private tripsApi: trips.ITripsApi,              // Interface to the trips database
        private teamsApi: teams.TeamsApi,               // Interface to the Teams Graph API
        private appDataStore: IAppDataStore,            // Interface to the app data store
        private appTeamsApi: teams.TeamsApi,            // Interface to the Teams Graph API, using app context
    ) {
        // Get the id of the user that owns all active teams
        this.activeTeamOwnerId = config.get("app.activeTeamOwnerId").toLowerCase();
        this.activeTeamOwnerUpn = config.get("app.activeTeamOwnerUpn");

        // Get the id of the user that owns all "archived" teams
        this.archivedTeamOwnerId = config.get("app.archivedTeamOwnerId").toLowerCase();
        this.archivedTeamOwnerUpn = config.get("app.archivedTeamOwnerUpn");

        // We allow for 2 different users here, in case there are scenarios where the active team owner ever has to
        // log in to Teams, which could be problematic if the user is part of thousands of teams. If both accounts are
        // treated as service accounts and never actually use Teams, then the active and archive owner accounts
        // can be the same user.
    }

    // Handle a trigger to update teams
    public async updateTeamsAsync(triggerTime: Date): Promise<void> {
        winston.info(`Updating teams based on trigger time ${triggerTime.toUTCString()}`);

        await this.resolveTeamOwnersAsync();

        await this.archiveTeamsAsync(triggerTime);
        await this.updateExistingTeamsAsync(triggerTime);
        await this.createTeamsAsync(triggerTime);
    }

    // Delete all the teams that we created (for testing)
    public async deleteAllTrackedTeamsAsync(): Promise<void> {
        let teams = await this.appDataStore.getAllGroupsAsync();
        winston.info(`Found ${teams.length} teams to delete`);

        let deleteTeamPromises = teams.map(async (groupData) => {
            let groupId = groupData.groupId;
            winston.info(`Deleting group ${groupId}`);

            try {
                await this.appTeamsApi.deleteGroupAsync(groupId);
            } catch (e) {
                if (e.statusCode === 404) {
                    // Not found, ok
                } else {
                    throw e;
                }
            }

            await this.appDataStore.deleteGroupDataAsync(groupId);
        });
        await Promise.all(deleteTeamPromises);
    }

    // Create new teams
    private async createTeamsAsync(triggerTime: Date): Promise<void> {
        // Create teams for trips departing up to "daysInAdvanceToCreateTrips" days in the future
        let maxDepartureTimeToCreate = moment(triggerTime).add(daysInAdvanceToCreateTrips, "d").toDate();
        let trips = await this.tripsApi.findTripsDepartingInRangeAsync(triggerTime, maxDepartureTimeToCreate);

        let departureTimes = trips.map(trip => trip.departureTime.toUTCString()).join(", ");
        winston.info(`Found ${trips.length} trips that depart at the following times ${departureTimes}`);

        let teamCreatePromises = trips.map(async (trip) => {
            try {
                // Create a team for the trip if we don't have one yet
                let groupData = await this.appDataStore.getGroupDataByTripAsync(trip.tripId);
                if (!groupData) {
                    let groupId = await this.createTeamForTripAsync(trip);

                    // Keep track of the teams that we have created
                    let newGroupData: GroupData = {
                        groupId:  groupId,
                        tripId: trip.tripId,
                        tripSnapshot: trip,
                        creationTime: triggerTime,
                    };
                    await this.appDataStore.addOrUpdateGroupDataAsync(newGroupData);

                    winston.info(`Team ${groupId} created for trip ${trip.tripId} departing on ${trip.departureTime.toUTCString()}`);
                }
            } catch (e) {
                winston.error(`Error creating team for trip ${trip.tripId}: ${e.message}`, e);
            }
        });
        await Promise.all(teamCreatePromises);
    }

    // Create a team for a trip
    private async createTeamForTripAsync(trip: trips.Trip): Promise<string> {
        let group: teams.Group;
        let team: teams.Team;

        // Create the team
        try {
            let displayName = this.getDisplayNameForTrip(trip);
            let description = this.getDescriptionForTrip(trip);

            // First create a modern group
            group = await this.teamsApi.createGroupAsync(displayName, description, trip.tripId);
            winston.info(`Created new group ${group.id}`);

            // If we're acting in app context, a user owner needs to be added to the group
            if (this.teamsApi.isInAppContext()) {
                await Promise.all([
                    this.teamsApi.addOwnerToGroupAsync(group.id, this.activeTeamOwnerId),
                    this.teamsApi.addMemberToGroupAsync(group.id, this.activeTeamOwnerId),
                ]);
            }

            // Convert the group into a team
            team = await this.teamsApi.createTeamFromGroupAsync(group.id, teamSettings);
        } catch (e) {
            winston.error(`Error creating team for trip ${trip.tripId}: ${e.message}`, e);

            // If we failed to convert the group to a team, clean up after ourselves by deleting the group
            if (group) {
                winston.error(`Error converting group ${group.id} to a team, will delete the group`);
                try {
                    await this.teamsApi.deleteGroupAsync(group.id);
                } catch (deleteError) {
                    winston.error(`Failed to delete the group ${group.id}: ${deleteError.message}`, deleteError);
                }
            }

            throw e;
        }
        winston.info(`Created a new group ${team.id} for trip ${trip.tripId}`);

        // Wait a few seconds for the team information to propagate
        winston.info(`Waiting ${teamCreationDelayInSeconds} seconds`);
        await new Promise((resolve, reject) => {
            setTimeout(() => { resolve(); },
            teamCreationDelayInSeconds * 1000);
        });

        // Add team members
        this.normalizeUpnCasing(trip.crewMembers);
        let memberAddPromises = trip.crewMembers.map(async crewMember => {
            try {
                let user = await this.teamsApi.getUserByUpnAsync(crewMember.userPrincipalName);
                await this.teamsApi.addMemberToGroupAsync(team.id, user.id);
            } catch (e) {
                winston.error(`Error adding ${crewMember.userPrincipalName}: ${e.message}`, e);
            }
        });
        await Promise.all(memberAddPromises);
        winston.info(`Added ${memberAddPromises.length} members to group ${team.id}`);

        return team.id;
    }

    // Get a legible display name for a trip
    private getDisplayNameForTrip(trip: trips.Trip): string {
        let flightNumbers = "EK" + _(trip.flights).map(flight => flight.flightNumber).uniq().join("/");
        let departureDate = moment(trip.departureTime).utcOffset(240).format("YYYY-MM-DD");
        return `${flightNumbers} ${departureDate}`;
    }

    // Get a legible description name for a trip
    private getDescriptionForTrip(trip: trips.Trip): string {
        let flightNumbers = "EK" + _(trip.flights).map(flight => flight.flightNumber).uniq().join("/");
        let route = _(trip.flights).map(flight => flight.destination).unshift(trip.flights[0].origin).join("-");
        let departureDate = moment(trip.departureTime).utcOffset(240).format("YYYY-MM-DD");
        return `${flightNumbers} (${route}) on ${departureDate}`;
    }

    // Update existing teams
    private async updateExistingTeamsAsync(triggerTime: Date): Promise<void> {
        // Monitor roster changes for trips departing up to "daysInPastToMonitorTrips" days ago
        let minDepartureTimeToUpdate = moment(triggerTime).subtract(daysInPastToMonitorTrips, "d").toDate();
        let activeGroups = await this.appDataStore.findActiveGroupsCreatedBeforeTimeAsync(triggerTime);
        let groupsToUpdate = activeGroups
            .filter(groupData => groupData.tripSnapshot.departureTime > minDepartureTimeToUpdate);

        let groupIdsToUpdate = groupsToUpdate.map(groupData => groupData.groupId).join(", ");
        winston.info(`Found ${groupsToUpdate.length} groups to update: ${groupIdsToUpdate}`);

        let teamUpdatePromises = groupsToUpdate.map(async (groupData) => {
            try
            {
                // Get trip info
                let groupId = groupData.groupId;
                let trip = await this.tripsApi.getTripAsync(groupData.tripId);
                this.normalizeUpnCasing(trip.crewMembers);

                // Get current members of the team
                let crewMembers = trip.crewMembers;
                let groupMembers = await this.teamsApi.getMembersOfGroupAsync(groupId);
                this.normalizeUpnCasing(groupMembers);

                // Add new crew members to group
                let crewMembersToAdd = crewMembers.filter(crewMember =>
                    !groupMembers.find(groupMember => groupMember.userPrincipalName === crewMember.userPrincipalName));
                if (crewMembersToAdd.length > 0) {
                    let memberAddPromises = crewMembersToAdd.map(async (crewMember) => {
                        let user = await this.teamsApi.getUserByUpnAsync(crewMember.userPrincipalName);
                        await this.teamsApi.addMemberToGroupAsync(groupId, user.id);
                    });
                    await Promise.all(memberAddPromises);
                    winston.info(`Added ${crewMembersToAdd.length} new members to group ${groupId}`);
                }

                // Remove deleted group members
                let groupMembersToRemove = groupMembers.filter(groupMember =>
                    !crewMembers.find(crewMember => groupMember.userPrincipalName === crewMember.userPrincipalName));
                if (groupMembersToRemove.length > 0) {
                    let memberRemovePromises = groupMembersToRemove.map(groupMember => {
                        return this.teamsApi.removeMemberFromGroupAsync(groupId, groupMember.id);
                    });
                    await Promise.all(memberRemovePromises);
                    winston.info(`Removed ${groupMembersToRemove.length} members from group ${groupId}`);
                }

                // Update team info in app database
                groupData.tripSnapshot = trip;
                await this.appDataStore.addOrUpdateGroupDataAsync(groupData);

                winston.info(`Synced membership of group ${groupId} with current roster of trip ${trip.tripId}`);
            }
            catch (e) {
                winston.error(`Error updating group ${groupData.groupId}: ${e.message}`, e);
            }
        });
        await Promise.all(teamUpdatePromises);
    }

    // Archive old teams
    private async archiveTeamsAsync(triggerTime: Date): Promise<void> {
        // Archive active teams created for trips that have departed more than "daysInPastToArchiveTrips" ago
        let maxDepartureTimeToArchive = moment(triggerTime).subtract(daysInPastToArchiveTrips, "d").toDate();
        let activeGroups = await this.appDataStore.findActiveGroupsCreatedBeforeTimeAsync(maxDepartureTimeToArchive);
        let groupsToArchive = activeGroups
            .filter(groupData => groupData.tripSnapshot.departureTime < maxDepartureTimeToArchive);

        let groupIdsToArchive = groupsToArchive.map(groupData => groupData.groupId).join(", ");
        winston.info(`Found ${groupsToArchive.length} groups to archive: ${groupIdsToArchive}`);

        let teamArchivePromises = groupsToArchive.map(async (groupData) => {
            try
            {
                // Archive team
                let groupId = groupData.groupId;
                await this.archiveTeamAsync(groupId);

                // Update team info in app database
                groupData.archivalTime = triggerTime;
                await this.appDataStore.addOrUpdateGroupDataAsync(groupData);

                winston.info(`Archived group ${groupId}`);
            }
            catch (e) {
                winston.error(`Error archiving group ${groupData.groupId}: ${e.message}`, e);
            }
        });
        await Promise.all(teamArchivePromises);
    }

    // "Archive" a team
    // Microsoft Teams doesn't support true archival yet, so instead we will:
    //  - remove all the users from the team
    //  - "park" it with an admin user (admins have no cap on the number of groups they can be part of)
    //  - rename the team to mark it as archived
    private async archiveTeamAsync(groupId: string): Promise<void> {
        // Get current members and owners (restricted to users only)
        let teamMembers = await this.teamsApi.getMembersOfGroupAsync(groupId);
        let teamOwners = await this.teamsApi.getOwnersOfGroupAsync(groupId);
        winston.info(`Found ${teamMembers.length} members and ${teamOwners.length} owners in the team ${groupId}`);

        // Add the archive owner to group, as both member and owner
        // Being a member is optional, but it makes it easier to query for all archived teams using a /me/memberOf query
        if (!teamMembers.find(member => member.id.toLowerCase() === this.archivedTeamOwnerId)) {
            await this.teamsApi.addMemberToGroupAsync(groupId, this.archivedTeamOwnerId);
        }
        if (!teamOwners.find(owner => owner.id.toLowerCase() === this.archivedTeamOwnerId)) {
            await this.teamsApi.addOwnerToGroupAsync(groupId, this.archivedTeamOwnerId);
        }
        winston.info(`Added ${this.archivedTeamOwnerId} to team ${groupId} as owner and member`);

        // Remove all existing members
        let memberRemovePromises = teamMembers
            .filter(member => member.id.toLowerCase() !== this.archivedTeamOwnerId)
            .map(async member => {
                try {
                    await this.teamsApi.removeMemberFromGroupAsync(groupId, member.id);
                } catch (e) {
                    winston.error(`Error removing member ${member.id}: ${e.message}`);
                }
            });
        await Promise.all(memberRemovePromises);
        winston.info(`Removed ${memberRemovePromises.length} members from team ${groupId}`);

        // Rename group to indicate that it has been archived
        let group = await this.teamsApi.getGroupAsync(groupId);
        if (!group.displayName.startsWith(archivedTag)) {
            await this.teamsApi.updateGroupAsync(groupId, {
                displayName: `${archivedTag} ${group.displayName}`,
            });
        }

        // Remove all other owners. This needs to be done last, as we cannot modify the team
        // once we have relinquished ownership over it (this happens when using user context;
        // in app context the app remains a group owner).
        let ownerRemovePromises = teamOwners
            .filter(owner => owner.id.toLowerCase() !== this.archivedTeamOwnerId)
            .map(async owner => {
                try {
                    await this.teamsApi.removeOwnerFromGroupAsync(groupId, owner.id);
                } catch (e) {
                    winston.error(`Error removing owner ${owner.id}: ${e.message}`);
                }
            });
        await Promise.all(ownerRemovePromises);
        winston.info(`Removed ${ownerRemovePromises.length} owners from team ${groupId}`);

        winston.info(`Finished archiving team ${groupId}`);
    }

    // Resolve the team owner accounts to their user ids
    private async resolveTeamOwnersAsync(): Promise<void> {
        if (!this.activeTeamOwnerId) {
            let user = await this.appTeamsApi.getUserByUpnAsync(this.activeTeamOwnerUpn);
            this.activeTeamOwnerId = user.id.toLowerCase();
        }
        if (!this.archivedTeamOwnerId) {
            let user = await this.appTeamsApi.getUserByUpnAsync(this.archivedTeamOwnerUpn);
            this.archivedTeamOwnerId = user.id.toLowerCase();
        }
    }

    // Normalize the casing of the UPN list, so we can compare directly
    // Note that this mutates the list
    private normalizeUpnCasing(users: (trips.CrewMember|teams.User)[]): void {
        users.forEach(user => {
            if (user.userPrincipalName) {
                user.userPrincipalName = user.userPrincipalName.toLowerCase();
            }
        });
    }
}
