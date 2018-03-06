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

const teamCreationDelayInSeconds = 10;          // Seconds to wait after creating a team, before adding members to it

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
    private archivedTeamOwnerId: string;

    constructor(
        private tripsApi: trips.ITripsApi,              // Interface to the trips database
        private teamsApi: teams.TeamsApi,               // Interface to the teams Graph API
        private appDataStore: IAppDataStore,            // Interface to the app data store
        private appTeamsApi: teams.TeamsApi,            // Interface to the teams Graph API, using app context
    ) {
        // Get the id of the user that owns all "archived" teams
        this.archivedTeamOwnerId = config.get("app.archivedTeamOwnerId").toLowerCase();
    }

    // Handle a trigger to update teams
    public async updateTeamsAsync(triggerTime: Date): Promise<void> {
        winston.info(`Updating teams based on trigger time ${triggerTime.toUTCString()}`);
        await this.archiveTeamsAsync(triggerTime);
        await this.updateExistingTeamsAsync(triggerTime);
        await this.createTeamsAsync(triggerTime);
    }

    // Delete all the teams that this updater created
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
            let groupData = await this.appDataStore.getGroupDataByTripAsync(trip.tripId);
            if (!groupData) {
                let groupId = await this.createTeamForTripAsync(trip);

                let newGroupData: GroupData = {
                    groupId:  groupId,
                    tripId: trip.tripId,
                    tripSnapshot: trip,
                    creationTime: triggerTime,
                };
                await this.appDataStore.addOrUpdateGroupDataAsync(newGroupData);

                winston.info(`Team ${groupId} created for trip ${trip.tripId} departing on ${trip.departureTime.toUTCString()}`);
            }
        });
        await Promise.all(teamCreatePromises);
    }

    // Create a team for a trip
    private async createTeamForTripAsync(trip: trips.Trip): Promise<string> {
        let team: teams.Team;

        // Create the team
        try {
            let displayName = this.getDisplayNameForTrip(trip);
            let description = this.getDescriptionForTrip(trip);
            team = await this.teamsApi.createTeamAsync(displayName, description, trip.tripId, teamSettings);
        } catch (e) {
            winston.error(`Error creating team for trip ${trip.tripId}: ${e.message}`, e);
            throw e;
        }
        winston.info(`Created a new group ${team.id} for trip ${trip.tripId}`);

        winston.info(`Waiting ${teamCreationDelayInSeconds} seconds`);
        await new Promise((resolve, reject) => {
            setTimeout(() => { resolve(); }, teamCreationDelayInSeconds * 1000);
        });

        // Add team members
        let memberAddPromises = trip.crewMembers.map(async crewMember => {
            try {
                await this.teamsApi.addMemberToGroupAsync(team.id, crewMember.aadObjectId);
            } catch (e) {
                winston.error(`Error adding ${crewMember.staffId} (${crewMember.aadObjectId}): ${e.message}`, e);
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
        let groupsToUpdate = (await this.appDataStore.findActiveGroupsCreatedBeforeTimeAsync(triggerTime))
            .filter(groupData => groupData.tripSnapshot.departureTime > minDepartureTimeToUpdate);

        let groupIdsToUpdate = groupsToUpdate.map(groupData => groupData.groupId).join(", ");
        winston.info(`Found ${groupsToUpdate.length} groups to update: ${groupIdsToUpdate}`);

        let teamUpdatePromises = groupsToUpdate.map(async (groupData) => {
            try
            {
                let groupId = groupData.groupId;
                let trip = await this.tripsApi.getTripAsync(groupData.tripId);

                let crewMembers = trip.crewMembers;
                let groupMembers = await this.teamsApi.getMembersOfGroupAsync(groupId);

                // Add new crew members to group
                let crewMembersToAdd = crewMembers.filter(crewMember =>
                    !groupMembers.find(groupMember => groupMember.id === crewMember.aadObjectId));
                let memberAddPromises = crewMembersToAdd.map((crewMember) => this.teamsApi.addMemberToGroupAsync(groupId, crewMember.aadObjectId));
                await Promise.all(memberAddPromises);
                winston.info(`Added ${crewMembersToAdd.length} new members to group ${groupId}`);

                // Remove deleted group members
                let groupMembersToRemove = groupMembers.filter(groupMember =>
                    !crewMembers.find(crewMember => groupMember.id === crewMember.aadObjectId));
                let memberRemovePromises = groupMembersToRemove.map((groupMember) => this.teamsApi.removeMemberFromGroupAsync(groupId, groupMember.id));
                await Promise.all(memberRemovePromises);
                winston.info(`Removed ${groupMembersToRemove.length} members from group ${groupId}`);

                groupData.tripSnapshot = trip;
                this.appDataStore.addOrUpdateGroupDataAsync(groupData);
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
        let groupsToArchive = (await this.appDataStore.findActiveGroupsCreatedBeforeTimeAsync(maxDepartureTimeToArchive))
            .filter(groupData => groupData.tripSnapshot.departureTime < maxDepartureTimeToArchive);

        let groupIdsToArchive = groupsToArchive.map(groupData => groupData.groupId).join(", ");
        winston.info(`Found ${groupsToArchive.length} groups to archive: ${groupIdsToArchive}`);

        let teamArchivePromises = groupsToArchive.map(async (groupData) => {
            try
            {
                await this.archiveTeamAsync(groupData.groupId);

                groupData.archivalTime = triggerTime;
                await this.appDataStore.addOrUpdateGroupDataAsync(groupData);
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
        // Get current members and owners
        let teamMembers = await this.teamsApi.getMembersOfGroupAsync(groupId);
        let teamOwners = await this.teamsApi.getOwnersOfGroupAsync(groupId);
        winston.info(`Found ${teamMembers.length} members and ${teamOwners.length} owners in the team ${groupId}`);

        // Add the archive owner to group, as both member and owner.
        // Being a member is optional, but it makes it easier to query for all archived teams using a /me/memberOf query.
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
        // once we have relinquished ownership over it.
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
}
