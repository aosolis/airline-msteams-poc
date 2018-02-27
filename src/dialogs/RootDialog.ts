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
import * as winston from "winston";
import * as config from "config";
import * as builder from "botbuilder";
import * as moment from "moment";
import * as constants from "../constants";
import * as utils from "../utils";
import * as teams from "../TeamsApi";
import * as trips from "../trips/TripsApi";
import * as sampledata from "../data/SampleData";
import { MongoDbTripsApi } from "../trips/MongoDbTripsApi";
import { GroupData, IAppDataStore } from "../storage/AppDataStore";
import { MongoDbAppDataStore } from "../storage/MongoDbAppDataStore";
import { AzureADv1Dialog } from "./AzureADv1Dialog";
let uuidv4 = require("uuid/v4");

const createTripsRegExp = /^createTrips(.*)$/i;
const showTripRegExp = /^showTrip (.*)$/i;
const addRemoveCrewRegExp = /^(add|remove)Crew (.*) (.*)$/i;
const triggerTimeRegExp = /^triggerTime(.*)$/i;

const daysInAdvanceToCreateTrips = 7;       // Create teams for trips departing X days in the future
const daysInPastToMonitorTrips = 7;         // Actively monitor future trips and trips that departed in the past Y days
const daysInPastToArchiveTrips = 14;        // Archive teams for trips that departed more that Z days ago
const archivedTag = "[ARCHIVED]";           // Tag prepended to team name when it is archived

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

// Root dialog provides choices in identity providers
export class RootDialog extends builder.IntentDialog
{
    private teamsApi: teams.TeamsApi = new teams.AppContextTeamsApi(config.get("app.tenantId"), config.get("bot.appId"), config.get("bot.appPassword"));
    private tripsApi: trips.ITripsApi = new MongoDbTripsApi(config.get("mongoDb.connectionString"));
    private appDataStore: IAppDataStore = new MongoDbAppDataStore(config.get("mongoDb.connectionString"));

    constructor() {
        super();
    }

    // Register the dialog with the bot
    public register(bot: builder.UniversalBot): void {
        // Register dialogs
        bot.dialog(constants.DialogId.Root, this);
        new AzureADv1Dialog().register(bot, this);

        // Commands to get user token for delegated consent
        this.matches(/login/i, constants.DialogId.AzureADv1, "login");
        this.matches(/logout/i, constants.DialogId.AzureADv1, "logout");

        // Commands to manipulate trip database and app state
        this.matches(createTripsRegExp, (session) => { this.handleCreateTrips(session); });
        this.matches(showTripRegExp, (session) => { this.handleShowTrip(session); });
        this.matches(addRemoveCrewRegExp, (session) => { this.handleAddRemoveCrew(session); });
        this.matches(/resetState/i, (session) => { this.handleResetState(session); });

        // Commands to simulate an update trigger at a given time
        this.matches(triggerTimeRegExp, (session) => { this.handleUpdateTrigger(session); });

        this.onDefault((session) => { session.send("I didn't understand that."); });
    }

    // Handle resumption of dialog
    public dialogResumed<T>(session: builder.Session, result: builder.IDialogResult<T>): void {
        // The only dialog we branch to is auth
        let userToken = utils.getUserToken(session, constants.IdentityProviders.azureADv1);
        if (userToken) {
            this.teamsApi = new teams.UserContextTeamsApi(userToken.accessToken, userToken.expirationTime);
        }
    }

    // Populate the trip database with fake sample trips
    private async handleCreateTrips(session: builder.Session): Promise<void> {
        let dateString = createTripsRegExp.exec(session.message.text)[1];
        let inputDate = moment.utc(dateString);

        // By default create trips that depart Dubai a week from now
        let baseDate = inputDate.isValid() ? inputDate.toDate() : new Date(new Date().valueOf() + (7 * 24 * 60 * 60 * 1000));

        let fakeTrips: trips.Trip[] = sampledata.tripTemplates.map((trip) => {
            let departureTime = trip.departureTime;
            return {
                ...trip,
                tripId: <string>uuidv4(),
                departureTime: new Date(Date.UTC(baseDate.getUTCFullYear(), baseDate.getUTCMonth(), baseDate.getUTCDate(), departureTime.getUTCHours(), departureTime.getUTCMinutes())),
            };
        });
        let addPromises = fakeTrips.map((trip) => {
            return (<trips.ITripsTest><any>this.tripsApi).addOrUpdateTripAsync(trip);
        });

        try {
            await Promise.all(addPromises);

            let tripInfo = fakeTrips.map(trip => `${this.createDisplayNameForTrip(trip)} (${trip.tripId})`).join(", ");
            session.send(`Created ${fakeTrips.length} trips: ${tripInfo}`);
        } catch (e) {
            winston.error(`Error creating trips: ${e.message}`, e);
            session.send(`An error occurred while creating trips: ${e.message}`);
        }
    }

    // Show the details of a trip in the database
    private async handleShowTrip(session: builder.Session): Promise<void> {
        let tripId = showTripRegExp.exec(session.message.text)[1].trim();
        let trip = await this.tripsApi.getTripAsync(tripId);
        if (!trip) {
            session.send(`Couldn't find a trip with id ${tripId}.`);
            return;
        }

        let message = `${this.createDisplayNameForTrip(trip)}<br/>\nRoster:<ol>\n` +
            trip.crewMembers.map(m => `<li>${m.displayName} (${m.userPrincipalName}) ${m.rosterGrade}</li>`).join("\n") +
            `\n</ol>`;
        session.send(message);
    }

    // Delete all trips in the trip database, and all teams that were created
    private async handleResetState(session: builder.Session): Promise<void> {
        // Delete all trips
        try {
            await (<trips.ITripsTest><any>this.tripsApi).deleteAllTripsAsync();
        } catch (e) {
            winston.error(`Error deleting trips: ${e.message}`, e);
            session.send(`An error occurred while deleting trips: ${e.message}`);
        }

        // Delete all teams
        try {
            let teams = await this.appDataStore.getAllGroupsAsync();
            let deleteTeamPromises = teams.map(async (groupData) => {
                let groupId = groupData.groupId;
                winston.info(`Deleting group ${groupId}`);

                await this.teamsApi.deleteGroupAsync(groupId);
                await this.appDataStore.deleteGroupDataAsync(groupId);
            });
            await Promise.all(deleteTeamPromises);
        } catch (e) {
            winston.error(`Error deleting teams: ${e.message}`, e);
            session.send(`An error occurred while deleting teams: ${e.message}`);
        }

        session.send("Deleted all trips and teams");
    }

    // Add or remove a crew member from a trip
    private async handleAddRemoveCrew(session: builder.Session): Promise<void> {
        let match = addRemoveCrewRegExp.exec(session.message.text);
        let command = match[1];
        let tripId = match[2];
        let crewMemberEmail = match[3];

        let trip = await this.tripsApi.getTripAsync(tripId);
        if (!trip) {
            session.send(`Couldn't find a trip with id ${tripId}.`);
            return;
        }

        let crewMember = sampledata.findCrewMemberByUpn(crewMemberEmail);
        if (!crewMember) {
            session.send(`Couldn't find a crew member with email address ${crewMemberEmail}`);
            return;
        }

        let tripName = this.createDisplayNameForTrip(trip);
        let testTripsApi = <trips.ITripsTest><any>this.tripsApi;

        switch (command) {
            case "add":
                trip.crewMembers = _(trip.crewMembers).push(crewMember).uniqBy("aadObjectId").value();
                await testTripsApi.addOrUpdateTripAsync(trip);
                session.send(`Added ${crewMember.displayName} to the trip roster for ${tripName}.`);
                break;

            case "remove":
                trip.crewMembers = trip.crewMembers.filter(member => member.aadObjectId !== crewMember.aadObjectId);
                await testTripsApi.addOrUpdateTripAsync(trip);
                session.send(`Removed ${crewMember.displayName} from the trip roster for ${tripName}.`);
                break;
        }
    }

    // Handle a trigger to update the teams for upcoming trips
    private async handleUpdateTrigger(session: builder.Session): Promise<void> {
        let dateString = triggerTimeRegExp.exec(session.message.text)[1];
        let inputDate = moment.utc(dateString);

        // By default simulate a time trigger for the current time, allowing override
        let triggerTime = inputDate.isValid() ? inputDate.toDate() : moment.utc().toDate();
        winston.info(`Simulating a time trigger for ${triggerTime.toUTCString()}`);

        try {
            // Archive old teams
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

            // Update existing teams
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

            // Create new teams
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

                    winston.info(`Team ${groupId} created for trip ${trip.tripId} departing DXB on ${trip.departureTime.toUTCString()}`);
                }
            });
            await Promise.all(teamCreatePromises);

            session.send(`Finished processing the time trigger for ${triggerTime.toUTCString()}`);
        } catch (e) {
            let errorMessage = `An error occurred while processing time trigger at ${triggerTime.toUTCString()}: ${e.message}`;
            winston.error(errorMessage, e);
            session.send(errorMessage);
        }
    }

    // Create a legible display name for a trip
    private createDisplayNameForTrip(trip: trips.Trip): string {
        let flightNumbers = "EK" + _(trip.flights).map(flight => flight.flightNumber).uniq().join("/");
        let route = _(trip.flights).map(flight => flight.destination).unshift(trip.flights[0].origin).join("-");
        let dubaiDepartureDate = moment(trip.departureTime).utcOffset(240).format("YYYY-MM-DD");
        return `${flightNumbers} ${route} ${dubaiDepartureDate}`;
    }

    // Create a team for a trip
    private async createTeamForTripAsync(trip: trips.Trip): Promise<string> {
        let team: teams.Team;

        // Create the team
        try {
            let displayName = this.createDisplayNameForTrip(trip);
            team = await this.teamsApi.createTeamAsync(displayName, null, trip.tripId, teamSettings);
        } catch (e) {
            winston.error(`Error creating team for trip ${trip.tripId}: ${e.message}`, e);
            throw e;
        }
        winston.info(`Created a new team ${team.id} for trip ${trip.tripId}`);

        // Add team members
        let memberAddPromises = trip.crewMembers.map(async crewMember => {
            try {
                await this.teamsApi.addMemberToGroupAsync(team.id, crewMember.aadObjectId);
            } catch (e) {
                winston.error(`Error adding ${crewMember.staffId} (${crewMember.aadObjectId}): ${e.message}`, e);
            }
        });
        await Promise.all(memberAddPromises);
        winston.info(`Added ${memberAddPromises.length} members to team ${team.id}`);

        return team.id;
    }

    // "Archive" a team
    // Microsoft Teams doesn't support true archival yet, so instead we will:
    //  - remove all the users from the team
    //  - "park" it with an admin user (admins have no cap on the number of groups they can be part of)
    //  - rename the team to mark it as archived 
    private async archiveTeamAsync(groupId: string): Promise<void> {
        // Get the id of the user that owns all "archived" teams
        let archivedTeamOwnerId = config.get("app.archivedTeamOwnerId").toLowerCase();

        // Get current members and owners
        let teamMembers = await this.teamsApi.getMembersOfGroupAsync(groupId);
        let teamOwners = await this.teamsApi.getOwnersOfGroupAsync(groupId);
        winston.info(`Found ${teamMembers.length} members and ${teamOwners.length} owners in the team ${groupId}`);

        // Add the archive owner to group, as both member and owner.
        // Being a member is optional, but it makes it easier to query for all archived teams using a /me/memberOf query.
        if (!teamMembers.find(member => member.id.toLowerCase() === archivedTeamOwnerId)) {
            await this.teamsApi.addMemberToGroupAsync(groupId, archivedTeamOwnerId);
        }
        if (!teamOwners.find(owner => owner.id.toLowerCase() === archivedTeamOwnerId)) {
            await this.teamsApi.addOwnerToGroupAsync(groupId, archivedTeamOwnerId);
        }
        winston.info(`Added ${archivedTeamOwnerId} to team ${groupId} as owner and member`);

        // Remove all existing members and owners apart from the archive owner
        let memberRemovePromises = teamMembers
            .filter(member => member.id.toLowerCase() !== archivedTeamOwnerId)
            .map(async member => {
                try {
                    await this.teamsApi.removeMemberFromGroupAsync(groupId, member.id);
                } catch (e) {
                    winston.error(`Error removing member ${member.id}: ${e.message}`);
                }
            });
        await Promise.all(memberRemovePromises);
        winston.info(`Removed ${memberRemovePromises.length} members from team ${groupId}`);

        let ownerRemovePromises = teamOwners
            .filter(owner => owner.id.toLowerCase() !== archivedTeamOwnerId)
            .map(async owner => {
                try {
                    await this.teamsApi.removeOwnerFromGroupAsync(groupId, owner.id);
                } catch (e) {
                    winston.error(`Error removing owner ${owner.id}: ${e.message}`);
                }
            });
        await Promise.all(ownerRemovePromises);
        winston.info(`Removed ${ownerRemovePromises.length} owners from team ${groupId}`);

        // Rename group to indicate that it has been archived
        let group = await this.teamsApi.getGroupAsync(groupId);
        if (!group.displayName.startsWith(archivedTag)) {
            await this.teamsApi.updateGroupAsync(groupId, {
                displayName: `${archivedTag} ${group.displayName}`,
            });
        }
        winston.info(`Finished archiving team ${groupId}`);
    }
}
