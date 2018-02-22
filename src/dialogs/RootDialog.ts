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
import * as config from "config";
import * as builder from "botbuilder";
import * as moment from "moment";
import * as constants from "../constants";
import * as utils from "../utils";
import * as teams from "../TeamsApi";
import * as trips from "../trips/TripsApi";
import { MongoDbTripsApi } from "../trips/MongoDbTripsApi";
import { GroupData, IAppDataStore } from "../storage/AppDataStore";
import { MongoDbAppDataStore } from "../storage/MongoDbAppDataStore";
import { AzureADv1Dialog } from "./AzureADv1Dialog";
let uuidv4 = require("uuid/v4");

const createTripsRegExp = /^createTrips(.*)$/i;
const triggerTimeRegExp = /^triggerTime(.*)$/i;

const daysInAdvanceToCreateTrips = 7;       // Create teams for trips departing X days in the future
const daysInPastToMonitorTrips = 7;         // Actively monitor future trips and trips that departed in the past Y days
const daysInPastToArchiveTrips = 14;        // Archive teams for trips that departed more that Z days ago

const tripTemplates: trips.Trip[] = [
    {
        tripId: null,
        departureTime: new Date("2018-02-08 14:25:00 UTC+4"),
        flights: [
            {
                flightNumber: "051",
                origin: "DXB",
                destination: "MUC",
            },
            {
                flightNumber: "052",
                origin: "MUC",
                destination: "DXB",
            },
        ],
        crewMembers: [
            {
                staffId: "292062",
                rosterGrade: "FG1",
                aadObjectId: "303b75b6-87e1-4f8e-8387-fbfe288356bf",
            },
            {
                staffId: "378718",
                rosterGrade: "FG1",
                aadObjectId: "0f429da5-2cbf-4d95-bc2c-16a1bef3ed1c",
            },
            {
                staffId: "431620",
                rosterGrade: "GR1",
                aadObjectId: "e5a7e50b-8005-4594-93aa-51f62912d1cd",
            },
            {
                staffId: "420501",
                rosterGrade: "GR1",
                aadObjectId: "2238636f-e037-47be-8cda-ca765ff96793",
            },
            {
                staffId: "431400",
                rosterGrade: "GR1",
                aadObjectId: "ecd0aca2-74ca-41fb-bbef-fd566b0e3aa2",
            },
            {
                staffId: "450986",
                rosterGrade: "GR2",
                aadObjectId: "4252dcaa-7a49-43e1-95e8-3db616da342d",
            },
            {
                staffId: "430109",
                rosterGrade: "GR1",
                aadObjectId: "0011a194-4d0e-4372-9016-40b11837f429",
            },
            {
                staffId: "381830",
                rosterGrade: "SFS",
                aadObjectId: "bf966547-37f8-43bc-b5b5-48cd7052ef75",
            },
            {
                staffId: "434722",
                rosterGrade: "GR1",
                aadObjectId: "2fdac1c4-69dd-418d-bf3c-9aafc42d950b",
            },
            {
                staffId: "422970",
                rosterGrade: "GR1",
                aadObjectId: "a2d06783-f918-4f4b-af32-5dacb94f1db4",
            },
            {
                staffId: "448210",
                rosterGrade: "GR2",
                aadObjectId: "07b8e33c-f86f-440c-ac1f-6725920dbe79",
            },
            {
                staffId: "380824",
                rosterGrade: "PUR",
                aadObjectId: "fff2cfa8-0eb6-4fdc-9902-fa0ba06219b3",
            },
        ],
    },
    {
        tripId: null,
        departureTime: new Date("2018-02-08 10:20:00 UTC+4"),
        flights: [
            {
                flightNumber: "209",
                origin: "DXB",
                destination: "ATH",
            },
            {
                flightNumber: "209",
                origin: "ATH",
                destination: "EWR",
            },
            {
                flightNumber: "210",
                origin: "EWR",
                destination: "ATH",
            },
            {
                flightNumber: "210",
                origin: "ATH",
                destination: "DXB",
            },
        ],
        crewMembers: [
            {
                staffId: "382244",
                rosterGrade: "PUR",
                aadObjectId: "d971021e-cc4d-4c7d-8076-aeaaad494fa7",
            },
            {
                staffId: "420873",
                rosterGrade: "GR1",
                aadObjectId: "1e45791d-0d96-404c-ac7d-9bf977362b1b",
            },
            {
                staffId: "429465",
                rosterGrade: "GR2",
                aadObjectId: "0a971a4f-b0bf-4ce4-8b39-944166165aeb",
            },
            {
                staffId: "442614",
                rosterGrade: "GR2",
                aadObjectId: "aeac155e-8202-472b-8a47-d5cf079e35f1",
            },
            {
                staffId: "441994",
                rosterGrade: "GR2",
                aadObjectId: "60e66497-bf32-4471-bcb2-253ac2fa20fc",
            },
        ],
    },
];

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

        // Commands to manipulate trip database
        this.matches(createTripsRegExp, (session) => { this.handleCreateTrips(session); });
        this.matches(/deleteTrips/i, (session) => { this.handleDeleteTrips(session); });

        // Commands to simulate a time trigger
        this.matches(triggerTimeRegExp, (session) => { this.handleTimeTrigger(session); });

        this.onDefault((session) => { this.onMessageReceived(session); });
    }

    // Handle resumption of dialog
    public dialogResumed<T>(session: builder.Session, result: builder.IDialogResult<T>): void {
        // The only dialog we branch to is auth
        let userToken = utils.getUserToken(session, constants.IdentityProviders.azureADv1);
        if (userToken) {
            this.teamsApi = new teams.UserContextTeamsApi(userToken.accessToken, userToken.expirationTime);
        }
    }

    private async handleCreateTrips(session: builder.Session): Promise<void> {
        let dateString = createTripsRegExp.exec(session.message.text)[1];
        let inputDate = moment.utc(dateString);

        // By default create trips that depart Dubai a week from now
        let baseDate = inputDate.isValid() ? inputDate.toDate() : new Date(new Date().valueOf() + (7 * 24 * 60 * 60 * 1000));

        let fakeTrips: trips.Trip[] = tripTemplates.map((trip) => {
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

            let departureTimes = fakeTrips.map(trip => trip.departureTime.toUTCString()).join(", ");
            session.send(`Created ${fakeTrips.length} trips that depart at the following times ${departureTimes}`);
        } catch (e) {
            console.error(`Error creating trips: ${e.message}`, e);
            session.send(`An error occurred while creating trips: ${e.message}`);
        }
    }

    // Delete all trips in the trip database
    private async handleDeleteTrips(session: builder.Session): Promise<void> {
        try {
            await (<trips.ITripsTest><any>this.tripsApi).deleteAllTripsAsync();
            session.send(`Deleted all trips from the trip database.`);
        } catch (e) {
            console.error(`Error deleting trips: ${e.message}`, e);
            session.send(`An error occurred while deleting trips: ${e.message}`);
        }
    }

    private async handleTimeTrigger(session: builder.Session): Promise<void> {
        let dateString = triggerTimeRegExp.exec(session.message.text)[1];
        let inputDate = moment.utc(dateString);

        // By default simulate a time trigger for the current time, allowing override
        let triggerTime = inputDate.isValid() ? inputDate.toDate() : moment.utc().toDate();
        console.log(`Simulating a time trigger for ${triggerTime.toUTCString()}`);

        try {
            // Archive old teams
            let maxDepartureTimeToArchive = moment(triggerTime).subtract(daysInPastToArchiveTrips, "d").toDate();
            let groupsToArchive = (await this.appDataStore.findActiveGroupsCreatedBeforeTimeAsync(maxDepartureTimeToArchive))
                .filter(groupData => groupData.tripSnapshot.departureTime < maxDepartureTimeToArchive);

            let groupIds = groupsToArchive.map(groupData => groupData.groupId).join(", ");
            console.log(`Found ${groupsToArchive.length} groups to archive: ${groupIds}`);

            let teamArchivePromises = groupsToArchive.map(async (groupData) => {
                try
                {
                    await this.archiveTeamAsync(groupData.groupId);

                    groupData.archivalTime = triggerTime;
                    await this.appDataStore.addOrUpdateGroupDataAsync(groupData);
                }
                catch (e) {
                    console.error(`Error archiving group ${groupData.groupId}: ${e.message}`, e);
                }
            });
            await Promise.all(teamArchivePromises);

            // Create new teams
            let maxDepartureTimeToCreate = moment(triggerTime).add(daysInAdvanceToCreateTrips, "d").toDate();
            let trips = await this.tripsApi.findTripsDepartingInRangeAsync(triggerTime, maxDepartureTimeToCreate);

            let departureTimes = trips.map(trip => trip.departureTime.toUTCString()).join(", ");
            console.log(`Found ${trips.length} trips that depart at the following times ${departureTimes}`);

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

                    console.log(`Team ${groupId} created for trip ${trip.tripId} departing DXB on ${trip.departureTime.toUTCString()}`);
                }
            });
            await Promise.all(teamCreatePromises);

            session.send(`Finished processing time trigger for ${triggerTime.toUTCString()}`);
        } catch (e) {
            let errorMessage = `An error occurred while processing time trigger at ${triggerTime.toUTCString()}: ${e.message}`;
            console.error(errorMessage, e);
            session.send(errorMessage);
        }
    }

    private createDisplayNameForTrip(trip: trips.Trip): string {
        let flightNumbers = _.uniq(trip.flights.map(flight => flight.flightNumber)).join("/");
        let route = _(trip.flights).map(flight => flight.destination).unshift(trip.flights[0].origin).join("-");
        let dxbDepartureDate = moment(trip.departureTime).utcOffset(240).format("YYYY-MM-DD");
        return `EK${flightNumbers} ${route} ${dxbDepartureDate}`;
    }

    // Create a team for a trip, returning the group id of the newly-created team
    private async createTeamForTripAsync(trip: trips.Trip): Promise<string> {
        // Create the team
        let team: teams.Team;
        try {
            let displayName = this.createDisplayNameForTrip(trip);
            team = await this.teamsApi.createTeamAsync(displayName, null, trip.tripId, teamSettings);
            console.log(`Created a new team, group id is ${team.id}`);
        } catch (e) {
            console.error(`Error creating team for trip ${trip.tripId}: ${e.message}`, e);
            throw e;
        }

        // Add team members
        let memberAddPromises = trip.crewMembers.map(async crewMember => {
            try {
                await this.teamsApi.addMemberToGroupAsync(team.id, crewMember.aadObjectId);
            } catch (e) {
                console.error(`Error adding ${crewMember.staffId} (${crewMember.aadObjectId}): ${e.message}`, e);
            }
        });
        await Promise.all(memberAddPromises);

        return team.id;
    }

    private async archiveTeamAsync(groupId: string): Promise<void> {
        // Remove all team members
        let teamMembers = await this.teamsApi.getMembersOfGroupAsync(groupId);
        console.log(`Found ${teamMembers.length} members in the team`);

        let memberRemovePromises = teamMembers.map(async member => {
            try {
                await this.teamsApi.removeMemberFromGroupAsync(groupId, member.id);
            } catch (e) {
                console.error(`Error removing member ${member.id}: ${e.message}`);
            }
        });
        await Promise.all(memberRemovePromises);

        // Rename group
        let group = await this.teamsApi.getGroupAsync(groupId);
        if (!group.displayName.startsWith("[ARCHIVED]")) {
            await this.teamsApi.updateGroupAsync(groupId, {
                displayName: "[ARCHIVED] " + group.displayName,
            });
        }
    }

    // Handle message
    private async onMessageReceived(session: builder.Session): Promise<void> {
        session.send("I didn't understand that.");
    }
}
