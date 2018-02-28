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
import { IAppDataStore } from "../storage/AppDataStore";
import { MongoDbAppDataStore } from "../storage/MongoDbAppDataStore";
import { AzureADv1Dialog } from "./AzureADv1Dialog";
import { TeamsUpdater } from "../TeamsUpdater";
let uuidv4 = require("uuid/v4");

const createTripsRegExp = /^createTrips(.*)$/i;
const showTripRegExp = /^showTrip (.*)$/i;
const addRemoveCrewRegExp = /^(add|remove)Crew (.*) (.*)$/i;
const triggerTimeRegExp = /^triggerTime(.*)$/i;

// Root dialog for driving the bot
export class RootDialog extends builder.IntentDialog
{
    private teamsApi: teams.TeamsApi = new teams.AppContextTeamsApi(config.get("app.tenantId"), config.get("bot.appId"), config.get("bot.appPassword"));
    private tripsApi: trips.ITripsApi = new MongoDbTripsApi(config.get("mongoDb.connectionString"));
    private appDataStore: IAppDataStore = new MongoDbAppDataStore(config.get("mongoDb.connectionString"));
    private teamsUpdater: TeamsUpdater = new TeamsUpdater(this.tripsApi, this.teamsApi, this.appDataStore);

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
            this.teamsUpdater = new TeamsUpdater(this.tripsApi, this.teamsApi, this.appDataStore);
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
        let triggerTimeString = triggerTime.toUTCString();
        winston.info(`Simulating a time trigger for ${triggerTimeString}`);

        try {
            await this.teamsUpdater.updateTeamsAsync(triggerTime);
            session.send(`Finished processing the time trigger for ${triggerTimeString}`);
        } catch (e) {
            let errorMessage = `Error processing time trigger for ${triggerTimeString}: ${e.message}`;
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
}
