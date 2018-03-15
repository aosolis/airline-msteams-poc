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

import * as config from "config";
import * as jwt from "jsonwebtoken";
import * as moment from "moment";
import * as constants from "./constants";
import * as sampledata from "./data/SampleData";
import { Request, Response } from "express";
import { Trip, ITripsApi, ITripsTest } from "./trips/TripsApi";
import { TeamsUpdater } from "./TeamsUpdater";
import { IAppDataStore } from "./storage/AppDataStore";
const uuidv4 = require("uuid/v4");

// Test dashboard
export class TestDashboard
{
    constructor(
        private tripsApi: ITripsApi,
        private appDataStore: IAppDataStore,
        private teamsUpdater: TeamsUpdater,
    ) {
    }

    // Render the test dashboard
    public async renderDashboard(res: Response): Promise<void> {
        let isUserContext = (config.get("app.apiContext") === "user");
        let locals = {
            appId: config.get("bot.appId"),
            tenantDomain: config.get("app.tenantDomain"),
            baseUri: config.get("app.baseUri"),
            apiKey: config.get("app.apiKey"),
            isUserContext: isUserContext,
        };

        // Get additional info for user context
        if (isUserContext) {
            let userToken = await this.appDataStore.getAppDataAsync(constants.AppDataKey.userToken);
            if (userToken && userToken.idToken) {
                let decodedToken = jwt.decode(userToken.idToken, { complete: true });
                locals["name"] = decodedToken.payload.name;
                locals["upn"] = decodedToken.payload.upn;
            }
        }

        res.render("test-dashboard", locals);
    }

    // Handle test commands
    public async handleCommand(req: Request, res: Response): Promise<void> {
        let tripsTest = <ITripsTest><any>this.tripsApi;

        switch (req.query["command"]) {
            case "deleteTeams":
                await this.teamsUpdater.deleteAllTrackedTeamsAsync();
                break;

            case "createTrips":
                await tripsTest.deleteAllTripsAsync();

                // first occurrence is the 15th of the next month
                let date = moment().add(1, "month").utc().date(15);
                for (let i = 0; i < 12; i++) {
                    await this.createTrips(tripsTest, date.toDate());
                    date = date.add(1, "month");
                }
                break;
        }

        res.sendStatus(200);
    }

    // Create trips in the database that leave on the given date
    private async createTrips(tripsTest: ITripsTest, date: Date): Promise<void> {
        let fakeTrips: Trip[] = sampledata.tripTemplates.map((trip) => {
            let departureTime = trip.departureTime;
            return {
                ...trip,
                tripId: <string>uuidv4(),
                departureTime: new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate(), departureTime.getUTCHours(), departureTime.getUTCMinutes())),
            };
        });
        let addPromises = fakeTrips.map(trip => tripsTest.addOrUpdateTripAsync(trip));
        await Promise.all(addPromises);
    }
}
