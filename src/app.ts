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

let express = require("express");
let exphbs  = require("express-handlebars");
import { Request, Response } from "express";
let bodyParser = require("body-parser");
let favicon = require("serve-favicon");
let http = require("http");
let path = require("path");
let logger = require("morgan");
import * as config from "config";
import * as msteams from "botbuilder-teams";
import * as moment from "moment";
import * as winston from "winston";
import * as storage from "./storage";
import * as providers from "./providers";
import * as teams from "./TeamsApi";
import { EmiratesBot } from "./EmiratesBot";
import { MongoDbTripsApi } from "./trips/MongoDbTripsApi";
import { TeamsUpdater } from "./TeamsUpdater";
import { TestDashboard } from "./TestDashboard";
import { UserContextLogin } from "./UserContextLogin";

let app = express();

app.set("port", process.env.PORT || 3978);
app.use(logger("dev"));
app.use(express.static(path.join(__dirname, "../../public")));
app.use(favicon(path.join(__dirname, "../../public/assets", "favicon.ico")));
app.use(bodyParser.json());

let handlebars = exphbs.create({
    extname: ".hbs",
});
app.engine("hbs", handlebars.engine);
app.set("view engine", "hbs");

// Configure API dependencies
let appDataStore = new storage.MongoDbAppDataStore(config.get("mongoDb.connectionString"));
let userTeamsApi = new teams.UserContextTeamsApi(appDataStore, config.get("app.appId"), config.get("app.appPassword"));
let appTeamsApi = new teams.AppContextTeamsApi(config.get("app.tenantDomain"), config.get("app.appId"), config.get("app.appPassword"));
let tripsApi = new MongoDbTripsApi(config.get("mongoDb.connectionString"));
let aadProvider = new providers.AzureADv1Provider(config.get("app.appId"), config.get("app.appPassword"));

let teamsApi = (config.get("app.apiContext") === "user") ? userTeamsApi : appTeamsApi;
let teamsUpdater = new TeamsUpdater(tripsApi, teamsApi, appDataStore, appTeamsApi);

// Update teams
let updateApiKey = config.get("app.updateApiKey");
app.post("/api/updateTeams", async (req, res) => {
    // This is a simple access control mechanism to prevent unauthorized users from triggering updates
    let apiKeyHeader = req.headers["x-api-key"];
    if (apiKeyHeader !== updateApiKey) {
        winston.error("Invalid api key");
        res.sendStatus(401);
        return;
    }

    try {
        let date = new Date();
        let dateParameter = req.query["date"];
        if (dateParameter) {
            winston.info(`Simulating update for date ${dateParameter}`);
            date = moment(dateParameter).toDate();
        }

        await teamsUpdater.updateTeamsAsync(date);
        res.status(200).send(date.toUTCString());
    } catch (e) {
        winston.error("Update teams failed", e);
        res.status(500, e);
    }
});

// Tenant admin consent
app.get("/adminconsent/login", async (req, res) => {
    let tenantDomain = config.get("app.tenantDomain");
    let appId = config.get("app.appId");
    let baseUri = config.get("app.baseUri");
    let endpoint = `https://login.microsoftonline.com/${tenantDomain}/adminconsent?client_id=${appId}&state=12345&redirect_uri=${baseUri}/adminconsent/callback`;
    res.redirect(endpoint);
});
app.get("/adminconsent/callback", (req, res) => {
    res.render("adminconsent-callback", {
        appId: config.get("app.appId"),
        baseUri: config.get("app.baseUri"),
    });
});

// User context management
let userContextLogin = new UserContextLogin(appDataStore, aadProvider);
app.get("/usercontext/login", async (req, res) => {
    await userContextLogin.handleLogin(req, res);
});
app.get("/usercontext/callback", async (req, res) => {
    await userContextLogin.handleCallback(req, res);
});

// Test dashboard
let testDashboard = new TestDashboard(tripsApi, appDataStore, teamsUpdater);
app.get("/test-dashboard", async (req, res) => {
    await testDashboard.renderDashboard(res);
});
app.post("/test-dashboard/execute", async (req, res) => {
    await testDashboard.handleCommand(req, res);
});

// Create bot
let connector = new msteams.TeamsChatConnector({
    appId: config.get("bot.appId"),
    appPassword: config.get("bot.appPassword"),
});
let botSettings = {
    storage: new storage.MongoDbBotStorage(config.get("mongoDb.botStateCollection"), config.get("mongoDb.connectionString")),
    azureADv1: aadProvider,
    appDataStore: appDataStore,
    tripsApi: tripsApi,
    teamsUpdater: teamsUpdater,
};
let bot = new EmiratesBot(connector, botSettings, app);

// Log bot errors
bot.on("error", (error: Error) => {
    winston.error(error.message, error);
});

// Bot routes
app.post("/api/messages", connector.listen());
app.get("/auth/azureADv1/callback", (req, res) => {
    bot.handleOAuthCallback(req, res, "azureADv1");
});

// Ping route
app.get("/ping", (req, res) => {
    res.status(200).send("OK");
});

// error handlers

// development error handler
// will print stacktrace
if (app.get("env") === "development") {
    app.use(function(err: any, req: Request, res: Response, next: Function): void {
        winston.error("Failed request", err);
        res.send(err.status || 500, err);
    });
}

// production error handler
// no stacktraces leaked to user
app.use(function(err: any, req: Request, res: Response, next: Function): void {
    winston.error("Failed request", err);
    res.sendStatus(err.status || 500);
});

http.createServer(app).listen(app.get("port"), function (): void {
    winston.verbose("Express server listening on port " + app.get("port"));
});
