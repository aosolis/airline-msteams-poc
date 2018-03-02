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
import * as builder from "botbuilder";
import * as msteams from "botbuilder-teams";
import * as winston from "winston";
import * as storage from "./storage";
import * as providers from "./providers";
import * as teams from "./TeamsApi";
import { EmiratesBot } from "./EmiratesBot";
import { MongoDbTripsApi } from "./trips/MongoDbTripsApi";
import { TeamsUpdater } from "./TeamsUpdater";

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

// Configure storage
let botStorageProvider = config.get("storage");
let botStorage = null;
switch (botStorageProvider) {
    case "mongoDb":
        botStorage = new storage.MongoDbBotStorage(config.get("mongoDb.botStateCollection"), config.get("mongoDb.connectionString"));
        break;
    case "memory":
        botStorage = new builder.MemoryBotStorage();
        break;
    case "null":
        botStorage = new storage.NullBotStorage();
        break;
}

// Configure APIs
let appDataStore = new storage.MongoDbAppDataStore(config.get("mongoDb.connectionString"));
let teamsApi = new teams.UserContextTeamsApi(appDataStore, config.get("bot.appId"), config.get("bot.appPassword"));
let tripsApi = new MongoDbTripsApi(config.get("mongoDb.connectionString"));
let teamsUpdater = new TeamsUpdater(tripsApi, teamsApi, appDataStore);

// Create chat bot
let connector = new msteams.TeamsChatConnector({
    appId: config.get("bot.appId"),
    appPassword: config.get("bot.appPassword"),
});
let botSettings = {
    storage: botStorage,
    azureADv1: new providers.AzureADv1Provider(config.get("bot.appId"), config.get("bot.appPassword")),
    appDataStore: appDataStore,
    tripsApi: tripsApi,
    teamsUpdater: teamsUpdater,
};
let bot = new EmiratesBot(connector, botSettings, app);

// Log bot errors
bot.on("error", (error: Error) => {
    winston.error(error.message, error);
});

// Configure bot routes
app.post("/api/messages", connector.listen());
app.get("/auth/azureADv1/callback", (req, res) => {
    bot.handleOAuthCallback(req, res, "azureADv1");
});

// Configure ping route
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
