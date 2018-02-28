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

import * as mongodb from "mongodb";
import * as winston from "winston";
import { GroupData, IAppDataStore } from "./AppDataStore";

const teamsCollectionName = "Teams";
const appDataCollectionName = "AppData";

export class MongoDbAppDataStore implements IAppDataStore {

    private initializePromise: Promise<void>;
    private mongoDb: mongodb.Db;
    private teamsCollection: mongodb.Collection;
    private appDataCollection: mongodb.Collection;

    constructor(
        private connectionString: string) {
    }

    public async addOrUpdateGroupDataAsync(groupData: GroupData): Promise<void> {
        await this.initialize();

        if (!groupData.creationTime) {
            groupData.creationTime = new Date();
        }

        let filter = { groupId: groupData.groupId };
        await this.teamsCollection.updateOne(filter, groupData, { upsert: true });
    }

    public async deleteGroupDataAsync(groupId: string): Promise<void> {
        await this.initialize();

        let filter = { groupId: groupId };
        return await this.teamsCollection.remove(filter);
    }

    public async getGroupDataByGroupAsync(groupId: string): Promise<GroupData> {
        await this.initialize();

        let filter = { groupId: groupId };
        return await this.teamsCollection.findOne(filter);
    }

    public async getGroupDataByTripAsync(tripId: string): Promise<GroupData> {
        await this.initialize();

        let filter = { tripId: tripId };
        return await this.teamsCollection.findOne(filter);
    }

    public async findActiveGroupsCreatedBeforeTimeAsync(endTime: Date): Promise<GroupData[]> {
        await this.initialize();

        let filter = {
            creationTime: {
                "$lte": endTime,
            },
            archivalTime: {
                "$exists": false,
            },
        };
        return await new Promise<GroupData[]>((resolve, reject) => {
            this.teamsCollection.find(filter).toArray((error, documents) => {
                if (error) {
                    reject(error);
                } else {
                    resolve(documents || []);
                }
            });
        });
    }

    public async getAllGroupsAsync(): Promise<GroupData[]> {
        await this.initialize();

        return await new Promise<GroupData[]>((resolve, reject) => {
            this.teamsCollection.find({}).toArray((error, documents) => {
                if (error) {
                    reject(error);
                } else {
                    resolve(documents || []);
                }
            });
        });
    }

    public async getAppDataAsync(key: string): Promise<any> {
        await this.initialize();

        let filter = { key: key };
        let document = await this.appDataCollection.findOne(filter);
        return document && document.data;
    }

    public async setAppDataAsync(key: string, data: any): Promise<void> {
        await this.initialize();

        let filter = { key: key };
        let document = {
            key: key,
            data: data,
        };
        await this.appDataCollection.updateOne(filter, document, { upsert: true });
    }

    // Returns a promise that is resolved when this instance is initialized
    private initialize(): Promise<void> {
        if (!this.initializePromise) {
            this.initializePromise = this.initializeWorker();
        }
        return this.initializePromise;
    }

    // Initialize this instance
    private async initializeWorker(): Promise<void> {
        if (!this.mongoDb) {
            try {
                this.mongoDb = await mongodb.MongoClient.connect(this.connectionString);
                this.teamsCollection = await this.mongoDb.collection(teamsCollectionName);
                this.appDataCollection = await this.mongoDb.collection(appDataCollectionName);

                // Set up indexes
                await this.teamsCollection.createIndex({ tripId: 1 });
                await this.teamsCollection.createIndex({ groupId: 1 });
                await this.teamsCollection.createIndex({ creationTime: 1 });
                await this.appDataCollection.createIndex({ key: 1 });
            } catch (e) {
                winston.error(`Error initializing MongoDB: ${e.message}`, e);
                this.close();
                this.initializePromise = null;
            }
        }
    }

    // Close the connection to the database
    private close(): void {
        this.teamsCollection = null;
        if (this.mongoDb) {
            this.mongoDb.close();
            this.mongoDb = null;
        }
    }
}
