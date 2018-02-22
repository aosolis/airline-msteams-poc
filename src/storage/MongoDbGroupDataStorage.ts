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
import { GroupData, IGroupDataStorage } from "./GroupDataStorage";

export class MongoDbGroupDataStorage implements IGroupDataStorage {

    private initializePromise: Promise<void>;
    private mongoDb: mongodb.Db;
    private tripsCollection: mongodb.Collection;

    constructor(
        private collectionName: string,
        private connectionString: string) {
    }

    public async addOrUpdateGroupDataAsync(groupData: GroupData): Promise<void> {
        await this.initialize();

        if (!groupData.creationTime) {
            groupData.creationTime = new Date();
        }

        let filter = { groupId: groupData.groupId };
        await this.tripsCollection.updateOne(filter, groupData, { upsert: true });
    }

    public async deleteGroupDataAsync(groupId: string): Promise<void> {
        await this.initialize();

        let filter = { groupId: groupId };
        return await this.tripsCollection.remove(filter);
    }

    public async getGroupDataByGroupAsync(groupId: string): Promise<GroupData> {
        await this.initialize();

        let filter = { groupId: groupId };
        return await this.tripsCollection.findOne(filter);
    }

    public async getGroupDataByTripAsync(tripId: string): Promise<GroupData> {
        await this.initialize();

        let filter = { tripId: tripId };
        return await this.tripsCollection.findOne(filter);
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
            this.tripsCollection.find(filter).toArray((error, documents) => {
                if (error) {
                    reject(error);
                } else {
                    resolve(documents || []);
                }
            });
        });
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
                this.tripsCollection = await this.mongoDb.collection(this.collectionName);

                // Set up indexes
                await this.tripsCollection.createIndex({ tripId: 1 });
                await this.tripsCollection.createIndex({ lastUpdate: 1 });
            } catch (e) {
                winston.error(`Error initializing MongoDB: ${e.message}`, e);
                this.close();
                this.initializePromise = null;
            }
        }
    }

    // Close the connection to the database
    private close(): void {
        this.tripsCollection = null;
        if (this.mongoDb) {
            this.mongoDb.close();
            this.mongoDb = null;
        }
    }
}
