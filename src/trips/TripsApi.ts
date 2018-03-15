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

// =========================================================
// Trips API
// =========================================================

export interface Trip {
    tripId: string;
    departureTime: Date;
    flights: Flight[];
    crewMembers: CrewMember[];
}

export interface Flight {
    flightNumber: string;
    origin: string;
    destination: string;
}

export interface CrewMember {
    userPrincipalName: string;
    displayName?: string;
}

// Interface to the trip database
export interface ITripsApi {
    // Get information about the trip with the given id
    getTripAsync(tripId: string): Promise<Trip>;

    // Find trips that are departing in the given time range
    findTripsDepartingInRangeAsync(startTime: Date, endTime: Date): Promise<Trip[]>;
}

// Test interface to the trip database
export interface ITripsTest {
    // Add or update a trip
    addOrUpdateTripAsync(trip: Trip): Promise<void>;

    // Delete all trips in the database
    deleteAllTripsAsync(): Promise<void>;
}
