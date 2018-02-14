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
// Flight API
// =========================================================

export interface Flight {
    id: string;
    flightNumber: string;
    origin: Airport;
    destination: Airport;
    aircraft: string;
    aircraftType: string;
    scheduledDeparture: DateTimeOffset;
    scheduledArrival: DateTimeOffset;
    crew: CrewMember[];
}

export interface Airport {
    code: string;
    name: string;
}

export interface DateTimeOffset {
    utcTime: number;
    timeZoneOffsetInMinutes: number;
    timeZoneName: string;
}

export interface CrewMember {
    objectId: string;
    name: string;
    role: Role;
}

export type Role = "captain" | "firstOfficer" | "purser" | "flightAttendant";

// Stub airline system API
export class AirlineApi {

    public async getFlightInfoAsync(id: string): Promise<Flight> {
        return {
            id: "03d705a0e772408eb74864ce8330e380",
            flightNumber: "EK764",
            origin: {
                code: "JNB",
                name: "OR Tambo Int'l (Johannesburg)",
            },
            destination: {
                code: "DXB",
                name: "Dubai Int'l",
            },
            aircraft: "A6-EEC",
            aircraftType: "A388",
            scheduledDeparture: {
                utcTime: new Date("2018-02-07 16:50 UTC").valueOf(),
                timeZoneName: "+02",
                timeZoneOffsetInMinutes: 120,
            },
            scheduledArrival: {
                utcTime: new Date("2018-02-08 1:05 UTC").valueOf(),
                timeZoneName: "+04",
                timeZoneOffsetInMinutes: 240,
            },
            crew: [
                {
                    objectId: "0f429da5-2cbf-4d95-bc2c-16a1bef3ed1c",
                    name: "Alex Wilber",
                    role: "captain",
                },
                {
                    objectId: "e5a7e50b-8005-4594-93aa-51f62912d1cd",
                    name: "Allan Deyoung",
                    role: "firstOfficer",
                },
                {
                    objectId: "fff2cfa8-0eb6-4fdc-9902-fa0ba06219b3",
                    name: "Adele Vance",
                    role: "purser",
                },
            ],
        };
    }

    public async findFlightsDepartingInRange(startTimeUTC: Date, endTimeUTC: Date): Promise<Flight[]> {
        return [
            {
                id: "03d705a0e772408eb74864ce8330e380",
                flightNumber: "EK764",
                origin: {
                    code: "JNB",
                    name: "OR Tambo Int'l (Johannesburg)",
                },
                destination: {
                    code: "DXB",
                    name: "Dubai Int'l",
                },
                aircraft: "A6-EEC",
                aircraftType: "A388",
                scheduledDeparture: {
                    utcTime: new Date("2018-02-07 16:50 UTC").valueOf(),
                    timeZoneName: "+02",
                    timeZoneOffsetInMinutes: 120,
                },
                scheduledArrival: {
                    utcTime: new Date("2018-02-08 1:05 UTC").valueOf(),
                    timeZoneName: "+04",
                    timeZoneOffsetInMinutes: 240,
                },
                crew: [
                    {
                        objectId: "0f429da5-2cbf-4d95-bc2c-16a1bef3ed1c",
                        name: "Alex Wilber",
                        role: "captain",
                    },
                    {
                        objectId: "e5a7e50b-8005-4594-93aa-51f62912d1cd",
                        name: "Allan Deyoung",
                        role: "firstOfficer",
                    },
                    {
                        objectId: "fff2cfa8-0eb6-4fdc-9902-fa0ba06219b3",
                        name: "Adele Vance",
                        role: "purser",
                    },
                ],
            },
        ];
    }
}
