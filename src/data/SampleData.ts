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
import * as trips from "../trips/TripsApi";

const tenantDomain = config.get("app.tenantDomain");

// Flight crew members
const crewMembers: trips.CrewMember[] = [
    {
        displayName: "Hannah Albrecht",
        userPrincipalName: "HannahA@" + tenantDomain,
    },
    {
        displayName: "Alex Wilber",
        userPrincipalName: "AlexW@" + tenantDomain,
    },
    {
        displayName: "Allan Deyoung",
        userPrincipalName: "AllanD@" + tenantDomain,
    },
    {
        displayName: "Harvey Rayford",
        userPrincipalName: "HarveyR@" + tenantDomain,
    },
    {
        displayName: "Debra Berger",
        userPrincipalName: "DebraB@" + tenantDomain,
    },
    {
        displayName: "Diego Siciliani",
        userPrincipalName: "DiegoS@" + tenantDomain,
    },
    {
        displayName: "Emily Braun",
        userPrincipalName: "EmilyB@" + tenantDomain,
    },
    {
        displayName: "Enrico Cattaneo",
        userPrincipalName: "EnricoC@" + tenantDomain,
    },
    {
        displayName: "Adriana Napolitani",
        userPrincipalName: "AdrianaN@" + tenantDomain,
    },
    {
        displayName: "Gebhard Stein",
        userPrincipalName: "GebhardS@" + tenantDomain,
    },
    {
        displayName: "Giorgia Angelo",
        userPrincipalName: "GiorgiaA@" + tenantDomain,
    },
    {
        displayName: "Adele Vance",
        userPrincipalName: "AdeleV@" + tenantDomain,
    },
    {
        displayName: "Christie Cline",
        userPrincipalName: "ChristieC@" + tenantDomain,
    },
    {
        displayName: "Henrietta Mueller",
        userPrincipalName: "HenriettaM@" + tenantDomain,
    },
    {
        displayName: "Irvin Sayers",
        userPrincipalName: "IrvinS@" + tenantDomain,
    },
    {
        displayName: "Isaiah Langer",
        userPrincipalName: "IsaiahL@" + tenantDomain,
    },
];

// Find a crew member by UPN
export function findCrewMemberByUpn(upn: string): trips.CrewMember {
    upn = upn.toLowerCase();
    return crewMembers.find(member => member.userPrincipalName.toLowerCase() === upn);
}

// Trip templates used when generating sample data
export const tripTemplates: trips.Trip[] = [
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
            findCrewMemberByUpn("HannahA@" + tenantDomain),
            findCrewMemberByUpn("AlexW@" + tenantDomain),
            findCrewMemberByUpn("AllanD@" + tenantDomain),
            findCrewMemberByUpn("HarveyR@" + tenantDomain),
            findCrewMemberByUpn("DebraB@" + tenantDomain),
            findCrewMemberByUpn("DiegoS@" + tenantDomain),
            findCrewMemberByUpn("EmilyB@" + tenantDomain),
            findCrewMemberByUpn("EnricoC@" + tenantDomain),
            findCrewMemberByUpn("GiorgiaA@" + tenantDomain),
            findCrewMemberByUpn("AdeleV@" + tenantDomain),
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
            findCrewMemberByUpn("ChristieC@" + tenantDomain),
            findCrewMemberByUpn("HenriettaM@" + tenantDomain),
            findCrewMemberByUpn("IrvinS@" + tenantDomain),
            findCrewMemberByUpn("IsaiahL@" + tenantDomain),
            findCrewMemberByUpn("AdrianaN@" + tenantDomain),
            findCrewMemberByUpn("GebhardS@" + tenantDomain),
        ],
    },
];
