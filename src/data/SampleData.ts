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

import * as trips from "../trips/TripsApi";

// Flight crew members
const crewMembers: trips.CrewMember[] = [
    {
        staffId: "292062",
        rosterGrade: "FG1",
        aadObjectId: "303b75b6-87e1-4f8e-8387-fbfe288356bf",
        displayName: "Hannah Albrecht",
        userPrincipalName: "HannahA@M365x146188.onmicrosoft.com",
    },
    {
        staffId: "378718",
        rosterGrade: "FG1",
        aadObjectId: "0f429da5-2cbf-4d95-bc2c-16a1bef3ed1c",
        displayName: "Alex Wilber",
        userPrincipalName: "AlexW@M365x146188.onmicrosoft.com",
    },
    {
        staffId: "431620",
        rosterGrade: "GR1",
        aadObjectId: "e5a7e50b-8005-4594-93aa-51f62912d1cd",
        displayName: "Allan Deyoung",
        userPrincipalName: "AllanD@M365x146188.onmicrosoft.com",
    },
    {
        staffId: "420501",
        rosterGrade: "GR1",
        aadObjectId: "2238636f-e037-47be-8cda-ca765ff96793",
        displayName: "Harvey Rayford",
        userPrincipalName: "HarveyR@M365x146188.onmicrosoft.com",
    },
    {
        staffId: "431400",
        rosterGrade: "GR1",
        aadObjectId: "ecd0aca2-74ca-41fb-bbef-fd566b0e3aa2",
        displayName: "Debra Berger",
        userPrincipalName: "DebraB@M365x146188.onmicrosoft.com",
    },
    {
        staffId: "450986",
        rosterGrade: "GR2",
        aadObjectId: "4252dcaa-7a49-43e1-95e8-3db616da342d",
        displayName: "Diego Siciliani",
        userPrincipalName: "DiegoS@M365x146188.onmicrosoft.com",
    },
    {
        staffId: "430109",
        rosterGrade: "GR1",
        aadObjectId: "0011a194-4d0e-4372-9016-40b11837f429",
        displayName: "Emily Braun",
        userPrincipalName: "EmilyB@M365x146188.onmicrosoft.com",
    },
    {
        staffId: "381830",
        rosterGrade: "SFS",
        aadObjectId: "bf966547-37f8-43bc-b5b5-48cd7052ef75",
        displayName: "Enrico Cattaneo",
        userPrincipalName: "EnricoC@M365x146188.onmicrosoft.com",
    },
    {
        staffId: "434722",
        rosterGrade: "GR1",
        aadObjectId: "2fdac1c4-69dd-418d-bf3c-9aafc42d950b",
        displayName: "Adriana Napolitani",
        userPrincipalName: "AdrianaN@M365x146188.onmicrosoft.com",
    },
    {
        staffId: "422970",
        rosterGrade: "GR1",
        aadObjectId: "a2d06783-f918-4f4b-af32-5dacb94f1db4",
        displayName: "Gebhard Stein",
        userPrincipalName: "GebhardS@M365x146188.onmicrosoft.com",
    },
    {
        staffId: "448210",
        rosterGrade: "GR2",
        aadObjectId: "07b8e33c-f86f-440c-ac1f-6725920dbe79",
        displayName: "Giorgia Angelo",
        userPrincipalName: "GiorgiaA@M365x146188.onmicrosoft.com",
    },
    {
        staffId: "380824",
        rosterGrade: "PUR",
        aadObjectId: "fff2cfa8-0eb6-4fdc-9902-fa0ba06219b3",
        displayName: "Adele Vance",
        userPrincipalName: "AdeleV@M365x146188.onmicrosoft.com",
    },
    {
        staffId: "382244",
        rosterGrade: "PUR",
        aadObjectId: "d971021e-cc4d-4c7d-8076-aeaaad494fa7",
        displayName: "Christie Cline",
        userPrincipalName: "ChristieC@M365x146188.onmicrosoft.com",
    },
    {
        staffId: "420873",
        rosterGrade: "GR1",
        aadObjectId: "1e45791d-0d96-404c-ac7d-9bf977362b1b",
        displayName: "Henrietta Mueller",
        userPrincipalName: "HenriettaM@M365x146188.onmicrosoft.com",
    },
    {
        staffId: "429465",
        rosterGrade: "GR2",
        aadObjectId: "0a971a4f-b0bf-4ce4-8b39-944166165aeb",
        displayName: "Humayd Zaher",
        userPrincipalName: "HumaydZ@M365x146188.onmicrosoft.com",
    },
    {
        staffId: "442614",
        rosterGrade: "GR2",
        aadObjectId: "aeac155e-8202-472b-8a47-d5cf079e35f1",
        displayName: "Irvin Sayers",
        userPrincipalName: "IrvinS@M365x146188.onmicrosoft.com",
    },
    {
        staffId: "441994",
        rosterGrade: "GR2",
        aadObjectId: "60e66497-bf32-4471-bcb2-253ac2fa20fc",
        displayName: "Isaiah Langer",
        userPrincipalName: "IsaiahL@M365x146188.onmicrosoft.com",
    },
];

// Find a crew member by AAD object ID
export function findCrewMemberByAadObjectId(aadObjectId: string): trips.CrewMember {
    aadObjectId = aadObjectId.toLowerCase();
    return crewMembers.find(member => member.aadObjectId.toLocaleLowerCase() === aadObjectId);
}

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
            findCrewMemberByUpn("HannahA@M365x146188.onmicrosoft.com"),
            findCrewMemberByUpn("AlexW@M365x146188.onmicrosoft.com"),
            findCrewMemberByUpn("AllanD@M365x146188.onmicrosoft.com"),
            findCrewMemberByUpn("HarveyR@M365x146188.onmicrosoft.com"),
            findCrewMemberByUpn("DebraB@M365x146188.onmicrosoft.com"),
            findCrewMemberByUpn("DiegoS@M365x146188.onmicrosoft.com"),
            findCrewMemberByUpn("EmilyB@M365x146188.onmicrosoft.com"),
            findCrewMemberByUpn("EnricoC@M365x146188.onmicrosoft.com"),
            findCrewMemberByUpn("AdrianaN@M365x146188.onmicrosoft.com"),
            findCrewMemberByUpn("GebhardS@M365x146188.onmicrosoft.com"),
            findCrewMemberByUpn("GiorgiaA@M365x146188.onmicrosoft.com"),
            findCrewMemberByUpn("AdeleV@M365x146188.onmicrosoft.com"),
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
            findCrewMemberByUpn("ChristieC@M365x146188.onmicrosoft.com"),
            findCrewMemberByUpn("HenriettaM@M365x146188.onmicrosoft.com"),
            findCrewMemberByUpn("HumaydZ@M365x146188.onmicrosoft.com"),
            findCrewMemberByUpn("IrvinS@M365x146188.onmicrosoft.com"),
            findCrewMemberByUpn("IsaiahL@M365x146188.onmicrosoft.com"),
        ],
    },
];
