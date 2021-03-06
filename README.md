# emirates-msteams-poc

## How it works
This demonstrates how to use [Microsoft Graph](https://developer.microsoft.com/en-us/graph/docs/concepts/overview) APIs to automatically provision and retire teams, in the context of an airline operations scenario.

### Problem statement
An airline has a set of cabin crew members that leave together on a trip, with the same group of people staying together throughout a trip. (A trip is comprised of multiple flight segments.) To facilitate collaboration between these staff members, we want to:
* automatically provision a team for the trip several days before the trip departs
* keep the team membership synced with the cabin crew roster while the trip is active
* at the end of the trip, remove all the team members
* keep the team contents accessible for later analysis

### Solution

#### Data model
A [`Trip`](https://github.com/aosolis/emirates-msteams-poc/blob/939c70b1f8c681d6b8a9c660af8c841b5618d653/src/trips/TripsApi.ts#L28) has a unique `tripId` and a `departureTime`. It is comprised of a list of `Flight` segments, each of which has a `flightNumber`, `origin` airport and `destination` airport. The trip also has a list of `CrewMember` entities, which are uniquely identified by their `userPrincipalName`.

Assume that the airline trip database has an [interface](https://github.com/aosolis/emirates-msteams-poc/blob/939c70b1f8c681d6b8a9c660af8c841b5618d653/src/trips/TripsApi.ts#L47) that supports the following operations:
* `getTripAsync(tripId: string): Promise<Trip>`
    * Get the details of a trip, given its trip id
* `findTripsDepartingInRangeAsync(startTime: Date, endTime: Date): Promise<Trip[]>`
    * Find all trips departing between `startTime` and `endTime` (inclusive) 

To keep track of the teams that the app has created, and the status of each team, the app maintains data for each team. [`TeamData`](https://github.com/aosolis/emirates-msteams-poc/blob/939c70b1f8c681d6b8a9c660af8c841b5618d653/src/storage/AppDataStore.ts#L30) has a `groupId`, `tripId`, `creationTime`,  `archivalTime` (not set if the team is active), and a snapshot of the trip details in `tripSnapshot`. The app keeps this in a [store](https://github.com/aosolis/emirates-msteams-poc/blob/939c70b1f8c681d6b8a9c660af8c841b5618d653/src/storage/AppDataStore.ts#L39) that supports the following operations:
* `addOrUpdateTeamDataAsync(teamData: TeamData)`
    * Add or update info about a team that app created
* `deleteTeamDataAsync(groupId: string)`
    * Delete info about a team that app created
* `getTeamDataByGroupAsync(groupId: string): Promise<TeamData>`
    * Get team info given a group (team) id
* `getTeamDataByTripAsync(tripId: string): Promise<TeamData>`
    * Get  team info given a trip id
* `findActiveTeamsCreatedBeforeTimeAsync(endTime: Date): Promise<TeamData[]>`
    * Find active (not archived) teams that were created before `endTime`
* `getAllTeamsAsync(): Promise<TeamData[]>`
    * Get all the teams that were created by this app

#### Core logic
At a periodic interval (e.g., 1 hour), depending on business needs, the app runs through the following steps to create and update teams:
1. [Archive old teams](https://github.com/aosolis/emirates-msteams-poc/blob/939c70b1f8c681d6b8a9c660af8c841b5618d653/src/TeamsUpdater.ts#L296)
    1. Get all active teams.
    2. Determine the teams that correspond to trips that have departed more than 14 days ago.
    3. ["Archive" each team.](https://github.com/aosolis/emirates-msteams-poc/blob/939c70b1f8c681d6b8a9c660af8c841b5618d653/src/TeamsUpdater.ts#L326) Microsoft Teams doesn't yet support true archive functionality, so instead:
        1. Remove all members from the team.
        2. Rename the team to add an "[ARCHIVED]" tag.
        3. Change the owner to an be "archive owner" user. This archive owner must be an administrator, as normal users are limited to being an owner/member of up to 250 groups only.
        4. Mark the team as "archived" in the app data store.
2. [Update active teams](https://github.com/aosolis/emirates-msteams-poc/blob/939c70b1f8c681d6b8a9c660af8c841b5618d653/src/TeamsUpdater.ts#L235)
    1. Get all active teams.
    2. Determine the teams that correspond to trips that haven't left yet, or have left up to 14 days ago.
    3. For each team, synchronize the team membership with the current crew roster:
        1. Get the current trip details.
        2. Get the current team members.
        3. Remove all users that are team members, but no longer in the cabin crew roster.
        4. Add users that are in the cabin crew roster, but are not members of the team.
        5. Update the trip snapshot in the app data store.
3. [Create new teams](https://github.com/aosolis/emirates-msteams-poc/blob/939c70b1f8c681d6b8a9c660af8c841b5618d653/src/TeamsUpdater.ts#L122)
    1. Get all trips departing within the next 7 days.
    3. For each trip, create a team if we don't have one yet:
        1. Check if we have already created a team for this trip, if so, skip this team (we would have updated it in step #2).
        2. Get the current trip details.
        3. [Create a team for the trip.](https://github.com/aosolis/emirates-msteams-poc/blob/939c70b1f8c681d6b8a9c660af8c841b5618d653/src/TeamsUpdater.ts#L156) Set the team name and description based on the trip details.
        4. Add users that are in the cabin crew roster.
        5. Update the team information in the app data store.

Most of these are straightforward member management operations. The only tricky operation is creating a team, where there are a few potential pitfalls:
1. To create a team, you first create a group, then convert ("migrate") it to a team.
2. It takes several seconds for the group information to propagate. During that time, attempting to convert the group to a team will fail. To handle this, the app [waits and retries the operation several times](https://github.com/aosolis/emirates-msteams-poc/blob/939c70b1f8c681d6b8a9c660af8c841b5618d653/src/TeamsApi.ts#L150).
3. If the operation ultimately fails, remember to [delete the group](https://github.com/aosolis/emirates-msteams-poc/blob/939c70b1f8c681d6b8a9c660af8c841b5618d653/src/TeamsApi.ts#L187) so that it is not left in an orphaned state.
4. There's a similar issue when adding members to a team right after it is created. The app works around this problem by [waiting for a few seconds before adding members to a team](https://github.com/aosolis/emirates-msteams-poc/blob/939c70b1f8c681d6b8a9c660af8c841b5618d653/src/TeamsUpdater.ts#L196). Without this delay, the members are added to the group successfully, but will not be visible in the team for several hours.

Team membership operations require the user's AAD object ID, but the trip database contains the user principal name (e.g., AdeleV@microsoft.com). To get the object ID from the UPN, the app uses the [Get user API](https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/user_get) from Microsoft Graph.

## Setting up the application

### Azure AD application
1. Go to the [Application Registration Portal](https://apps.dev.microsoft.com) and sign in.
2. Under "Converged applications", click on "Add an app".
3. Give your app a name, the click "Create". This takes you to the application registration page.
4. Note the application's "Application Id".
4. Under "Microsoft Graph Permissions", add the following:
    * Delegated permissions:
        * offline_access
        * User.Read
        * Group.ReadWrite.All
        * User.Read.All
    * Application permissions:
        * Group.ReadWrite.All
        * User.Read.All
5. Under "Platforms", click on "Add platform", choose "Web", then add the following redirect URLs:
     * `https://<your_ngrok_url>/usercontext/callback`
     * `https://<your_ngrok_url>/adminconsent/callback`
6. Under "Application Secrets", click on "Generate New Password", and remember the generated password.
7. Click "Save".

### Database
This sample uses a Mongo database to:
* store app-level settings
* track the teams that it has created
* simulate a database of airplane trips

If you are using Azure:
1. Create a [Cosmos DB](https://docs.microsoft.com/en-us/azure/cosmos-db/mongodb-introduction) instance with API = "Mongo DB".
2. Go to the "Collections > Browse" panel, select "Add Database" and enter a unique name.
3. Go to "Settings > Connection String" panel, and get a read-write connection string, which will look like `mongodb://<instance>:<password>@<instance>.documents.azure.com:10255/?ssl=true&replicaSet=globaldb`.
4. Insert the database name right before the query string `mongodb://<instance>:<password>@<instance>.documents.azure.com:10255/<databaseName>?ssl=true&replicaSet=globaldb`. This the connection string that you will use in the app configuration.

### Office 365 tenant
Select 2 users and make them administrators:
1. a user that will be used to create teams (see "Establish user context" section below)
2. a user that will be the owner of archived teams (see `ARCHIVEDTEAM_OWNER_UPN` below)

These can be the same user, if desired. Note that as teams are archived they will be "parked" on the archived teams owner,
so that user will end owning many many teams.

### Sample data
Go to `src\data\SampleData.ts` and edit the user names to correspond to users in your test tenant.
_Tip:_ The names in the file correspond to auto-generated users in Microsoft demo tenants.

### Application environment

Set the following environment variables:
* BASE_URI: Base URI of the site, e.g., `https://16a685b5.ngrok.io`
* MICROSOFT_APP_ID: Your app's application id
* MICROSOFT_APP_PASSWORD: Your app's password
* MONGODB_CONNECTION_STRING: Your Mongo database connection string (remember to include the database name) 
* TENANT_DOMAIN: The domain of your tenant, e.g., `M365x263448.onmicrosoft.com`
* API_CONTEXT: Set to either `user` or `app`, for user context or app context, respectively
* UPDATE_API_KEY: Set to a string secret that controls access to the `/api/updateTeams` endpoint
* ARCHIVEDTEAM_OWNER_UPN: The UPN of the user that will be the owner of archived teams (must be an admin)

If you are using app context:
* ACTIVETEAM_OWNER_UPN: The UPN of the user that will set as the owner of active teams (must be an admin if there will be more than 250 teams at a time)

For example, if you're using Visual Studio Code, you would add the section to your `launch.json` file:
```
    "env": {
        "BASE_URI": "https://16a685b5.ngrok.io",
        "MICROSOFT_APP_ID": "19b9213e-2835-4c5c-bdae-7793b4f41774",
        "MICROSOFT_APP_PASSWORD": "<secret>",
        "MONGODB_CONNECTION_STRING": "mongodb://emirates-poc-mongo:<secret>@emirates-poc-mongo.documents.azure.com:10255/M365x263448?ssl=true&replicaSet=globaldb",
        "TENANT_DOMAIN": "M365x263448.onmicrosoft.com",
        "ACTIVETEAM_OWNER_UPN": "TeresaS@M365x263448.onmicrosoft.com",
        "ARCHIVEDTEAM_OWNER_UPN": "PradeepG@M365x263448.onmicrosoft.com",
        "UPDATE_API_KEY": "<secret>",
        "API_CONTEXT": "user",
    }
```

### Build and run
1. Run `npm install`
2. Run `gulp build`
3. Launch the application.
4. Go to the test dashboard at `<BASE_URI>/test-dashboard`. For example, `https://16a685b5.ngrok.io/test-dashboard`.

## Using the test application

### Initial setup

#### 1. Grant administrator consent
The sample creates and manages teams, which needs permissions that require tenant administrator consent.
1. Go to the test dashboard and click on "Grant admin consent".
2. When prompted, log in as a tenant administrator and authorize the application.

#### 2. Establish user context
If the app is running in user context, the test dashboard will have a "User context" section that indicates the user corresponding to the token that will be used by the app. To set or change the user, click on "Set user" or "Change user".

This user must be an admin, as it will create multiple teams, and a normal user can only create 250 teams.

#### 3. Populate the trips database
To populate the trips database with fake flights, click on "Create trips". This deletes all existing trips, then adds new trips. The simulated trips leave on the 15th day of each month, starting with the next month, and continue for the next 12 months. For example, if the current month is March 2018, trips will be created that depart on 15 Apr 2018, 15 May 2018, ..., 15 Mar 2019.

### Simulate updates
To simulate an update trigger, enter the date and time and click on "Simulate".

### Reset the app state
The app tracks the teams that it has created, so the app will create a team for a trip only once. When you've reached the end of the 12 months, reset the state of the simulation:
1. Archive all the created teams by simulating a trigger for a date far into the future, for example, 15 Dec 2030.
2. Click on "Delete created teams" to delete all the teams and clear the database.

This does not reset the trips database. The trips pre-populated previously will still be there. To create a new set of trips, click on "Create trips".

## Graph references
Microsoft Graph has APIs to [manage groups](https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/resources/group) and to [manage teams](https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/resources/team).
* Groups
    * [Create group](https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/api/group_post_groups)
    * [Get group](https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/api/group_get)
    * [Add owner](https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/api/group_post_members)
    * [Add member](https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/api/group_post_members)
    * [Get owners](https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/api/group_list_owners)
    * [Get members](https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/api/group_list_members)
    * [Remove owner](https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/api/group_delete_owners)
    * [Remove member](https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/api/group_delete_members)
    * [Update group](https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/api/group_update)
* Team
    * [Create team](https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/api/team_put_teams)
    * [Get team](https://developer.microsoft.com/en-us/graph/docs/api-reference/beta/api/team_get)
* User
    * [Get user](https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/user_get)