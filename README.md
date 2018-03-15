# emirates-msteams-poc

## Setup

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
        * User.ReadWrite.All
5. Under "Platforms", click on "Add platform", choose "Web", then add the following redirect URLs:
     * `https://<your_ngrok_url>/usercontext/callback`
     * `https://<your_ngrok_url>/adminconsent/callback`
6. Under "Application Secrets", click on "Generate New Password", and remember the generated password.
7. Click "Save".

### Application environment

Set the following environment variables:
* BASE_URI: Base URI of the site, e.g., `https://16a685b5.ngrok.io`
* MICROSOFT_APP_ID: Your app's application id
* MICROSOFT_APP_PASSWORD: Your app's password
* MONGODB_CONNECTION_STRING: Your Mongo database connection string (remember to include the database name) 
* TENANT_DOMAIN: The domain of your tenant, e.g., `M365x263448.onmicrosoft.com`
* ARCHIVEDTEAM_OWNER_UPN: The UPN of the user that will be the owner of archived teams (must be an admin)
* API_CONTEXT: Set to either `user` or `app`, for user context or app context, respectively
* API_KEY: Set to a secret string

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
        "API_KEY": "<secret>",
        "API_CONTEXT": "user",
    }
```

### Build and run
1. Run `npm install`
2. Run `gulp build`
3. Launch the application.
4. Go to the test dashboard at `<BASE_URI>/test-dashboard`. For example, `https://16a685b5.ngrok.io/test-dashboard`.

## Usage

### Initial setup

#### Grant administrator consent
The sample creates and manages teams, which needs permissions that require tenant administrator consent.
1. Go to the test dashboard and click on "Grant admin consent".
2. When prompted, log in as a tenant administrator and authorize the application.

#### Establish user context
If the app is running in user context, the test dashboard will have a "User context" section that indicates the user corresponding to the token that will be used by the app. To set or change the user, click on "Set user" or "Change user".

This user must be an admin, as it will create multiple teams, and a normal user can only create 250 teams.

#### Populate the trips database
To populate the trips database with fake flights, click on "Create trips". This deletes all existing trips, then adds new trips. The simulated trips leave on the 15th day of each month, starting with the next month, and continue for the next 12 months. For example, if the current month is March 2018, trips will be created that depart on 15 Apr 2018, 15 May 2018, ..., 15 Mar 2019.

### Running the sample
To simulate an update trigger, enter the date and time and click on "Simulate".

If the trigger time is T, the update logic does the following:
1. Look for trips departing within the range [T, (T + 7d)], and create teams for each trip. 
    * The team name and description are be based on the trip information.
    * The team roster comprises the cabin crew members working on that flight.
2. Look for teams that correspond to trips depart in the range [(T - 7d), (T + 7d)], and update their roster so that they reflect any crew changes. That is, if crew members were added, they are added to the team, and if crew members were removed from the trip, they are removed from the team.
3. Look for teams that correspond to trips that have departed before (T - 14d), and "archive" those teams. Microsoft Teams does not support true archive functionality yet, so to "archive" a team:
    * Remove all members from the team.
    * Rename the team to add an "[ARCHIVED]" tag.
    * Change the owner to an be "archive owner" user.

### Resetting the sample
The app tracks the teams that it has created, so the app will create a team for a trip only once. When you've reached the end of the 12 months, reset the state of the simulation:
1. Archive all the created teams by simulating a trigger for a date far into the future, for example, 15 Dec 2030.
2. Click on "Delete created teams" to delete all the teams and clear the database.

This does not reset the trips database. The trips pre-populated previously will still be there. To create a new set of trips, click on "Create trips".
