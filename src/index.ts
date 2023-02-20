import * as path from 'path';
import * as dotenv from 'dotenv';
import express from 'express';
import { botFrameworkAuth } from "./auth";
import { CloudAdapter, ConfigurationBotFrameworkAuthentication, ConfigurationServiceClientCredentialFactory, MemoryStorage, ShowTypingMiddleware, UserState } from 'botbuilder';
import { TenantSsoBot } from './bots/teamsTenantSsoBot';

// Read config from our .env file
const env_file = path.join(__dirname, "..", ".env");
dotenv.config({path: env_file});

// Setup express
const app = express();

// Configure CORS
app.use((req, res, next) => {
    res.header("Access-Control-Allow-Origin", "*");
    res.header("Access-Control-Allow-Headers", "X-Requested-With");
    return next();
});

// Allow express to read json content
app.use(express.json());

// Setup our static web content
const staticViewsPath = path.join(__dirname, "..", "src\\StaticViews");
app.use("/StaticViews", express.static(staticViewsPath));

// Create our cloud adapter
// This is what is used to tal to the bot service API
const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
    MicrosoftAppId: process.env.MicrosoftAppId,
    MicrosoftAppPassword: process.env.MicrosoftAppPassword,
    MicrosoftAppTenantId: process.env.MicrosoftDirectoryId
});

const botFrameworkAuthentication = new ConfigurationBotFrameworkAuthentication(undefined, credentialsFactory);

// Create the cloud adapter with our app credentials (used to get a bot service token)
const adapter = new CloudAdapter(botFrameworkAuthentication);

// Add the built-in middleware that shows the typing indicator when a message is being processed
adapter.use(new ShowTypingMiddleware());

// Handle any turn context errors
adapter.onTurnError = async (context, error) => {
    console.error(`\n [onTurnError] unhandled error: ${ error }`);

    // Send a trace activity, which will be displayed in Bot Framework Emulator
    await context.sendTraceActivity(
        'OnTurnError Trace',
        `${ error }`,
        'https://www.botframework.com/schemas/error',
        'TurnError'
    );

    // Send a message to the user
    await context.sendActivity('The bot encountered an error or bug.');
    await context.sendActivity('To continue to run this bot, please fix the bot source code.');
}

// You can persist state across requests using Blob storage or CosmosDB using the built-in storage implementations
// ... or build your own!
const memoryStorage = new MemoryStorage();

// User state allows us to maintain state for each user across requests, regardless of where they came from
const userState = new UserState(memoryStorage);

// Create our bot instance
const bot = new TenantSsoBot(userState);

app.post('/api/messages', botFrameworkAuth, async (req, res) => {
    await adapter.process(req, res, context => bot.run(context));
});

const port = process.env.port || process.env.PORT || 3978;

// Start the server
app.listen(port, () => {
    console.log(`\nListening to ${ port }`);
});



