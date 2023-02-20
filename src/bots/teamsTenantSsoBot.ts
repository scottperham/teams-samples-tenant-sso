import { TeamsActivityHandler, TurnContext, UserState, MessageFactory, CardFactory, SigninStateVerificationQuery } from "botbuilder";
import { randomUUID } from "crypto";
import jwt, { JwtPayload } from 'jsonwebtoken';

export class TenantSsoBot extends TeamsActivityHandler {

    userState: UserState;

    constructor(userState: UserState) {
        super();

        this.userState = userState;
        
        this.onMessage(async (context, next) : Promise<void> => {
            const commandText = context.activity.text.replace(/\s+/g, "").toLowerCase();

            if (commandText === "signin")
            {
                // This is a very specific Adaptive Card that Teams knows how to
                // handle. It will get a token for the current user and check
                // that they have consented to your access_as_user scope
                const activity = MessageFactory.attachment({
                    contentType: CardFactory.contentTypes.oauthCard,
                    content: {
                        tokenExchangeResource: {
                            id: randomUUID()
                        },
                        connectionName: process.env.OAuthConnectionName
                    }
                });
    
                await context.sendActivity(activity);
            }
            else {
                await context.sendActivity("Sorry, I didn't recognise that command. Type 'sign in'");
            }

            await next();
        });

        this.onInstallationUpdate(async (context, next): Promise<void> => {
            // If the app was updated or uninstalled, clear the welcome message state for the current user
            if (context.activity.action == "add") {
                await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard({
                    "type": "AdaptiveCard",
                    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                    "version": "1.4",
                    "body": [
                        {
                            "type": "TextBlock",
                            "text": "Welcome message! At this point, no SSO has taken place - the information below is taken from the conversation context",
                            "wrap": true
                        },
                        {
                            "type": "FactSet",
                            "facts": [
                                {
                                    "title": "Tenant Id",
                                    "value": context.activity.conversation.tenantId
                                },
                                {
                                    "title": "AAD Object ID",
                                    "value": context.activity.from.aadObjectId
                                }
                            ]
                        }
                    ]
                })));
            }
            await next();
        });
    }

    // This is the entry point for the bot processing pipeline
    // Generally we want the base class to handle the initial processing
    // but this is a great place to save any state changes we've set
    // during the turn
    async run(context: TurnContext): Promise<void> {
        await super.run(context);

        await this.userState.saveChanges(context);
    }

    // Handles the callback from a signin and consent attempt - the token is in `context.activity.value.token`
    protected async handleTeamsSigninTokenExchange(context: TurnContext, query: SigninStateVerificationQuery): Promise<void> {
        
        var token = context.activity.value.token;
        
        const decoded = <JwtPayload>jwt.decode(token);

        await context.sendActivity(MessageFactory.attachment(CardFactory.adaptiveCard({
            "type": "AdaptiveCard",
            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
            "version": "1.4",
            "body": [
                {
                    "type": "TextBlock",
                    "text": "Post sign-in message! The information below is taken from the AAD token",
                    "wrap": true
                },
                {
                    "type": "FactSet",
                    "facts": [
                        {
                            "title": "Tenant Id",
                            "value": decoded.tid
                        },
                        {
                            "title": "AAD Object ID",
                            "value": decoded.oid
                        }
                    ]
                }
            ]
        })));

    }
}