import { BotFrameworkAdapter, MessageFactory, ConversationParameters, Activity } from "botbuilder";
import * as debug from "debug";

// tslint:disable-next-line:no-var-requires
require("dotenv").config();

// Initialize debug logging module
const log = debug("msteams");

// tslint:disable-next-line:no-var-requires
const BotConnector = require("botframework-connector");

let adapter: BotFrameworkAdapter;
let connectorClient;


const initConnectorClient = () => {
    adapter = new BotFrameworkAdapter({
        appId: process.env.MICROSOFT_APP_ID,
        appPassword: process.env.MICROSOFT_APP_PASSWORD
    });

    BotConnector.MicrosoftAppCredentials.trustServiceUrl(
        process.env.SERVICE_URL
    );

    connectorClient = adapter.createConnectorClient(process.env.SERVICE_URL as string);

    log("connector client initialized");

};


const sendUserMessage = async (messageTxt: string, userId: string) => {

        const message = MessageFactory.text(messageTxt) as Activity;

        // User Scope
        const conversationParameters = {
            isGroup: false,
            channelData: {
                tenant: {
                    id: process.env.TENANT_ID
                }
            },
            bot: {
                id: process.env.BOT_ID,
                name: process.env.BOT_NAME
            },
            members: [
                {
                    id: userId
                }
            ]
        };

        const parametersTalk = conversationParameters as ConversationParameters;
        const response = await connectorClient.conversations.createConversation(parametersTalk);
        await connectorClient.conversations.sendToConversation(response.id, message);
        log("user message sent");
};


const sendChannelMessage = async (messageTxt: string) => {

        const message = MessageFactory.text(messageTxt) as Activity;

        // Channel Scope
        const conversationParameters = {
            isGroup: true,
            channelData: {
                channel: {
                    id: process.env.TEAMS_CHANNEL_ID
                }
            },
            activity: message
        };

        const conversationParametersReference = conversationParameters as ConversationParameters;
        await connectorClient.conversations.createConversation(conversationParametersReference);
        log("channel message sent");
};


export {
    sendUserMessage,
    sendChannelMessage,
    initConnectorClient
};

