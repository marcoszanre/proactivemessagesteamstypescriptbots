import * as Express from "express";
import * as http from "http";
import * as path from "path";
import * as morgan from "morgan";
import { MsTeamsApiRouter, MsTeamsPageRouter } from "express-msteams-host";
import * as debug from "debug";
import * as compression from "compression";
// import {  } from "../app/tableService";



// Initialize debug logging module
const log = debug("msteams");

log(`Initializing Microsoft Teams Express hosted App...`);

// Initialize dotenv, to use .env file settings if existing
// tslint:disable-next-line:no-var-requires
require("dotenv").config();



// The import of components has to be done AFTER the dotenv config
import * as allComponents from "./TeamsAppsComponents";
import { getUserReference, IUserReference, getconversationID, IConversationID, getuserID, IUserID, insertConversationID } from "./tableService";
import { BotFrameworkAdapter, MessageFactory, ConversationParameters, Activity } from "botbuilder";

// Create the Express webserver
const express = Express();
const port = process.env.port || process.env.PORT || 3007;

// Inject the raw request body onto the request object
express.use(Express.json({
    verify: (req, res, buf: Buffer, encoding: string): void => {
        (req as any).rawBody = buf.toString();
    }
}));
express.use(Express.urlencoded({ extended: true }));

// Express configuration
express.set("views", path.join(__dirname, "/"));

// Add simple logging
express.use(morgan("tiny"));

// Add compression - uncomment to remove compression
express.use(compression());

// Add /scripts and /assets as static folders
express.use("/scripts", Express.static(path.join(__dirname, "web/scripts")));
express.use("/assets", Express.static(path.join(__dirname, "web/assets")));

// routing for bots, connectors and incoming web hooks - based on the decorators
// For more information see: https://www.npmjs.com/package/express-msteams-host
express.use(MsTeamsApiRouter(allComponents));

// routing for pages for tabs and connector configuration
// For more information see: https://www.npmjs.com/package/express-msteams-host
express.use(MsTeamsPageRouter({
    root: path.join(__dirname, "web/"),
    components: allComponents
}));

// Set default web page
express.use("/", Express.static(path.join(__dirname, "web/"), {
    index: "index.html"
}));


// Set the port
express.set("port", port);

// Start the webserver
http.createServer(express).listen(port, () => {
    log(`Server running on ${port}`);
});

// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about adapters.
const adapter = new BotFrameworkAdapter({
    appId: process.env.MICROSOFT_APP_ID,
    appPassword: process.env.MICROSOFT_APP_PASSWORD
});

// tslint:disable-next-line:no-var-requires
const BotConnector = require("botframework-connector");
const credentials = new BotConnector.MicrosoftAppCredentials(
  process.env.MICROSOFT_APP_ID,
  process.env.MICROSOFT_APP_PASSWORD
);

BotConnector.MicrosoftAppCredentials.trustServiceUrl(
  process.env.SERVICE_URL
);

// Send user proactive message
express.get("/api/proactive", async (req, res, next) => {

    const message = MessageFactory.text("This is a user message 😀") as Activity;
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
                id: "29:1r48gyAgyrbiAeDNnSVcd99hKNML6XcwBorYH4OOxZjzBCFYHtRKZMW3c2at7SLedQCCYvGTYWbvbw8VT5fBAjA",
                name: "MOD Administrator"
            }
        ]
    };

    const parametersTalk = conversationParameters as ConversationParameters;
    const connectorClient = adapter.createConnectorClient(process.env.SERVICE_URL as string);
    const response = await connectorClient.conversations.createConversation(parametersTalk);
    await connectorClient.conversations.sendToConversation(response.id, message);

    res.send("Message sent");
    next();
});


// Send channel proactive message
express.post("/api/notification", async (req, res, next) => {

    const nametxt = req.body.name;
    const teamsChannelId = process.env.TEAMS_CHANNEL_ID;
    const message = MessageFactory.text(`This is the name of the request: ${nametxt}`);

    const conversationParameters = {
        isGroup: true,
        channelData: {
            channel: {
                id: teamsChannelId
            }
        },
        activity: message
    };

    const conversationParametersReference = conversationParameters as ConversationParameters;
    const connectorClient = adapter.createConnectorClient(process.env.SERVICE_URL as string);
    await connectorClient.conversations.createConversation(conversationParametersReference);

    res.send("Message sent");
    next();
});

// Store and retrieve conversation id
express.post("/createUserConversation", (req, res, next) => {
    // log(req.body.userid);
    getconversationID(req.body.userid as string).then(async (convID: IConversationID) => {
        if (convID.conversationID !== "nouser") {
            // Conversation ID Found, send the message
            const message = MessageFactory.text("This is a user message 😀") as Activity;
            const connectorClient = adapter.createConnectorClient(process.env.SERVICE_URL as string);
            await connectorClient.conversations.sendToConversation(convID.conversationID, message);

            res.send("conversation continued");
            next();
        } else {
            log("user NOT found");
            // Get user id
            getuserID(req.body.userid as string).then(async (userID: IUserID) => {
                const message = MessageFactory.text("This is a user message 😀") as Activity;

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
                            id: userID.id,
                            name: userID.name
                        }
                    ]
                };

                const parametersTalk = conversationParameters as ConversationParameters;
                const connectorClient = adapter.createConnectorClient(process.env.SERVICE_URL as string);
                const response = await connectorClient.conversations.createConversation(parametersTalk);
                await connectorClient.conversations.sendToConversation(response.id, message);

                insertConversationID(req.body.userid, response.id);

                res.send("Message sent");
                next();
            });

        }

    });
});

