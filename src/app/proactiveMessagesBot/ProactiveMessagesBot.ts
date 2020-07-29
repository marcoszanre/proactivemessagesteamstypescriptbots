import { BotDeclaration } from "express-msteams-host";
import { TurnContext, MemoryStorage, TeamsActivityHandler, MessageFactory, ConversationParameters, Activity, BotFrameworkAdapter, TeamsInfo } from "botbuilder";
import { initTableSvc, insertUserReference, insertUserID } from "../tableService";

/**
 * Implementation for proactive messages Bot
 */
@BotDeclaration(
    "/api/messages",
    new MemoryStorage(),
    process.env.MICROSOFT_APP_ID,
    process.env.MICROSOFT_APP_PASSWORD)

export class ProactiveMessagesBot extends TeamsActivityHandler {

    /**
     * The constructor
     * @param conversationState
     */
    public constructor() {
        super();

        // Init table service
        initTableSvc();

        // Set up the Activity processing
        this.onMessage(async (context: TurnContext): Promise<void> => {

            if (context.activity.conversation.conversationType === "personal") {
                switch (context.activity.text) {
                    case 'get':
                        await this.createUserConversation(context);
                        break;
                    case 'channel':
                        await this.teamsCreateConversation(context);
                        break;
                    default:
                        insertUserReference(context);
                        break;
                }
            } else {
                // Channel conversation
                let text = TurnContext.removeRecipientMention(context.activity);
                text = text.toLowerCase();
                text = text.trim();

                if (text === "users") {

                    // Store user ids
                    const teamMembers = await TeamsInfo.getTeamMembers(context);
                    teamMembers.forEach(teamMember => {
                        insertUserID(teamMember.aadObjectId as string, teamMember.name, teamMember.id);
                    });
                }
            }

            await context.sendActivity("thanks for your message 😀");
        });

        this.onConversationUpdate(async (context: TurnContext): Promise<void> => {
            if (context.activity.membersAdded && context.activity.membersAdded.length !== 0) {
                for (const idx in context.activity.membersAdded) {
                    if (context.activity.membersAdded[idx].id === context.activity.recipient.id) {
                        await context.sendActivity("thanks for your message 😀");
                    }
                }
            }
        });

        this.onMessageReaction(async (context: TurnContext): Promise<void> => {
            const added = context.activity.reactionsAdded;
            if (added && added[0]) {
                await context.sendActivity({
                    textFormat: "xml",
                    text: `That was an interesting reaction (<b>${added[0].type}</b>)`
                });
            }
        });;
   }

   async teamsCreateConversation(context: TurnContext) {
    // Create Channel conversation
    let message = MessageFactory.text("This is the first channel message") as Activity;
   
    const conversationParameters: ConversationParameters = {
        isGroup: true,
        channelData: {
            channel: {
                id: process.env.TEAMS_CHANNEL_ID
            }
        },
        bot: {
            id: context.activity.recipient.id,
            name: context.activity.recipient.name
        },
        activity: message
    };

    const notifyAdapter = context.adapter as BotFrameworkAdapter;

    const connectorClient = notifyAdapter.createConnectorClient(context.activity.serviceUrl);
    const response = await connectorClient.conversations.createConversation(conversationParameters);
    await context.sendActivity("conversation sent to channel");

    // Send reply to channel
    message = MessageFactory.text("This is the second channel message") as Activity;
    await connectorClient.conversations.sendToConversation(response.id, message);
    
    }

    async createUserConversation(context: TurnContext) {
        // Create User Conversation
        const message = MessageFactory.text("This is a proactive message") as Activity;
       
        const conversationParameters = {
            isGroup: false,
            channelData: {
                tenant: {
                    id: process.env.TENANT_ID
                }
            },
            members: [
                {
                    id: context.activity.from.id,
                    name: context.activity.from.name
                }
            ]
        };
    
        const notifyAdapter = context.adapter as BotFrameworkAdapter;
        const parametersTalk = conversationParameters as ConversationParameters;
    
        const connectorClient = notifyAdapter.createConnectorClient(context.activity.serviceUrl);
        const response = await connectorClient.conversations.createConversation(parametersTalk);
        await connectorClient.conversations.sendToConversation(response.id, message);
        await context.sendActivity("conversation sent to user");
        
        }

}
