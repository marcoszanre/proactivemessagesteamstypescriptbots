import { BotDeclaration, MessageExtensionDeclaration, PreventIframe } from "express-msteams-host";
import * as debug from "debug";
import { DialogSet, DialogState } from "botbuilder-dialogs";
import { StatePropertyAccessor, CardFactory, TurnContext, MemoryStorage, ConversationState, ActivityTypes, TeamsActivityHandler, ConversationReference, MessageFactory, ConversationParameters, Activity, BotFrameworkAdapter, TeamsInfo } from "botbuilder";
import HelpDialog from "./dialogs/HelpDialog";
import WelcomeCard from "./dialogs/WelcomeDialog";
import { initTableSvc, insertUserReference, getUserReference, IUserReference, insertUserID } from "../tableService";


// Initialize debug logging module
const log = debug("msteams");

/**
 * Implementation for proactive messages Bot
 */
@BotDeclaration(
    "/api/messages",
    new MemoryStorage(),
    process.env.MICROSOFT_APP_ID,
    process.env.MICROSOFT_APP_PASSWORD)

export class ProactiveMessagesBot extends TeamsActivityHandler {
    private readonly conversationState: ConversationState;
    private readonly dialogs: DialogSet;
    private dialogState: StatePropertyAccessor<DialogState>;

    /**
     * The constructor
     * @param conversationState
     */
    public constructor(conversationState: ConversationState) {
        super();

        // Init table service
        initTableSvc();
        
        this.conversationState = conversationState;
        this.dialogState = conversationState.createProperty("dialogState");
        this.dialogs = new DialogSet(this.dialogState);
        this.dialogs.add(new HelpDialog("help"));

        // Set up the Activity processing

        this.onMessage(async (context: TurnContext): Promise<void> => {
            // Save conversation reference
            if (context.activity.conversation.conversationType === "personal") {
                // log(context.activity);
                switch (context.activity.text) {
                    case 'get':
                        await getUserReference(context.activity.from.aadObjectId as string).then(async (userRefence: IUserReference) => {
                            // log(userRefence);
                            const notifyConversationReference: Partial<ConversationReference> = JSON.parse(userRefence.reference);
                            await context.adapter.continueConversation(notifyConversationReference, async turnContext => {
                                await turnContext.sendActivity("this is a continued conversation");
                            });
                        });
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
            
            // Save state changes
            return this.conversationState.saveChanges(context);
        });

        this.onConversationUpdate(async (context: TurnContext): Promise<void> => {
            if (context.activity.membersAdded && context.activity.membersAdded.length !== 0) {
                for (const idx in context.activity.membersAdded) {
                    if (context.activity.membersAdded[idx].id === context.activity.recipient.id) {
                        const welcomeCard = CardFactory.adaptiveCard(WelcomeCard);
                        await context.sendActivity({ attachments: [welcomeCard] });
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
    const message = MessageFactory.text("This is a channel message") as Activity;
   
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
    await connectorClient.conversations.createConversation(conversationParameters);
    await context.sendActivity("conversation sent to channel");
    }


}
