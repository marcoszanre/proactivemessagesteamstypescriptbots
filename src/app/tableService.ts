import { TurnContext } from "botbuilder";
import * as debug from "debug";

// tslint:disable-next-line:no-var-requires
const azure = require("azure-storage");

// Initialize debug logging module
const log = debug("msteams");

const tableSvc = azure.createTableService(process.env.STORAGE_ACCOUNT_NAME, process.env.STORAGE_ACCOUNT_ACCESSKEY);

const initTableSvc = () => {
    tableSvc.createTableIfNotExists("proactiveTable", (error, result, response) => {
        if (!error) {
          // Table exists or created
          log("table service done");
        }
    });
};

const insertUserReference = (context: TurnContext) => {

        const conversationReference = JSON.stringify(TurnContext.getConversationReference(context.activity));

        const userReference = {
            PartitionKey: {_: "userReference"},
            RowKey: {_: context.activity.from.aadObjectId},
            user: {_: context.activity.from.name},
            reference: {_: conversationReference},
        };

        tableSvc.insertEntity("proactiveTable", userReference, (error, result, response) => {
            if (!error) {
              // Entity inserted
              log("success!");
            } else {
                log(error);
            }
        });
};


const getUserReference = async (rowkey: string) => {

    return new Promise((resolve, reject) => {

        tableSvc.retrieveEntity("proactiveTable", "userReference", rowkey, (error, result, response) => {
            if (!error) {
                // result contains the entity
                const userReference: IUserReference = {
                    user: result.user._,
                    reference: result.reference._,
                };
                resolve(userReference);
            }
        });
    });
};

interface IUserReference {
    user: string;
    reference: string;
}

const insertConversationID = (user: string, conversationID: string) => {

    const conversationIDReference = {
        PartitionKey: {_: "conversationIDs"},
        RowKey: {_: user},
        conversationID: {_: conversationID}
    };

    tableSvc.insertEntity("proactiveTable", conversationIDReference, (error, result, response) => {
        if (!error) {
          // Entity inserted
          log("success!");
        } else {
            log(error);
        }
    });
};

const getconversationID = async (rowkey: string) => {

    return new Promise((resolve, reject) => {

        tableSvc.retrieveEntity("proactiveTable", "conversationIDs", rowkey, (error, result, response) => {
            if (!error) {
                // result contains the entity
                const convID: IConversationID = {
                    conversationID: result.conversationID._
                };
                resolve(convID);
            } else {
                // conversation ID doesn't exist yet
                const convID: IConversationID = {
                    conversationID: "nouser"
                };
                resolve(convID);
            }
        });
    });
};

const insertUserID = (aadObjectId: string, name: string, id: string) => {

    const userIDReference = {
        PartitionKey: {_: "userIDs"},
        RowKey: {_: aadObjectId},
        name: {_: name},
        id: {_: id}
    };

    tableSvc.insertEntity("proactiveTable", userIDReference, (error, result, response) => {
        if (!error) {
          // Entity inserted
          log("success!");
        } else {
            log(error);
        }
    });
};

const getuserID = async (rowkey: string) => {

    return new Promise((resolve, reject) => {

        tableSvc.retrieveEntity("proactiveTable", "userIDs", rowkey, (error, result, response) => {
            if (!error) {
                // result contains the entity
                const userID: IUserID = {
                    name: result.name._,
                    id: result.id._
                };
                resolve(userID);
            }
        });
    });
};

interface IConversationID {
    conversationID: string;
}

interface IUserID {
    name: string;
    id: string;
}

export {
    initTableSvc,
    insertUserReference,
    getUserReference,
    insertConversationID,
    getconversationID,
    insertUserID,
    getuserID,
    IUserID,
    IUserReference,
    IConversationID
};
