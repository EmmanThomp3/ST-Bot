// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityHandler, MessageFactory } = require("botbuilder");
const { LuisRecognizer, QnAMaker } = require("botbuilder-ai");
const { ActionTypes, ActivityTypes } = require("botframework-schema");
const firebaseAdmin = require("firebase-admin");
const firestore = firebaseAdmin.firestore();
const scoreData = firestore.collection("ScoreData");
const summaryData = firestore.collection("SummaryData");
const { v4: uuidv4 } = require("uuid");

const intensityScores = {
    "Happy": 1,
    "Sad": 8
}

const sessions = {};

class DispatchBot extends ActivityHandler {
    constructor() {
        super();

        // If the includeApiResults parameter is set to true, as shown below, the full response
        // from the LUIS api will be made available in the properties  of the RecognizerResult
        const dispatchRecognizer = new LuisRecognizer(
            {
                applicationId: process.env.LuisAppId,
                endpointKey: process.env.LuisAPIKey,
                endpoint: `https://${process.env.LuisAPIHostName}.api.cognitive.microsoft.com`,
            },
            {
                includeAllIntents: true,
                includeInstanceData: true,
            },
            true
        );

        const qnaMaker = new QnAMaker({
            knowledgeBaseId: process.env.QnAKnowledgebaseId,
            endpointKey: process.env.QnAEndpointKey,
            host: process.env.QnAEndpointHostName,
        });

        this.dispatchRecognizer = dispatchRecognizer;
        this.qnaMaker = qnaMaker;

        this.onMessage(async (context, next) => {
            console.log("Processing Message Activity.");

            if (context.activity.channelData.postBack && context.activity.text == "Finish") {
                const id = context.activity.recipient.id;
                const avgIntensity = sessions[id].map(s => s.intensity).reduce((a, b) => a + b) / sessions[id].length;
                const avgScore = sessions[id].map(s => s.score).reduce((a, b) => a + b) / sessions[id].length;
                const keywords = sessions[id].map(s => s.query);
                await summaryData.add({
                    avgIntensity,
                    avgScore,
                    keywords
                })

                sessions[id] = [];

                await context.sendActivity({
                    type: ActivityTypes.EndOfConversation
                });
            } else {
                // First, we use the dispatch model to determine which cognitive service (LUIS or QnA) to use.
                const recognizerResult = await dispatchRecognizer.recognize(
                    context
                );
    
                // Next, we call the dispatcher with the top intent.
                await this.dispatchToTopIntentAsync(
                    context,
                    recognizerResult
                );
                await this.sendFinishAction(context);
            }

            await next();
        });

        this.onMembersAdded(async (context, next) => {
            const id = context.activity.recipient.id;
            sessions[id] = [];

            await this.sendWelcomeMessage(context);

            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
    }

    /**
     * Send a welcome message along with suggested actions for the user to click.
     * @param {TurnContext} turnContext A TurnContext instance containing all the data needed for processing this conversation turn.
     */
    async sendWelcomeMessage(turnContext) {
        const { activity } = turnContext;

        // Iterate over all new members added to the conversation.
        for (const idx in activity.membersAdded) {
            if (activity.membersAdded[idx].id !== activity.recipient.id) {
                const welcomeMessage = `Hello, this is ST Bot! How are you today?`;
                await turnContext.sendActivity(welcomeMessage);
            }
        }
    }

    /**
     * Send suggested actions to the user.
     * @param {TurnContext} turnContext A TurnContext instance containing all the data needed for processing this conversation turn.
     */
    async sendFinishAction(turnContext) {
        const cardActions = [
            {
                type: ActionTypes.PostBack ,
                title: "Finish",
                value: "Finish"
            }
        ];

        var reply = MessageFactory.suggestedActions(cardActions);
        await turnContext.sendActivity(reply);
    }

    async dispatchToTopIntentAsync(context, recognizerResult) {
        // Top intent tell us which cognitive service to use.
        const { topScoringIntent, query } = recognizerResult.luisResult;
        const { score, intent } = topScoringIntent;
        const intensity = intensityScores[intent] | 0;
        const id = context.activity.recipient.id;

        await this.processSTQnA(context);

        const data = { query, intent, score, intensity };
        const docId = (8 - intensity) + "_" + uuidv4();
        sessions[id].push( data );
        await scoreData.doc(docId).set(data);
    }

    async processSTQnA(context) {
        console.log("processSTQnA");

        const results = await this.qnaMaker.getAnswers(context);

        if (results.length > 0) {
            await context.sendActivity(`${results[0].answer}`);
        } else {
            await context.sendActivity(
                "Sorry, could not find an answer in the Q and A system."
            );
        }
    }
}

module.exports.DispatchBot = DispatchBot;
