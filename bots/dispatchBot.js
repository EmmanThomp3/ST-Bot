// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityHandler } = require("botbuilder");
const { LuisRecognizer, QnAMaker } = require("botbuilder-ai");
const firebaseAdmin = require("firebase-admin");
const firestore = firebaseAdmin.firestore();
const scoreData = firestore.collection("ScoreData");

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

            // First, we use the dispatch model to determine which cognitive service (LUIS or QnA) to use.
            const recognizerResult = await dispatchRecognizer.recognize(
                context
            );

            await scoreData.add(recognizerResult.luisResult);

            // Top intent tell us which cognitive service to use.
            const intent = LuisRecognizer.topIntent(recognizerResult);

            // Next, we call the dispatcher with the top intent.
            await this.dispatchToTopIntentAsync(
                context,
                intent,
                recognizerResult
            );

            await next();
        });

        this.onMembersAdded(async (context, next) => {
            const welcomeText = "Hello, this is ST Bot! How are you today?";
            const membersAdded = context.activity.membersAdded;

            for (const member of membersAdded) {
                if (member.id !== context.activity.recipient.id) {
                    await context.sendActivity(welcomeText);
                }
            }

            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
    }

    async dispatchToTopIntentAsync(context, intent, recognizerResult) {
        switch (intent) {
            case "l_STLUIS":
                await this.processSTLUIS(context, recognizerResult.luisResult);
                break;
            case "q_STKB":
                await this.processSTQnA(context);
                break;
            default:
                await this.processSTQnA(context);
                break;
        }
    }

    async processSTLUIS(context, luisResult) {
        console.log("processSTLUIS");

        // Retrieve LUIS result for Process Automation.
        const result = luisResult.connectedServiceResult;
        const intent = result.topScoringIntent.intent;

        await context.sendActivity(`STLUIS top intent ${intent}.`);
        await context.sendActivity(
            `STLUIS intents detected:  ${luisResult.intents
                .map((intentObj) => intentObj.intent)
                .join("\n\n")}.`
        );

        if (luisResult.entities.length > 0) {
            await context.sendActivity(
                `STLUIS entities were found in the message: ${luisResult.entities
                    .map((entityObj) => entityObj.entity)
                    .join("\n\n")}.`
            );
        }
    }

    async processIntent(context, luisResult) {
        console.log("processIntent");

        const intent = luisResult.topScoringIntent.intent;
        const score = luisResult.topScoringIntent.score;

        await context.sendActivity(`Top intent ${intent}. Score ${score}`);

        if (luisResult.entities.length > 0) {
            await context.sendActivity(
                `Happy entities were found in the message: ${luisResult.entities
                    .map((entityObj) => entityObj.entity)
                    .join("\n\n")}.`
            );
        }
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
