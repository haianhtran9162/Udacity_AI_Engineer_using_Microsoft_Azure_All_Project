// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const { ActivityHandler, MessageFactory } = require('botbuilder');

const { QnAMaker } = require('botbuilder-ai');
const DentistScheduler = require('./dentistscheduler');
const IntentRecognizer = require('./intentrecognizer');

class DentaBot extends ActivityHandler {
    constructor(configuration, qnaOptions) {
        // call the parent constructor
        super();
        if (!configuration) throw new Error('[QnaMakerBot]: Missing parameter. configuration is required');
        // create a QnAMaker connector
        this.QnAMaker = new QnAMaker(configuration.QnAConfiguration, qnaOptions);
        // create a DentistScheduler connector
        this.DentistScheduler = new DentistScheduler(configuration.SchedulerConfiguration);
        // create a IntentRecognizer connector
        this.IntentRecognizer = new IntentRecognizer(configuration.LuisConfiguration);

        this.onMessage(async (context, next) => {
            // send user input to QnA Maker and collect the response in a variable
            // don't forget to use the 'await' keyword
            const qna = await this.QnAMaker.getAnswers(context);
            const luis = await this.IntentRecognizer.executeLuisQuery(context);
            // send user input to IntentRecognizer and collect the response in a variable
            // don't forget 'await'
            // determine which service to respond with based on the results from LUIS //
            if (luis.luisResult.prediction.topIntent === 'ScheduleAppointment' && luis.intents.ScheduleAppointment.score > 0.5) {
                if (luis.entities.$instance && luis.entities.$instance.time) {
                    const time = luis.entities.$instance.time[0].text;
                    const setupAppointment = await this.DentistScheduler.scheduleAppointment(time);
                    await context.sendActivity(setupAppointment);
                    await next();
                    return;
                }
            }
            if (luis.luisResult.prediction.topIntent === 'GetAvailability' && luis.intents.GetAvailability.score > 0.5) {
                const available = await this.DentistScheduler.getAvailability();
                await context.sendActivity(available);
                await next();
                return;
            }
            if (qna[0]) {
                await context.sendActivity(`${ qna[0].answer }`);
            } else {
                await context.sendActivity('I\'m sorry, I don\'t understand. Please try again.');
            }
            await next();
        });

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            // write a custom greeting
            const welcomeText = 'Hi there! I\'m the Contoso Dentistry Bot. I can help you schedule an appointment or answer any questions you have about our services.';
            for (let cnt = 0; cnt < membersAdded.length; ++cnt) {
                if (membersAdded[cnt].id !== context.activity.recipient.id) {
                    await context.sendActivity(MessageFactory.text(welcomeText, welcomeText));
                }
            }
            // by calling next() you ensure that the next BotHandler is run.
            await next();
        });
    }
}

module.exports.DentaBot = DentaBot;
