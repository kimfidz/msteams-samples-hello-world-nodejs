'use strict';

module.exports.setup = function(app) {
    var builder = require('botbuilder');
    var teams = require('botbuilder-teams');
    var excuse = require('huh');
    var greeting = require('greeting');
    var randomNumber = require('random-number');
    var config = require('config');
    var botConfig = config.get('bot');
    
    // Create a connector to handle the conversations
    var connector = new teams.TeamsChatConnector({
        // It is a bad idea to store secrets in config files. We try to read the settings from
        // the environment variables first, and fallback to the config file.
        // See node config module on how to create config files correctly per NODE environment
        appId: process.env.MICROSOFT_APP_ID || botConfig.microsoftAppId,
        appPassword: process.env.MICROSOFT_APP_PASSWORD || botConfig.microsoftAppPassword
    });

    var inMemoryStorage = new builder.MemoryBotStorage();
    
    // Define a simple bot with the above connector that echoes what it received
    var bot = new builder.UniversalBot(connector, function(session) {
        // Message might contain @mentions which we would like to strip off in the response
        var text = teams.TeamsMessage.getTextWithoutMentions(session.message);
        text = text.trim().toLowerCase();
        var response = "Thank you for sending " + text + ", but the problem is " + excuse.get().toLowerCase();
        session.send(response);

    }).set('storage', inMemoryStorage);

    var luisAppId = process.env.LuisAppId;
    var luisAPIKey = process.env.LuisAPIKey;
    var luisAPIHostName = process.env.LuisAPIHostName;

    const LuisModelUrl = 'https://' + luisAPIHostName + '/luis/v2.0/apps/' + luisAppId + '?subscription-key=' + luisAPIKey;

    // Create a recognizer that gets intents from LUIS, and add it to the bot
    var recognizer = new builder.LuisRecognizer(LuisModelUrl);
    bot.recognizer(recognizer);

    // Add a dialog for each intent that the LUIS app recognizes.
    // Greeting dialog
    bot.dialog('GreetingDialog',
        (session) => {
            session.send(greeting.random() + " " + session.message.user.name);
            session.endDialog();
        }
    ).triggerAction({
        matches: 'Greeting'
    })

    // Create Maintenance Task dialog
    bot.dialog('CreateMaintenanceTask', [
        function (session) {
            var maintenanceTask = session.dialogData.maintenanceTask = {
                assetName: null ? description : null,
            };

            // Prompt for asset
            builder.Prompts.text(session, 'Ok, creating a new maintenance task. What is the associated asset name?');
        },
        function (session, results) {
            var maintenanceTask = session.dialogData.maintenanceTask;
            if (results.response) {
                maintenanceTask.assetName = teams.TeamsMessage.getTextWithoutMentions(session.message);
            }

            // Prompt for the task description
            builder.Prompts.text(session, 'Please describe the task.');
        },
        function (session, results) {
            var maintenanceTask = session.dialogData.maintenanceTask;
            if (results.response) {
                maintenanceTask.description = teams.TeamsMessage.getTextWithoutMentions(session.message);
            }

            // Prompt for the task priority
            builder.Prompts.choice(session, "What is the task priority?", ["high","medium","low"]);
        },
        function (session, results) {
            var maintenanceTask = session.dialogData.maintenanceTask;
            if (results.response) {
                maintenanceTask.priority = teams.TeamsMessage.getTextWithoutMentions(session.message);
            }

            // TODO: Call Camunda BPMN here and get back a task id.

            // Use a random number for the task id for now.
            var options = {
                min:  100
              , max:  199
              , integer: true
              }
            var taskId = randomNumber(options);

            // Send confirmation to user
            session.endDialog('Maintenance task created.<br/>Task Id: %s<br/>Asset name: %s<br/>Description: %s<br/>Priority: %s',
                taskId, maintenanceTask.assetName, maintenanceTask.description, maintenanceTask.priority);
        }
    ]).triggerAction({ 
        matches: 'MaintenanceTask.Create',
        confirmPrompt: "This will cancel the creation of the maintenance task you started. Are you sure?" 
    }).cancelAction('cancelCreateMaintenanceTask', "Maintenance task cancelled.", {
        matches: /^(cancel|nevermind)/i,
        confirmPrompt: "Are you sure?"
    });

    // Subscribe dialog
    var savedAddress;
    bot.dialog('SubscribeDialog',
        (session) => {
            savedAddress = session.message.address;
            // (Save this information somewhere that it can be accessed later, such as in a database, or session.userData)
            session.userData.savedAddress = savedAddress;
            session.send("Subscribed to Line 1 Filler stopped events");
            session.endDialog();
        }
    ).triggerAction({
        matches: 'Subscribe'
    })

    // Line 1 Filler stopped dialog
    bot.dialog('line1FillerStoppedDialog', function (session, args, next) {
        session.endDialog();
        session.send("Line 1 Filler stopped!");
    });

    // send simple proactive notification
    function sendProactiveMessage(address) {
        var msg = new builder.Message().address(address);
        msg.text('Line 1 Filler stopped!');
        msg.textLocale('en-US');
        bot.send(msg);
    }

    // initiate a proactive dialog  
    function startProactiveDialog(address) {

        // new conversation address, copy without conversationId
        var newConversationAddress = Object.assign({}, address);
        delete newConversationAddress.conversation;

        // start survey dialog
        bot.beginDialog(newConversationAddress, 'line1FillerStoppedDialog', null, function (err) {
            if (err) {
                // error ocurred while starting new conversation. Channel not supported?
                bot.send(new builder.Message()
                    .text('This channel does not support this operation: ' + err.message)
                    .address(address));
            }
        });
    }

    // Setup an endpoint on the router for the bot to listen.
    // NOTE: This endpoint cannot be changed and must be api/messages
    app.post('/api/messages', connector.listen());

    // Do GET this endpoint to send simple proactive notification
    app.get('/api/send-proactive-message', (req, res, next) => {
        sendProactiveMessage(savedAddress);
        res.send('triggered');
        next();
    });

    // Do GET this endpoint to initiate a proactive dialog
    app.get('/api/start-proactive-dialog', (req, res, next) => {
        startProactiveDialog(savedAddress);
        res.send('triggered');
        next();
    });

    // Export the connector for any downstream integration - e.g. registering a messaging extension
    module.exports.connector = connector;
};
