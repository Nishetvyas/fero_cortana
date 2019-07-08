// This loads the environment variables from the .env file
require('dotenv').config();

const builder = require('botbuilder');
const restify = require('restify');
const Store = require('./store');
const spellService = require('./spell-service');

// Setup Restify Server
const server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, () => {
    console.log(`${server.name} listening to ${server.url}`);
});
// Create connector and listen for messages
const connector = new builder.ChatConnector({
    appId: process.env.MICROSOFT_APP_ID,
    appPassword: process.env.MICROSOFT_APP_PASSWORD
});
server.post('/api/messages', connector.listen());


// Default store: volatile in-memory store - Only for prototyping!
var inMemoryStorage = new builder.MemoryBotStorage();
var bot = new builder.UniversalBot(connector, function (session) {
    session.send('Sorry, I did not understand \'%s\'. Type \'help\' if you need assistance.', session.message.text);
}).set('storage', inMemoryStorage); // Register in memory storage


// You can provide your own model by specifing the 'LUIS_MODEL_URL' environment variable
// This Url can be obtained by uploading or creating your model from the LUIS portal: https://www.luis.ai/
const recognizer = new builder.LuisRecognizer(process.env.LUIS_MODEL_URL);
bot.recognizer(recognizer);

// bot.dialog('SearchHotels', [
//     (session, args, next) => {
//         session.send(`Welcome to the Hotels finder! We are analyzing your message: 'session.message.text'`);
//         // try extracting entities
//         const cityEntity = builder.EntityRecognizer.findEntity(args.intent.entities, 'builtin.geography.city');
//         const airportEntity = builder.EntityRecognizer.findEntity(args.intent.entities, 'AirportCode');
//         if (cityEntity) {
//             // city entity detected, continue to next step
//             session.dialogData.searchType = 'city';
//             next({ response: cityEntity.entity });
//         } else if (airportEntity) {
//             // airport entity detected, continue to next step
//             session.dialogData.searchType = 'airport';
//             next({ response: airportEntity.entity });
//         } else {
//             // no entities detected, ask user for a destination
//             builder.Prompts.text(session, 'Please enter your destination');
//         }
//     },
//     (session, results) => {
//         const destination = results.response;
//         let message = 'Looking for hotels';
//         if (session.dialogData.searchType === 'airport') {
//             message += ' near %s airport...';
//         } else {
//             message += ' in %s...';
//         }
//         session.send(message, destination);
//         // Async search
//         Store
//             .searchHotels(destination)
//             .then(hotels => {
//                 // args
//                 session.send(`I found ${hotels.length} hotels:`);
//                 let message = new builder.Message()
//                     .attachmentLayout(builder.AttachmentLayout.carousel)
//                     .attachments(hotels.map(hotelAsAttachment));
//                 session.send(message);
//                 // End
//                 session.endDialog();
//             });
//     }
// ]).triggerAction({
//     matches: 'SearchHotels',
//     onInterrupted:  session => {
//         session.send('Please provide a destination');
//     }
// });
//
// bot.dialog('ShowHotelsReviews', (session, args) => {
//     // retrieve hotel name from matched entities
//     const hotelEntity = builder.EntityRecognizer.findEntity(args.intent.entities, 'Hotel');
//     if (hotelEntity) {
//         session.send(`Looking for reviews of '${hotelEntity.entity}'...`);
//         Store
//             .searchHotelReviews(hotelEntity.entity)
//             .then(reviews => {
//                 let message = new builder.Message()
//                     .attachmentLayout(builder.AttachmentLayout.carousel)
//                     .attachments(reviews.map(reviewAsAttachment));
//                 session.endDialog(message);
//             });
//     }
// }).triggerAction({
//     matches: 'ShowHotelsReviews'
// });

bot.dialog('Help', session => {
    session.endDialog(`Hi! Try asking me things like 'Tell about Fero ', 'What is Tia' or 'What is bidding'`);
}).triggerAction({
    matches: 'Help'
});

bot.dialog('test_intent', session => {
    session.send(`Only Fero brings high-end technologies and industry expertise together to deliver a modern,
                  dramatically better freight execution experience that enables businesses of all kinds to import
                  and export globally, using their mode of choice, be it Road, Air or Sea.`);
}).triggerAction({
    matches: 'test_intent'
});
bot.dialog('bid_def', session => {
  session.send(`Bidding is an offer to set a price by an individual or business for a product or service or a demand that something be done.
                Bidding is used to determine the cost or value of something.
                Bidding can be performed by a "buyer" or "supplier" of a product or service based on the context of the situation.
                Anything else you would like to know `);
}).triggerAction({
  matches: 'bid_def'

});
bot.dialog('Client_Blockchain', session => {
  session.send(`Blockchain is a shared , unchangeable ledger that facilitates the process of
                recording commercial contracts. It helps to secure your contracts. `);
}).triggerAction({
  matches: 'Client_Blockchain'
});
bot.dialog('Explore_TIA', session => {
  session.send(`TIA stands for Transport Interactive Assistant. It is the Log Bot that works as Virtual Transport Coordinators for your shipment.
                You can ask for current status, share documents and bills, type a text or send a recording to the Driver.
                It’s a multi-directional communication channel in which you will be able to view input from Driver. Would you like to know more ?`);
}).triggerAction({
  matches: 'Explore_TIA'
});
bot.dialog('goodby', session => {
  session.endDialog(`Ohk goodby.See you again. `);
}).triggerAction({
  matches: 'goodby'
});

bot.dialog('welcome_tia', session => {
  session.send(` Welcome to Tia cortana bot.How may i help you? `);
}).triggerAction({
  matches: 'welcome_tia'
});
// Spell Check
if (process.env.IS_SPELL_CORRECTION_ENABLED === 'true') {
    bot.use({
        botbuilder: (session, next) => {
            spellService
                .getCorrectedText(session.message.text)
                .then(text => {
                    session.message.text = text;
                    next();
                })
                .catch(error => {
                    console.error(error);
                    next();
                });
        }
    });
}

// Helpers
// const hotelAsAttachment = hotel => {
//     return new builder.HeroCard()
//         .title(hotel.name)
//         .subtitle('%d stars. %d reviews. From $%d per night.', hotel.rating, hotel.numberOfReviews, hotel.priceStarting)
//         .images([new builder.CardImage().url(hotel.image)])
//         .buttons([
//             new builder.CardAction()
//                 .title('More details')
//                 .type('openUrl')
//                 .value('https://www.google.com/search?q=hotels+in+' + encodeURIComponent(hotel.location))
//         ]);
// }
//
// const reviewAsAttachment = review => {
//     return new builder.ThumbnailCard()
//         .title(review.title)
//         .text(review.text)
//         .images([new builder.CardImage().url(review.image)]);
// }
