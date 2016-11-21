'use strict';

require('dotenv').config();

const botauth = require("botauth");
const restify = require('restify');
const builder = require('botbuilder');
const OIDCStrategy = require('passport-azure-ad').OIDCStrategy;
const envx = require("envx");
const expressSession = require('express-session');
//const crypto = require('crypto');
//const querystring = require('querystring');

const WEBSITE_HOSTNAME = envx("WEBSITE_HOSTNAME");
const PORT = envx("PORT", 3998);
const BOTAUTH_SECRET = envx("BOTAUTH_SECRET");

//=========================================================
// Bot Setup
//=========================================================

// Setup Restify Server
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3979, function () {
  console.log('%s listening to %s', server.name, server.url); 
});
  
// Create chat bot
console.log('started...')
console.log(process.env.MICROSOFT_APP_ID);
var connector = new builder.ChatConnector({
  appId: process.env.MICROSOFT_APP_ID,
  appPassword: process.env.MICROSOFT_APP_PASSWORD
});
var bot = new builder.UniversalBot(connector);
server.post('/api/messages', connector.listen());
server.get('/', restify.serveStatic({
  'directory': __dirname,
  'default': 'index.html'
}));
//=========================================================
// Auth Setup
//=========================================================

server.use(restify.queryParser());
server.use(restify.bodyParser());
server.use(expressSession({ secret: BOTAUTH_SECRET, resave: true, saveUninitialized: false }));
//server.use(passport.initialize());

// Use the v2 endpoint (applications configured by apps.dev.microsoft.com)
// For passport-azure-ad v2.0.0, had to set realm = 'common' to ensure authbot works on azure app service


var ba = new botauth.BotAuthenticator(server, bot, { session: true, baseUrl: `https://${WEBSITE_HOSTNAME}`, secret : BOTAUTH_SECRET });

ba.provider("aadv2", (options) => {
    var realm = process.env.MICROSOFT_REALM; 
    console.log(options);
    let oidStrategyv2 = {
      redirectUrl: options.callbackURL, //  redirect: /botauth/aadv2/callback
      realm: realm,
      clientID: process.env.MICROSOFT_APP_ID,
      clientSecret: process.env.MICROSOFT_APP_PASSWORD,
      identityMetadata: 'https://login.microsoftonline.com/' + realm + '/v2.0/.well-known/openid-configuration',
      skipUserProfile: true,
      validateIssuer: false,
      //allowHttpForRedirectUrl: true,
      responseType: 'code',
      responseMode: 'query',
      scope: ['email', 'profile'],
      passReqToCallback: true
    };

    let strategy = oidStrategyv2;

    return new OIDCStrategy(strategy,
        (req, iss, sub, profile, accessToken, refreshToken, done) => {
          if (!profile.displayName) {
            return done(new Error("No oid found"), null);
          }
          profile.accessToken = accessToken;
          profile.refreshToken = refreshToken;

          done(null, profile);
    });
});


//=========================================================
// Bots Dialogs
//=========================================================

bot.dialog("/", new builder.IntentDialog()
    .matches(/logout/, "/logout")
    .matches(/signin/, "/signin")
    .onDefault((session, args) => {
        session.endDialog("welcome");
    })
);

bot.dialog("/logout", (session) => {
    ba.logout(session, "aadv2");
    session.endDialog("logged_out");
});

bot.dialog("/signin", [].concat(
    ba.authenticate("aadv2"),
    (session, args, skip) => {
        let user = ba.profile(session, "aadv2");
        session.send(user.displayName);
        // console.log(user.accessToken);
        // console.log(user.refreshToken);
    }
));