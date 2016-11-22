'use strict';

const botauth = require("botauth");
const restify = require('restify');
const builder = require('botbuilder');
const OIDCStrategy = require('passport-azure-ad').OIDCStrategy;
const envx = require("envx");
const expressSession = require('express-session');

const WEBSITE_HOSTNAME = envx("WEBSITE_HOSTNAME");
const PORT = envx("PORT", 3998);
const BOTAUTH_SECRET = envx("BOTAUTH_SECRET");

//bot application identity
const MICROSOFT_APP_ID = envx("MICROSOFT_APP_ID");
const MICROSOFT_APP_PASSWORD = envx("MICROSOFT_APP_PASSWORD");

//oauth details for dropbox
const AZUREAD_APP_ID = envx("AZUREAD_APP_ID");
const AZUREAD_APP_PASSWORD = envx("AZUREAD_APP_PASSWORD");
const AZUREAD_APP_REALM = envx("AZUREAD_APP_REALM");

//=========================================================
// Bot Setup
//=========================================================

// Setup Restify Server
var server = restify.createServer();
server.listen(PORT, function () {
  console.log('%s listening to %s', server.name, server.url); 
});
  
// Create chat bot
console.log('started...')
console.log(MICROSOFT_APP_ID);
var connector = new builder.ChatConnector({
  appId: MICROSOFT_APP_ID,
  appPassword: MICROSOFT_APP_PASSWORD
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

var ba = new botauth.BotAuthenticator(server, bot, { session: true, baseUrl: `https://${WEBSITE_HOSTNAME}`, secret : BOTAUTH_SECRET });

ba.provider("aadv2", (options) => {
    // Use the v2 endpoint (applications configured by apps.dev.microsoft.com)
    // For passport-azure-ad v2.0.0, had to set realm = 'common' to ensure authbot works on azure app service
    let oidStrategyv2 = {
      redirectUrl: options.callbackURL, //  redirect: /botauth/aadv2/callback
      realm: AZUREAD_APP_REALM,
      clientID: AZUREAD_APP_ID,
      clientSecret: AZUREAD_APP_PASSWORD,
      identityMetadata: 'https://login.microsoftonline.com/' + AZUREAD_APP_REALM + '/v2.0/.well-known/openid-configuration',
      skipUserProfile: false,
      validateIssuer: false,
      //allowHttpForRedirectUrl: true,
      responseType: 'code',
      responseMode: 'query',
      scope: ['email', 'profile', 'offline_access', 'https://graph.microsoft.com/mail.read'],
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
        session.endDialog(user.displayName);
        console.log(user.accessToken);
        console.log(user.refreshToken);
    }
));