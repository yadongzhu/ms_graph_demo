require("dotenv").config();
const { Client } = require("@microsoft/microsoft-graph-client");
require('isomorphic-fetch');
const msal = require('@azure/msal-node');
const fs = require("fs");
const path=require("path");
const msalConfig = {
    auth: {
      clientId:process.env.CLIENT_ID,
      clientSecret:process.env.CLIENT_SECRET,
      authority:`${process.env.AAD_ENDPOINT}/${process.env.TENANT_ID}`
    },
  };
  const tokenRequest = {
    scopes: [`${process.env.GRAPH_ENDPOINT}/.default`],
  };


const client = Client.init({
    // Implement an auth provider that gets a token
    // from the app's MSAL instance
    authProvider: async (done) => {
      try {
        // Get the user's account
        // const account = await msalClient
        //   .getTokenCache()
        //   .getAccountByHomeId(userId);
        const cca = new msal.ConfidentialClientApplication(msalConfig);
        let token = await cca.acquireTokenByClientCredential(tokenRequest);
        if (token) {
          // Attempt to get the token silently
          // This method uses the token cache and
          // refreshes expired tokens as needed
        //   const response = await msalClient.acquireTokenSilent({
        //     scopes: process.env.OAUTH_SCOPES.split(','),
        //     redirectUri: process.env.OAUTH_REDIRECT_URI,
        //     account: account
        //   });

          // First param to callback is the error,
          // Set to null in success case
          done(null, token.accessToken);
        }
      } catch (err) {
        console.log(JSON.stringify(err, Object.getOwnPropertyNames(err)));
        done(err, null);
      }
    }
  });

(async ()=>{
    const mail = {
        subject:'this is demo email for test purpose',
        from: {
            emailAddress: {
              address: process.env.EMAIL_USER,
            },
          },
          toRecipients:[
              {
                emailAddress:{address:'yd.zhu@biosensors.com'}
              }
          ],
          body: {
            content:'this is sample email body',
            contentType: 'html',
          },
          attachments:[
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "contentBytes": fs.readFileSync(path.join(__dirname,`f11.pdf`),{encoding:'base64'}),
                    "name": "f11.pdf"
                },
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "contentBytes": fs.readFileSync(path.join(__dirname,`m66.pdf`),{encoding:'base64'}),
                    "name": 'm66.pdf'
                }
    
          ]
    }    
    // const res = await client.api("/users/").get();
    let response = await client.api(`/users/${process.env.EMAIL_USER}/sendMail`).post({ message: mail,saveToSentItems: true });
    // let response = await client.api("/me/sendMail").post({ message: mail,saveToSentItems: true });
    console.log(response);
})()