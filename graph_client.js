require("dotenv").config();
const { ClientSecretCredential } = require("@azure/identity");
const { Client } = require("@microsoft/microsoft-graph-client");
const { TokenCredentialAuthenticationProvider } = require("@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials");

// Create an instance of the TokenCredential class that is imported
const credential = new ClientSecretCredential(process.env.TENANT_ID, process.env.CLIENT_ID, process.env.CLIENT_SECRET);

// Set your scopes and options for TokenCredential.getToken (Check the ` interface GetTokenOptions` in (TokenCredential Implementation)[https://github.com/Azure/azure-sdk-for-js/blob/master/sdk/core/core-auth/src/tokenCredential.ts])

const authProvider = new TokenCredentialAuthenticationProvider(credential, { scopes: [`${process.env.GRAPH_ENDPOINT}/.default`] });

const client = Client.initWithMiddleware({
	debugLogging: true,
	authProvider,
});
(async ()=>{
    const res = await client.api("/users/").get();
    console.log(res);
})()