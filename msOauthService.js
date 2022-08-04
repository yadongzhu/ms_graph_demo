'use strict';
const msal = require('@azure/msal-node');

class MsOAuthService{
    constructor(){
        this.msalConfig = {
          auth: {
            clientId:process.env.CLIENT_ID,
            clientSecret:process.env.CLIENT_SECRET,
            authority:`${process.env.AAD_ENDPOINT}/${process.env.TENANT_ID}`
          },
        };
    
        this.tokenRequest = {
          scopes: [`${process.env.GRAPH_ENDPOINT}/.default`],
        };
    }

    async getTokenInfo(){
        try {
            if (!this.tokenInfo){
                const cca = new msal.ConfidentialClientApplication(this.msalConfig);
                this.tokenInfo = await cca.acquireTokenByClientCredential(this.tokenRequest);
            }
            return this.tokenInfo;
        } catch (error) {
            console.error(error);
        }
    }

    async getAuthHeader(){
        let tokenInfo=await this.getTokenInfo();
        return {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${tokenInfo.accessToken}`
        }
    }
}

module.exports=MsOAuthService