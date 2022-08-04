'use strict';
const fs = require("fs");
const MsOAuthSvc = require('./msOauthService');
const axios = require("axios");

class MsGraphApiService{
    constructor(msOAuth,options){
        this.msOAuth=msOAuth;
    }
    async listOfMailFolders(){
        let headers=await this.msOAuth.getAuthHeader();
        let options={
            headers: headers,
            method:'GET',
            url:`${process.env.GRAPH_ENDPOINT}/v1.0/users/${process.env.EMAIL_USER}/mailfolders`
         };

        let response;
        try {
            response=await axios(options);
            return response.data.value;
        } catch (error) {
            console.error(error);
            return [];
        }
    }
    async listOfMailsInFolder(folder){
        let headers=await this.msOAuth.getAuthHeader();
        let options={
            headers: headers,
            method:'GET',
            url:`${process.env.GRAPH_ENDPOINT}/v1.0/users/${process.env.EMAIL_USER}/mailfolders/${folder.id}/messages?$top=${folder.totalItemCount}&$count=true`
         };

        let response;
        try {
            response=await axios(options);
            console.log('status is ',response.status);
            let list=response.data.value||[];
            return list;
        } catch (error) {
            console.error(error);
            return [];
        }
    }

    async listOfAttachmentsInMail(mail){
        let headers=await this.emailAuth.getAuthHeader();
        let options={
            headers: headers,
            method:'GET',
            url:`${process.env.GRAPH_ENDPOINT}/v1.0/users/${process.env.EMAIL_USER}/mailfolders/${mail.parentFolderId}/messages/${mail.id}/attachments`
         };

        let response;
        try {
            response=await axios(options);
            console.log('status is ',response.status);
            let list=response.data.value||[];
            return list;
        } catch (error) {
            console.error(error);
            return [];
        }        
    }

    async moveMail(mail,destFolder){
        let headers=await this.emailAuth.getAuthHeader();

        let options={
          headers: headers,
          method:'POST',
          url:`${process.env.GRAPH_ENDPOINT}/v1.0/users/${process.env.EMAIL_USER}/mailFolders/${mail.parentFolderId}/messages/${mail.id}/move`,
          data:{"destinationId": destFolder.id}
       };

      let response;
      try {
          response=await axios(options);
          console.log('status is ',response.status);
        } catch (error) {
            console.error(error);
        }        
    }
}

module.exports=MsGraphApiService