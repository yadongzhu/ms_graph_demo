require("dotenv").config();
const MsGraphApi=require('./msGraphApiService');
const MsOAuth=require('./msOauthService');
let apiSvc=new MsGraphApi(new MsOAuth(),{});
const EmailAuthSvc = require('./emailAuthService');
const fs = require("fs");
const path=require("path");
const axios = require("axios");
(async ()=>{
    // try {
    //     let list = await apiSvc.listOfMailFolders();
    //     console.log(JSON.stringify(list,null,2));
    // } catch (error) {
    //     console.error('error for apiSvc.listOfMailFolders:',error);
    // }
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
    let emailAuth=new EmailAuthSvc()
    let headers=await emailAuth.getAuthHeader();
    
    let options={
      maxContentLength: Infinity,
      maxBodyLength: Infinity,
      headers: headers,
      method:'POST',
      url:`${process.env.GRAPH_ENDPOINT}/v1.0/users/${process.env.EMAIL_USER}/sendMail`,
      data:JSON.stringify({ message: mail, saveToSentItems: true })
    };
    let response;
            try {
                response=await axios(options);
                console.log(`status is:${response.status}`)
    
            } catch (error) {
                console.error(`error is:${error}`)
            }
})()

