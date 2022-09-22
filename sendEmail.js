require("dotenv").config();
const MsGraphApi=require('./msGraphApiService');
const MsOAuth=require('./msOauthService');
const EmailAuthSvc = require('./emailAuthService');
const fs = require("fs");
const path=require("path");
const axios = require("axios");
(async ()=>{
    let mail = {
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
          }
    }
    let attachmentFiles=[
      {
        file:'bpm_script.sql'
      },
      {
        file: `f11.pdf`
      },
      {
        file: `m66.pdf`
      }
    ]
    let attachments =[];
    attachmentFiles.forEach(item => {
      attachments.push({
        "@odata.type": "#microsoft.graph.fileAttachment",
          "contentBytes": fs.readFileSync(path.join(__dirname,item.file),{encoding:'base64'}),
          "name": item.file,
          "size": fs.statSync(item.file).size
      })
    });
    
    
    let emailAuth=new EmailAuthSvc()
    let headers=await emailAuth.getAuthHeader();
    
    //send draft message
    let options={
      maxContentLength: Infinity,
      maxBodyLength: Infinity,
      headers: headers,
      method:'POST',
      url:`${process.env.GRAPH_ENDPOINT}/v1.0/users/${process.env.EMAIL_USER}/messages`,
      data:mail
      // data:JSON.stringify({ message: mail, saveToSentItems: true })
    };
    let response,messageId;
            try {
              console.log(options.url);
                response=await axios(options);
                console.log(`status is:${response.status}`)
              messageId=response.data.id;
              //loop over attachment
              for (let i = 0; i < attachments.length; i++) {
                const att = attachments[i];
                if (att.size<1024*1024*3){//if size is under 3MB
                  options.url=`${process.env.GRAPH_ENDPOINT}/v1.0/users/${process.env.EMAIL_USER}/messages/${messageId}/attachments`
                  // delete att.size;
                  options.data=att;
                  response=await axios(options);
                  console.log(`attachment upload:${response.status}`)
                } else { //if size is greater than or equal to 3MB
                  //create upload session
                  options.url=`${process.env.GRAPH_ENDPOINT}/v1.0/users/${process.env.EMAIL_USER}/messages/${messageId}/attachments/createUploadSession`
                  options.data={
                    "AttachmentItem": {
                      "attachmentType": "file",
                      "name": att.name,
                      "size": att.size
                    }
                  };
                  response=await axios(options);
                  console.log(`create upload session for file "${att.name}":${response.status}`)
                  let uploadSessionOptions={
                    maxContentLength: Infinity,
                    maxBodyLength: Infinity,
                    method:'PUT',
                    url:response.data.uploadUrl,
                    headers:{
                        'Content-Type': 'application/octet-stream',
                        'Content-Length':  att.size,
                        'Content-Range':`bytes 0-${att.size-1}/${att.size}`
                    },
                    data:att
                  }
                  response=await axios(uploadSessionOptions);
                  console.log(`upload file "${att.name}":${response.status}`)
                }

              }
              options.url=`${process.env.GRAPH_ENDPOINT}/v1.0/users/${process.env.EMAIL_USER}/messages/${messageId}/send`;
              delete options.data;
              response=await axios(options);
              console.log(`send message:${response.status}`)
            } catch (error) {
                console.error(`error is:${error}`)
            }
})()

