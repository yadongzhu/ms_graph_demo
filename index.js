require("dotenv").config();
const MsGraphApi=require('./msGraphApiService');
const MsOAuth=require('./msOauthService');
let apiSvc=new MsGraphApi(new MsOAuth(),{});
(async ()=>{
    try {
        let list = await apiSvc.listOfMailFolders();
        console.log(JSON.stringify(list,null,2));
    } catch (error) {
        console.error('error for apiSvc.listOfMailFolders:',error);
    }
})()