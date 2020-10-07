const fs = require('fs');
const readline = require('readline');
const {google} = require('googleapis');
const atob = require('atob');
const xlsx = require('xlsx');
const { auth } = require('googleapis/build/src/apis/abusiveexperiencereport');
const users = require('./users');

// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/gmail.readonly'];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = 'token.json';

// Load client secrets from a local file.
fs.readFile('credentials.json', (err, content) => {
  if (err) return console.log('Error loading client secret file:', err);
  // Authorize a client with credentials, then call the Gmail API.
//   authorize(JSON.parse(content), listLabels);
  authorize(JSON.parse(content), getUnreadedMessages);
});

/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials, callback) {
  const {client_secret, client_id, redirect_uris} = credentials.installed;
  const oAuth2Client = new google.auth.OAuth2(
      client_id, client_secret, redirect_uris[0]);

  // Check if we have previously stored a token.
  fs.readFile(TOKEN_PATH, (err, token) => {
    if (err) return getNewToken(oAuth2Client, callback);
    oAuth2Client.setCredentials(JSON.parse(token));
    callback(oAuth2Client);
  });
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback for the authorized client.
 */
function getNewToken(oAuth2Client, callback) {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
  });
  console.log('Authorize this app by visiting this url:', authUrl);
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
  rl.question('Enter the code from that page here: ', (code) => {
    rl.close();
    oAuth2Client.getToken(code, (err, token) => {
      if (err) return console.error('Error retrieving access token', err);
      oAuth2Client.setCredentials(token);
      // Store the token to disk for later program executions
      fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
        if (err) return console.error(err);
        console.log('Token stored to', TOKEN_PATH);
      });
      callback(oAuth2Client);
    });
  });
}

/**
 * Lists the labels in the user's account.
 *
 * @param {google.auth.OAuth2} auth An authorized OAuth2 client.
 */
function listLabels(auth) {
  const gmail = google.gmail({version: 'v1', auth});
  gmail.users.labels.list({
    userId: 'me',
  }, (err, res) => {
    if (err) return console.log('The API returned an error: ' + err);
    const labels = res.data.labels;
    if (labels.length) {
      console.log('Labels:');
      labels.forEach((label) => {
        console.log(`- ${label.name}`);
      });
    } else {
      console.log('No labels found.');
    }
  });
}

function getUnreadedMessages(auth){
  const gmail = google.gmail({version:'v1',auth});
  gmail.users.messages.list({
    userId:'me',
    labelIds:['UNREAD']
  },(err,res)=>{
    if(err) return console.log(`the API is returned an err ${err}`);
    const list = res.data;
    if(list){
      console.log(list);
      list.messages.forEach((message,index)=>{
        getEmail(auth,message.id)
      })
    }else{
      console.log('no messages found');
    }
  })
}

 function getEmail(auth,messageId){
    const gmail = google.gmail({version:'v1',auth});
    gmail.users.messages.get({
        userId:'me',
        id:messageId
    }, (err,res)=>{
        if(err) return console.log(`the api returned an error ${err}`);
        const message = res.data;
        if(message){
          let sender = message.payload.headers.find((header)=>{
             return header.name === 'Return-Path'
          })
          sender = sender.value.replace('<','').replace('>','');
          user = verifyUser(sender);
          if(user){
            var fileName = message.payload.parts[1].filename;
            attachmentId = message.payload.parts[1].body.attachmentId
            let attachment =  getAttachment(auth,messageId,attachmentId,sender);
            console.log(message);
            console.log(sender);
          }
        }else{
            console.log(`no message found`);
        }
    });
}

function getAttachment(auth,messageId,attachmentId,sender){
    const gmail = google.gmail({version:'v1',auth});
    gmail.users.messages.attachments.get({
        userId:'me',
        messageId:messageId,
        id:attachmentId
    },async (err,res)=>{
        if(err) return console.log('the api returned a error ' + err);
        const attachment = res.data;
        if(attachment){
          var file = xlsx.read(attachment.data.replace(/_/g, "/").replace(/-/g, "+"), {type:'base64'})  
          var dir = await fs.promises.readdir('./src/public');
          var currentDir = dir.find((single)=>{
             return single == file.Sheets["הזמנה"].B3.w;
          })
          if(currentDir == undefined){
             let err = await fs.promises.mkdir(`src/public/${file.Sheets["הזמנה"].B3.w}`);
          }
          xlsx.writeFile(file,`./src/public/${file.Sheets["הזמנה"].B3.w}/${sender}.xlsb`);
            console.log(file);
            singAsReaded(auth,messageId);
        }else{
            console.log('attachment not found');
        }
    })
}

function verifyUser(userName){
  user = users.find((singleUser)=>{
    return singleUser.userName === userName
  });
  if(user === undefined){
    return console.log(`${user.userName}is not a registred user`);
  }else if(user.userType !== 'email'){
    return console.log(`the order method of user ${userName} is by ${user.userType}`);
  }else{
    return user;
  }
}

function singAsReaded(auth,messageId){
  const gmail = google.gmail({version:'v1',auth});
  gmail.users.messages.modify({
    userId:'me',
    id:messageId,
    requestBody:{
      removeLabelIds:['UNREAD']
    }
  },(err,res)=>{
    debugger;
  })
}
