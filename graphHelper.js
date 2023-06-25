require('isomorphic-fetch');
const azure = require('@azure/identity');
const graph = require('@microsoft/microsoft-graph-client');
const authProviders =
  require('@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials');

let _settings = undefined;
let _deviceCodeCredential = undefined;
let _userClient = undefined;


async function getUserAsync() {
    // Ensure client isn't undefined
    if (!_userClient) {
        throw new Error('Graph has not been initialized for user auth');
    }
    
    return _userClient.api('/me')
    // Only request specific properties
    .select(['displayName', 'mail', 'userPrincipalName'])
    .get();
}
module.exports.getUserAsync = getUserAsync;

function initializeGraphForUserAuth(settings, deviceCodePrompt) {
  // Ensure settings isn't null
  if (!settings) {
    throw new Error('Settings cannot be undefined');
  }
  
  _settings = settings;
  
  _deviceCodeCredential = new azure.DeviceCodeCredential({
      clientId: settings.clientId,
    tenantId: settings.tenantId,
    userPromptCallback: deviceCodePrompt
});

const authProvider = new authProviders.TokenCredentialAuthenticationProvider(
    _deviceCodeCredential, {
        scopes: settings.graphUserScopes
    });
    
    _userClient = graph.Client.initWithMiddleware({
        authProvider: authProvider
    });
}

async function getUserTokenAsync() {
    // Ensure credential isn't undefined
    if (!_deviceCodeCredential) {
      throw new Error('Graph has not been initialized for user auth');
    }
    
    // Ensure scopes isn't undefined
    if (!_settings?.graphUserScopes) {
        throw new Error('Setting "scopes" cannot be undefined');
    }
  
    // Request token with given scopes
    const response = await _deviceCodeCredential.getToken(_settings?.graphUserScopes);
    return response.token;
}
module.exports.getUserTokenAsync = getUserTokenAsync;

async function sendMailAsync(subject, body, recipient) {
    // Ensure client isn't undefined
    if (!_userClient) {
      throw new Error('Graph has not been initialized for user auth');
    }
  
    // Create a new message
    const message = {
      subject: subject,
      body: {
        content: body,
        contentType: 'text'
      },
      toRecipients: [
        {
          emailAddress: {
            address: recipient
          }
        }
      ]
    };
  
    // Send the message
    return _userClient.api('me/sendMail')
      .post({
        message: message
      });
  }
  module.exports.sendMailAsync = sendMailAsync;

async function getInboxAsync() {
    // Ensure client isn't undefined
    if (!_userClient) {
      throw new Error('Graph has not been initialized for user auth');
    }
  
    return _userClient.api('/me/mailFolders/inbox/messages')
      .select(['id','from', 'isRead', 'receivedDateTime', 'subject'])
      .top(25)
      .orderby('receivedDateTime DESC')
      .get();
  }
module.exports.getInboxAsync = getInboxAsync;
module.exports.initializeGraphForUserAuth = initializeGraphForUserAuth;

// This function serves as a playground for testing Graph snippets
// or other code
async function createNewFolderAsync() {
    // INSERT YOUR CODE HERE
    // for creating a new folder
    // const authProvider = new authProviders.TokenCredentialAuthenticationProvider(
    //     _deviceCodeCredential, {
    //         scopes: settings.graphUserScopes
    //     });
        
    //     _userClient = graph.Client.initWithMiddleware({
    //         authProvider: authProvider
    //     });
    
    // const options = {
    //     authProvider,
    // };
    
    // const client = Client.initWithMiddleware(options);
    
    // const mailFolder = {
    //   displayName: 'Clutter',
    //   isHidden: false
    // };
    
    // await client.api('/me/mailFolders')
    //     .post(mailFolder);

    if (!_userClient){
        throw new Error('Graph has not been initialised for user auth')
    }

    // create new folder
    let folname = 'test'
    const mailFolder = {
        displayName: folname,
        isHidden: false
    }

    return _userClient.api('/me/mailFolders')
        .post(mailFolder)
  }
  module.exports.createNewFolderAsync = createNewFolderAsync;


  async function moveMessageAsync(){
    if(!_userClient){
        throw new Error('Graph has not been initialised for user auth')
    }

    // move a message from inbox folder to test folder
    // update the folder name afterwards when we get to know the categorising

    let folname = 'test'
    let folderId='AAMkADRhMTUwZWE3LWY5MDMtNDkwOC1iYmU2LTI1OGY0NWQ3ODI4OQAuAAAAAABKDb0Jr_LGTaiaquJ3ENDYAQD3DFw5relpQIUkHY160tSmAABm7o2cAAA='
    const message = {
        destinationId: folderId
    }

    // fill up..
    let messageId = 'AAMkADRhMTUwZWE3LWY5MDMtNDkwOC1iYmU2LTI1OGY0NWQ3ODI4OQBGAAAAAABKDb0Jr_LGTaiaquJ3ENDYBwD3DFw5relpQIUkHY160tSmAAAAAAEMAAD3DFw5relpQIUkHY160tSmAABluzolAAA='

    return _userClient.api(`/me/messages/${messageId}/move`)
        .post(message)
  }
  module.exports.moveMessageAsync = moveMessageAsync;