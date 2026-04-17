// In order to re-link an account with Plaid, you have to:
//   1. Call getLinkToken()
//   2. Call linkSetup() and provide the Link Token from step 1
//   3. Call getAccessToken() and provide the Public Token from step 2 (will be saved as properties service values)

function printServices()
{
  Logger.log(PropertiesService.getUserProperties().getProperty('plaidAccessToken'));
}

function setPlaidAccessProperties()
{
  // PropertiesService.getUserProperties().setProperty('plaidAccessToken', 'access-development-280dedbc-0c96-4de0-b40e-9a37b969e110');
  // PropertiesService.getUserProperties().setProperty('plaidClientId', '5b12dd7b16769b00124f1353');
  // PropertiesService.getUserProperties().setProperty('plaidSecretId', '14f5b848518e17d1087995c601f287');
}

function getLinkToken() {
  const CLIENT_ID = '5b12dd7b16769b00124f1353';
  const SECRET    = '14f5b848518e17d1087995c601f287';

  // Headers are a parameter plaid requires for the post request
  // Plaid takes a contentType parameter
  // Google app script takes a content-type parameter
  var headers = {                                         
    'contentType' : 'application/json',                                        
    'Content-Type': 'application/json',
  };
  
  // Data is a parameter plaid requires for the post request
  var data = {
    'client_id'     : CLIENT_ID,
    'secret'        : SECRET,
    'user'          : {'client_user_id' : CLIENT_ID},
    'client_name'   : "Plaid App",
    "products"      : ["auth"],
    'country_codes' : ['US'],
    'language'      : 'en'
  };

  // Pass in the necessary headers
  // Pass the payload as a json object
  var parameters = {                                                                                                             
    'headers'           : headers,            
    'payload'           : JSON.stringify(data),                            
    'method'            : 'post',
    'muteHttpExceptions': true,
  };

  // API host + endpoint
  var url      = "https://development.plaid.com/link/token/create";
  var response = UrlFetchApp.fetch(url, parameters);

  // Parse the response into a JSON object
  var json_data = JSON.parse(response);

  const linkToken = json_data.link_token;

  PropertiesService.getDocumentProperties().setProperty('plaidLinkToken', linkToken);

  Logger.log(PropertiesService.getDocumentProperties().getProperty('plaidLinkToken'));
}

function linkSetup() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('PlaidLink').setWidth(400).setHeight(800);
  SpreadsheetApp.getUi().alert(PropertiesService.getDocumentProperties().getProperty('plaidLinkToken'));
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, ' ');
}

function getAccessToken() {
  const CLIENT_ID    = '5b12dd7b16769b00124f1353';
  const SECRET       = '14f5b848518e17d1087995c601f287';
  const PUBLIC_TOKEN = 'public-development-abef714c-b43f-40b2-b7b6-c8f0b8755208';

  // Headers are a parameter plaid requires for the post request
  // Plaid takes a contentType parameter
  // Google app script takes a content-type parameter
  var headers = {                                         
    'contentType' : 'application/json',                                        
    'Content-Type': 'application/json',
  };
  
  // Data is a parameter plaid requires for the post request
  var data = {
    "client_id"    : CLIENT_ID,
    "secret"       : SECRET,
    "public_token" : PUBLIC_TOKEN
  };

  // Pass in the necessary headers
  // Pass the payload as a json object
  var parameters = {                                                                                                             
    'headers'           : headers,            
    'payload'           : JSON.stringify(data),                            
    'method'            : 'post',
    'muteHttpExceptions': true,
  };

  // API host + endpoint
  var url      = "https://development.plaid.com/item/public_token/exchange";
  var response = UrlFetchApp.fetch(url, parameters);

  // Parse the response into a JSON object
  var json_data = JSON.parse(response);

  const accessToken = json_data.access_token;

  PropertiesService.getDocumentProperties().setProperty('plaidAccessToken', accessToken);

  Logger.log(PropertiesService.getDocumentProperties().getProperty('plaidAccessToken'));
}

// function getPlaidAccessToken() {
//   // Set your Plaid credentials
//   const CLIENT_ID = '5b12dd7b16769b00124f1353';
//   const SECRET = '14f5b848518e17d1087995c601f287';
  
//   // Set the Plaid environment (sandbox, development, production)
//   const PLAID_ENV = 'development';
  
//   // Get a Link Token from the Plaid API
//   const linkTokenUrl = `https://${PLAID_ENV}.plaid.com/link/token/create`;
//   const linkTokenOptions = {
//     method: 'POST',
//     headers: {
//       'contentType' : 'application/json',
//       'Content-Type': 'application/json'
//     },
//     payload: JSON.stringify({
//       client_id: CLIENT_ID,
//       secret: SECRET,
//       user: {
//         client_user_id: CLIENT_ID // Unique user ID
//       },
//       client_name: 'Plaid App', // App name
//       products: ['auth'], // Replace with the Plaid products you want to use
//       country_codes: ['US'], // Replace with the country codes you want to support
//       language: 'en' // Replace with the language you want to use
//     }),
//     'muteHttpExceptions': true,
//   };
//   const linkTokenResponse = UrlFetchApp.fetch(linkTokenUrl, linkTokenOptions);
//   const linkToken = JSON.parse(linkTokenResponse.getContentText()).link_token;
  
//   // Use the Link Token to obtain a Public Token
//   const publicTokenUrl = `https://${PLAID_ENV}.plaid.com/link/item/public_token/exchange`;
//   const publicTokenOptions = {
//     method: 'POST',
//     headers: {
//       'Content-Type': 'application/json'
//     },
//     payload: JSON.stringify({
//       client_id: CLIENT_ID,
//       secret: SECRET,
//       link_token: linkToken
//     })
//   };
//   const publicTokenResponse = UrlFetchApp.fetch(publicTokenUrl, publicTokenOptions);
//   const publicToken = JSON.parse(publicTokenResponse.getContentText()).public_token;
  
//   // Exchange the Public Token for an Access Token
//   const accessTokenUrl = `https://${PLAID_ENV}.plaid.com/item/public_token/exchange`;
//   const accessTokenOptions = {
//     method: 'POST',
//     headers: {
//       'Content-Type': 'application/json'
//     },
//     payload: JSON.stringify({
//       client_id: CLIENT_ID,
//       secret: SECRET,
//       public_token: publicToken
//     })
//   };
//   const accessTokenResponse = UrlFetchApp.fetch(accessTokenUrl, accessTokenOptions);
//   const accessToken = JSON.parse(accessTokenResponse.getContentText()).access_token;
  
//   // Save the Access Token as a properties service value
//   PropertiesService.getDocumentProperties().setProperty('plaidAccessToken', accessToken);
// }