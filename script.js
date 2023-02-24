// Get spreadsheet
var sheet = SpreadsheetApp.openById('19dVknnxzC4yvB2bxuyQUdfTlubtvlbZldTwHXhM9nX4');
var sheetName = 'Training Data'

var links = ['DOC LINK(s)'];

function listComments() {
  // clearRange()
  // addHeadersToSheet()
  for (var i = 0; i < links.length; i++) {
    try {
      var doc = DocumentApp.openByUrl(links[i]);
      var docId = doc.getId();
      var comments = Drive.Comments.list(docId);
      if (comments.items && comments.items.length > 0) {
        for (var z = 0; z < comments.items.length; z++) {
          var comment = comments.items[z];
          if (comment.context) {
            if (comment.author.displayName.includes('Robin') || comment.author.displayName.includes('Ben') || comment.author.displayName.includes('Abby')) {
              if (comment.content.length > 15) {
                if (comment.context.value.length > 10) {
                  appendRowToSheet(doc.getName(), comment.author.displayName, comment.context.value, comment.content, doc.getUrl())
                }
              }
            }
          }
        }
      } else {
        Logger.log('No comment found.');
      }
    }
    catch (e) { break }
  }
}

function addHeadersToSheet() {
  sheet.getSheetByName(sheetName).appendRow(["Doc Name", "Author", "Commmented On", "Comment", "Properly formatted comment for training", "Doc Link"]);
}

function appendRowToSheet(docName, commentAuthor, commentContext, commentContent, docURL) {
  sheet.getSheetByName(sheetName).appendRow([docName, commentAuthor, commentContext, commentContent, "", docURL]);
}

function clearRange() {
  sheet.getSheetByName(sheetName).getRange("A1:F").clearContent();
}

function makeOpenAIRequest(comment) {
  // Set the API endpoint URL
  var apiUrl = 'https://api.openai.com/v1/completions';

  // Set the API key
  var apiKey = '';

  // Set the request headers
  var headers = {
    'Content-Type': 'application/json',
    Authorization: 'Bearer ' + apiKey,
  };

  // Set the request body
  var requestBody = {
    model: 'text-davinci-003',
    prompt:
      `I want you to help me analyze a comment from a Google doc. I want you to help me classify what action the comment was requesting. Just for context, I am doing a comment audit for Aspire Software, a landscaping business management software by looping through a bunch of Google Doc Comments. I want to classify the action the comment was telling us to do in less than 7 word phrase. For example if a comment says "I don't think I would say Aspire "automates home business operations" for in this case marketing. We don't really have that capability." The output should be something similar to: "Suggest a different wording for the marketing statement" Here's the actual comment I need you to classify: ${comment}`,
    max_tokens: 256
  }

  // Encode the request body as a JSON string
  var payload = JSON.stringify(requestBody);

  // Set the options for the HTTP request
  var options = {
    method: 'POST',
    headers: headers,
    payload: payload,
  };

  var numAttempts = 0;
  var maxAttempts = 3;
  var success = false;

  while (numAttempts < maxAttempts && !success) {
    try {
      Logger.log(`API Request Attempt #: ${numAttempts + 1}`)
      var response = UrlFetchApp.fetch(apiUrl, options);
      success = true;
    } catch (error) {
      numAttempts++;
      Utilities.sleep(3000)
      if (numAttempts == maxAttempts) {
        throw error;
      }
    }
  }

  // Parse the response
  var responseJson = response.getContentText();
  var responseObj = JSON.parse(responseJson);

  // Get the completion text from the response
  var completionText = responseObj.choices[0].text;
  completionText = completionText.replace(/\n/g, " ");

  Utilities.sleep(4000)

  Logger.log(completionText)

  return completionText
}
