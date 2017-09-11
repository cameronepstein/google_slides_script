/*
// AdWords Script: Add a Slide with AdWords Data
// --------------------------------------------------------------
// Copyright 2017 Optmyzr Inc., All Rights Reserved
//
// This script takes a Google Presentation as input and appends a slide with basic AdWords metrics.
// Use this to automate creating an appendix of AdWords data to existing PPC report slides.
// The AW data we append is basic but can easily be tweaked to your own needs.
//
// For more PPC management tools and reports, visit www.optmyzr.com
//
*/

// Update this line with the presentation you want to edit.
// E.g. this is for presentation https://docs.google.com/presentation/d/1RxIzTJC6Jwwd3H5aaRjA-zj3d5IhcG9uOTuOfwk8PUg/edit#slide=id.optmyzr_slide_a1f911e6-9538-427d-9e2f-12fdc951f752
var PRESENTATION_ID = "18PrgEeC2B2pudL7PnXQ3Sad9uIccGXwzfv4VPOM_JD4"

function main() {

  var pageId = createSlide(PRESENTATION_ID);

  // Get the page element IDs for a basic TITLE_AND_BODY layout
  var baseElementId = readPageElementIds(PRESENTATION_ID, pageId);
  var titleId = baseElementId + "_0";
  var textId = baseElementId + "_1";

  // Edit the following with the text for the slide's title
  var titleText = "Automatically Fetched AdWords Data";
  updateElement(PRESENTATION_ID, titleId, titleText);

  // The next line gets text for the body section
  var dataForSlide = getLastMonthData();
  updateElement(PRESENTATION_ID, textId, dataForSlide);

  Logger.log("Done updating slides at https://docs.google.com/presentation/d/" + PRESENTATION_ID);

}

function getLastMonthData() {
  var currentAccount = AdWordsApp.currentAccount();
  //Logger.log('Customer ID: ' + currentAccount.getCustomerId() +
  //    ', Currency Code: ' + currentAccount.getCurrencyCode() +
  //    ', Timezone: ' + currentAccount.getTimeZone());
  var stats = currentAccount.getStatsFor('LAST_MONTH');
  var clicks = stats.getClicks();
  var impressions = stats.getImpressions();
  var text = clicks + " clicks from " + impressions + " impressions.";
  return(text);
}

function createSlide(presentationId) {
  // You can specify the ID to use for the slide, as long as it's unique.
  var pageId = Utilities.getUuid();

  var requests = [{
    "createSlide": {
      "objectId": pageId,
      //"insertionIndex": 1,
      "slideLayoutReference": {
        "predefinedLayout": "TITLE_AND_BODY"
      }
    }
  }];
  var slide =
      Slides.Presentations.batchUpdate({'requests': requests}, presentationId);
  //Logger.log(slide);
  //Logger.log("Created Slide with ID: " + slide.replies[0].createSlide.objectId);

  return (pageId);
}

function updateElement(presentationId, elementId, textToAdd) {

  var requests = [{
      "insertText": {
        "objectId": elementId,
        "text": textToAdd,
      }
    }];
  var result =
      Slides.Presentations.batchUpdate({'requests': requests}, presentationId);
  //Logger.log(result);
}

function readPageElementIds(presentationId, pageId) {
  // You can use a field mask to limit the data the API retrieves
  // in a get request, or what fields are updated in an batchUpdate.
  var response = Slides.Presentations.Pages.get(
      presentationId, pageId, {"fields": "pageElements.objectId"});
  //Logger.log(response);
 var objectIds = response.pageElements[0].objectId;
  var parts = objectIds.split("_");
  var objectIdBase = parts[0] + "_" + parts[1];
  //Logger.log("objectIdBase: " + objectIdBase);
  return(objectIdBase);
}

module.exports = {
  main: main,
  updateElement: updateElement,
  createSlide: createSlide,
  getLastMonthData: getLastMonthData,
  readPageElementIds: readPageElementIds
}
