/*
MIT License

Copyright (c) 2019 Takashi Iwamoto

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/

var ss = SpreadsheetApp.getActive();
var properties = PropertiesService.getUserProperties();

// Run tests

function run() {
  var activeSheet = ss.getActiveSheet();
  if (activeSheet.getSheetName() == 'Config') {
    throw new Error('Can not run tests on Config sheet.');
  }

  resetTestDetailsSheet(activeSheet);
  var params = makeTestParams();
  var urls = getTestUrls(activeSheet);
  runTests(activeSheet, params, urls);
}

function resetTestDetailsSheet(sheet) {
  sheet.getRange(2, 2, sheet.getLastRow() - 1, sheet.getLastColumn() - 1).clearContent();
}

function makeTestParams() {
  var configSheet = ss.getSheetByName('Config');
  var configValues = getValuesOfColumn(configSheet, 2);
  return {
    f: 'json',
    fvonly: '1',
    runs: '1',
    video: '1',
    k: configValues[0],
    location: configValues[1] + '.' + configValues[2],
    mobile: configValues[3]
  };
}

function getTestUrls(sheet) {
  return getValuesOfColumn(sheet, 1, 2);
}

function runTests(sheet, params, urls) {
  var requests = makeRunTestRequests(params, urls);
  var responses = UrlFetchApp.fetchAll(requests);
  var jsonResultUrls = collectJsonResultUrls(responses);
  var testDetails = makeTestDetails(jsonResultUrls);
  setProperty(sheet.getSheetId(), testDetails);
  updateSheets();
}

function makeRunTestRequests(params, urls) {
  return urls.map(function(url) {
    var payload = JSON.parse(JSON.stringify(params));
    payload.url = url;
    return {
      url: 'http://www.webpagetest.org/runtest.php',
      method: 'post',
      payload: payload
    };
  });
}

function collectJsonResultUrls(responses) {
  return responses.map(function(response) {
    var runTestResponse = JSON.parse(response);
    if (runTestResponse.statusCode !== 200) {
      throw new Error(runTestResponse.statusText);
    }
    return runTestResponse.data.jsonUrl + '&requests=0&pagespeed=0&domains=0&breakdown=0';
  });
}

function makeTestDetails(jsonResultUrls) {
  return jsonResultUrls.map(function(jsonResultUrl, i) {
    return {
      index: i,
      jsonResultUrl: jsonResultUrl,
      running: true
    };
  });
}

function updateSheets() {
  deleteTriggers();

  var sheets = ss.getSheets();
  var runningSheets = sheets.filter(function(sheet) {
    return updateSheet(sheet);
  });
  if (runningSheets.length > 0) {
    createTrigger();
    ss.toast('Please wait for a while until the test is completed.', 'Genjo v1.0', 10);
  }
}

// Update sheets

function updateSheet(sheet) {
  sheetId = sheet.getSheetId();

  var testDetails = getProperty(sheetId);
  if (!testDetails) {
    return false;
  }

  var requests = makeJsonResultRequests(testDetails);
  var responses = UrlFetchApp.fetchAll(requests);

  updateTestDetails(sheet, testDetails, responses);

  if (collectRunningTestDetails(testDetails).length > 0) {
    setProperty(sheetId, testDetails);
    return true;
  } else {
    setProperty(sheetId, null);
    return false;
  }
}

function makeJsonResultRequests(testDetails) {
  var runningTestDetails = collectRunningTestDetails(testDetails);
  return runningTestDetails.map(function(testDetail) {
    return {
      url: testDetail.jsonResultUrl,
      method: 'get'
    };
  });
}

function updateTestDetails(sheet, testDetails, responses) {
  var runningTestDetails = collectRunningTestDetails(testDetails);
  var runningTestDetailsIndice = runningTestDetails.map(function(testDetail) {
    return testDetail.index;
  });

  responses.forEach(function(response, i) {
    var testDetailIndex = runningTestDetailsIndice[i];
    var testDetail = testDetails[testDetailIndex];
    var testResult = JSON.parse(response);
    updateTestDetail(testDetail, testResult);
    showTestDetail(sheet, testDetailIndex, testResult);
  });
}

function updateTestDetail(testDetail, testResult) {
  var statusCode = testResult.statusCode;
  if (statusCode === 200 || statusCode >= 400) {
    testDetail.running = false;
  }
}

function showTestDetail(sheet, testDetailIndex, testResult) {
  var values = [testResult.statusText];
  if (testResult.statusCode === 200) {
    var metrics = testResult.data.median.firstView;
    values = values.concat([
      testResult.data.summary,
      metrics.loadTime,
      metrics.TTFB,
      metrics.render,
      metrics.SpeedIndex,
      metrics.FirstInteractive || metrics.LastInteractive || '-',
      metrics.docTime,
      metrics.requestsDoc,
      Math.round(metrics.bytesInDoc / 1024),
      metrics.fullyLoaded,
      metrics.requestsFull,
      Math.round(metrics.bytesIn / 1024)
    ]);
  }
  sheet.getRange(testDetailIndex + 2, 2, 1, values.length).setValues([values]);
}

function collectRunningTestDetails(testDetails) {
  return testDetails.filter(function(testDetail) {
    return testDetail.running;
  });
}

// Properties

function setProperty(sheetId, testDetails) {
  properties.setProperty(sheetId, JSON.stringify(testDetails));
}

function getProperty(sheetId) {
  var testDetailsString = properties.getProperty(sheetId);
  return testDetailsString ? JSON.parse(testDetailsString) : null;
}

// Trigger

function createTrigger() {
  ScriptApp.newTrigger('updateSheets').timeBased().after(10000).create();
}

function deleteTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    ScriptApp.deleteTrigger(trigger);
  });
}

// Utility

function getValuesOfColumn(sheet, column, startRow) {
  startRow = startRow || 1;
  return sheet.getRange(startRow, column, sheet.getLastRow() - startRow + 1).getDisplayValues().map(function(r) {
    return r[0];
  });
}
