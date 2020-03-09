function myFunction() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var symbols = sheet.getSheetValues(51,2,1,22);
  Logger.log(symbols);
  var arrObj = {
    finalArr: [],
    priceArr: [],
    divYArr: [],
    divPArr: []
  };
  
  var genericGet = function(symbol, getCB) {
    var url = "https://query.yahooapis.com/v1/public/yql?q=select%20*%20from%20yahoo.finance.quotes%20where%20symbol%20in%20(%22" 
      + symbol + "%22)&format=json&env=store%3A%2F%2Fdatatables.org%2Falltableswithkeys&callback=";

    var request = UrlFetchApp.fetch(url);

    // On complete data fetch
    if (request.getResponseCode() >= 200 && request.getResponseCode() < 400) {
      var stringified = request.getContentText();
      getCB(null, JSON.parse(stringified));
    } else {
      Logger.log('[ERROR] Server returned invalid response: ' + request.getResponseCode());
      getCB('Error: ' + request.getResponseCode(), false);
    }
};

// risk model algorithms and related
var pEG = function(peg) {
	var p = parseFloat(peg);

  if ((peg === 'null') || (peg === null) || (typeof peg === 'undefined') || (p > 20.00) || (p < -20.00)) {
  	return 4;
  } else if (((p > 5.00) && (p <= 20.00)) || ((p < 0) && (p >= -20.00))){
  	return 3;
  } else if ((p > 2.00) && (p <= 5.00)) {
  	return 2;
  } else if ((p >= 1.00) && (p <= 2.00)) {
		return 1;
  } else {
  	return 0;
  }
};
var dividend = function(div) {
	var d = parseFloat(div);

  if ((div === 'null') || (div === null) || (typeof div === 'undefined') || (d < 1.00)) {
  	return 4;
  } else if ((d >= 1.00) && (d <= 1.99)) {
  	return 3;
  } else if ((d >= 2.00) && (d <= 2.99)) {
  	return 2;
  } else if ((d >= 3.00) && (d <= 4.99)) {
		return 1;
  } else {
  	return 0;
  }
};
var shortRatio = function(short) {
	var s = parseFloat(short);

  if ((short === 'null') || (short === null) || (typeof short === 'undefined') || (s > 20.00)) {
  	return 4;
  } else if ((s > 10.00) && (s <= 20.00)) {
  	return 3;
  } else if ((s > 5.00) && (s <= 10.00)) {
  	return 2;
  } else if ((s >= 2.00) && (s <= 5.00)) {
		return 1;
  } else {
  	return 0;
  }
};

var controller = function(symArr, arrObj, contCB) {
  var symbol = symArr.pop();
  if (symbol) {
    genericGet(symbol, function(err, results) {
      if (err) {
        Logger.log('Error: ' + err);
        return; 
      }
      var dataObj = results.query.results.quote;
      var symbolReturned = dataObj.symbol;
      var price = dataObj.LastTradePriceOnly;
      var divY = dataObj.DividendYield/100;
      var divP = price ? (divY*price) : 0; 
      var modelArr = [
        pEG(dataObj.PEGRatio),
        dividend(dataObj.DividendYield),
        shortRatio(dataObj.ShortRatio)
      ];
      Logger.log('model array: ' + symbolReturned + ': ' + modelArr);
      var riskFigure = modelArr.reduce(function(prev, curr) {
      	return prev + curr;
      },0);
 
      arrObj.finalArr.push(riskFigure);
      arrObj.priceArr.push([price]);
      arrObj.divYArr.push([divY]);
      arrObj.divPArr.push([divP]);
      
      
      // add raw data to individual sheet
      var symSheet = ss.getSheetByName(symbol);
      var symValLastRow = symSheet.getLastRow();
      var date = results.query.created;
      var vals = [date];
      
      for (var key in dataObj) {
        vals.push(dataObj[key]);
      }

      var symValLastRow = symSheet.getLastRow();
      var symValRange = symSheet.getRange(symValLastRow+1, 1, 1,vals.length);
      symValRange.setValues([vals]);

      if (symArr.length > 0) {
        // recursively call this function until array is empty
        controller(symArr, arrObj, contCB);
      } else {
        arrObj.finalArr.unshift(date);
        contCB(arrObj);
      }
    });
   } else {
   	contCB(arrObj);
   }
};
  
  controller(symbols[0].reverse(), arrObj, function(arrObj) {
  	var lastRow = sheet.getLastRow();
    var lastRange = sheet.getRange(lastRow+1,1,1,23);
    var priceRange = sheet.getRange(2,8,22,1);
    var divYRange = sheet.getRange(2,17,22,1);
    var divPRange = sheet.getRange(2,18,22,1);
    var armRange = sheet.getRange(2,12,22,1);
    var arms = arrObj.finalArr.map(function(el, index) {
      if (index > 0) {
        return [el];
      }
    });
    // add totals to Time Series sheet
    var tsSheet = ss.getSheetByName("timeseries");
    var tsLastRow = tsSheet.getLastRow();
    var tsLastRange = tsSheet.getRange(tsLastRow+1,1,1,23);
    var shares = sheet.getSheetValues(2, 10, 22, 1);
    var pTotals = arrObj.priceArr.map(function(el, index) {
      return el[0] * shares[index];
    });
    pTotals.unshift(arrObj.finalArr[0]);
    Logger.log("tsSheet: " + tsSheet);
    Logger.log("tsLastRow: " + tsLastRow);
    Logger.log("pTotals: " + pTotals);
    
    lastRange.setValues([arrObj.finalArr]);
    priceRange.setValues(arrObj.priceArr);
    divYRange.setValues(arrObj.divYArr);
    divPRange.setValues(arrObj.divPArr);
    armRange.setValues(arms.slice(1));
    tsLastRange.setValues([pTotals]);
  });
}