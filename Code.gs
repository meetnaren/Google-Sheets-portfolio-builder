API_KEY='YOUR_API_KEY_HERE';

function readAlpha(ticker){
  var response = UrlFetchApp.fetch("https://www.alphavantage.co/query?function=TIME_SERIES_DAILY&symbol="+ticker+"&apikey="+API_KEY+"&outputsize=full");
  var json = response.getContentText();
  var data = JSON.parse(json);
  return data;
}

function createHistorySheet(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var hist=ss.getSheetByName('History');
  if(hist!= null){
    ss.deleteSheet(hist);
  }
  
  ss.insertSheet('History');
  
  var txns=ss.getSheetByName('Transactions');
  var hist=ss.getSheetByName('History');
  
  hist.getRange(1,1).setValue('Date');
  hist.getRange(2,1).setValue("=MIN(Transactions!C:C)");
  
  minDate=new Date(hist.getRange(2,1).getValue());
  
  var today=new Date();
  
  if(minDate >= today){
    Logger.log('Invalid date in Transactions Sheet')
  }
  
  hist.insertRowsAfter(hist.getMaxRows(), 4000);
  
  hist.getRange(3,1).activate();
  hist.getRange(3,1).setValue("=A2+1");
  
  var dateDiff=dateDiffInDays(minDate, today)-1;
  
  hist.getRange(3,1,dateDiff,1).activate();
  hist.getRange('A3').copyTo(hist.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  
  data=txns.getRange(2,1,txns.getLastRow()-1,1).getValues();
  
  var stocks=new Array();
  
  for(i in data){
    var dupe=false;
    for(j in stocks){
      if(data[i].join()==stocks[j].join()){
        dupe=true;
      }
    }
    if(!dupe){
      stocks.push(data[i]);
    }
  }
  
  Logger.log(stocks.length+" "+hist.getMaxColumns());
  
  if(hist.getMaxColumns()-2<stocks.length){
    hist.insertColumnsAfter(hist.getMaxColumns(), stocks.length-hist.getMaxColumns()+2);
  }
  
  hist.getRange(1,3).setValue("=TRANSPOSE(UNIQUE(Transactions!A2:A))");
  
  hist.getRange(2,3).setValue('=SUMIFS(Transactions!$D:$D,Transactions!$A:$A,C$1,Transactions!$C:$C,"<="&$A2)')
  hist.getRange(2,3,hist.getLastRow()-1,stocks.length).activate();
  hist.getRange(2,3).copyTo(hist.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  
  
/*  
  var d=minDate;
  
  while(d<today){
    d.setDate(d.getDate()+1);
    hist.getRange(hist.getLastRow()+1,1).setValue(d);
  }
*/  
}

function getAlphaData(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var txns=ss.getSheetByName('Transactions');
  var hist=ss.getSheetByName('History');
  
  stockPrices=new Object();
  
  tickers=hist.getRange(1,3,1,hist.getMaxColumns()-2).getValues();
  
  for(i in tickers[0]){
    ticker=tickers[0][i]
    stockPrices[ticker]=readAlpha(ticker)
    //Logger.log(ticker+" prices retrieved.");
  }
  Logger.log(stockPrices['AAPL']);
  //var cache = CacheService.getScriptCache();
  //cache.put("Stock Prices", stockPrices, 1500);
  
  var dates=hist.getRange(2,1,hist.getLastRow()-1,1).getValues();
  var shares=hist.getRange(1,3,hist.getLastRow(),hist.getLastColumn()-2).getValues();
  
  values = new Array();
  LR=hist.getLastRow();
  for(i=2;i<=LR;i++){
    try{
      d=new Date(dates[i-2]).toISOString().slice(0,10);
      var value=0;
      for(j in shares[0]){
        ticker=shares[0][j];
        //Logger.log(ticker);
        //logs.getRange(1,1).setValue(stockPrices[ticker]['Time Series (Daily)'][0]);
        try{
          closePrice=parseFloat(stockPrices[ticker]['Time Series (Daily)'][d]['4. close']);
        }
        catch(error){
          closePrice=0;
        }
        value+=(closePrice) * shares[i-1][j];
      }
      values.push([value]);
    }
    catch(error){
      Logger.log(error+" on row number "+i);
    }
  }
  
  for(i in values){
    if(i>0){
      if(values[i][0]==0){
        values[i][0]=values[i-1][0]
      }
      
    }
    
  }
  hist.getRange(2,2,hist.getLastRow()-1,1).setValues(values);
  
  /*
  for(i=1;i++;i<=hist.getMaxColumns()-2){
    ticker=hist.getRange(1,i+2).getValue();
    stockPrices[ticker]=readAlpha(ticker);
    Logger.log(ticker+" prices retrieved.");
  }
  */
  
}

function buildChart(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var hist=ss.getSheetByName('History');
  var chart=ss.getSheetByName('Chart');
  if(chart!= null){
    ss.deleteSheet(chart);
  }
  
  ss.insertSheet('Chart');
  var chart=ss.getSheetByName('Chart');

  chart.getRange(1, 1).setValue('Date');
  chart.getRange(1, 2).setValue('Portfolio value');
  chart.getRange(1, 3).setValue('Days to plot');
  chart.getRange(1, 4).setValue(10);
  chart.getRange('B:B').setNumberFormat('"$"#,##0.00');
  chart.getRange(2, 1).setValue('=ARRAYFORMULA(INDIRECT("History!A"&(rows(filter(History!A:A,not(isblank(History!A:A))))-D1)&":B"&rows(filter(History!A:A,not(isblank(History!A:A))))))');
  
  var lineChart = chart.newChart()
      .asLineChart()
      .addRange(chart.getRange("A:B"))
      .setPosition(3, 4, 0, 0)
      .setOption('legend.position', 'none')
      .setOption('animation.duration', 1000)
      .setYAxisTitle('Portfolio value')
      .build();
  
  chart.insertChart(lineChart);
  chart.hideColumns(1, 2);
  chart.setHiddenGridlines(true);
  hist.hideSheet();
}

function buildSummary(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var summary=ss.getSheetByName('Summary');
  if(summary!= null){
    ss.deleteSheet(summary);
  }
  
  ss.insertSheet('Summary');
  var summary=ss.getSheetByName('Summary');
  
  summary.getRange(2,1).setValue("=UNIQUE(Transactions!A:A)");
  
  var headings = {2:"Shares", 
                  3:"Average cost", 
                  4:"Current price", 
                  5:"Total investment", 
                  6:"Market value", 
                  7:"Gain", 
                  8:"Gain %", 
                  9:"Last price", 
                  10:"Today's gain ($)", 
                  11:"Today's gain (%)", 
                  12:'=L1&"-day trend"'};
  
  for(i in headings){
    summary.getRange(2,i).setValue(headings[i]);
  }
  
  var firstRow = {2:"=sumif(Transactions!A:A,A3,Transactions!D:D)",
                  3:"=E3/B3",
                  4:"=alphaVantage(A3,today(),$A$1)",
                  5:"=sumif(Transactions!A:A,A3,Transactions!G:G)",
                  6:"=D3*B3",
                  7:"=F3-E3",
                  8:"=G3/E3",
                  9:'=alphaVantage(A3,today()-1,$A$1)',
                  10:"=(D3-I3)*B3",
                  11:"=(D3-I3)/I3",
                  12:'=SPARKLINE(GOOGLEFINANCE(A3,"price",today()-$L$1,today()))'};
  
  for(i in firstRow){
    summary.getRange(3,i).setValue(firstRow[i]);
  }
  
  summary.getRange(3,2,summary.getLastRow()-3,11).activate();
  summary.getRange('B3:L3').copyTo(summary.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  
  summary.getRange(1,5).setValue("=SUBTOTAL(9,E3:E)");
  summary.getRange(1,6).setValue("=SUBTOTAL(9,F3:F)");
  summary.getRange(1,7).setValue("=F1-E1");
  summary.getRange(1,8).setValue("=G1/E1");
  
  summary.getRange(1,10).setValue("=SUBTOTAL(9,J3:J)");
  summary.getRange(1,11).setValue("=J1/(F1-J1)");
  summary.getRange(1,12).setValue(60);
  
  summary.getRangeList(['C:G', 'I:J']).setNumberFormat('"$"#,##0.00');
  summary.getRangeList(['H:H', 'K:K']).setNumberFormat('0.00%');
  
  summary.setHiddenGridlines(true);
  
  summary.getRange('2:2').setFontWeight('bold');
  //summaryRange=summary.getRange(3,2,summary.getLastRow()-3,11);
  summary.getRangeList(['A2:L2', 'E1:H1', 'J1:L1']).setBorder(true, true, true, true, true, true);
  
  var rangeOne = summary.getRange('J3:J');
  var rangeTwo = summary.getRange('K3:K');
  
  var ranges={
    'rangeOne':summary.getRange('J3:J'),
    'rangeTwo':summary.getRange('K3:K')
  }
  
  for(r in ranges){  
    var rule = SpreadsheetApp.newConditionalFormatRule()
    .setRanges([ranges[r]])
    .setGradientMinpoint('#E67C73')
    .setGradientMidpointWithValue('#FFFFFF', SpreadsheetApp.InterpolationType.NUMBER, '0')
    .setGradientMaxpoint('#57BB8A')
    .build();
  
    var rules = summary.getConditionalFormatRules();
    rules.push(rule);
    summary.setConditionalFormatRules(rules);
  }
  
  summary.getRange('A2:L2').createFilter();
  summary.autoResizeColumns(1, 12);
  summary.hideColumns(9,1);
  
  summary.getRange('A2:L'+(summary.getLastRow()-1)).applyRowBanding(SpreadsheetApp.BandingTheme.CYAN);
}

function dateDiffInDays(a, b) {
  // Discard the time and time-zone information.
  var _MS_PER_DAY = 1000 * 60 * 60 * 24;
  var utc1 = Date.UTC(a.getFullYear(), a.getMonth(), a.getDate());
  var utc2 = Date.UTC(b.getFullYear(), b.getMonth(), b.getDate());

  return Math.floor((utc2 - utc1) / _MS_PER_DAY);
}

function alphaVantage(ticker, date, dummyDate){
  inDate = new Date(date);
  var today=new Date();
  if(inDate>today){
    return "Invalid date";
  }
  var response = UrlFetchApp.fetch("https://www.alphavantage.co/query?function=TIME_SERIES_DAILY&symbol="+ticker+"&apikey="+API_KEY+"&outputsize=full");
  var json = response.getContentText();
  var data = JSON.parse(json);
  
  dateString=inDate.toISOString().slice(0,10);
  Logger.log(dateString);
  while(!(dateString in data['Time Series (Daily)'])){
    inDate.setDate(inDate.getDate()+1);
    dateString=inDate.toISOString().slice(0,10);
  }
  return parseFloat(data['Time Series (Daily)'][dateString]['4. close']);
  
}

function dummyUpdate(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var summary=ss.getSheetByName('Summary');
  
  summary.getRange(1,1).setValue(new Date());
  summary.getRange(1,1).setNumberFormat(';;;');
  //summary.getRange(1,1).setValue(new Date());
}

function trackPortfolio(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var summary=ss.getSheetByName('Summary');
  var history=ss.getSheetByName('History');
  
  today_value=summary.getRange('F1').getValue();
  
  history.getRange(history.getLastRow()+1,1).setValue(new Date());
  history.getRange(history.getLastRow()+1,1).setValue(today_value);
}
