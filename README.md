# Google-Sheets-portfolio-builder
The script lets you build a portfolio similar to Google Finance through Google Sheets.

## Enter your transactions
Open this Google Sheet (https://docs.google.com/spreadsheets/d/1TRq_RFyqTCnMUIKgFMsfMh2AAO3aYYructAokw343sQ/edit?usp=sharing) and make a copy (File → Make a copy…) and save it to your Google Drive. There are some sample transactions that I have entered in the Transactions tab of the sheet. Remove them and enter all your stock transactions. I think the column names are self-explanatory.

## Get an AlphaVantage API key
Alpha Vantage provides a free API that gives real time and historical prices of stocks. We will use this service to build the portfolio history starting from the first date you purchased a stock. Get your free API key here: https://www.alphavantage.co/support/#api-key

## Run scripts
You may be familiar with how Excel lets you automate tasks with VBA macros. Google Sheets does with same with JavaScript. You can write scripts to do almost anything you do manually. These are part of a larger framework of script support that Google provides not just for Sheets, but for all GApps that come with the Google App Suite. The reference is fantastic and is very easy to understand and implement.

Go back to your Google Sheet and open the Script Editor (Tools → Script Editor). In the code you see here, enter your AlphaVantage key here:

API_KEY='YOUR_API_KEY_HERE';
You are ready to run the scripts now. Google Apps scripts have a maximum run time of 6 minutes (for the free tier), so I had to break this routine up into separate scripts.

In the toolbar of the Google Apps window, you should see a drop down box that says ‘Select function’. Select the functions listed below one by one and run them. You may be prompted to give permissions for the script to access your Sheets.

createHistorySheet — this function inserts a new tab called History to keeps track of the portfolio value from the date of the first purchase till today.

getAlphaData — this function issues API calls to Alpha Vantage to get the historical prices. This will enable calculation of the portfolio value. This may take some time, depending on how many stocks are in your portfolio.

buildChart — this function will build a chart showing how the portfolio value has progressed from the first date of purchase, similar to the chart found on Google Finance. This will be available in a separate tab called Chart, which will also let you modify the no. of days to plot on the chart.

buildSummary — this function will build a summary of your portfolio in a separate tab called Summary. The table will show a summarized view of each stock in the portfolio, the overall gain (% and $), the daily gain (% and $) and also a 60-day trendline of how each stock has been performing.

Now that the portfolio has been set up, we need to set up triggers to update the real time prices, and to update the history with the daily portfolio value. In the Script Editor, go to Edit → Current project’s triggers. Click on ‘Add a new trigger’ and select ‘dummyUpdate’ function to run on Time-driven event to run every 30 minutes. Add another trigger and select ‘trackPortfolio’ to run every day after the stock market closes.

That’s it! The only manual work is entering the transactions in the Transactions tab. If your stock broker has APIs to feed the transactional data into Google Sheets, you may be able to eliminate this work too.

P.S: owing to the 6 minute execution time limit, the script MAY fail if there are too many stocks in the portfolio. Good luck! :)
