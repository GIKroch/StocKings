# StocKings - C# Script for Stocks Tracking

## What is it? 
I created StocKings to help me with my personal trading activities. StocKings collects (through web-scraping) financial details (prices, dividends, etc.) of selected companies stocks from Yahoo Finance, introduces simple measures and saves it in a user friendly excel file. With that I can easily narrow down my stock picks - focusing on companies with the highest price changes, good dividend rate etc. As I'm not really a risk taker, I limited the stocks available to only the largest caps in a few mature markets. However, as long as you have a list with proper ticker names, StocKings can be run on it to give you financial details of the stocks of your interest. 

One of the main reasons I started this project, was to play around with C# and the way it interacts with Excel. Usually, to access Excel programatically I used ```openpyxl``` in Python. However, I sometimes felt that it doesn't give me a fully integrated control over Excel - or that there was too much coding for something which appeared to be a simple task. Thus, I chose C#, assuming there can't be a better tool to work with Microsoft Excel than their own programming language. And I was right on that - it really seems that C# API for Excel is well connected with a spreadsheet tool, so that it is feasible to customize your spreadsheet programatically. 


## How does it work? 
1. Extract details of large cap companies from: [https://companiesmarketcap.com/](https://companiesmarketcap.com/)
Currently the task is performed for 4 countries: Germany, France, UK, US. This can be easily adjusted in the code. 

2. For large cap companies collected in the first step, web-scraping of Yahoo Finance delivers: 
    * Historical Prices
    * Dividends info

3. All the requested details are saved to a user friendly excel file - which can be later filtered and analysed to get further insights. 

![Process Map](/Assets/process.drawio.svg)

