using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using System.Reflection;
using System.Globalization;


namespace StocKings
{
    //Yahoo finance parser retrieves information about historical prices and calculate desired measures.
    public class YahooFinanceParser
    {
        public string YahooDateMaker(int daysShift)
        {
            var day = DateTime.UtcNow.AddDays(-daysShift);
            var dayMidnight = new DateTime(day.Year, day.Month, day.Day);
            var dayMidnightYahoo = (int)dayMidnight.Subtract(new DateTime(1970, 1, 1)).TotalSeconds;
            var dayString = dayMidnightYahoo.ToString();


            return dayString;
        }

        public float ChangeRatio(float startPrice, float endPrice)
        {
            
            var changeRatio = (endPrice - startPrice) / endPrice;
            return changeRatio;

        }
        public List<List<float>> Parser(string companyTicker)
        {
            // As Yahoo Finance loads its data dynamically we cannot query too long periods 
            // As only part of it will be captured by our HTML respone. 
            // The focus of this script however is only to retrieve prices for specific dates. 
            // Thus, we can create multiple date links to get small html responses and be sure we capture relevant data
            // It is safe to assume that html response will capture 3 months history

            // As of now - the script only focuses on a 1 year period. 
            // We generate today date and a date a year before. 
            // Yahoo Finance URL uses unix timestamp dates, so we have to convert our dates into this format. 
            // We must ensure we operate on midnight timestamps, otherwise the link won't work

            var today = YahooDateMaker(0);
            var quarterAgo = YahooDateMaker(150);
            var yearAgo = YahooDateMaker(365);
            var yearAgoDayBefore = YahooDateMaker(370);

            // Now we can define URL for a particular ticker and a year timespan
            var urlQuarterly = string.Format("https://finance.yahoo.com/quote/{0}/history?period1={1}&period2={2}&interval=1d&filter=history&frequency=1d&includeAdjustedClose=true",
                companyTicker, quarterAgo, today);

            var urlYearly = string.Format("https://finance.yahoo.com/quote/{0}/history?period1={1}&period2={2}&interval=1d&filter=history&frequency=1d&includeAdjustedClose=true",
                companyTicker, yearAgoDayBefore, yearAgo);

            var TableParser = new TableParser(urlQuarterly);
            var TableParser2 = new TableParser(urlYearly);

            var historicalPricesTableQuarterly = TableParser.GetYahooHistoricalPrices;
            var historicalPricesTableYearly = TableParser2.GetYahooHistoricalPrices;

            //Now we extract dates we are interested in (only a few are useful for our analysis, 1year ago, 3 months ago, 1 month ago, 3 weeks ago, 2 weeks ago, 1 week ago, today)
            // As the prices are scraped from HTML they come as strings, we need to format them properly
            
            var price1y = float.Parse(historicalPricesTableYearly[^1][5], CultureInfo.InvariantCulture);
            var price3m = float.Parse(historicalPricesTableQuarterly[90][5], CultureInfo.InvariantCulture);
            var price1m = float.Parse(historicalPricesTableQuarterly[30][5], CultureInfo.InvariantCulture);
            var price3w = float.Parse(historicalPricesTableQuarterly[21][5], CultureInfo.InvariantCulture);
            var price2w = float.Parse(historicalPricesTableQuarterly[14][5], CultureInfo.InvariantCulture);
            var price1w = float.Parse(historicalPricesTableQuarterly[7][5], CultureInfo.InvariantCulture);
            var pricetoday = float.Parse(historicalPricesTableQuarterly[0][5], CultureInfo.InvariantCulture);

            var change1y = ChangeRatio(price1y, pricetoday);
            var change3m = ChangeRatio(price3m, pricetoday);
            var change1m = ChangeRatio(price1m, pricetoday);
            var change3w = ChangeRatio(price3w, pricetoday);
            var change2w = ChangeRatio(price2w, pricetoday);
            var change1w = ChangeRatio(price1w, pricetoday);

            // We define two lists to pass the historical prices and calculated values to the output excel file
            var historicalPrices = new List<float>()
            {
                price1y, price3m, price1m, price3w, price2w, price1m, pricetoday
            };

            var calculatedRatios = new List<float>()
            {
                change1y, change3m, change1m, change3w, change2w, change1w
            };

            var outputList = new List<List<float>>()
            {
                historicalPrices, calculatedRatios
            };

            
            return outputList;
            

        }
    }
}

