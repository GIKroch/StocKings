using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HtmlAgilityPack;

namespace StocKings
{
    public class YahooTickerParser
    {
        //Yahoo Ticker Parser finds matching tickers to the company names which are obtained through LargeCapParser - without tickers we are not able to query yahoo finance
        public string GetTickers(string companyName) 
        {
            var lookupLink = String.Format("https://finance.yahoo.com/lookup?s={0}", companyName);
            var parser = new TableParser(lookupLink);
            var table = parser.GetTickerTable;

            if (table.Count == 0)
            {
                var ticker = "NULL";
                return ticker;

            }
            else
            {
                var ticker = table[0][0];
                return ticker;

            }
            // The assumption is our ticker will be the first one to appear in the table in the first row. Therefore we only retrieve this element. 


        }


    }
}
