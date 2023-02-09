using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using CsvHelper;
using System.Security.Cryptography.X509Certificates;
using CsvHelper.Configuration;
using System.Globalization;

namespace StocKings
{
    public class LargeCaps
    {
        public string CompanyName { get; set; }
        public string Ticker { get; set; }
        public float MarketCap { get; set; }
        public string Country { get; set; }
    }
    public class Program
    {
        public static void Main()
        {
            var watch = Stopwatch.StartNew();
            // Before running the large cap parser we define our working directory. 
            string myDirectory = new FileInfo(Assembly.GetEntryAssembly().Location).Directory.ToString();
            Console.WriteLine(myDirectory);
            var largeCapsFilePath = myDirectory + @"\LargeCaps.csv";

            //First we check if the files exists and what is its creation date. 
            //There is no need to run Large Cap parser often, as market caps don't change intesively. 
            //Therefore, by assumption if its last change date is less than 3 months ago, this step will be omitted. 

            var quarterAgo = DateTime.Now.AddDays(-90);

            if (File.Exists(largeCapsFilePath))
            {
                var lastOverWriteDate = File.GetLastWriteTime(largeCapsFilePath);
                //Console.WriteLine(lastOverWriteDate);   
                if (lastOverWriteDate <= quarterAgo)
                {
                    Console.WriteLine("Existing Large Caps File is older than a quarter. We must refresh it.");
                    var Parser = new ParseResult();
                    Parser.LargeCapParser(myDirectory);
                }
                else
                {
                    Console.WriteLine("Existing Large Caps File is fresher than a quarter. We don't refresh it.");
                }
            }
            else
            {
                var Parser = new ParseResultCsv();
                Parser.LargeCapParser(myDirectory);
            }

            var configuration = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Encoding = Encoding.UTF8, // Our file uses UTF-8 encoding.
                Delimiter = "," // The delimiter is a comma.
            };

            // Initialize Yahoo Finance Parser which will obtain historical prices and their ratio for tickers
            var Yahoo = new YahooFinanceParser();

            // Intializing output file
            var csvOutput = new StringBuilder();
            var titleLine = string.Format(
                $"{"Company Name"}," +
                $"{"Ticker"}," +
                $"{"Market Cap"}," +
                $"{"Country"}," +
                $"{"Price 1 year ago"}," +
                $"{"Price 3 months ago"}," +
                $"{"Price 1 month ago"}," +
                $"{"Price 3 weeks ago"}," +
                $"{"Price 2 weeks ago"}," +
                $"{"Price 1 week ago"}," +
                $"{"Price Today"}," +
                $"{"1 year ratio"}," +
                $"{"3 months ratio"}," +
                $"{"1 month ratio"}," +
                $"{"3 weeks ratio"}," +
                $"{"2 weeks ratio"}," +
                $"{"1 week ratio"}," +
                $"{"Forward Annual Dividend Rate"}," +
                $"{"Forward Annual Dividend Yield"}," +
                $"{"Dividend Date"}," +
                $"{"Ex-Dividend Date"}"
                );
            csvOutput.AppendLine( titleLine );  

            // Reading CSV file with LargeCap details 
            using (var fs = File.Open(largeCapsFilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                using (var textReader = new StreamReader(fs, Encoding.UTF8))
                using (var csv = new CsvReader(textReader, configuration))
                {
                    var data = csv.GetRecords<LargeCaps>();

                    foreach (var largeCap in data)
                    {                        
                        Console.WriteLine(largeCap.Ticker);

                        // The output of parser is list of list of floats
                        var financialsList = Yahoo.Parser(largeCap.Ticker, largeCap.CompanyName);
                        var historicalPrices = financialsList[0];
                        var calculatedRatios = financialsList[1];
                        var dividendList = financialsList[2];

                        var newLine = string.Format(
                            $"{largeCap.CompanyName}," +
                            $"{largeCap.Ticker}," +
                            $"{largeCap.MarketCap}," +
                            $"{largeCap.Country}," +
                            $"{historicalPrices[0]}," +
                            $"{historicalPrices[1]}," +
                            $"{historicalPrices[2]}," +
                            $"{historicalPrices[3]}," +
                            $"{historicalPrices[4]}," +
                            $"{historicalPrices[5]}," +
                            $"{historicalPrices[6]}," +
                            $"{calculatedRatios[0]}," +
                            $"{calculatedRatios[1]}," +
                            $"{calculatedRatios[2]}," +
                            $"{calculatedRatios[3]}," +
                            $"{calculatedRatios[4]}," +
                            $"{calculatedRatios[5]}," +
                            $"{dividendList[0]}," +
                            $"{dividendList[1]}," +
                            $"{dividendList[2]}," +
                            $"{dividendList[3]}"
                        );

                        csvOutput.AppendLine( newLine );

                    }
                }
            }
            
            File.WriteAllText("LargeCapsWithPrices.csv", csvOutput.ToString() );

            watch.Stop();
            var elapsedTime = watch.Elapsed;

            Console.WriteLine("Execution Time was: ", elapsedTime.ToString());
        }
    }
}
