using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;


namespace StocKings
{
    public class Program
    {
        public static void Main()
        {
            var watch = Stopwatch.StartNew();
            // Before running the large cap parser we define our working directory. 
            string myDirectory = new FileInfo(Assembly.GetEntryAssembly().Location).Directory.ToString();
            Console.WriteLine(myDirectory);
            var largeCapsFilePath = myDirectory + @"\LargeCaps.xlsx";

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
                var Parser = new ParseResult();
                Parser.LargeCapParser(myDirectory);
            }


            //Now we reopen the file created by parser
            var excelApp = new Excel.Application();
            var excelWorkbook = excelApp.Workbooks.Open(largeCapsFilePath);
            excelApp.Visible = true;
            var ws = (Excel.Worksheet)excelWorkbook.Worksheets["LargeCaps"];

            //In the next step we iterrate through tickers, they are saved in column B 
            //For iterration purpose we need rowIndex variable
            var rowIndex = 2;
            var isEmpty = false;

            // We also initialize Yahoo Finance Parser which will obtain historical prices and their ratio for tickers

            var Yahoo = new YahooFinanceParser();


            while (isEmpty == false)
            {

                var ticker = ws.Range["B" + rowIndex.ToString()].Value2;
                var companyName = ws.Range["A" + rowIndex.ToString()].Value2;
                if (ticker == null)
                {
                    Console.WriteLine("All tickers processed");
                    isEmpty = true;
                    break;
                }
                // We must ensure our ticker and company arguments are in String format. 
                ticker = ticker.ToString();
                companyName = companyName.ToString();
                Console.WriteLine(ticker);

                // The output of parser is list of list of floats
                var financialsList = Yahoo.Parser(ticker, companyName);
                var historicalPrices = financialsList[0];
                var calculatedRatios = financialsList[1];
                var dividendList = financialsList[2];

                // Now we save obtained values to excel sheet
                ws.Range["E" + rowIndex.ToString()].Value2 = historicalPrices[0];
                ws.Range["F" + rowIndex.ToString()].Value2 = historicalPrices[1];
                ws.Range["G" + rowIndex.ToString()].Value2 = historicalPrices[2];
                ws.Range["H" + rowIndex.ToString()].Value2 = historicalPrices[3];
                ws.Range["I" + rowIndex.ToString()].Value2 = historicalPrices[4];
                ws.Range["J" + rowIndex.ToString()].Value2 = historicalPrices[5];
                ws.Range["K" + rowIndex.ToString()].Value2 = historicalPrices[6];

                ws.Range["L" + rowIndex.ToString()].Value2 = calculatedRatios[0];
                ws.Range["M" + rowIndex.ToString()].Value2 = calculatedRatios[1];
                ws.Range["N" + rowIndex.ToString()].Value2 = calculatedRatios[2];
                ws.Range["O" + rowIndex.ToString()].Value2 = calculatedRatios[3];
                ws.Range["P" + rowIndex.ToString()].Value2 = calculatedRatios[4];
                ws.Range["Q" + rowIndex.ToString()].Value2 = calculatedRatios[5];

                ws.Range["R" + rowIndex.ToString()].Value2 = dividendList[0];
                ws.Range["S" + rowIndex.ToString()].Value2 = dividendList[1];
                ws.Range["T" + rowIndex.ToString()].Value2 = dividendList[2];
                ws.Range["U" + rowIndex.ToString()].Value2 = dividendList[3];

                rowIndex++;


            }

            // Disabling alerts
            excelApp.DisplayAlerts = false;

            excelWorkbook.SaveAs(myDirectory + @"\LargeCapsPrices.xlsx");
            excelWorkbook.Close();
            excelApp.Quit();

            watch.Stop();
            var elapsedTime = watch.Elapsed;

            Console.WriteLine("Execution Time was: ", elapsedTime.ToString());


        }
    }
}
