using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;


namespace StocKings
{
    public class Program
    {
        public static void Main()
        {
            // Before running the large cap parser we define our working directory. 
            string myDirectory = new FileInfo(Assembly.GetEntryAssembly().Location).Directory.ToString();
            Console.WriteLine(myDirectory);

            //var Parser = new ParseResult();
            //Parser.LargeCapParser(myDirectory);
            //Now we reopen the file created by parser

            var largeCapsFilePath = myDirectory + @"\LargeCaps.xlsx";
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


            //while (isEmpty == false)
            while(rowIndex < 10)
            {
                var ticker = ws.Range["B" + rowIndex.ToString()].Value2;
                
                // The output of parser is list of list of floats
                var financialsList  = Yahoo.Parser(ticker);
                var historicalPrices = financialsList[0];
                var calculatedRatios = financialsList[1];

                Console.WriteLine(historicalPrices[0] + " " + calculatedRatios[0]);

                if (ticker == null)
                {
                    Console.WriteLine("All tickers processed");
                    isEmpty = true;
                    break;
                }
                Console.WriteLine(ticker);
                rowIndex++;


            }




            

        }
    }
}
