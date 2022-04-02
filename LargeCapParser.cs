using HtmlAgilityPack;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;


namespace StocKings
{
   
    public class ParseResult
    {
        public void LargeCapParser()
        {
            
            // Initiating excel
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
            Excel.Worksheet ws = (Excel.Worksheet)excelWorkbook.Sheets.Add();
            ws.Name = "LargeCaps";
            excelApp.Visible = true;

            // Defining list of countries - necessary to create parsing links properly
            var countries = new List<List<string>> ();
            countries.Add(new List<string> { "germany", "germany" });
            countries.Add(new List<string> { "france", "france" });
            countries.Add(new List<string> { "usa", "the-usa" });
            countries.Add(new List<string> { "united-kingdom", "the-uk" });
            
            // First we define flow control variable to write down values row by row
            var rowIndex = 1;

            // Now we specify column names
            ws.Range["A" + rowIndex.ToString()].Value = "Company Name";
            ws.Range["B" + rowIndex.ToString()].Value = "Ticker";
            ws.Range["C" + rowIndex.ToString()].Value = "Market Cap";
            ws.Range["D" + rowIndex.ToString()].Value = "Country";
            ws.Range["A:D"].ColumnWidth = 23;
            rowIndex++;


            foreach (var country in countries)
            {
                var countryLink = String.Format("https://companiesmarketcap.com/{0}/largest-companies-in-{1}-by-market-cap/?page=", country[0], country[1]);
                var isNotNull = false;
                var i = 1;
                
                while (isNotNull == false)
                {
                    
                    var countryPageLink = countryLink + i.ToString();
                    TableParser parser = new TableParser(countryPageLink);
                    var table = parser.GetTable;
                    Console.WriteLine(countryPageLink);
                    
                    if (table.Count == 0)
                    {
                        isNotNull = true;
                        continue;

                    }

                    else
                    {
                        
                        foreach (var companyInfo in table)
                        {
                            //First we have to extract ticker and company name which are within one string.
                            //Sometimes there is an additional line between strings, so that we need count variable to identify such occurrences
                            var companyNameTickerLength = companyInfo[1].Split("\n").Count();
                            var companyName = companyInfo[1].Split("\n")[0];
                            string ticker;

                            if(companyNameTickerLength == 2)
                            {
                                ticker = companyInfo[1].Split("\n")[1];
                            }
                            else if(companyNameTickerLength == 3)
                            {
                                ticker = companyInfo[1].Split("\n")[2];
                            }
                            else
                            {
                                ticker = "NULL";
                            }
                            
                            var companyMarketCap = companyInfo[2];                            
                            
                            ws.Range["A" + rowIndex.ToString()].Value = companyName;
                            ws.Range["B" + rowIndex.ToString()].Value = ticker;
                            ws.Range["C" + rowIndex.ToString()].Value = companyMarketCap;
                            ws.Range["D" + rowIndex.ToString()].Value = country[0];

                            rowIndex++;
                            
                        }
                        i++;
                        Thread.Sleep(1000);
                    }
                }                           

            }

            // Disabling alerts
            excelApp.DisplayAlerts = false;

            excelWorkbook.SaveAs(@"LargeCaps.xlsx");
            excelWorkbook.Close();
            excelApp.Quit();

        }


    }
}