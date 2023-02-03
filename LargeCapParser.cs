using HtmlAgilityPack;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;


namespace StocKings
{
   
    public class ParseResult
    {
        public void LargeCapParser(string myDirectory)
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
            ws.Range["A:R"].ColumnWidth = 23;

            // We also specify column names which will remain empty in this step - will be filled with yahoo data 
            ws.Range["E" + rowIndex.ToString()].Value2 = "Price 1 year ago";
            ws.Range["F" + rowIndex.ToString()].Value2 = "Price 3 months ago";
            ws.Range["G" + rowIndex.ToString()].Value2 = "Price 1 month ago";
            ws.Range["H" + rowIndex.ToString()].Value2 = "Price 3 weeks ago";
            ws.Range["I" + rowIndex.ToString()].Value2 = "Price 2 weeks ago";
            ws.Range["J" + rowIndex.ToString()].Value2 = "Price 1 week ago";
            ws.Range["K" + rowIndex.ToString()].Value2 = "Price Today";

            ws.Range["L" + rowIndex.ToString()].Value2 = "1 year ratio";
            ws.Range["M" + rowIndex.ToString()].Value2 = "3 months ratio";
            ws.Range["N" + rowIndex.ToString()].Value2 = "1 month ratio";
            ws.Range["O" + rowIndex.ToString()].Value2 = "3 weeks ratio";
            ws.Range["P" + rowIndex.ToString()].Value2 = "2 weeks ratio";
            ws.Range["Q" + rowIndex.ToString()].Value2 = "1 week ratio";
            ws.Range["R" + rowIndex.ToString()].Value2 = "Forward Annual Dividend Rate";
            ws.Range["S" + rowIndex.ToString()].Value2 = "Forward Annual Dividend Yield";
            ws.Range["T" + rowIndex.ToString()].Value2 = "Dividend Date";
            ws.Range["U" + rowIndex.ToString()].Value2 = "Ex-Dividend Date";

            // Formatting percentage columns
            ws.Range["L:Q"].NumberFormat = "###,##.00%";

            // Bolding column headers
            ws.Range["A1:R1"].Font.Bold = true;
            rowIndex++;


            foreach (var country in countries)
            {
                var countryLink = String.Format("https://companiesmarketcap.com/{0}/largest-companies-in-{1}-by-market-cap/?page=", country[0], country[1]);
                var isNotNull = false;
                var isSmall = false;
                var i = 1;
                
                while (isNotNull == false && isSmall == false)
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
                            
                            // Converting companyMarketCap to numerical value
                            string companyMarketCap = companyInfo[2];
                            float companyMarketCapAdjusted = 0;
                            if (companyMarketCap.Contains("T"))
                            {
                                companyMarketCapAdjusted = float.Parse(companyMarketCap.Replace("T", string.Empty).Replace("$", string.Empty)) * 1000000000000;
                            }
                            else if (companyMarketCap.Contains("B"))
                            {
                                companyMarketCapAdjusted = float.Parse(companyMarketCap.Replace("B", string.Empty).Replace("$", string.Empty)) * 1000000000;
                            }
                            else if (companyMarketCap.Contains("M"))
                            {
                                companyMarketCapAdjusted = float.Parse(companyMarketCap.Replace("B", string.Empty).Replace("$", string.Empty)) * 1000000;
                            }

                            // To reduce the number of parsed entries, we limit ourself to large companies > 1bn 

                            if (companyMarketCapAdjusted < 1000000000)
                            {
                                isSmall = true;
                                break;
                            }

                            ws.Range["A" + rowIndex.ToString()].Value = companyName;
                            ws.Range["B" + rowIndex.ToString()].Value = ticker;
                            ws.Range["C" + rowIndex.ToString()].Value = companyMarketCapAdjusted;
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

            excelWorkbook.SaveAs(myDirectory + @"\LargeCaps.xlsx");
            excelWorkbook.Close();
            excelApp.Quit();

        }


    }
}