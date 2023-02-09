using HtmlAgilityPack;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;
using CsvHelper;
using System.Text;

namespace StocKings
{

    public class ParseResultCsv
    {
        public void LargeCapParser(string myDirectory)
        {

            var csv = new StringBuilder();

            // Defining list of countries - necessary to create parsing links properly
            var countries = new List<List<string>>();
            countries.Add(new List<string> { "germany", "germany" });
            countries.Add(new List<string> { "france", "france" });
            countries.Add(new List<string> { "usa", "the-usa" });
            countries.Add(new List<string> { "united-kingdom", "the-uk" });

            // First we define flow control variable to write down values row by row
            var rowIndex = 1;

            // Now we specify column names
            //var titleLine = string.Format(
            //    $"{"Company Name"}," +
            //    $"{"Ticker"}," +
            //    $"{"Market Cap"}," +
            //    $"{"Country"}," 
            //    +
            //    $"{"Price 1 year ago"}," +
            //    $"{"Price 3 months ago"}," +
            //    $"{"Price 1 month ago"}," +
            //    $"{"Price 3 weeks ago"}," +
            //    $"{"Price 2 weeks ago"}," +
            //    $"{"Price 1 week ago"}," +
            //    $"{"Price Today"}," +
            //    $"{"1 year ratio"}," +
            //    $"{"3 months ratio"}," +
            //    $"{"1 month ratio"}," +
            //    $"{"3 weeks ratio"}," +
            //    $"{"2 weeks ratio"}," +
            //    $"{"1 week ratio"}," +
            //    $"{"Forward Annual Dividend Rate"}," +
            //    $"{"Forward Annual Dividend Yield"}," +
            //    $"{"Dividend Date"}," +
            //    $"{"Ex-Dividend Date"}"
            //    );
            var titleLine = string.Format(
                $"{"CompanyName"}," +
                $"{"Ticker"}," +
                $"{"MarketCap"}," +
                $"{"Country"}"
                );
            csv.AppendLine(titleLine);
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

                            if (companyNameTickerLength == 2)
                            {
                                ticker = companyInfo[1].Split("\n")[1];
                            }
                            else if (companyNameTickerLength == 3)
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
                            var newLine = string.Format($"{companyName},{ticker},{companyMarketCapAdjusted},{country[0]}");
                            csv.AppendLine(newLine);

                        }
                        i++;
                        Thread.Sleep(1000);
                    }
                }

            }

            File.WriteAllText("LargeCaps.csv", csv.ToString());

        }


    }
}