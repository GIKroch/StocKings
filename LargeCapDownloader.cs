using CsvHelper;
using CsvHelper.Configuration;
using System.Globalization;
using System.Text;

namespace StocKings
{
    public class LargeCapDownloader
    {
        

        private void SaveFile(string fileUrl, string pathToSave)
        {
            // See https://learn.microsoft.com/en-us/dotnet/api/system.net.http.httpclient
            // for why, in the real world, you want to use a shared instance of HttpClient
            // rather than creating a new one for each request
            var client = new HttpClient();
            var s = client.GetStreamAsync(fileUrl);
            using (FileStream fs = new FileStream(pathToSave, FileMode.OpenOrCreate, FileAccess.Write))
            {
                s.Result.CopyTo(fs);
            }
                //var fileStream = File.Create(pathToSave);
            
        }

        private List<string> GetCSV(string outputDirectory)
        {
            
            // Here we define our list of links to download csv files
            var links = new List<string>();
            links.Add("https://companiesmarketcap.com/germany/largest-companies-in-germany-by-market-cap/?download=csv");
            links.Add("https://companiesmarketcap.com/usa/largest-companies-in-the-usa-by-market-cap/?download=csv");
            links.Add("https://companiesmarketcap.com/united-kingdom/largest-companies-in-the-uk-by-market-cap/?download=csv");
            links.Add("https://companiesmarketcap.com/france/largest-companies-in-france-by-market-cap/?download=csv");

            //
            string fileOutputDirectory;
            var paths = new List<string>();
            var i = 1;
            foreach (var link in links)
            {
                Console.WriteLine(link);
                fileOutputDirectory= Path.Combine(outputDirectory, String.Format("largecaps{0}.csv",i));
                SaveFile(link, fileOutputDirectory);
                paths.Add(fileOutputDirectory);
                i++;
            }

            return paths;
            
        }

        private string nameFixer(string companyName) { 

            return companyName;
        }

        public void Download(string outputDirectory)
        {
            // Creating and adding header of the output file
            var outputCsv = new StringBuilder();
            var configuration = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Encoding = Encoding.UTF8,
                Delimiter = ",",
                MissingFieldFound = null
            };

            var titleLine = string.Format(
                $"{"CompanyName"}|" +
                $"{"Ticker"}|" +
                $"{"MarketCap"}|" +
                $"{"Country"}"
                );
            outputCsv.AppendLine(titleLine);    

            var filePaths = GetCSV(outputDirectory);
            foreach (var file in filePaths)
            {
                using (var reader = new StreamReader(file))
                using (var csv = new CsvReader(reader, configuration))
                {
                    csv.Read();
                    csv.ReadHeader();
                    while (csv.Read())
                    {
                        try
                        {
                          
                            var companyName = csv.GetField("Name");

                            // Sometime there is a data quality issue, where the name contains &amp
                            // and there are more than one columns because of that 
                            // it is so rare that we can just skip those records
                            if (companyName.Contains("&amp")){
                                continue;
                            };

                            var ticker = csv.GetField("Symbol");
                            var marketCap = csv.GetField("marketcap");
                            var country = csv.GetField("country");
                            var newLine = string.Format($"{companyName}|{ticker}|{marketCap}|{country}");
                            outputCsv.AppendLine(newLine);  
                        }
                        catch
                        {
                            continue;
                        }
                        
                    }
                }

            }
            File.WriteAllText("LargeCaps.csv", outputCsv.ToString());
            Console.WriteLine("LargeCaps created");
        }
    }


}
