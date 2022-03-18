using HtmlAgilityPack;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading.Tasks;


namespace StockTickerParser
{
    public class Parser
    {
        //List<List<string>> table = new List<string>();
        HtmlWeb web = new HtmlWeb();
        private string url;
        private HtmlDocument htmlDoc;
        public Parser(string urlLink)
        {
            url = urlLink;
            htmlDoc = web.Load(url);
        }

        public List<List<string>> GetTable
        {
            get
            {
                var table = htmlDoc.DocumentNode.SelectSingleNode("//table[@class='table marketcap-table dataTable']")
                    .Descendants("tr")
                     .Skip(1)
                    .Where(tr => tr.Elements("td").Count() > 1)
                    .Select(tr => tr.Elements("td")
                    .Select(td => td.InnerText.Trim()).ToList())
                    .ToList();

                return table;
            }
        }



        public List<List<string>> GetHeaders
        {
            get
            {
                var htmlDoc = web.Load(url);
                //var headers = htmlDoc.DocumentNode.
                //    SelectSingleNode("//table[@class='table marketcap-table dataTable']/thead").
                //    InnerText;

                var headers = htmlDoc.DocumentNode.
                  SelectSingleNode("//table[@class='table marketcap-table dataTable']/thead")
                  .Descendants("tr")
                  .Where(tr => tr.Elements("th").Count() > 1)
                  .Select(tr => tr.Elements("th").Select(th => th.InnerText.Trim()).ToList())
                  .ToList();


                return headers;
            }
        }

    }

    public class InitiateExcel
    {
      
        
        public Excel.Worksheet InitiateWorksheet
        {
            get
            {
                Excel.Application excelApp  = new Excel.Application();
                Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets.Add();
                excelWorksheet.Name = "SEXY";
                excelApp.Visible = true;
                excelWorksheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                return excelWorksheet;

            }
        }

       
    }

    public class ParseResult
    {
        public static void Main()
        {
            
            // Initiating excel
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
            Excel.Worksheet ws = (Excel.Worksheet)excelWorkbook.Sheets.Add();
            ws.Name = "test";
            excelApp.Visible = true;
            
            
            var urlLink = "https://companiesmarketcap.com/germany/largest-companies-in-germany-by-market-cap/?page=1";
            Parser parser = new Parser(urlLink);

            var table = parser.GetTable;

            // First we define flow control variable to write down values row by row
            var rowIndex = 0;
           
            // Now we specify column names
            ws.Range["A" + rowIndex.ToString()].Value = "Company Name";
            ws.Range["B" + rowIndex.ToString()].Value = "Market Cap";

            rowIndex++;
            
            
            

            foreach (var companyInfo in table)
            {

                var companyName = companyInfo[1].Split("\n")[0];
                var companyMarketCap = companyInfo[2];

                ws.Range["A" + rowIndex.ToString()].Value = companyName;
                ws.Range["B" + rowIndex.ToString()].Value = companyMarketCap;

            }

            ws.SaveAs(@"C:\Users\grzeg\Desktop\test.xlsx", Excel.XlFileFormat.xlWorkbookNormal);
            excelWorkbook.Close();
            excelApp.Quit();
        }
    }
}