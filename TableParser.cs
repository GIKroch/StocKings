﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using HtmlAgilityPack;


namespace StocKings
{

    //Table Parser is a class used to parse html table for our solution. It contains method for different pages, to capture differences in XPATHs and table structure.
    //
    public class TableParser
    {

        HtmlWeb web = new HtmlWeb();
        private string url;
        private HtmlDocument htmlDoc;
        public TableParser(string urlLink)
        {
            url = urlLink;
            htmlDoc = web.Load(url);
        }

        public List<List<string>> GetTable
        {
            get
            {
                try
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
                catch (Exception e)
                {
                    var table = new List<List<string>>();
                    return table;
                }



            }
        }

        public List<List<string>> GetTickerTable
        {
            get
            {
                try
                {
                    var table = htmlDoc.DocumentNode.SelectSingleNode("//table[@class='lookup-table W(100%) Pos(r) BdB Bdc($seperatorColor) smartphone_Mx(20px)']")
                    .Descendants("tr")
                     .Skip(1)
                    .Where(tr => tr.Elements("td").Count() > 1)
                    .Select(tr => tr.Elements("td")
                    .Select(td => td.InnerText.Trim()).ToList())
                    .ToList();
                    return table;
                }
                catch (Exception e)
                {
                    var table = new List<List<string>>();
                    return table;
                }

            }
        }
    }
}
