using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StocKings
{
    public class Program
    {
        public static void Main()
        {
            var Parser = new ParseResult();
            Parser.LargeCapParser();

            //var TickerParser = new YahooTickerParser();
            //TickerParser.GetTickers("basf");

        }
    }
}
