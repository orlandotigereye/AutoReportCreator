using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoReportCreator {
    public static class CmdOptions {
        public static bool StockForecastReport = false;
        public static bool DailySaleReport = false;
        public static bool DailyStoreReport = false;
        public static bool AutoSendMail = false;
        public static bool AutoClose = false;

        public static bool ShowHelp = false;
        public static List<string> Extra = new List<string>();
    }
}
