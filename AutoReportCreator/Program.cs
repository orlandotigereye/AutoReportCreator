using NDesk.Options;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AutoReportCreator {
    static class Program {
        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
        static void Main(string[] args) {
            var p = new OptionSet() {
                { "StockForecastReport", "", v => CmdOptions.StockForecastReport = v != null },
                { "DailySaleReport", "", v => CmdOptions.DailySaleReport = v != null },
                { "DailyStoreReport", "", v => CmdOptions.DailyStoreReport = v != null },
                { "AutoSendMail", "", v => CmdOptions.AutoSendMail = v != null },
                { "AutoClose", "", v => CmdOptions.AutoClose = v != null },
                { "h|help",  "", v => CmdOptions.ShowHelp = v != null },
            };

            try {
                CmdOptions.Extra = p.Parse(args);
                if (CmdOptions.ShowHelp) {
                    AllocConsole();
                    //ShowCmdHelp();
                    Console.ReadKey();
                    return;
                }
            }
            catch (OptionException e) {
                AllocConsole();
                Console.Write("greet: ");
                Console.WriteLine(e.Message);
                Console.WriteLine("Try `greet --help' for more information.");
                Console.ReadKey();
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FrmMain());
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();
    }
}
