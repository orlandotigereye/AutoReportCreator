using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using SnweiCom.Data;
using SnweiCom.IO;
using System.Reflection;
using SnweiCom.Office.Excel;
using InteropExcel = Microsoft.Office.Interop.Excel;
using System.IO;
using SnweiCom.Net;
using System.Net.Mail;

namespace AutoReportCreator
{
    public class DailySaleReport
    {
        CMail mail = null;

        public class Data {
            public string Date { get; set; }
            public decimal Amount { get; set; }
        }

        public DailySaleReport(CMail mail) {
            this.mail = mail;
        }

        public List<Data> GetDataList(string TG001, string TG005, DateTime dtStart, DateTime dtEnd) {
            string strSQL = $@"
USE NEWIPO;
select 
	TG003 [Date], sum(TG045+TG046) [Amount]
from COPTG
where
	TG001 like '{TG001}' and
	TG003 >= '{dtStart.ToString("yyyyMMdd")}' and TG003 <= '{dtEnd.ToString("yyyyMMdd")}' and
	TG005 = '{TG005}' and TG023 = 'Y'
group by
	TG003
order by 
	TG003;
";
            return DacGet.dacErpDB.QueryDataTable(strSQL).ToList<Data>();
        }

        private void FillListToExcel(Cell StartCell, List<Data> list) {
            string name = StartCell.CellName;
            Cell cell = StartCell;
            foreach (var item in list) {
                cell.CellName = name;
                cell.CellPos.Col += Convert.ToInt32(item.Date.Substring(6, 2)) - 1;
                cell.Value = item.Amount;
            }
        }

        public void Create(string[] mail_receiver) {
            EmbeddedResource er = new EmbeddedResource(Assembly.GetExecutingAssembly(), "MailTemplate.每日銷貨統計表.html");
            StreamReader sr = new StreamReader(er.GetStream());
            string mailbody = sr.ReadToEnd();

            er = new EmbeddedResource(Assembly.GetExecutingAssembly(), "ExcelTemplate.每日銷貨統計表.xlsx");
            string FileName = er.CopyToTempFile();
            ExcelAppManager.ExcelItem ExcelItem = Global.MyExcelAppManager.OpenExcel(FileName);
            InteropExcel.Application ExcelApp = ExcelItem.App;
            InteropExcel.Worksheet sheet = ExcelItem.Workbook.Worksheets[1];
            ExcelApp.ScreenUpdating = false;
            ExcelApp.Visible = false;
            ExcelApp.DisplayAlerts = false;
            Cell cell = new Cell(sheet, "A1");

            //==Test reg
            DateTime dtEnd = DateTime.Now.AddDays(-1);
            dtEnd = new DateTime(dtEnd.Year, dtEnd.Month, dtEnd.Day);
            //DateTime dtEnd = new DateTime(2020, 07, 31);
            //==

            DateTime dtStart = new DateTime(dtEnd.Year, dtEnd.Month, 1);

            mailbody = mailbody.Replace("{Year}", dtStart.Year.ToString());
            mailbody = mailbody.Replace("{Month}", dtStart.Month.ToString());
            mailbody = mailbody.Replace("{Month/Day}", dtEnd.ToString("MM/dd"));

            cell.CellName = "C1";
            int currentMonth = dtStart.Month;
            DateTime dt = dtStart;
            while (dt.Month == currentMonth) {
                cell.Value = dt.ToString("MM/dd");
                cell.CellPos.Col++;
                dt = dt.AddDays(1);
            }

            decimal total = 0;
            decimal sum = 0;

            List<Data> list = null;
            //單別 2302--B10000 一芳
            list = GetDataList("%", "B10000", dtStart, dtEnd);
            cell.CellName = "C2";
            FillListToExcel(cell, list);
            sum = list.Sum(x => x.Amount);
            total += sum;
            mailbody = mailbody.Replace("{TW-001}", sum.ToString("###,##0"));

            //單別 2302--G10000 霜江
            list = GetDataList("%", "G10000", dtStart, dtEnd);
            cell.CellName = "C3";
            FillListToExcel(cell, list);
            sum = list.Sum(x => x.Amount);
            total += sum;
            mailbody = mailbody.Replace("{TW-002}", sum.ToString("###,##0"));

            //單別 2302--A10000 喬治
            list = GetDataList("%", "A10000", dtStart, dtEnd);
            cell.CellName = "C4";
            FillListToExcel(cell, list);
            sum = list.Sum(x => x.Amount);
            total += sum;
            mailbody = mailbody.Replace("{TW-003}", sum.ToString("###,##0"));

            //單別 2302--J10000 美濃
            list = GetDataList("%", "J10000", dtStart, dtEnd);
            cell.CellName = "C5";
            FillListToExcel(cell, list);
            sum = list.Sum(x => x.Amount);
            total += sum;
            mailbody = mailbody.Replace("{TW-004}", sum.ToString("###,##0"));

            mailbody = mailbody.Replace("{TW-ALL}", total.ToString("###,##0"));

            total = 0;
            //單別 2302--B30000 大陸一芳
            list = GetDataList("%", "B30000", dtStart, dtEnd);
            cell.CellName = "C8";
            FillListToExcel(cell, list);
            sum = list.Sum(x => x.Amount);
            total += sum;
            mailbody = mailbody.Replace("{OUT-001}", sum.ToString("###,##0"));

            //單別 2302--B20000 海外一芳
            list = GetDataList("%", "B20000", dtStart, dtEnd);
            cell.CellName = "C9";
            FillListToExcel(cell, list);
            sum = list.Sum(x => x.Amount);
            total += sum;
            mailbody = mailbody.Replace("{OUT-002}", sum.ToString("###,##0"));

            mailbody = mailbody.Replace("{OUT-ALL}", total.ToString("###,##0"));

            //Output ChartObject
            string chart_filename = $"{Path.GetTempPath()}\\{DateTime.Now.ToString("yyyyMMdd")}_chart.png";
            InteropExcel.ChartObjects chartObjects = (InteropExcel.ChartObjects)(sheet.ChartObjects(Type.Missing));
            foreach (InteropExcel.ChartObject co in chartObjects) {
                InteropExcel.Chart chart = (InteropExcel.Chart)co.Chart;
                chart.Export(chart_filename, "PNG", false);
                break;
            }

            ExcelItem.Workbook.Save();
            ExcelItem.Workbook.Close(false);
            try {
                Global.MyExcelAppManager.KillNotUseItem();
            }
            catch (Exception) { }

            FileInfo fi = new FileInfo(FileName);
            FileName = Path.GetTempPath() + $"{DateTime.Now.ToString("yyyyMMdd")} 每日銷貨統計表.xlsx";
            fi.CopyTo(FileName, true);

            if (mail_receiver != null) {
                CMailLinkResource linkRes = new CMailLinkResource();
                linkRes.FileName = chart_filename;
                linkRes.ContentType = "image/png";
                linkRes.ContentID = "chart.png";

                var result = mail.SendMail(
                    mail_receiver,
                    $"{DateTime.Now.ToString("yyyyMMdd")} 每日銷貨統計表",
                    mailbody,
                    true,
                    new string[] { FileName },
                    false,
                    new CMailLinkResource[] { linkRes }
                );
            }
            else {
                System.Diagnostics.Process.Start(@"C:\Windows\explorer.exe", FileName);
            }
        }
    }
}
