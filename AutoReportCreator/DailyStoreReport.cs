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
using System.Data.Common;

namespace AutoReportCreator
{
    public class DailyStoreReport
    {
        CMail mail = null;

        public enum Brand {
            YiFang,
            ShuangJiang,
            Mino
        }

        public class Data {
            public string StoreNo { get; set; }
            public string StoreName { get; set; }
            public int StoreType { get; set; }
            public DateTime RetailDate { get; set; }
            public decimal UberEats { get; set; }
            public int UberEatsCount { get; set; }
            public decimal FoodPanda { get; set; }
            public int FoodPandaCount { get; set; }
            public decimal Delivery { get; set; }
            public int DeliveryCount { get; set; }
            public decimal TakeOut { get; set; }
            public int TakeOutCount { get; set; }
            public decimal DineIn { get; set; }
            public int DineInCount { get; set; }
            public decimal TakeSelf { get; set; }
            public int TakeSelfCount { get; set; }
            public decimal TotalAmount { get; set; }
        }

        public DailyStoreReport(CMail mail) {
            this.mail = mail;
        }

        public List<Data> GetDataList(string StoreBrand, string StoreCom, string StoreType, DateTime DateStart, DateTime DateEnd) {
            string strSQL = $@"
declare @StoreTypeList table (StoreType int)
insert into @StoreTypeList values @pStoreType

declare @StoreBrand varchar(10) = '@pStoreBrand'
declare @StoreCom varchar(10) = '@pStoreCom'
declare @DateStart date = '@pDateStart'
declare @DateEnd date = '@pDateEnd'

declare @StoreData TABLE(
	StoreNo varchar(30) NOT NULL,
	StoreName nvarchar(100) NOT NULL,
	StoreType varchar(30) NOT NULL
)

insert into @StoreData
select StoreNo, StoreName, StoreType from StoreData 
where StoreBrand=@StoreBrand and StoreCom=@StoreCom and StoreName not like '%(關)%' and StoreType in (select * from @StoreTypeList)
order by StoreType, StoreNo

select
	SD.StoreNo,
	SD.StoreName,
	SD.StoreType,
	Retail.RetailDate,
	sum(UberEats) UberEats,
	count(case when UberEats > 0 then 1 end) UberEatsCount,
	sum(FoodPanda) FoodPanda,
	count(case when FoodPanda > 0 then 1 end) FoodPandaCount,
	sum(case when OrderType='外送' and UberEats+FoodPanda<=0 then TotalPrice else 0 end) Delivery,
	sum(case when OrderType='外送' and UberEats+FoodPanda<=0 then 1 else 0 end) DeliveryCount,
	sum(case when OrderType='外帶' and UberEats+FoodPanda<=0 then TotalPrice else 0 end) TakeOut,
	sum(case when OrderType='外帶' and UberEats+FoodPanda<=0 then 1 else 0 end) TakeOutCount,
	sum(case when OrderType='內用' and UberEats+FoodPanda<=0 then TotalPrice else 0 end) DineIn,
	sum(case when OrderType='內用' and UberEats+FoodPanda<=0 then 1 else 0 end) DineInCount,
	sum(case when OrderType='自取' and UberEats+FoodPanda<=0 then TotalPrice else 0 end) TakeSelf,
	sum(case when OrderType='自取' and UberEats+FoodPanda<=0 then 1 else 0 end) TakeSelfCount,
    sum(TotalPrice) TotalAmount
from
	@StoreData SD
left join Retail on Retail.StoreNo=SD.StoreNo and Retail.RetailDate between @DateStart and @DateEnd and IsCancel=0
group by
	SD.StoreNo,
	SD.StoreName,
	SD.StoreType,
	Retail.RetailDate
order by
	SD.StoreType, SD.StoreNo, Retail.RetailDate

";
            strSQL = strSQL.Replace("@pStoreBrand", StoreBrand);
            strSQL = strSQL.Replace("@pStoreCom", StoreCom);
            strSQL = strSQL.Replace("@pStoreType", StoreType);
            strSQL = strSQL.Replace("@pDateStart", DateStart.ToString("yyyy/MM/dd"));
            strSQL = strSQL.Replace("@pDateEnd", DateEnd.ToString("yyyy/MM/dd"));
            return DacGet.dacInkismMIS.QueryDataTable(strSQL).ToList<Data>();
        }

        public void CreateAndSend(string[] mail_receiver, Brand brand, string StoreType) {
            string StoreBrand = "";
            DateTime DateEnd = DateTime.Now.AddDays(-1);
            DateTime DateStart = new DateTime(DateEnd.Year, DateEnd.Month, 1);
            int daysInMonth = DateTime.DaysInMonth(DateStart.Year, DateStart.Month);
            string BrandName = "";

            if (brand == Brand.YiFang) {
                StoreBrand = "0001";
                BrandName = "一芳";
            }
            else if (brand == Brand.ShuangJiang) {
                StoreBrand = "0012";
                BrandName = "霜江";
            }
            else if (brand == Brand.Mino) {
                StoreBrand = "0015";
                BrandName = "美濃";
            }

            var list = GetDataList(StoreBrand, $"{StoreBrand}001", StoreType, DateStart, DateEnd);

            EmbeddedResource er;

            er = new EmbeddedResource(Assembly.GetExecutingAssembly(), "ExcelTemplate.每日業績表.xlsx");
            string FileName = er.CopyToTempFile();
            ExcelAppManager.ExcelItem ExcelItem = Global.MyExcelAppManager.OpenExcel(FileName);
            InteropExcel.Application ExcelApp = ExcelItem.App;
            InteropExcel.Workbook workbook = ExcelItem.Workbook;
            InteropExcel.Worksheet sheet_summary = workbook.Worksheets["彙總"];
            InteropExcel.Worksheet sheet_store = workbook.Worksheets["單店"];
            ExcelApp.ScreenUpdating = false;
            ExcelApp.Visible = false;
            ExcelApp.DisplayAlerts = false;
            Cell cell = new Cell(sheet_summary, "A1");

            var store_list = list.GroupBy(x => x.StoreNo).Select(x => new {
                StoreNo = x.Key,
                StoreName = x.First().StoreName,
                Data = x
            }).ToList();

            //Copy Sheets
            for (int i = store_list.Count()-1; i >= 0 ; i--) {
                InteropExcel.Worksheet sheet = workbook.Sheets.Add(Type.Missing, sheet_store, 1, Type.Missing);
                sheet_store.Cells.Copy();
                sheet.Paste();
                sheet.Name = $"{store_list[i].StoreName}";
                new Cell(sheet, "A1").Range.Select();
            }
            sheet_store.Delete();

            cell = new Cell(sheet_summary, "A1");
            cell.Value = $"{DateStart.Year}年{DateStart.Month}月 {BrandName} 門市營業額總表";
            cell = new Cell(sheet_summary, "A2");
            cell.Value = $"製表時間：{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}";

            cell = new Cell(sheet_summary, "D4");
            InteropExcel.Range range_from = sheet_summary.Range["C:C"];
            for (int i = store_list.Count() - 1; i >= 0; i--) {
                InteropExcel.Range range_to = sheet_summary.Range["D:D"];
                range_to.Insert(InteropExcel.XlInsertShiftDirection.xlShiftToRight, range_from.Copy());
                cell.Value = $"{store_list[i].StoreName}";
            }
            range_from.Delete();

            cell = new Cell(sheet_summary, "A5");
            string[] chtDayOfWeek = { "日", "一", "二", "三", "四", "五", "六" };
            for (int i = 1; i <= 31; i++) {
                if (i <= daysInMonth) {
                    cell.Value = i;
                    cell.CellPos.Col++;
                    cell.Value = chtDayOfWeek[(int)(new DateTime(DateStart.Year, DateStart.Month, i).DayOfWeek)];
                }
                else {
                    cell.Value = "";
                    cell.CellPos.Col++;
                    cell.Value = "";
                }
                cell.CellPos.Row++;
                cell.CellPos.Col--;
            }

            cell = new Cell(sheet_summary, "C5");
            foreach (var store in store_list) {
                foreach (var item in store.Data) {
                    int row_index = item.RetailDate.Day + 4;
                    cell.CellPos.Row = row_index;
                    if (item.TotalAmount > 0) cell.Value = item.TotalAmount;
                }
                cell.CellPos.Col++;
            }

            cell = new Cell(sheet_summary, "C5");
            cell.CellPos.Col += store_list.Count;
            for (int i = 1; i <= 32; i++) {
                var cellStart = new CellPos(cell.CellName);
                var cellEnd = new CellPos(cell.CellName);
                cellStart.Col -= store_list.Count;
                cellEnd.Col -= 1;
                cell.SetFormula($"=SUM({cellStart.CellName}:{cellEnd.CellName})");
                cell.CellPos.Row++;
            }

            int sheet_index = 2;
            foreach (var store in store_list) {
                InteropExcel.Worksheet sheet = workbook.Worksheets[sheet_index];
                cell = new Cell(sheet, "A1");
                cell.Value = $"{store.StoreNo} {store.StoreName}";
                cell.CellName = "A2";
                cell.Value = $"{DateStart.Month} 月";

                cell.CellName = "A4";
                for (int i = 1; i <= 31; i++) {
                    if (i > daysInMonth) break;
                    cell.Value = i;
                    cell.CellPos.Col++;
                    cell.Value = chtDayOfWeek[(int)(new DateTime(DateStart.Year, DateStart.Month, i).DayOfWeek)];
                    cell.CellPos.Row++;
                    cell.CellPos.Col--;
                }

                foreach (var item in store.Data) {
                    int row_index = item.RetailDate.Day + 3;
                    cell.CellName = $"E{row_index}";
                    if (item.UberEats > 0) cell.Value = item.UberEats;
                    cell.CellName = $"F{row_index}";
                    if (item.UberEatsCount > 0) cell.Value = item.UberEatsCount;
                    cell.CellName = $"H{row_index}";
                    if (item.FoodPanda > 0) cell.Value = item.FoodPanda;
                    cell.CellName = $"I{row_index}";
                    if (item.FoodPandaCount > 0) cell.Value = item.FoodPandaCount;
                    cell.CellName = $"K{row_index}";
                    if (item.Delivery > 0) cell.Value = item.Delivery;
                    cell.CellName = $"L{row_index}";
                    if (item.DeliveryCount > 0) cell.Value = item.DeliveryCount;
                    cell.CellName = $"N{row_index}";
                    if (item.TakeOut > 0) cell.Value = item.TakeOut;
                    cell.CellName = $"O{row_index}";
                    if (item.TakeOutCount > 0) cell.Value = item.TakeOutCount;
                    cell.CellName = $"Q{row_index}";
                    if (item.DineIn > 0) cell.Value = item.DineIn;
                    cell.CellName = $"R{row_index}";
                    if (item.DineInCount > 0) cell.Value = item.DineInCount;
                    cell.CellName = $"T{row_index}";
                    if (item.TakeSelf > 0) cell.Value = item.TakeSelf;
                    cell.CellName = $"U{row_index}";
                    if (item.TakeSelfCount > 0) cell.Value = item.TakeSelfCount;
                }

                sheet.Activate();
                sheet.Range["C4"].Select();
                ExcelApp.ActiveWindow.FreezePanes = true;

                sheet_index++;
            }

            //Output ChartObject
            //string chart_filename = $"{Path.GetTempPath()}\\{DateTime.Now.ToString("yyyyMMdd")}_chart.png";
            //InteropExcel.ChartObjects chartObjects = (InteropExcel.ChartObjects)(sheet.ChartObjects(Type.Missing));
            //foreach (InteropExcel.ChartObject co in chartObjects) {
            //    InteropExcel.Chart chart = (InteropExcel.Chart)co.Chart;
            //    chart.Export(chart_filename, "PNG", false);
            //    break;
            //}
            workbook.Worksheets[1].Activate();

            ExcelItem.Workbook.Save();
            ExcelItem.Workbook.Close(false);
            try {
                Global.MyExcelAppManager.KillNotUseItem();
            }
            catch (Exception) { }

            FileInfo fi = new FileInfo(FileName);
            FileName = Path.GetTempPath() + $"{BrandName}每日業績表_{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx";
            fi.CopyTo(FileName, true);

            er = new EmbeddedResource(Assembly.GetExecutingAssembly(), "MailTemplate.每日門市營業額.html");
            StreamReader sr = new StreamReader(er.GetStream());
            string mailbody = sr.ReadToEnd();

            mailbody = mailbody.Replace("{Year}", DateEnd.Year.ToString());
            mailbody = mailbody.Replace("{Month}", DateEnd.Month.ToString());
            mailbody = mailbody.Replace("{Day}", DateEnd.Day.ToString());
            mailbody = mailbody.Replace("{BrandName}", BrandName);

            var tt = "<tr><td width=\"180px\" align=\"right\">{StoreName}</td><td style=\"width: 100px;\" align=\"right\">{Amount} 元</td></tr>";
            var TableDataList = "";
            foreach (var store in store_list) {
                var StoreName = store.StoreName;
                decimal Amount = 0;
                if (store.Data.LastOrDefault().RetailDate.Date == DateEnd.Date) {
                    Amount = store.Data.LastOrDefault().TotalAmount;
                }
                TableDataList += tt.Replace("{StoreName}", StoreName).Replace("{Amount}", Amount.ToString("#,#")) + "\r\n";
            }
            mailbody = mailbody.Replace("{TableDataList}", TableDataList);

            if (mail_receiver != null) {
                CMailLinkResource linkRes = new CMailLinkResource();
                //linkRes.FileName = chart_filename;
                //linkRes.ContentType = "image/png";
                //linkRes.ContentID = "chart.png";

                var result = mail.SendMail(
                    mail_receiver,
                    $"{DateTime.Now.AddDays(-1).ToString("yyyyMMdd")} {BrandName}每日門市營業額",
                    mailbody,
                    true,
                    new string[] { FileName },
                    false,
                    null  //new CMailLinkResource[] { linkRes }
                );
            }
            else {
                System.Diagnostics.Process.Start(@"C:\Windows\explorer.exe", FileName);
            }
        }
    }
}
