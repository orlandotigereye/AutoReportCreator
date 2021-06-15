using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using InteropExcel = Microsoft.Office.Interop.Excel;
using SnweiCom.Data;
using SnweiCom.Office.Excel;
using System.Reflection;
using SnweiCom.IO;
using System.IO;
using SnweiCom.Net;

namespace AutoReportCreator {
    public partial class FrmMain : Form {
        public DAC dacInkismMIS = DacGet.dacInkismMIS;
        public DAC dacErpDB = DacGet.dacErpDB;
        CMail mail = new CMail("192.168.2.111", "", "");

        public FrmMain() {
            InitializeComponent();
            mail.MailFrom = "report@inkism.com.tw";
        }

        public class StockForecastReportData {
            public string MaterialID { get; set; }
            public string MaterialName { get; set; }
            public string StockUnit { get; set; }
            public decimal StockAmount { get; set; }
            public decimal MonthInStock { get; set; }
            public decimal MonthOutStock { get; set; }
            public decimal Requesting { get; set; }
            public decimal Requested { get; set; }
            public decimal Purchasing { get; set; }
            public decimal Purchased { get; set; }
            public decimal ThisMonthForecast { get; set; }
            public decimal NextMonthForecast { get; set; }
            public decimal SafetyStock { get; set; }
            public decimal MaxStock { get; set; }
            public decimal MinPurchase { get; set; }
            public decimal ArriveDays { get; set; }
        }

        private List<StockForecastReportData> GetStockForecastReportDataList() {
            string strSQL = $@"
USE NEWIPO;
declare @DepotList table (DepotID nvarchar(10))
insert into @DepotList values ('HQ011'), ('HQ012'), ('HQ071'), ('HQ072')
declare @ThisYearMonth as nvarchar(10) = '@@ThisYearMonth'
declare @NextYearMonth as nvarchar(10) = '@@NextYearMonth'
declare @DateStart as nvarchar(10) = '@@DateStart'
declare @DateEnd as nvarchar(10) = '@@DateEnd'
select 
	--品號
	rtrim(MB001) MaterialID,
	--品名
	MB002 MaterialName,
    --單位
	MB004 StockUnit,
	--庫存量
	isnull((select sum(MC007) from INVMC where MC001=MB001 and MC002 in (select * from @DepotList)), 0) StockAmount,
	--入庫
	isnull((select sum(LA011) from INVLA where LA001=MB001 and LA005= 1 and LA004 between @DateStart and @DateEnd and LA009 in (select * from @DepotList)), 0) MonthInStock,
	--出庫
	isnull((select sum(LA011) from INVLA where LA001=MB001 and LA005=-1 and LA004 between @DateStart and @DateEnd and LA009 in (select * from @DepotList)), 0) MonthOutStock,

	--請購簽核中
	isnull((
		select sum(TB009) from PURTA
		left join PURTB on TA001=TB001 and TA002=TB002
		where
			TA016 in ('1', 'N') and TA007='N' and TB039='N' and
			TB004=MB001 and TB008 in (select * from @DepotList)
	), 0) Requesting,

	--已請購
	isnull((
		select sum(TB009) from PURTA
		left join PURTB on TA001=TB001 and TA002=TB002
		where
			TA007='Y' and TB039='N' and
			TB004=MB001 and TB008 in (select * from @DepotList)
	), 0) Requested,
	--採購簽核中
	isnull((
		select sum(TD008-TD015)
		from PURTC
		left join PURTD on TC001=TD001 and TC002=TD002
		where
			TC030 in ('1') and
			TC014='N' and TD016='N' and
			TD004=MB001 and TD007 in (select * from @DepotList)
	), 0) Purchasing,
	--己採未進
	isnull((
		select sum(TD008-TD015)
		from PURTC
		left join PURTD on TC001=TD001 and TC002=TD002
		where
			TC014='Y' and TD016='N' and
			TD004=MB001 and TD007 in (select * from @DepotList)
	), 0) Purchased,

	--本月銷售預估
	F1.ForecastAmount ThisMonthForecast,
	--次月鎖量預估
	F2.ForecastAmount NextMonthForecast,
	--尚需補貨數量
	--本月預估與實際差異
	--本月預估與實際差異%

	--安全存量
	A.SafetyStock,
	--最高庫存上限
	A.MaxStock,
	--最小訂購量(MOQ)
	A.MinPurchase,
	--交期(天)
	A.ArriveDays
from [192.168.2.37].[InkismMIS].[dbo].ReportMaterial A
left join INVMB on MB001=A.MaterialID
left join [192.168.2.37].[InkismMIS].[dbo].ReportMaterialForecast F1 on F1.MaterialID = A.MaterialID and F1.YearMonth = @ThisYearMonth
left join [192.168.2.37].[InkismMIS].[dbo].ReportMaterialForecast F2 on F2.MaterialID = A.MaterialID and F2.YearMonth = @NextYearMonth
order by
	A.OrderIndex
";
            strSQL = strSQL.Replace("@@ThisYearMonth", DateTime.Now.ToString("yyyyMM"));
            strSQL = strSQL.Replace("@@NextYearMonth", DateTime.Now.AddMonths(1).ToString("yyyyMM"));
            strSQL = strSQL.Replace("@@DateStart", DateTime.Now.ToString("yyyyMM") + "01");
            strSQL = strSQL.Replace("@@DateEnd", DateTime.Now.ToString("yyyyMMdd"));

            return dacErpDB.QueryDataTable(strSQL).ToList<StockForecastReportData>();
        }

        private void btnStockForecastReport_Click(object sender, EventArgs e) {
            var mail_receiver = new string[] {
                "george@inkism.com.tw",         //董事長
                "222001@inkism.com.tw",         //採購部
                "danny.chen@inkism.com.tw",     //陳堯鑫
                "jason.shen@inkism.com.tw",     //沈鉑淳
                "225001@inkism.com.tw",         //儲運部
                "amberlai@inkism.com.tw",       //賴俞汶
                "mayzhou@inkism.com.tw",        //周美君
                "tina.chen@inkism.com.tw",      //陳玟均
                "flora.chiang@inkism.com.tw",   //江秀惠
                "claire.cheng@inkism.com.tw",   //鄭鈺璉
                "snwei@inkism.com.tw",          //黃勝威
                "alex.yang@inkism.com.tw"       //楊志豪
            };


            var list = GetStockForecastReportDataList();
            pb.Maximum = list.Count;
            pb.Value = 0;

            EmbeddedResource er = new EmbeddedResource(Assembly.GetExecutingAssembly(), "ExcelTemplate.庫存採購預測報表.xlsx");
            string FileName = er.CopyToTempFile();
            ExcelAppManager.ExcelItem ExcelItem = Global.MyExcelAppManager.OpenExcel(FileName);
            InteropExcel.Application ExcelApp = ExcelItem.App;
            InteropExcel.Worksheet sheet = ExcelItem.Workbook.Worksheets[1];
            ExcelApp.ScreenUpdating = false;
            ExcelApp.Visible = false;
            ExcelApp.DisplayAlerts = false;
            Cell cell = new Cell(sheet, "A1");

            ExcelTools.DuplicateRow(sheet, 6, 1, list.Count);

            cell.CellName = "A3";
            cell.Value = $"資料日期：{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}";

            cell.CellName = "A6";
            int i = 1;
            foreach (var item in list) {
                cell.Value = item.MaterialID; cell.CellPos.Col++;
                cell.Value = item.MaterialName; cell.CellPos.Col++;
                cell.Value = item.StockUnit; cell.CellPos.Col++;

                cell.Value = item.StockAmount; cell.CellPos.Col++;
                cell.Value = item.MonthInStock; cell.CellPos.Col++;
                cell.Value = item.MonthOutStock; cell.CellPos.Col++;
                cell.Value = item.Requesting; cell.CellPos.Col++;
                cell.Value = item.Requested; cell.CellPos.Col++;
                cell.Value = item.Purchasing; cell.CellPos.Col++;
                cell.Value = item.Purchased; cell.CellPos.Col++;
                cell.Value = item.ThisMonthForecast; cell.CellPos.Col++;
                cell.Value = item.NextMonthForecast; cell.CellPos.Col++;

                //尚需補貨數量
                cell.CellPos.Col++;
                //本月預估與實際差異
                cell.CellPos.Col++;
                //本月預估與實際差異 %
                cell.CellPos.Col++;

                cell.Value = item.SafetyStock; cell.CellPos.Col++;
                cell.Value = item.MaxStock; cell.CellPos.Col++;
                cell.Value = item.MinPurchase; cell.CellPos.Col++;
                cell.Value = item.ArriveDays; cell.CellPos.Col++;

                cell.CellPos.Row++;
                cell.CellPos.Col = 1;
                pb.Value++;
                i++;
                Application.DoEvents();
            }

            ExcelItem.Workbook.Save();
            ExcelItem.Workbook.Close(false);
            try {
                Global.MyExcelAppManager.KillNotUseItem();
            }
            catch (Exception) { }

            FileInfo fi = new FileInfo(FileName);
            //FileName = $"R:\\{DateTime.Now.ToString("yyyyMMdd")} 庫存採購預測報表.xlsx";
            FileName = Path.GetTempPath() + $"{DateTime.Now.ToString("yyyyMMdd")} 庫存採購預測報表.xlsx";
            fi.CopyTo(FileName, true);

            if (chkSendMail.Checked) {
                var result = mail.SendMail(
                    mail_receiver,
                    $"{DateTime.Now.ToString("yyyyMMdd")} 庫存採購預測報表",
                    "詳如附件！此信此為系統自動寄出，請勿直接回覆！謝謝。",
                    true,
                    new string[] { FileName }
                );
            }
            else {
                System.Diagnostics.Process.Start(@"C:\Windows\explorer.exe", FileName);
            }
        }

        private void FrmMain_Load(object sender, EventArgs e) {
            chkSendMail.Checked = CmdOptions.AutoSendMail;

            this.Show();
            if (CmdOptions.StockForecastReport)
                btnStockForecastReport_Click(btnStockForecastReport, new EventArgs());
            if (CmdOptions.DailySaleReport)
                btnDailySaleReport_Click(btnDailySaleReport, new EventArgs());
            if (CmdOptions.DailyStoreReport)
                btnDailyStoreReport_Click(btnDailyStoreReport, new EventArgs());

            if (CmdOptions.AutoClose)
                this.Close();
        }

        private void btnDailySaleReport_Click(object sender, EventArgs e)
        {
            var mail_receiver = new string[] {
                "george@inkism.com.tw",         //董事長
                "alex.yang@inkism.com.tw",      //楊志豪
                "weihan.chen@inkism.com.tw",    //陳維漢
                "amberlai@inkism.com.tw",       //賴俞汶
                "mayzhou@inkism.com.tw",        //周美君
                "felix.kuo@inkism.com.tw",      //郭志傑
                "claire.cheng@inkism.com.tw",   //鄭鈺璉
                "snwei@inkism.com.tw"           //黃勝威
            };

            //mail_receiver = new string[] {
            //    "snwei@inkism.com.tw"           //黃勝威
            //};

            if (!chkSendMail.Checked) mail_receiver = null;
            this.Enabled = false;
            DailySaleReport report = new DailySaleReport(mail);
            report.Create(mail_receiver);
            this.Enabled = true;
        }

        private void btnDailyStoreReport_Click(object sender, EventArgs e)
        {
            //一芳
            var mail_receiver_yifang = new string[] {
                "george@inkism.com.tw",         //董事長
                "weihan.chen@inkism.com.tw",    //陳維漢
                "felix.kuo@inkism.com.tw",      //郭志傑

                "yuanxin.liao@inkism.com.tw",   //廖源鑫
                "david.chen@inkism.com.tw",     //陳弘衛
                "jamie.pan@inkism.com.tw",      //潘徵萍
                "steven@inkism.com.tw",         //李岳峰

                "ryan.lee@inkism.com.tw",       //李軏瀚
                "snwei@inkism.com.tw"
            };

            //霜江
            var mail_receiver_shuangjiang = new string[] {
                "george@inkism.com.tw",         //董事長
                "weihan.chen@inkism.com.tw",    //陳維漢
                "felix.kuo@inkism.com.tw",      //郭志傑

                "stanley@inkism.com.tw",        //黃明順
                "mia.shen@inkism.com.tw",       //沈郁欣

                "ryan.lee@inkism.com.tw",       //李軏瀚
                "snwei@inkism.com.tw"
            };

            //美濃
            var mail_receiver_mino = new string[] {
                "george@inkism.com.tw",         //董事長
                "weihan.chen@inkism.com.tw",    //陳維漢
                "felix.kuo@inkism.com.tw",      //郭志傑

                "taku0207@inkism.com.tw",       //李明耀
                "jennie.chiang@inkism.com.tw",  //蔣岷玹

                "ryan.lee@inkism.com.tw",       //李軏瀚
                "snwei@inkism.com.tw"
            };
            
            this.Enabled = false;
            var report = new DailyStoreReport(mail);

            if (!chkSendMail.Checked) mail_receiver_yifang = null;
            report.CreateAndSend(mail_receiver_yifang, DailyStoreReport.Brand.YiFang, "(1)");

            if (!chkSendMail.Checked) mail_receiver_shuangjiang = null;
            report.CreateAndSend(mail_receiver_shuangjiang, DailyStoreReport.Brand.ShuangJiang, "(1), (2)");

            if (!chkSendMail.Checked) mail_receiver_mino = null;
            report.CreateAndSend(mail_receiver_mino, DailyStoreReport.Brand.Mino, "(1), (2)");

            this.Enabled = true;
        }
    }
}
