namespace AutoReportCreator {
    partial class FrmMain {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent() {
            this.btnStockForecastReport = new System.Windows.Forms.Button();
            this.pb = new System.Windows.Forms.ProgressBar();
            this.chkSendMail = new System.Windows.Forms.CheckBox();
            this.btnDailySaleReport = new System.Windows.Forms.Button();
            this.btnDailyStoreReport = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnStockForecastReport
            // 
            this.btnStockForecastReport.Location = new System.Drawing.Point(30, 48);
            this.btnStockForecastReport.Margin = new System.Windows.Forms.Padding(4);
            this.btnStockForecastReport.Name = "btnStockForecastReport";
            this.btnStockForecastReport.Size = new System.Drawing.Size(269, 59);
            this.btnStockForecastReport.TabIndex = 0;
            this.btnStockForecastReport.Text = "庫存採購預估報表";
            this.btnStockForecastReport.UseVisualStyleBackColor = true;
            this.btnStockForecastReport.Click += new System.EventHandler(this.btnStockForecastReport_Click);
            // 
            // pb
            // 
            this.pb.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.pb.Location = new System.Drawing.Point(0, 271);
            this.pb.Name = "pb";
            this.pb.Size = new System.Drawing.Size(330, 23);
            this.pb.TabIndex = 1;
            // 
            // chkSendMail
            // 
            this.chkSendMail.AutoSize = true;
            this.chkSendMail.Location = new System.Drawing.Point(30, 21);
            this.chkSendMail.Name = "chkSendMail";
            this.chkSendMail.Size = new System.Drawing.Size(123, 20);
            this.chkSendMail.TabIndex = 3;
            this.chkSendMail.Text = "直接寄出信件";
            this.chkSendMail.UseVisualStyleBackColor = true;
            // 
            // btnDailySaleReport
            // 
            this.btnDailySaleReport.Location = new System.Drawing.Point(30, 115);
            this.btnDailySaleReport.Margin = new System.Windows.Forms.Padding(4);
            this.btnDailySaleReport.Name = "btnDailySaleReport";
            this.btnDailySaleReport.Size = new System.Drawing.Size(269, 59);
            this.btnDailySaleReport.TabIndex = 4;
            this.btnDailySaleReport.Text = "每日銷貨統計報表";
            this.btnDailySaleReport.UseVisualStyleBackColor = true;
            this.btnDailySaleReport.Click += new System.EventHandler(this.btnDailySaleReport_Click);
            // 
            // btnDailyStoreReport
            // 
            this.btnDailyStoreReport.Location = new System.Drawing.Point(30, 181);
            this.btnDailyStoreReport.Name = "btnDailyStoreReport";
            this.btnDailyStoreReport.Size = new System.Drawing.Size(269, 59);
            this.btnDailyStoreReport.TabIndex = 5;
            this.btnDailyStoreReport.Text = "一芳、霜江、美濃門市每日報表";
            this.btnDailyStoreReport.UseVisualStyleBackColor = true;
            this.btnDailyStoreReport.Click += new System.EventHandler(this.btnDailyStoreReport_Click);
            // 
            // FrmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(330, 294);
            this.Controls.Add(this.btnDailyStoreReport);
            this.Controls.Add(this.btnDailySaleReport);
            this.Controls.Add(this.chkSendMail);
            this.Controls.Add(this.pb);
            this.Controls.Add(this.btnStockForecastReport);
            this.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.Name = "FrmMain";
            this.Text = "報表自動產出程式";
            this.Load += new System.EventHandler(this.FrmMain_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnStockForecastReport;
        private System.Windows.Forms.ProgressBar pb;
        private System.Windows.Forms.CheckBox chkSendMail;
        private System.Windows.Forms.Button btnDailySaleReport;
        private System.Windows.Forms.Button btnDailyStoreReport;
    }
}

