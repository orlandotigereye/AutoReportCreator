using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using MySql.Data.MySqlClient;
using SnweiCom.Data;

namespace AutoReportCreator {
    public static class DacGet {
        public static string ErpDbIP = "192.168.2.121";
        public static string EZFlowDbIP = "192.168.2.121";
        public static string InkismMisIP = "192.168.2.37";

        public static DAC dacDSCSYS = new DAC(new SqlConnection(
            ConnStrFactory.SqlServerConnString(ErpDbIP, "DSCSYS", "sa", "Dsc123")
            ));

        public static DAC dacErpDB = new DAC(new SqlConnection(
            ConnStrFactory.SqlServerConnString(ErpDbIP, "DSC_CHT", "sa", "Dsc123")
            ));

        public static DAC dacInkismMIS = new DAC(new SqlConnection(
            ConnStrFactory.SqlServerConnString(InkismMisIP, "InkismMIS", "InkismMIS", "InkismMIS727")
            ));
    }
}
