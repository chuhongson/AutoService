using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.IO;
using DataAccess;

namespace AutoService
{
    public partial class ExportDataService : ServiceBase
    {
        private bool isFlg = true;
        private Timer timer = null;
        private DataAccess.DataAccess dataAccess;
        private DataAccess.BusinessLogic businessLogic;
        //private string sql = "select h.ma_so_thue , h.ten_kh , 'IV' as hang_bd, m.ngay_lct, m.so_ct, d.line_nbr, d.ma_vt, vt.ten_vt, d.dvt, d.so_luong from m81$201806 m inner join dmkh h on m.ma_kh = h.ma_kh inner join d81$201806 d on m.stt_rec = d.stt_rec inner join dmvt vt on d.ma_vt = vt.ma_vt";
        private string procName = "sp_ExportDataInvoce";
        private string path_folder_sql = " select val from options where ma_phan_he = 'GL' and name = 'm_path' ";
        private string count_dvcs_sql = " select count(ma_dvcs) as sl_ma_dvcs from dmdvcskb where status = '1' ";
        private string ma_so_thue_dl_sql = " select ma_so_thue as ma_so_thue from dmdvcskb where ma_dvcs = 'CTY' and status = '1' ";
        private string ma_so_thue_cn_sql = " select ma_so_thue as ma_so_thue from dmdvcskb where ma_dvcs <> 'CTY' and status = '1' ";

        private string pathFileFolder = null;
        private string rootPath = null;
        private string rootPathTest = null;
        private string count_dvcs = "0";

        private int ctest = 0;
        DateTime datetimeXX = DateTime.Now;

        public ExportDataService()
        {
            InitializeComponent();
            dataAccess = new DataAccess.DataAccess();
            businessLogic = new DataAccess.BusinessLogic();
        }

        protected override void OnStart(string[] args)
        {
            timer = new Timer();
            //timer.Interval = 86400000;
            
            timer.Interval = 3000;
            timer.Elapsed += timer_Tick;
            timer.Enabled = true;
        }

        private void timer_Tick(object sender, ElapsedEventArgs args)
        {
            ctest++;
            DateTime datetime = datetimeXX.AddDays(ctest);

            DataTable da = new DataTable();
            pathFileFolder = dataAccess.GetData(path_folder_sql).Rows[0][0].ToString().Trim() + "\\";
            if ("".Equals(pathFileFolder) || pathFileFolder ==  null)
            {
                pathFileFolder = AppDomain.CurrentDomain.BaseDirectory + "\\";
            }

            count_dvcs = dataAccess.GetData(count_dvcs_sql).Rows[0][0].ToString().Trim();
            if ("0".Equals(count_dvcs))
            {
                businessLogic.logbug("Not find madvcs");
            }
            else if ("1".Equals(count_dvcs) && pathFileFolder != null)
            {
                
                string ma_so_thue_a = dataAccess.GetData(ma_so_thue_dl_sql).Rows[0][0].ToString().Trim();
                rootPath = @"" + pathFileFolder + "Sales_Detail_" + ma_so_thue_a + "_A_" + datetime.AddDays(-1).ToString("yyyyMMdd") + ".txt";
                rootPathTest = @"" + pathFileFolder + "Sales_Detail_" + ma_so_thue_a + "_A_" + datetime.AddDays(-1).ToString("yyyyMMdd") + ".xlsx";
                da = getDataTable(datetime);
                businessLogic.writeResult(rootPath, da);
                
                if (isFlg == true)
                 {
                    //businessLogic.writeResultToExcel(da, pathFileFolder);
                    //businessLogic.writeResultToExcel(da, rootPathTest);
                    businessLogic.sendMail("hongcuong206@gmail.com", "hongcuong205", "hongsonbk1@gmail.com", "hongson5018@gmail.com", "hongson5018@gmail.com", "Test send auto email of fast system", da, rootPathTest);
                    isFlg = false;
                }
                    
            }
            else
            {
                DataTable daTable = dataAccess.GetData(ma_so_thue_cn_sql);
                string ma_so_thue_a = dataAccess.GetData(ma_so_thue_dl_sql).Rows[0][0].ToString().Trim();

                for (int i = 0; i < daTable.Rows.Count; i ++)
                {
                    string ma_so_thue_b = daTable.Rows[i][0].ToString().Trim();
                    rootPath = @"" + pathFileFolder + "Sales_Detail_" + ma_so_thue_a + "_" + ma_so_thue_b + "_" + datetime.AddDays(-1).ToString("yyyyMMdd") + ".txt";
                }
                
            }

        }


        private DataTable getDataTable(DateTime datetime)
        {
            DataTable da = new DataTable();
            int secondNow = datetime.Second;
            int hourNow = datetime.Hour;
            int dateNow = datetime.Day;
            int monthNow = datetime.Month;
            int yearNow = datetime.Year;

            DateTime dtPre = datetime.AddMonths(-1);
            int datePre = dtPre.Day;
            int monthPre = dtPre.Month;
            int yearPre = dtPre.Year;

            int lastDateOfMonthNow = businessLogic.getLastDayInMonth(yearNow, monthNow);

            // Set hour to export data
            //if (((hourNow * 60) + secondNow) == 780)
            //{
                if (dateNow >= 1 && dateNow <= 5)
                {
                    DateTime fistDayInMonth = businessLogic.getFistDayInMonth(yearPre, monthPre);
                    da = dataAccess.ExecuteProc(procName, fistDayInMonth, datetime);
                    businessLogic.logbug(procName + fistDayInMonth+ datetime);
                }
                else if (dateNow >= 6 && dateNow <= lastDateOfMonthNow)
                {
                    DateTime fistDayInMonth = businessLogic.getFistDayInMonth(yearNow, monthNow);
                    da = dataAccess.ExecuteProc(procName, fistDayInMonth, datetime);
                    businessLogic.logbug(procName + fistDayInMonth + datetime);
            }
            //}
            return da;
        }

        protected override void OnStop()
        {
            
        }

    }
}
