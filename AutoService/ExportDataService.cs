﻿using System;
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
        private Timer timer = null;
        private DataAccess.DataAccess dataAccess;
        private DataAccess.BusinessLogic businessLogic;
        private string procName = "sp_ExportDataInvoce";
        private string path_folder_sql = " select val from options where ma_phan_he = 'GL' and name = 'm_auto_service_path' ";
        private string path_folder_sql_excel = " select val from options where ma_phan_he = 'GL' and name = 'm_auto_service_pathexcel' ";
        private string count_dvcs_sql = " select count(ma_dvcs) as sl_ma_dvcs from dmdvcskb where status = '1' ";
        private string ma_so_thue_dl_sql = " select ma_so_thue as ma_so_thue from dmdvcskb where ma_dvcs = 'CTY' and status = '1' ";
        private string ma_so_thue_cn_sql = " select ma_so_thue as ma_so_thue from dmdvcskb where ma_dvcs <> 'CTY' and status = '1' ";
        private string m_auto_service_h = " select val from options where ma_phan_he = 'GL' and name = 'm_auto_service_h'";
        private string m_auto_service_from_name = " select val from options where ma_phan_he = 'GL' and name = 'm_auto_service_from_name'";
        private string m_auto_service_from_pass = " select val from options where ma_phan_he = 'GL' and name = 'm_auto_service_from_pass'";
        private string m_auto_service_to = " select val from options where ma_phan_he = 'GL' and name = 'm_auto_service_to'";
        private string m_auto_service_ccID = " select val from options where ma_phan_he = 'GL' and name = 'm_auto_service_ccID'";
        private string m_auto_service_bccID = " select val from options where ma_phan_he = 'GL' and name = 'm_auto_service_bccID'";
        private string m_auto_service_title = " select val from options where ma_phan_he = 'GL' and name = 'm_auto_service_title'";
        private string m_auto_service_body = " select val from options where ma_phan_he = 'GL' and name = 'm_auto_service_body'";
        private string m_auto_service_from_host = " select val from options where ma_phan_he = 'GL' and name = 'm_auto_service_from_host'";
        private string m_auto_service_from_port = " select val from options where ma_phan_he = 'GL' and name = 'm_auto_service_from_port'";

        private string pathFileFolder = null;
        private string pathFileFolderExcel = null;
        private string rootPath = null;
        private string rootPathExcel = null;
        private string count_dvcs = "0";
        private string fistDayInMonthFi = null;
        private string dayNowFi = null;
        private int hourExportData = 0;

        //DateTime datetimeXX = new DateTime(2017, 04, 27, 13, 00, 00);

        public ExportDataService()
        {
            InitializeComponent();
            dataAccess = new DataAccess.DataAccess();
            businessLogic = new DataAccess.BusinessLogic();
        }

        protected override void OnStart(string[] args)
        {
            timer = new Timer();
            // 1h run time 1 step
            timer.Interval = 3600000;
            //timer.Interval = 10000;
            timer.Elapsed += timer_Tick;
            timer.Enabled = true;
        }

        private void timer_Tick(object sender, ElapsedEventArgs args)
        {
            //DateTime datetime = datetimeXX.AddDays(ctest);
            DateTime datetimeXX = DateTime.Now;
            DateTime datetime = datetimeXX.AddDays(-1);
            string hourExport = dataAccess.GetData(m_auto_service_h).Rows[0][0].ToString().Trim();
            
            try
            {
                hourExportData = Int32.Parse(hourExport);
            } catch (Exception exx)
            {
                hourExportData = 13;
                businessLogic.logbug(exx.ToString());
            }
            
            if (datetime.Hour == hourExportData)
            {
                DataTable da = new DataTable();
                pathFileFolder = dataAccess.GetData(path_folder_sql).Rows[0][0].ToString().Trim() + "\\";
                pathFileFolderExcel = dataAccess.GetData(path_folder_sql_excel).Rows[0][0].ToString().Trim() + "\\";
                if ("".Equals(pathFileFolder) || pathFileFolder == null)
                {
                    pathFileFolder = "D:\\Michelin\\FTP\\Upload\\";
                }
                else
                {
                    if (!System.IO.Directory.Exists(pathFileFolder))
                    {
                        string folderNametxt = @"D:\Michelin\FTP";

                        string pathString1 = System.IO.Path.Combine(folderNametxt, "Upload");
                        string pathString2 = System.IO.Path.Combine(folderNametxt, "Encrypted");
                        string pathString3 = System.IO.Path.Combine(folderNametxt, "Sent");
                        System.IO.Directory.CreateDirectory(pathString1);
                        System.IO.Directory.CreateDirectory(pathString2);
                        System.IO.Directory.CreateDirectory(pathString3);
                        pathFileFolder = "D:\\Michelin\\FTP\\Upload\\";
                    }
                }
                if ("".Equals(pathFileFolderExcel) || pathFileFolderExcel == null)
                {
                    pathFileFolderExcel = "D:\\FAST\\AutoService\\Excel\\";
                }
                else
                {
                    if (!System.IO.Directory.Exists(pathFileFolderExcel))
                    {
                        string folderNameExcel = @"D:\FAST\AutoService\Excel";
                        System.IO.Directory.CreateDirectory(folderNameExcel);
                        pathFileFolderExcel = "D:\\FAST\\AutoService\\Excel\\";
                    }
                }

                count_dvcs = dataAccess.GetData(count_dvcs_sql).Rows[0][0].ToString().Trim();
                if ("0".Equals(count_dvcs))
                {
                    businessLogic.logbug("Not find madvcs");
                }
                else if ("1".Equals(count_dvcs) && pathFileFolder != null)
                {

                    string ma_so_thue_a = dataAccess.GetData(ma_so_thue_dl_sql).Rows[0][0].ToString().Trim();
                    rootPath = @"" + pathFileFolder + "Sales_Detail_" + ma_so_thue_a + "_A_" + datetime.ToString("yyyyMMdd") + ".txt";
                    rootPathExcel = @"" + pathFileFolderExcel + "Sales_Detail_" + ma_so_thue_a + "_A_" + datetime.ToString("yyyyMMdd") + ".xlsx";
                    da = getDataTable(datetime);
                    businessLogic.writeResult(rootPath, da);

                        //businessLogic.writeResultToExcel(da, pathFileFolder);
                        //businessLogic.writeResultToExcel(da, rootPathTest);
                        string from_name = dataAccess.GetData(m_auto_service_from_name).Rows[0][0].ToString().Trim();
                        string from_pass = dataAccess.GetData(m_auto_service_from_pass).Rows[0][0].ToString().Trim();
                        string from_host = dataAccess.GetData(m_auto_service_from_host).Rows[0][0].ToString().Trim();
                        int from_port = 0;
                        try
                        {
                            from_port = Int32.Parse(dataAccess.GetData(m_auto_service_from_port).Rows[0][0].ToString().Trim());
                        }
                        catch (Exception ex)
                        {
                            businessLogic.logbug(ex.ToString());
                        }

                        string to = dataAccess.GetData(m_auto_service_to).Rows[0][0].ToString().Trim();
                        string ccID = dataAccess.GetData(m_auto_service_ccID).Rows[0][0].ToString().Trim();
                        string bccID = dataAccess.GetData(m_auto_service_bccID).Rows[0][0].ToString().Trim();
                        string title = dataAccess.GetData(m_auto_service_title).Rows[0][0].ToString().Trim();
                        string service_body = dataAccess.GetData(m_auto_service_body).Rows[0][0].ToString().Trim();

                        businessLogic.sendMail(from_name, from_pass, from_host, from_port, to, ccID, bccID, title, service_body, da, rootPathExcel, fistDayInMonthFi, dayNowFi);
                        dayNowFi = null;
                        fistDayInMonthFi = null;
                }
                else
                {
                    DataTable daTable = dataAccess.GetData(ma_so_thue_cn_sql);
                    string ma_so_thue_a = dataAccess.GetData(ma_so_thue_dl_sql).Rows[0][0].ToString().Trim();

                    for (int i = 0; i < daTable.Rows.Count; i++)
                    {
                        string ma_so_thue_b = daTable.Rows[i][0].ToString().Trim();
                        rootPath = @"" + pathFileFolder + "Sales_Detail_" + ma_so_thue_a + "_" + ma_so_thue_b + "_" + datetime.ToString("yyyyMMdd") + ".txt";
                    }

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
                    fistDayInMonthFi = businessLogic.formatDate(fistDayInMonth.ToString().Substring(0, 9));
                    dayNowFi = businessLogic.formatDate(datetime.ToString().Substring(0, 9));
                    da = dataAccess.ExecuteProc(procName, fistDayInMonth, datetime);
                    //businessLogic.logbug(procName + fistDayInMonth+ datetime);
                }
                else if (dateNow >= 6 && dateNow <= lastDateOfMonthNow)
                {
                    DateTime fistDayInMonth = businessLogic.getFistDayInMonth(yearNow, monthNow);
                    fistDayInMonthFi = businessLogic.formatDate(fistDayInMonth.ToString().Substring(0, 9));
                    dayNowFi = businessLogic.formatDate(datetime.ToString().Substring(0, 9));
                    da = dataAccess.ExecuteProc(procName, fistDayInMonth, datetime);
                    //businessLogic.logbug(procName + fistDayInMonth + datetime);
            }
            //}
            return da;
        }

        protected override void OnStop()
        {
            
        }

    }
}
