using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using System.Net.Mail;
using OfficeOpenXml;
using System.Xml;
using System.Drawing;
using OfficeOpenXml.Style;
using DataAccess;

namespace DataAccess
{
    public class BusinessLogic
    {
        /// <summary>
        /// getLastDayInMonth
        /// </summary>
        /// <param>y</param>
        /// <param>m</param>
        /// <returns>int</returns>
        public int getLastDayInMonth(int y, int m)
        {
            DateTime dateTimeX = new DateTime(y, m, 1);
            DateTime dateTimeY = dateTimeX.AddMonths(1).AddDays(-1);
            return dateTimeY.Day;
        }

        /// <summary>
        /// getFistDayInMonth
        /// </summary>
        /// <param>y</param>
        /// <param>m</param>
        /// <returns>DateTime</returns>
        public DateTime getFistDayInMonth(int y, int m)
        {
            DateTime dt1 = new DateTime(y, m, 1);
            return dt1;
        }

        /// <summary>
        /// Write to file logbug.txt
        /// </summary>
        /// <param>String err </param>
        /// <returns>file logbug.txt</returns>
        public void logbug(string str)
        {
            DataAccess dataAccess = new DataAccess();
            string path_folder_sql = "select val from options where ma_phan_he = 'GL' and name = 'm_auto_service_path'";
            string errPathDefault = @"" + AppDomain.CurrentDomain.BaseDirectory + "\\" + "logbug.txt";
            string errPath = dataAccess.GetData(path_folder_sql).Rows[0][0].ToString().Trim() + "\\" + "logbug.txt";
            FileStream fs = new FileStream(errPath, FileMode.Create);
            StreamWriter swer = new StreamWriter(fs, Encoding.Unicode);
            try
            {
                swer.WriteLine(DateTime.Now + str);
                swer.Flush();
                swer.Close();
            }
            catch (IOException ioerr)
            {
                FileStream fss = new FileStream(errPath, FileMode.Create);
                StreamWriter swerr = new StreamWriter(fs, Encoding.Unicode);
                swer.WriteLine(DateTime.Now + "Bug Default : " + ioerr.ToString());
                swerr.Flush();
                swerr.Close();
            }
        }

        /// <summary>
        /// Export data to file txt
        /// </summary>
        /// <param>string rootPath, DataTable da</param>
        /// <returns>void</returns>
        public void writeResult(string rootPath, System.Data.DataTable da)
        {
            FileStream fs = new FileStream(rootPath, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs, Encoding.Unicode);
            StringBuilder result = new StringBuilder("");
            for (int i = 0; i < da.Rows.Count; i++)
            {
                for (int j = 0; j < da.Columns.Count; j++)
                {
                    if (j == 3)
                    {
                        result.Append(da.Rows[i][j].ToString().Substring(0, 9) + "|");
                    }
                    else if (j == 9)
                    {
                        string str_so_luong = da.Rows[i][j].ToString();
                        if (str_so_luong.IndexOf(".") != -1)
                        {
                            result.Append(str_so_luong.Substring(0, str_so_luong.IndexOf(".")));
                        }
                        else
                        {
                            result.Append(Int32.Parse(str_so_luong).ToString());
                        }
                    }
                    else
                    {
                        result.Append(da.Rows[i][j].ToString().Trim() + "|");
                    }
                }

                if (!"".Equals(result))
                {
                    sw.WriteLine(result);
                    result.Clear();
                }
                else
                {
                    logbug("Server can't read data!");
                }
            }
            sw.Flush();
            sw.Close();
        }


        /// <summary>
        /// Send file file excel to email
        /// </summary>
        /// <param name="from">from</param>
        /// <param name="password">pass</param>
        /// <param name="to">to</param>
        /// <param name="ccID">ccID</param>
        /// <param name="bccID">bccID</param>
        /// <param name="da">da</param>
        /// <returns>bool</returns>
        public bool sendMail(string from, string password, string host, int post, string to, string ccID, string bccID,string subjectTitle, System.Data.DataTable da, string path)
        {
            try
            {
                MailMessage mailMessage = new MailMessage();
                mailMessage.From = new MailAddress(from);

                mailMessage.Subject = subjectTitle.Trim();
                mailMessage.SubjectEncoding = System.Text.Encoding.UTF8;
                mailMessage.Body = "Body Test send auto email of fast system";
                mailMessage.BodyEncoding = System.Text.Encoding.UTF8;
                mailMessage.IsBodyHtml = false;
                mailMessage.Priority = MailPriority.High;
                writeResultToExcel(da, path);
                //string excelPath = @"" + AppDomain.CurrentDomain.BaseDirectory + "\\tmpSendEmail.xls";
                mailMessage.Attachments.Add(new Attachment(path));

                // send with to
                string[] toEmails = to.Split(',');
                foreach (string toEmail in toEmails)
                {
                    mailMessage.To.Add(new MailAddress(toEmail));
                }

                // send with cc
                string[] ccIDs = ccID.Split(',');
                foreach (string ccIDmail in ccIDs)
                {
                    mailMessage.CC.Add(new MailAddress(ccIDmail));
                }

                // send with bcc
                string[] bccIDs = bccID.Split(',');
                foreach (string bccIDmail in bccIDs)
                {
                    mailMessage.Bcc.Add(new MailAddress(bccIDmail));
                }

                SmtpClient smtpClient = new SmtpClient();
                smtpClient.Host = host;
                smtpClient.Port = post;
                smtpClient.Credentials = new System.Net.NetworkCredential(from, password);
                smtpClient.EnableSsl = true;
                smtpClient.Timeout = 500000;

                try
                {
                    smtpClient.Send(mailMessage);
                    return true;
                }
                catch (System.Net.Mail.SmtpException smtpErr)
                {
                    logbug(smtpErr.ToString());
                    return false;
                }
            } catch (Exception ex)
            {
                logbug(ex.ToString());
                return false;
            }
        }

        public void writeResultToExcel(System.Data.DataTable da, string path)
        {
            using (var excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("FastAutoService");

                worksheet.Cells[1, 1].Value = "MST";
                worksheet.Cells[1, 2].Value = "Ten KH";
                worksheet.Cells[1, 3].Value = "Hang Ban|Tra";
                worksheet.Cells[1, 4].Value = "Ngay";
                worksheet.Cells[1, 5].Value = "So HD";
                worksheet.Cells[1, 6].Value = "So Dong";
                worksheet.Cells[1, 7].Value = "MA Hang";
                worksheet.Cells[1, 8].Value = "Ten Hang";
                worksheet.Cells[1, 9].Value = "Ma CAI";
                worksheet.Cells[1, 10].Value = "So Luong";

                excelPackage.Workbook.Properties.Author = "FastAutoService";
                excelPackage.Workbook.Properties.Title = "FastAutoService";
                excelPackage.Workbook.Properties.Comments = "FastAutoService for send email";
                
                var workSheet = excelPackage.Workbook.Worksheets[1];
                for (int i = 0; i < da.Rows.Count; i ++)
                {
                    workSheet.Cells[i + 2, 1].Value = da.Rows[i][0].ToString();
                    workSheet.Cells[i + 2, 2].Value = da.Rows[i][1].ToString();
                    workSheet.Cells[i + 2, 3].Value = da.Rows[i][2].ToString();
                    workSheet.Cells[i + 2, 4].Value = da.Rows[i][3].ToString();
                    workSheet.Cells[i + 2, 5].Value = da.Rows[i][4].ToString();
                    workSheet.Cells[i + 2, 6].Value = da.Rows[i][5].ToString();
                    workSheet.Cells[i + 2, 7].Value = da.Rows[i][6].ToString();
                    workSheet.Cells[i + 2, 8].Value = da.Rows[i][7].ToString();
                    workSheet.Cells[i + 2, 9].Value = da.Rows[i][8].ToString();
                    if (da.Rows[i][9].ToString().IndexOf(".") != -1)
                    {
                        string str_so_luong = da.Rows[i][9].ToString();
                        workSheet.Cells[i + 2, 10].Value = str_so_luong.Substring(0, str_so_luong.IndexOf("."));
                    }
                    else
                    {
                        workSheet.Cells[i + 2, 10].Value = da.Rows[i][9].ToString();
                    }
                }
                workSheet.Protection.IsProtected = false;
                workSheet.Protection.AllowSelectLockedCells = false;
                excelPackage.SaveAs(new FileInfo(path));
            }
        }
    }
}
