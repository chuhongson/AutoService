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
                        result.Append(formatDate(da.Rows[i][j].ToString().Substring(0, 9)).Trim() + "|");
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
        public bool sendMail(string from, string password, string host, int post, string to, string ccID, string bccID,string subjectTitle , string contentBody, System.Data.DataTable da, string path, string fistDayInMonthFi, string datetime)
        {
            try
            {
                MailMessage mailMessage = new MailMessage();
                mailMessage.From = new MailAddress(from);

                mailMessage.Subject = subjectTitle.Trim();
                mailMessage.SubjectEncoding = System.Text.Encoding.UTF8;
                mailMessage.Body = contentBody;
                mailMessage.BodyEncoding = System.Text.Encoding.UTF8;
                mailMessage.IsBodyHtml = false;
                mailMessage.Priority = MailPriority.High;
                writeResultToExcel(da, path, fistDayInMonthFi, datetime);
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

        public void writeResultToExcel(System.Data.DataTable da, string path, string fistDayInMonthFi, string datetime)
        {
            using (var excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("FastAutoService");

                worksheet.Cells[1, 1, da.Rows.Count + 4, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[1, 1, da.Rows.Count + 4, 10].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 255, 255));

                worksheet.Row(1).Height = 30;
                worksheet.Row(2).Height = 30;
                worksheet.Cells[1, 1].Value = "BẢNG KÊ HÓA ĐƠN BÁN HÀNG";
                worksheet.Cells[1, 1, 1, 10].Merge = true;
                worksheet.Cells[1, 1, 1, 10].Style.Font.Bold = true;
                worksheet.Cells[1, 1, 1, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[1, 1, 1, 10].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheet.Cells[1, 1, 1, 10].Style.Font.Size = 16;

                worksheet.Cells[2, 1].Value = "Từ ngày " + fistDayInMonthFi.Trim() + "  đến ngày " + datetime.Trim();
                worksheet.Cells[2, 1, 2, 10].Merge = true;
                worksheet.Cells[2, 1, 2, 10].Style.Font.Size = 10;
                worksheet.Cells[2, 1, 2, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Cells[2, 1, 2, 10].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                worksheet.Cells[3, 1].Value = "Mã số thuế";
                worksheet.Cells[3, 2].Value = "Tên khách";
                worksheet.Cells[3, 3].Value = "Hàng bán(IV)| Trả lại(CN)";
                worksheet.Cells[3, 4].Value = "Ngày hóa đơn";
                worksheet.Cells[3, 5].Value = "Số hóa đơn";
                worksheet.Cells[3, 6].Value = "Số dòng";
                worksheet.Cells[3, 7].Value = "Mã hàng";
                worksheet.Cells[3, 8].Value = "Tên mặt hàng";
                worksheet.Cells[3, 9].Value = "Mã CAI Michelin";
                worksheet.Cells[3, 10].Value = "Số lượng";

                excelPackage.Workbook.Properties.Author = "FastAutoService";
                excelPackage.Workbook.Properties.Title = "FastAutoService";
                excelPackage.Workbook.Properties.Comments = "FastAutoService for send email";
                // Format style excel
                for (int i = 1; i <= 10; i ++)
                {
                    worksheet.Cells[3, i, 4, i].Merge = true;
                    worksheet.Cells[3, i, 4, i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells[3, i, 4, i].Style.Font.Bold = true;
                    worksheet.Cells[3, i, 4, i].AutoFitColumns();
                    worksheet.Cells[3, i, 4, i].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells[3, i, 4, i].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    worksheet.Cells[3, i, 4, i].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(237, 245, 255));
                    worksheet.Cells[3, i, 4, i].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[3, i, 4, i].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[3, i, 4, i].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells[3, i, 4, i].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                }


                // Write data to excel
                var workSheet = excelPackage.Workbook.Worksheets[1];
                for (int i = 0; i < da.Rows.Count; i ++)
                {
                    workSheet.Cells[i + 5, 1].Value = da.Rows[i][0].ToString().Trim();
                    worksheet.Cells[i + 5, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    worksheet.Cells[i + 5, 1].Style.Border.Top.Style = ExcelBorderStyle.Dotted;
                    worksheet.Cells[i + 5, 1].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                    workSheet.Cells[i + 5, 2].Value = da.Rows[i][1].ToString().Trim();
                    worksheet.Cells[i + 5, 2].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    worksheet.Cells[i + 5, 2].Style.Border.Top.Style = ExcelBorderStyle.Dotted;
                    worksheet.Cells[i + 5, 2].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                    workSheet.Cells[i + 5, 3].Value = da.Rows[i][2].ToString().Trim();
                    worksheet.Cells[i + 5, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    worksheet.Cells[i + 5, 3].Style.Border.Top.Style = ExcelBorderStyle.Dotted;
                    worksheet.Cells[i + 5, 3].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                    workSheet.Cells[i + 5, 4].Value = formatDate(da.Rows[i][3].ToString().Substring(0, 9)).Trim();
                    worksheet.Cells[i + 5, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    worksheet.Cells[i + 5, 4].Style.Border.Top.Style = ExcelBorderStyle.Dotted;
                    worksheet.Cells[i + 5, 4].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                    workSheet.Cells[i + 5, 5].Value = da.Rows[i][4].ToString().Trim();
                    worksheet.Cells[i + 5, 5].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    worksheet.Cells[i + 5, 5].Style.Border.Top.Style = ExcelBorderStyle.Dotted;
                    worksheet.Cells[i + 5, 5].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                    workSheet.Cells[i + 5, 6].Value = da.Rows[i][5].ToString().Trim();
                    worksheet.Cells[i + 5, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    worksheet.Cells[i + 5, 6].Style.Border.Top.Style = ExcelBorderStyle.Dotted;
                    worksheet.Cells[i + 5, 6].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                    workSheet.Cells[i + 5, 7].Value = da.Rows[i][6].ToString().Trim();
                    worksheet.Cells[i + 5, 7].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    worksheet.Cells[i + 5, 7].Style.Border.Top.Style = ExcelBorderStyle.Dotted;
                    worksheet.Cells[i + 5, 7].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                    workSheet.Cells[i + 5, 8].Value = da.Rows[i][7].ToString().Trim();
                    worksheet.Cells[i + 5, 8].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    worksheet.Cells[i + 5, 8].Style.Border.Top.Style = ExcelBorderStyle.Dotted;
                    worksheet.Cells[i + 5, 8].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                    workSheet.Cells[i + 5, 9].Value = da.Rows[i][8].ToString().Trim();
                    worksheet.Cells[i + 5, 9].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    worksheet.Cells[i + 5, 9].Style.Border.Top.Style = ExcelBorderStyle.Dotted;
                    worksheet.Cells[i + 5, 9].Style.Border.Right.Style = ExcelBorderStyle.Thin;

                    string str_so_luong = da.Rows[i][9].ToString().Trim();
                    if (str_so_luong.IndexOf(".") != -1)
                    {
                        workSheet.Cells[i + 5, 10].Value =  str_so_luong.Substring(0, str_so_luong.IndexOf(".")) + ".00";
                    }
                    else
                    {
                        workSheet.Cells[i + 5, 10].Value =  str_so_luong + ".00";
                    }
                    worksheet.Cells[i + 5, 10].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    worksheet.Cells[i + 5, 10].Style.Border.Top.Style = ExcelBorderStyle.Dotted;
                    worksheet.Cells[i + 5, 10].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                }

                workSheet.Protection.IsProtected = false;
                workSheet.Protection.AllowSelectLockedCells = false;
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                worksheet.Cells[worksheet.Dimension.Address].Style.Font.Name = "Times New Roman";
                worksheet.Cells[5, 1, worksheet.Dimension.End.Row, worksheet.Dimension.End.Column].Style.Font.Size = 10;
                worksheet.Cells[da.Rows.Count + 5, 1, da.Rows.Count + 5, 10].Style.Border.Top.Style = ExcelBorderStyle.Thin;

                if (System.IO.File.Exists(path))
                {
                    try
                    {
                        logbug("File da ton tai");
                    }
                    catch (System.IO.IOException e)
                    {
                        logbug(e.ToString());
                    }
                } else
                {
                    //Stream stream = File.Create(path);
                    //excelPackage.SaveAs(stream);
                    //stream.Close();
                    excelPackage.SaveAs(new FileInfo(path));
                }
            }
        }

        public string formatDate(string tt)
        {
            if ("".Equals(tt) || tt == null)
            {
                tt = "";
                return tt;
            }
            else
            {
                string[] sy = tt.Split('/');
                string fi = "";

                if (sy[1].Length == 1)
                {
                    fi = "0" + sy[1] + "/";
                }
                else
                {
                    fi = sy[1] + "/";
                }

                if (sy[0].Length == 1)
                {
                    fi = fi + "0" + sy[0] + "/";
                }
                else
                {
                    fi = fi + sy[0] + "/";
                }

                return fi + sy[2];
            }
        }
    }
}
