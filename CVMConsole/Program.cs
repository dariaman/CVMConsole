using MySql.Data.MySqlClient;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Configuration;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;

namespace CVMConsole
{
    class Program
    {
        static string constring = ConfigurationManager.AppSettings["DefaultDB"];
        static string serverftp = ConfigurationManager.AppSettings["ServerFtp"];
        static string userftp = ConfigurationManager.AppSettings["UserFtp"];
        static string pwdftp = ConfigurationManager.AppSettings["PwdFtp"];
        static string filepathftp = ConfigurationManager.AppSettings["FilePathFtp"];
        static string Backup = ConfigurationManager.AppSettings["BackupLeads"];
        static string TempFile = ConfigurationManager.AppSettings["TempFile"];
        static string Log = ConfigurationManager.AppSettings["FilePathFtpLog"];

        static string EmailTo = ConfigurationManager.AppSettings["EmailTo"];
        static string EmailCC = ConfigurationManager.AppSettings["EmailCC"];
        static string EmailBCC = ConfigurationManager.AppSettings["EmailBCC"];
        static DateTime tgl = DateTime.Now.AddHours(-1);

        static void Main(string[] args)
        {
            string tgl = DateTime.Now.ToString("ddMMMyyyy_hhmmss");
            var Errorlogfile = TempFile + "GenerateFileError" + tgl + ".txt";

            var filePresales = "Presales_" + tgl + ".xls";
            var fileGeneral = "General_" + tgl + ".xls";
            /// Create File PreSales
            try
            {
                var jlh = GenerateExcel(TempFile + filePresales, "PresalesDataCreateFile");
                UploadFileToFtp(serverftp, TempFile + filePresales, userftp, pwdftp, filepathftp);
                UploadFileToFtp(serverftp, TempFile + filePresales, userftp, pwdftp, Backup);
            }
            catch (Exception ex)
            {
                var ts = new StreamWriter(Errorlogfile);
                ts.WriteLine("Error generate PreSales = " + ex.Message);
                ts.Close();
            }

            /// Create File General
            try
            {
                var jlh = GenerateExcel(TempFile + fileGeneral, "GeneralDataCreateFile");
                UploadFileToFtp(serverftp, TempFile + fileGeneral, userftp, pwdftp, filepathftp);
                UploadFileToFtp(serverftp, TempFile + fileGeneral, userftp, pwdftp, Backup);

                SendEmail(filePresales, fileGeneral);

                if (System.IO.File.Exists(TempFile + filePresales)) System.IO.File.Delete(TempFile + filePresales);
                if (System.IO.File.Exists(TempFile + fileGeneral)) System.IO.File.Delete(TempFile + fileGeneral);
            }
            catch (Exception ex)
            {
                var ts = new StreamWriter(Errorlogfile);
                ts.WriteLine("Error generate General = " + ex.Message);
                ts.Close();
            }
        }

        public static void UploadFileToFtp(string url, string filePath, string username, string password, string destination)
        {
            //Uri serverUri = new Uri(url);
            var fileName = Path.GetFileName(filePath);
            var request = (FtpWebRequest)WebRequest.Create("ftp://" + url + destination + fileName);

            request.Method = WebRequestMethods.Ftp.UploadFile;
            request.Credentials = new NetworkCredential(username, password);
            request.UsePassive = true;
            request.UseBinary = true;
            request.KeepAlive = false;

            using (var fileStream = File.OpenRead(filePath))
            {
                using (var requestStream = request.GetRequestStream())
                {
                    fileStream.CopyTo(requestStream);
                    requestStream.Close();
                }
                fileStream.Close();
            }

            var response = (FtpWebResponse)request.GetResponse();
            Console.WriteLine("Upload done: {0}", response.StatusDescription);
            response.Close();
        }

        public static int GenerateExcel(string filename, string spName)
        {
            HSSFWorkbook hssfwb = new HSSFWorkbook();
            MySqlConnection con = new MySqlConnection(constring);
            // connection must be openned for command
            con.Open();
            MySqlCommand cmd = new MySqlCommand(spName, con);
            cmd.CommandType = System.Data.CommandType.StoredProcedure;
            int i = 0;
            using (FileStream file = new FileStream(filename, FileMode.Create, FileAccess.Write))
            {
                try
                {
                    using (MySqlDataReader reader = cmd.ExecuteReader())
                    {
                        int jlhKolom = reader.FieldCount;
                        ISheet sheet = hssfwb.CreateSheet("sheet1");
                        IRow row = sheet.GetRow(0);
                        if (row == null) row = sheet.CreateRow(0);
                        for (int x = 0; x < jlhKolom; x++)
                        {
                            row.CreateCell(x).SetCellValue(reader.GetName(x));
                        }
                        i = 1;
                        while (reader.Read())
                        {
                            row = sheet.GetRow(i);
                            if (row == null) row = sheet.CreateRow(i);
                            for (int x = 0; x < jlhKolom; x++)
                            {
                                row.CreateCell(x).SetCellValue(reader[x].ToString());
                            }
                            i++;
                        }
                    }
                }
                catch (Exception ex)
                {
                    //throw new Exception(ex.Message);
                    var ts = new StreamWriter(TempFile + spName + tgl + ".txt");
                    ts.WriteLine("Fungsi GenerateExcel => " + cmd.CommandText + " \n" + ex.Message);
                    ts.Close();
                }
                finally
                {
                    hssfwb.Write(file);
                    file.Close();
                    con.Close();
                }
            }
            return i;
        }

        public static void SendEmail(string file1, string file2)
        {
            var presales = TempFile + file1;
            var general = TempFile + file2;

            var attachment1 = new Attachment(File.Open(presales, FileMode.Open), file1)
            {
                ContentType = new ContentType("application/vnd.ms-excel")
            };

            var attachment2 = new Attachment(File.Open(general, FileMode.Open), file2)
            {
                ContentType = new ContentType("application/vnd.ms-excel")
            };

            SmtpClient mailClient = new SmtpClient
            {
                Host = "mail.caf.co.id",
                UseDefaultCredentials = true,
                Port = 25
            };

            MailAddress from = new MailAddress("no-reply@jagadiri.co.id");
            MailMessage message = new MailMessage()
            {
                IsBodyHtml = true,
                Subject = "Data Leads",
                From = from,
                Body = "Terlampir Data Leads"
            };
            message.To.Add(EmailTo);
            message.CC.Add(EmailCC);
            message.Bcc.Add(EmailBCC);
            message.Attachments.Add(attachment1);
            message.Attachments.Add(attachment2);

            mailClient.Send(message);
            attachment1.Dispose();
            attachment2.Dispose();
        }
    }
}
