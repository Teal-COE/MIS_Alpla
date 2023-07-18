using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Net.Mail;
using System.Text;

namespace ProjectT_MIS
{
    class SendMail
    {

        public void SendEmail(String connStr, DataSet MailToset, DataSet CCset, DataSet BCCset, string htmlstring,DateTime today,string LineCode)
        {
            //string connectionstring = ConfigurationManager.ConnectionStrings["conn"].ToString();
            try
            {
                using (SqlConnection con = new SqlConnection(connStr))
                {
                    MailMessage mail = new MailMessage();
                    DataTable dt = new DataTable();
                    SqlCommand cmd_mail = new SqlCommand("SELECT * FROM tbl_gmail_settings", con);
                    SqlDataAdapter da = new SqlDataAdapter(cmd_mail);
                    da.Fill(dt);
                    SmtpClient smtp = new SmtpClient();
                    smtp.Host = dt.Rows[0]["Smtp_host"].ToString();
                    smtp.Port = Convert.ToInt32(dt.Rows[0]["Smtp_port"].ToString());
                    smtp.UseDefaultCredentials = true;
                    smtp.Credentials = new System.Net.NetworkCredential(dt.Rows[0]["Smtp_user"].ToString(), dt.Rows[0]["Smtp_pass"].ToString());
                    smtp.EnableSsl = false;
                    //foreach (DataRow Row in MailToset.Tables[0].Rows)
                    //{
                    //    string MailTo = Row["Email_ID"].ToString();
                    //    mail.To.Add(MailTo);                             
                    //}
                    //mail.To.Add("ravidasb@titan.co.in");
                    //mail.To.Add("sankararaman@titan.co.in");
                    //mail.To.Add("veeravalavan@titan.co.in");

                    //mail.To.Add("tamilmozhimj@titan.co.in");

                    //mail.To.Add("sanjaybaswa@titan.co.in");
                    //mail.To.Add("faj@titan.co.in");
                    //mail.To.Add("punyashree@titan.co.in");
                    //mail.To.Add("annefebronia@titan.co.in");

                    mail.To.Add("Naseruddin.Mohammed@alpla.com");
                    //mail.To.Add("M.Suraj @alpla.com");
                    

                    mail.From = new MailAddress(dt.Rows[0]["Smtp_user"].ToString());
                    string dts = today.ToString("dd-MM-yyyy");
                    mail.Subject = "Daily Production Summary of DPAL Assembly Report on " + dts + "";
                   

                    mail.Body = htmlstring;
                    mail.IsBodyHtml = true;

                    //List<string> li = new List<string>();
                    //foreach (DataRow Row in CCset.Tables[0].Rows)
                    //{
                    //    string cc = Row["Email_ID"].ToString();
                    //    li.Add(cc);
                    //}
                    //mail.CC.Add(string.Join<string>(",", li));        //--------------- Sending CC  

                    //List<string> bli = new List<string>();
                    //foreach (DataRow Row in BCCset.Tables[0].Rows)
                    //{
                    //    string cc = Row["Email_ID"].ToString();
                    //    bli.Add(cc);
                    //}

                    //foreach (string address in bli)
                    //{
                    //    MailAddress bcc = new MailAddress(address);
                    //    mail.Bcc.Add(bcc);                          //----------- Sending Bcc 
                    //}

                    String MonthLabel = today.ToString("MMMM_yyyy");

                    String day = today.ToString("dd");

                    string AppLocation = "";
                    AppLocation = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase);
                    AppLocation = AppLocation.Replace("file:\\", "");
                    string file = "";
                    file = AppLocation + "\\ExcelFiles\\" + MonthLabel + "_Day_" + day + "_"+ LineCode + "_Report.xlsx";

                    System.Net.Mail.Attachment attachment;
                    attachment = new System.Net.Mail.Attachment(file); //Attaching File to Mail  
                    mail.Attachments.Add(attachment);
                    Console.WriteLine("Attachment appended");

                    smtp.Send(mail);
                    Console.WriteLine("Email sent successfully");

                    //Console.ReadLine();
                    //Console.ReadLine();
                }
            }
            catch (Exception ex)
            {
                ExceptionSetting.SendErrorTomail(ex, connStr);
                Console.WriteLine("Failed to send email" + ex);
            }
        }


    }
}
