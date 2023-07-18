using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using ClosedXML.Excel;

namespace ProjectT_MIS
{
    class Program
    {
        public static string CompanyCode = "Teal_ALPLA";
        public static string PlantCode = "Teal_ALPLA01";
        public static string LineCode = "Proj_DPAL";

        static void Main(string[] args)
        {
   
            SqlConnectionStringBuilder builder1 = new SqlConnectionStringBuilder();

            ////----------Project-T server-----------////

            builder1.DataSource = "INPAS1VMS013";
            builder1.UserID = "sa";
            builder1.Password = "Sqldb@123";
            builder1.InitialCatalog = "Teal_Project_DPAL_N";


            //builder1.DataSource = @"DESKTOP-1S5C74D\MSSQLSERVER01";
            //builder1.InitialCatalog = "ProjectT_New";
            //builder1.Remove("UserID");
            //builder1.Remove("Password");
            //builder1.IntegratedSecurity = true;


            string connStrProjectT = builder1.ConnectionString;
            string connStr = builder1.ConnectionString;

            DateTime today = DateTime.Today.AddDays(-1);
            var dat = today.ToString("yyyy-MM-dd");
            DateTime nxttoday = today.AddDays(1);
            var nxtdat = nxttoday.ToString("yyyy-MM-dd");
            string holidayname = "";
            SqlConnection con = new SqlConnection(connStr);
            SqlCommand cmd = new SqlCommand("SELECT [HolidayReason],[Date] FROM [dbo].[tbl_Holiday] where Date= @date and CompanyCode='" + CompanyCode + "' and PlantID='" + PlantCode + "'", con);
            cmd.Parameters.AddWithValue("@date", dat);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);
            con.Open();



            var first = new DateTime(today.Year, today.Month, 1);
            var frstdaystr = first.ToString("yyyy-MM-dd");



            if (dat == frstdaystr)
            {
                SqlCommand cmd555 = new SqlCommand("truncate table [dbo].[DailyReport_Time];", con);
                cmd555.ExecuteNonQuery();
            }
            else
            {
                SqlCommand cmd555 = new SqlCommand("delete from [dbo].[DailyReport_Time] where Date=@date;", con);
                cmd555.Parameters.AddWithValue("@date", dat);
                cmd555.ExecuteNonQuery();
            }


            string AppLocation = "";
            AppLocation = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase);
            AppLocation = AppLocation.Replace("file:\\", "");
            string file = AppLocation + "\\ExcelFiles";


            DirectoryInfo d = new DirectoryInfo(file);//Assuming Test is your Folder
            FileInfo[] Files = d.GetFiles("*.xlsx"); //Getting Text files

            foreach (FileInfo file11 in d.EnumerateFiles())
            {
                if (file11 != null)
                {
                    file11.Delete();
                }

                Console.WriteLine("Old files deleted");
            }

            if (dt.Rows.Count != 0)
            {
                holidayname = dt.Rows[0][0].ToString();

                //string dts = DateTime.Today.AddDays(-1).ToString("dd-MM-yyyy");
                //string dte = DateTime.Today.ToString("dd-MM-yyyy");

                bool attachment;
                string messageBody = "<font> "+ PlantCode + "- "+LineCode + " MIS Report on :" + dat + " </font><br><br>";
                messageBody += "<p><b>Day Start Time: " + dat + " 08:15:00 AM</b></p><p><b>Day End Time  : " + nxtdat + " 08:15:00 AM</p></b>";
                DataSet edataset = getemailDataSet(connStr);
                DataSet ccdataset = getccDataSet(connStr);
                DataSet bccdataset = getbccDataSet(connStr);

                //default locale
                //System.DateTime.Now.DayOfWeek.ToString();
                //localized version
                //var day = System.DateTime.Now.ToString("dddd");

                attachment = false;
                messageBody = messageBody + "<font><b>--" + holidayname + "--<b></font><br><br>";
                messageBody = messageBody + "***Mail generated from TEAL IIOT Portal Email App Service***";
                Console.WriteLine("Created mail body with no data message");
                SendMail Sm = new SendMail();

                Sm.SendEmail(connStr, edataset, ccdataset, bccdataset, messageBody,today, LineCode);

            }
            else
            {

                SqlConnection conn = new SqlConnection(connStr);
                SqlCommand cmdd = new SqlCommand("select AssetID as Machine_code, AssetName as MachineName,f.FunctionID as Line_code" +
                    " from tbl_Assets a inner join tbl_function f on a.FunctionName=f.FunctionID ", conn);
                //cmd.Parameters.AddWithValue("@date", dat);
                SqlDataAdapter da1 = new SqlDataAdapter(cmdd);
                DataTable dt1 = new DataTable();
                da1.Fill(dt1);

                
                string[] name = new string[dt1.Rows.Count];
                string[] line = new string[dt1.Rows.Count];
                string[] machinecode = new string[dt1.Rows.Count];


                for (int i = 0; i < dt1.Rows.Count; i++)
                {

                    machinecode[i] = dt1.Rows[i][0].ToString();
                    name[i] = dt1.Rows[i][1].ToString();
                    line[i] = dt1.Rows[i][2].ToString();

                }


                //string dts = DateTime.Today.AddDays(-1).ToString("dd-MM-yyyy");
                //string dte = DateTime.Today.ToString("dd-MM-yyyy");
                bool attachment;
                string messageBody = "<font>Production count of <b> "+ PlantCode +" - " + LineCode + ":" + dat + " </font><br><br>";
                messageBody += "<p><b>Day Start Time: " + dat + " 08:15:00 AM</b></p><p><b>Day End Time  : " + nxtdat + " 08:15:00 AM</p></b>";
                DataSet edataset = getemailDataSet(connStr);
                DataSet ccdataset = getccDataSet(connStr);
                DataSet bccdataset = getbccDataSet(connStr);


                string htmlTableStart = "<table style=\"border-collapse:collapse; text-align:center;\" >";
                string htmlTableEnd = "</table>";
                string htmlHeaderRowStart = "<tr style =\"background-color:#6FA1D2; color:#ffffff;\">";
                string htmlHeaderRowEnd = "</tr>";
                string htmlTrStart = "<tr style =\"color:#555555;\">";
                string htmlTrEnd = "</tr>";
                string htmlTdStart = "<td style=\" border-color:#5c87b2; border-style:solid; border-width:thin; padding: 5px;\">";
                string htmlTdEnd = "</td>";

                messageBody += htmlTableStart;
                messageBody += htmlHeaderRowStart;
                messageBody += htmlTdStart + "Machine Name " + htmlTdEnd;
                messageBody += htmlTdStart + "Variant Name " + htmlTdEnd;
                //messageBody += htmlTdStart + "Target Production Qty " + htmlTdEnd;
                messageBody += htmlTdStart + "Actual Ok Parts " + htmlTdEnd;
                messageBody += htmlTdStart + "Actual NOk Parts " + htmlTdEnd;
                messageBody += htmlTdStart + " Rejection %" + htmlTdEnd;
                messageBody += htmlTdStart + "UpTime(in Mins / in %) " + htmlTdEnd;
                messageBody += htmlTdStart + "Downtime(in Mins / in %) " + htmlTdEnd;
                //messageBody += htmlTdStart + "IdleTime(in Mins) " + htmlTdEnd;
               
               // messageBody += htmlTdStart + "BreakTime(in Mins) " + htmlTdEnd;
               // messageBody += htmlTdStart + "LossTime(in Mins) " + htmlTdEnd;

                //messageBody += htmlTdStart + "UpTime(%) " + htmlTdEnd;
                //messageBody += htmlTdStart + "Downtime(%) " + htmlTdEnd;
                messageBody += htmlHeaderRowEnd;


                //default locale
                System.DateTime.Now.DayOfWeek.ToString();
                //localized version
                var day = today.ToString("dddd");

                string path = AppLocation + "\\Template\\" + "Template_3.xlsx";

                String MonthLabel = today.ToString("MMMM_yyyy");

                String dayy = today.ToString("dd");


                string filepath = AppLocation + "\\ExcelFiles\\" + MonthLabel + "_Day_" + dayy + "_" + LineCode + "_Report.xlsx";

                XLWorkbook workbook = new XLWorkbook(path);

                workbook.SaveAs(filepath);

                DataSet dataseteol = getEOL(connStr);


                for (int i = 0; i < dt1.Rows.Count; i++)
                {

                    Console.WriteLine(machinecode[i]+" "+ dataseteol.Tables[0].Rows[0][0].ToString());

                    if (dataseteol.Tables[0].Rows[0][0].ToString() != machinecode[i].ToString())
                    {
                        continue;
                    }

                    Console.WriteLine(machinecode[i]);
                    DataSet dataset = getDataSet(connStr, line[i], machinecode[i],today);
                   

                    Console.WriteLine("-------Machine Name: " + name[i] + " started-------");

                    //if (i == 0)
                    //{
                        LoopAllMachines GF1 = new LoopAllMachines();
                        //Class1 GF = new Class1();
                        GF1.getData(connStr, dataset, name[i], path, filepath,line[i]);
                    //}

                    bool noproduction = true;
                    //bool upt = true;

                    foreach (DataRow dr in dataset.Tables[1].Rows)
                    {
                        var a = dr["ActualProduction"].ToString();
                        var b = dr["UpTime(Min)"].ToString();

                        //if uptime 0 no production
                        if (b != "0")
                        {
                            noproduction = false;
                        }
                        //else
                        //{
                        //    production = true;
                        //}

                    }

                    

                    if (day == "Sunday")
                    {
                        if (dataset.Tables[1].Rows.Count != 0 && !noproduction)
                        {

                            //GetExcelFile GF = new GetExcelFile();
                            ////Class1 GF = new Class1();
                            //GF.getData(connStr, dataset, name[i],path,filepath);
                            //attachment = true;


                            foreach (DataRow Row in dataset.Tables[1].Rows)
                            {
                                messageBody = messageBody + htmlTrStart;
                                messageBody = messageBody + htmlTdStart + name[i] + htmlTdEnd;
                                messageBody = messageBody + htmlTdStart + Row["VariantCode"] + htmlTdEnd;
                                //messageBody = messageBody + htmlTdStart + Row["PlannedProduction"] + htmlTdEnd;
                                //messageBody = messageBody + htmlTdStart + "10000" + htmlTdEnd;
                                messageBody = messageBody + htmlTdStart + Row["ActualProduction"] + htmlTdEnd;

                                messageBody = messageBody + htmlTdStart + Row["actual_nok"] + htmlTdEnd;
                                if ((Convert.ToInt32(Row["ActualProduction"]) + Convert.ToInt32(Row["actual_nok"])) != 0)
                                {
                                    messageBody = messageBody + htmlTdStart + ((float)Math.Round((((float)(Convert.ToInt32(Row["actual_nok"])) / (float)(Convert.ToInt32(Row["ActualProduction"]) + Convert.ToInt32(Row["actual_nok"]))) * 100) * 100f) / 100f) + htmlTdEnd;
                                }
                                else
                                {
                                    messageBody = messageBody + htmlTdStart + "0" + htmlTdEnd;
                                }
                                messageBody = messageBody + htmlTdStart + Row["UpTime(Min)"] + " Mins / " + Row["UPTime%"] + " %" + htmlTdEnd;
                                messageBody = messageBody + htmlTdStart + Row["DownTime(Min)"] + " Mins / " + Row["DownTime%"] + " %" + htmlTdEnd;
                                //messageBody = messageBody + htmlTdStart + Row["IDLETime(Min)"] + htmlTdEnd;
                               
                                //messageBody = messageBody + htmlTdStart + Row["BreakTime(min)"] + " Mins / " + Row["BreakTime%"] + " %" + htmlTdEnd;

                                var a = ((Row["IDLETime(Min)"].ToString()));
                                var b = ((Row["LossTime(min)"].ToString()));

                                var cc = string.IsNullOrEmpty(a) ? "0" : a;
                                var dd = string.IsNullOrEmpty(b) ? "0" : b;

                                var aa = int.Parse(cc) + int.Parse(dd);

                                var a1 = ((Row["IDLETime%"].ToString()));
                                var b1 = ((Row["LossTime%"].ToString()));

                                var cc1 = string.IsNullOrEmpty(a) ? "0" : a1;
                                var dd1 = string.IsNullOrEmpty(b) ? "0" : b1;

                                var aa1 = float.Parse(cc1) + float.Parse(dd1);

                               // messageBody = messageBody + htmlTdStart + aa + " Mins / " + aa1 + " %" + htmlTdEnd;

                                //messageBody = messageBody + htmlTdStart + Row["UPTime%"] + htmlTdEnd;
                                //messageBody = messageBody + htmlTdStart + Row["DownTime%"] + htmlTdEnd;
                                messageBody = messageBody + htmlTrEnd;

                            }

                        }
                        else
                        {
                                attachment = false;
                            //messageBody = messageBody + "<font>---" + name[i] +"---Sunday--Holiday--</font><br><br>";
                            //foreach (DataRow Row in dataset.Tables[0].Rows)
                            //{
                                messageBody = messageBody + htmlTrStart;
                                messageBody = messageBody + htmlTdStart + name[i] + htmlTdEnd;
                                messageBody = messageBody + htmlTdStart + "<font> --Sunday--Holiday-- </font >" + htmlTdEnd;
                               // messageBody = messageBody + htmlTdStart + "-" + htmlTdEnd;
                                //messageBody = messageBody + htmlTdStart + "8000" + htmlTdEnd;
                                messageBody = messageBody + htmlTdStart + "-" + htmlTdEnd;
                                messageBody = messageBody + htmlTdStart + "-" + htmlTdEnd;
                                messageBody = messageBody + htmlTdStart + "-" + htmlTdEnd;
                                //messageBody = messageBody + htmlTdStart + "-" + htmlTdEnd;
                                messageBody = messageBody + htmlTdStart + "-" + htmlTdEnd;
                                messageBody = messageBody + htmlTdStart + "-" + htmlTdEnd;
                                //messageBody = messageBody + htmlTdStart + "-" + htmlTdEnd;
                                //messageBody = messageBody + htmlTdStart + "-" + htmlTdEnd;
                                messageBody = messageBody + htmlTrEnd;

                            //}
                            //messageBody = messageBody + "***Mail generated from TEAL IIOT Portal Email App Service***";
                            Console.WriteLine("Created mail body with no data message");
                            //SendMail Sm0 = new SendMail();
                            //Sm0.SendEmail(connStr, edataset, ccdataset, bccdataset, messageBody, dte, attachment);
                        }

                    }
                    else
                    {
                        if (dataset.Tables[1].Rows.Count == 0)
                        {
                            attachment = false;
                            //messageBody = messageBody + "<font>---" + name[i] +"---Data Not Logged--</font><br><br>";
                            ////foreach (DataRow Row in dataset.Tables[0].Rows)
                            ////{
                            messageBody = messageBody + htmlTrStart;
                            messageBody = messageBody + htmlTdStart + name[i] + htmlTdEnd;
                            messageBody = messageBody + htmlTdStart + "<font> --Data Not Logged-- </font >" + htmlTdEnd;
                            //messageBody = messageBody + htmlTdStart + "-" + htmlTdEnd;
                            //messageBody = messageBody + htmlTdStart + "8000" + htmlTdEnd;
                            messageBody = messageBody + htmlTdStart + "-" + htmlTdEnd;
                            messageBody = messageBody + htmlTdStart + "-" + htmlTdEnd;
                            messageBody = messageBody + htmlTdStart + "-" + htmlTdEnd;
                            messageBody = messageBody + htmlTdStart + "-" + htmlTdEnd;
                            messageBody = messageBody + htmlTdStart + "-" + htmlTdEnd;
                            //messageBody = messageBody + htmlTdStart + "-" + htmlTdEnd;
                            //messageBody = messageBody + htmlTdStart + "-" + htmlTdEnd;
                            //messageBody = messageBody + htmlTdStart + "-" + htmlTdEnd;
                            messageBody = messageBody + htmlTrEnd;

                            ////}
                            ////messageBody = messageBody + "<font>--Data Not Logged--</font><br><br>";
                            //// messageBody = messageBody + "***Mail generated from TEAL IIOT Portal Email App Service***";
                            Console.WriteLine("Created mail body with no data message");
                            ////SendMail Sm0 = new SendMail();
                            ////Sm0.SendEmail(connStr, edataset, ccdataset, bccdataset, messageBody, dts, attachment);

                 
                        }

                        else if (noproduction)
                        {
                            //GetExcelFile GF = new GetExcelFile();
                            //Class1 GF = new Class1();
                            //GF.getData(connStr, dataset);

                            attachment = false;
                            foreach (DataRow Row in dataset.Tables[1].Rows)
                            {
                                messageBody = messageBody + htmlTrStart;
                                messageBody = messageBody + htmlTdStart + name[i] + htmlTdEnd;
                                messageBody = messageBody + htmlTdStart + Row["VariantCode"] + htmlTdEnd;
                                //messageBody = messageBody + htmlTdStart + "-" + htmlTdEnd;
                                //messageBody = messageBody + htmlTdStart + "8000" + htmlTdEnd;
                                messageBody = messageBody + htmlTdStart + "<font> --No Production-- </font >" + htmlTdEnd;
                                messageBody = messageBody + htmlTdStart + "-" + htmlTdEnd;
                                messageBody = messageBody + htmlTdStart + "-" + htmlTdEnd;
                                //messageBody = messageBody + htmlTdStart + "-" + htmlTdEnd;
                                messageBody = messageBody + htmlTdStart + "-" + htmlTdEnd;
                                messageBody = messageBody + htmlTdStart + "-" + htmlTdEnd;
                                //messageBody = messageBody + htmlTdStart + "-" + htmlTdEnd;
                                //messageBody = messageBody + htmlTdStart + "-" + htmlTdEnd;
                                messageBody = messageBody + htmlTrEnd;

                            }
                            //messageBody = messageBody + "<font>--No Production--</font><br><br>";
                            //messageBody = messageBody + "***Mail generated from TEAL IIOT Portal Email App Service***";
                            Console.WriteLine("Created mail body with no data message");
                            //SendMail Sm0 = new SendMail();
                            //Sm0.SendEmail(connStr, edataset, ccdataset, bccdataset, messageBody, dts, attachment);
                        }
                        else
                        {
                            
                            foreach (DataRow Row in dataset.Tables[1].Rows)
                            {
                                messageBody = messageBody + htmlTrStart;
                                messageBody = messageBody + htmlTdStart + name[i] + htmlTdEnd;
                                messageBody = messageBody + htmlTdStart + Row["VariantCode"] + htmlTdEnd;
                               // messageBody = messageBody + htmlTdStart + Row["PlannedProduction"] + htmlTdEnd;
                                //messageBody = messageBody + htmlTdStart + "8000" + htmlTdEnd;
                                messageBody = messageBody + htmlTdStart + Row["ActualProduction"] + htmlTdEnd;
                                messageBody = messageBody + htmlTdStart + Row["actual_nok"] + htmlTdEnd;
                                if ((Convert.ToInt32(Row["ActualProduction"])+ Convert.ToInt32(Row["actual_nok"])) != 0)
                                {
                                    messageBody = messageBody + htmlTdStart + ((float)Math.Round((((float)(Convert.ToInt32(Row["actual_nok"])) / (float)(Convert.ToInt32(Row["ActualProduction"]) + Convert.ToInt32(Row["actual_nok"]))) * 100) * 100f) / 100f) + htmlTdEnd;
                                }
                                else 
                                { 
                                    messageBody = messageBody + htmlTdStart + "0" + htmlTdEnd; 
                                }
                               
                                messageBody = messageBody + htmlTdStart + Row["UpTime(Min)"] + " Mins / " + Row["UPTime%"] + " %" + htmlTdEnd;
                                messageBody = messageBody + htmlTdStart + Row["DownTime(Min)"] + " Mins / " + Row["DownTime%"] + " %" + htmlTdEnd;
                                //messageBody = messageBody + htmlTdStart + Row["IDLETime(Min)"] + htmlTdEnd;
                               
                                //messageBody = messageBody + htmlTdStart + Row["BreakTime(min)"] + " Mins / " + Row["BreakTime%"] + " %" + htmlTdEnd;

                                var a = ((Row["IDLETime(Min)"].ToString()));
                                var b = ((Row["LossTime(min)"].ToString()));

                                var cc = string.IsNullOrEmpty(a) ? "0" : a;
                                var dd = string.IsNullOrEmpty(b) ? "0" : b;


                                var aa = int.Parse(cc) + int.Parse(dd);

                                var a1 = ((Row["IDLETime%"].ToString()));
                                var b1 = ((Row["LossTime%"].ToString()));

                                var cc1 = string.IsNullOrEmpty(a) ? "0" : a1;
                                var dd1 = string.IsNullOrEmpty(b) ? "0" : b1;


                                var aa1 = float.Parse(cc1) + float.Parse(dd1);

                           //     messageBody = messageBody + htmlTdStart + aa + " Mins / " + aa1 + " %" + htmlTdEnd;
                                //messageBody = messageBody + htmlTdStart + Row["UPTime%"] + htmlTdEnd;
                                //messageBody = messageBody + htmlTdStart + Row["DownTime%"] + htmlTdEnd;

                            }
                            messageBody = messageBody + htmlTrEnd;

                        }
                    }
                    
                    GetExcelFile GF = new GetExcelFile();
                    //Class1 GF = new Class1();
                    GF.getData(connStr, dataset, name[i], path, filepath, line[i],CompanyCode,PlantCode,machinecode[i],dat,i);
                    attachment = true;


                    //string dateforAddProcessParameter = DateTime.Today.AddDays(-1).ToString("yyyy-MM-dd");

                   
                    AddProcessParameter GF2 = new AddProcessParameter();
                    //Class1 GF = new Class1();
                    GF2.getData(connStrProjectT, dataset, machinecode[i], path, filepath, line[i], today, CompanyCode, PlantCode);
                    attachment = true;


                    Console.WriteLine("*******Machine Name: " + name[i] + " ended*******");
                    Console.WriteLine("");

                }


                messageBody = messageBody + htmlTableEnd;
                // messageBody = messageBody + "<mark>Note: V0 - No Variant Selected.</mark><br/>";
                messageBody = messageBody + "<br/>For more details refer the attachment!!<br/>";
                SqlConnectionStringBuilder builder2 = new SqlConnectionStringBuilder();

                ////----------Project-T server-----------////
             
                builder2.DataSource = "INPAS1VMS013";
                builder2.UserID = "sa";
                builder2.Password = "Sqldb@123";
                builder2.InitialCatalog = "Master_DB";

                string MasterconnStr = builder2.ConnectionString;
                DataTable dt11 = new DataTable();
                SqlConnection masCon = new SqlConnection(MasterconnStr);
                SqlCommand cmd_mail1 = new SqlCommand("SELECT distinct [URL] FROM [dbo].[Portal_URL] where [CompanyCode]='"+ CompanyCode +"' and [PlantCode]='"+ PlantCode + "' ", masCon);
                cmd_mail1.Parameters.AddWithValue("@CompanyCode", CompanyCode);
                cmd_mail1.Parameters.AddWithValue("@PlantCode", PlantCode);
                SqlDataAdapter da11 = new SqlDataAdapter(cmd_mail1);
                da11.Fill(dt11);
                var s = dt11.Rows[0]["URL"].ToString();
                //var s = "https://teali4metricstest.azurewebsites.net/";
                messageBody = messageBody + "Refer the portal for more info " + "<a href='" + s + "'>click to login</a><br/>";
                messageBody = messageBody + "***Mail generated from TEAL IIOT Portal Email App service***";
                Console.WriteLine("Created mail body table");
                SendMail Sm = new SendMail();

                Sm.SendEmail(MasterconnStr, edataset, ccdataset, bccdataset, messageBody,today, LineCode);


            }


        }


        public static DataSet getDataSet(string connStr, string line, string machinecode,DateTime today)
        {
            DataSet ds = new DataSet();
            using (SqlConnection con = new SqlConnection(connStr))
            {
                try
                {
                    con.Open();
                    SqlCommand cmd = new SqlCommand("sp_Project_T_MIS_Report_Email", con);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 0;
                    cmd.Parameters.Add("@Machine_Code", SqlDbType.NVarChar, 150).Value = machinecode;
                    cmd.Parameters.Add("@CompanyCode", SqlDbType.NVarChar, 150).Value = CompanyCode;
                    cmd.Parameters.Add("@Line_code", SqlDbType.NVarChar, 150).Value = line;                      // added for enter the date manually
                    cmd.Parameters.Add("@Date", SqlDbType.NVarChar, 150).Value = today.ToString("yyyy-MM-dd");  // current date parameter
                    //cmd.Parameters.Add("@Date", SqlDbType.NVarChar, 150).Value = "2022-12-30";
                    cmd.Parameters.Add("@PlantCode", SqlDbType.NVarChar, 150).Value = PlantCode;
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(ds);
                    Console.WriteLine("Data required for mail body table has been collected");
                    return (ds);
                }
                catch (Exception ex)
                {
                    ExceptionSetting.SendErrorTomail(ex, connStr);
                    Console.WriteLine("Failed to collect data required for mail body table" + ex);
                    throw;
                }
                finally
                {
                    ds.Dispose();
                }
            }
        }

        public static DataSet getemailDataSet(string connStr)
        {
            DataSet ds = new DataSet();
            using (SqlConnection con = new SqlConnection(connStr))
            {
                try
                {
                    con.Open();
                    SqlCommand cmd = new SqlCommand("select * from tbl_emails " +
                        "where Companycode=@CompanyCode and Plantcode=@PlantCode and Status=@to and line_code=@linecode", con);
                    cmd.Parameters.Add("@to", SqlDbType.Int).Value = 1;
                    cmd.Parameters.Add("@CompanyCode", SqlDbType.VarChar).Value = CompanyCode;
                    cmd.Parameters.Add("@PlantCode", SqlDbType.VarChar).Value = PlantCode;
                    cmd.Parameters.Add("@linecode", SqlDbType.VarChar).Value = LineCode;
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(ds);
                    Console.WriteLine("Recipient data required to send e-mail has been collected ");
                    return (ds);
                }
                catch (Exception ex)
                {
                    ExceptionSetting.SendErrorTomail(ex, connStr);
                    Console.WriteLine("Failed to collect recipient data required to send e-mail " + ex);
                    throw;
                }
                finally
                {
                    ds.Dispose();
                }
            }
        }

        public static DataSet getccDataSet(string connStr)
        {
            DataSet ds = new DataSet();
            using (SqlConnection con = new SqlConnection(connStr))
            {
                try
                {
                    con.Open();
                    SqlCommand cmd = new SqlCommand("select * from tbl_emails" +
                        " where Companycode=@CompanyCode and Plantcode=@PlantCode and Status=@cc and line_code=@linecode", con);
                    cmd.Parameters.Add("@CompanyCode", SqlDbType.VarChar).Value = CompanyCode;
                    cmd.Parameters.Add("@PlantCode", SqlDbType.VarChar).Value = PlantCode;
                    cmd.Parameters.Add("@cc", SqlDbType.Int).Value = 0;
                    cmd.Parameters.Add("@linecode", SqlDbType.VarChar).Value = LineCode;
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(ds);
                    Console.WriteLine("CC data required to send e-mail has been collected ");
                    return (ds);
                }
                catch (Exception ex)
                {
                    ExceptionSetting.SendErrorTomail(ex, connStr);
                    Console.WriteLine("Failed to collect CC data required to send e-mail " + ex);
                    throw;
                }
                finally
                {
                    ds.Dispose();
                }
            }
        }


        public static DataSet getbccDataSet(string connStr)
        {
            DataSet ds = new DataSet();
            using (SqlConnection con = new SqlConnection(connStr))
            {
                try
                {
                    con.Open();
                    SqlCommand cmd = new SqlCommand("select * from tbl_emails" +
                        " where Companycode=@CompanyCode and Plantcode=@PlantCode and Status=@bcc and line_code=@linecode", con);
                    cmd.Parameters.Add("@CompanyCode", SqlDbType.VarChar).Value = CompanyCode;
                    cmd.Parameters.Add("@PlantCode", SqlDbType.VarChar).Value = PlantCode;
                    cmd.Parameters.Add("@bcc", SqlDbType.Int).Value = 2;
                    cmd.Parameters.Add("@linecode", SqlDbType.VarChar).Value = LineCode;
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(ds);
                    Console.WriteLine("BCC data required to send e-mail has been collected ");
                    return (ds);
                }
                catch (Exception ex)
                {
                    ExceptionSetting.SendErrorTomail(ex, connStr);
                    Console.WriteLine("Failed to collect CC data required to send e-mail " + ex);
                    throw;
                }
                finally
                {
                    ds.Dispose();
                }
            }
        }


        public static DataSet getEOL(string connStr)
        {
            DataSet ds = new DataSet();
            using (SqlConnection con = new SqlConnection(connStr))
            {
                try
                {
                    con.Open();
                    SqlCommand cmd = new SqlCommand("select AssetID from tbl_Assets where EOL='1' ", con);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(ds);
                    Console.WriteLine("EOL MachineID has been collected ");
                    return (ds);
                }
                catch (Exception ex)
                {
                    ExceptionSetting.SendErrorTomail(ex, connStr);
                    Console.WriteLine("Failed to collect EOL MachineID" + ex);
                    throw;
                }
                finally
                {
                    ds.Dispose();
                }
            }
        }


    }
}
