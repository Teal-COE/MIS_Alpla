using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Text;
using SQL = System.Data;
using System.Text.RegularExpressions;
using System.Net;

namespace ProjectT_MIS
{
    class AddProcessParameter
    {
        public void getData(string connStr, DataSet ds, string machinecode, string path, string filepath, string linecode, DateTime today, string CompanyCode, string PlantCode)
        {
            //DataSet ds = new DataSet();

            DataSet ds1 = new DataSet();
            //DataSet ds2 = new DataSet();

              //ProcessParameter(machinecode, connStr, path, filepath,"","", today);


            //using (SqlConnection con = new SqlConnection(connStr))
            //{
            //    try
            //    {
            //        con.Open();

            //        DataTable data = new DataTable();

            //        SqlCommand cmd = new SqlCommand("select distinct ParameterName from [dbo].[Tbl_Raw_Parameters] " +
            //            "where [Companycode] = @Company and [PlantCode] = @Plant and [Line_code] = @line and [Machine_code]=@machine and [Date]=@date", con);
            //        cmd.Parameters.AddWithValue("@date", today.ToString("yyyy-MM-dd"));
            //        cmd.Parameters.AddWithValue("@line", linecode);
            //        cmd.Parameters.AddWithValue("@machine", machinecode);
            //        cmd.Parameters.AddWithValue("@Company", CompanyCode);
            //        cmd.Parameters.AddWithValue("@Plant", PlantCode);
            //        cmd.ExecuteNonQuery();
            //        SqlDataAdapter da = new SqlDataAdapter(cmd);
            //        da.Fill(data);

            //        string col = "I";

            //        for (int i = 0; i < data.Rows.Count; i++)
            //        {
            //            //ProcessParameter(machinecode, connStr, path, filepath, data.Rows[i][0].ToString(),col,today);

            //            int add = sum(col);

            //            add += 10;

            //            col = calculation(add);
            //        }

            //        //ExportDataSetToExcel(ds);
            //        Console.WriteLine("Excel Chart has been generated");


            //    }
            //    catch (SqlException ex)
            //    {
            //        ExceptionSetting.SendErrorTomail(ex, connStr);
            //        Console.WriteLine("SQL Error: " + ex.Message);
            //    }
            //    catch (Exception e)
            //    {
            //        ExceptionSetting.SendErrorTomail(e, connStr);
            //        Console.WriteLine("Failed to generate Excel File" + e);
            //    }

            //}



        }

        
        public static void ProcessParameter(string machinecode, String connStr, String path, String filepath, string parameter,string col,DateTime today)
        {
           // string date = DateTime.Today.AddDays(-1).ToString("yyyy-MM-dd");

            DataTable data = new DataTable();

            SqlConnection con = new SqlConnection(connStr);
            SqlCommand cmd = new SqlCommand("SELECT top 500 [Shift_ID],[Variant_code],[Batch_no],[Part_Counter],[OK_Part],[NOK_Part],[TimeStamp]," +
                "[Min_Value],[Max_Value],[Part1],[Part1_status],[Part2],[Part2_status],[Part3],[Part3_status],[Part4],[Part4_status],[ParameterName]" +
                "[Line_code],[Machine_code],[Companycode],[PlantCode],[Date] " +
                "FROM[dbo].[Tbl_Raw_Parameters] " +
                "where [Companycode] = 'Teal_SLGN' and [PlantCode] = 'Teal_SLGN01' and [Line_code] = 'Proj_Trigger' and [Machine_code]=@machine " +
                " and [date] between @from and @to and ([Part1_status]=2 or [Part2_status]=2 or [Part3_status]=2 or [Part4_status]=2) " +
                "order by [Date] desc, [TimeStamp]", con);

            //DateTime today = DateTime.Today.AddDays(-3);
            string First = "2022-02-17" + " 06:00:00.000";         // added fromdate manually
            string second = "2022-02-18" + " 06:00:00.000";        // added Todate manually  

            string frst = today.ToString("yyyy-MM-dd") + " 06:00:00.000";
            string scnd = today.AddDays(1).ToString("yyyy-MM-dd") + " 06:00:00.000";

            cmd.Parameters.AddWithValue("@from", First);
            cmd.Parameters.AddWithValue("@to", second);
            //cmd.Parameters.AddWithValue("@Parameter", parameter);
            cmd.Parameters.AddWithValue("@machine", machinecode);
            
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(data);


            //Started reading the Excel file.  
            using (XLWorkbook workbook = new XLWorkbook(filepath))
            {

                ////stationwise process parameter sheet
                
                if (data.Rows.Count > 0)
                {

                    var machine = data.Rows[0][18].ToString();

                    int sheetno = 0;

                    if (machine == "M1")
                    {
                        sheetno = 8;

                    }
                    else if (machine == "M2")
                    {
                        sheetno = 13;

                    }
                    else if (machine == "M3")
                    {
                        sheetno = 18;

                    }
                    else if (machine == "M4")
                    {
                        sheetno = 23;

                    }
                    else
                    {
                        sheetno = 28;

                    }

                    IXLWorksheet ws8 = workbook.Worksheet(sheetno);


                    int aa2 = 6;
                    //int aa3 = 7;


                    // Adding DataRows.
                    for (int i = 0; i < data.Rows.Count; i++)
                    {

                        ws8.Cell("A" + (aa2)).Value = i+1;
                        ws8.Cell("B" + (aa2)).Value = data.Rows[i][0];
                        ws8.Cell("C" + (aa2)).Value = data.Rows[i][1];
                        ws8.Cell("D" + (aa2)).Value = data.Rows[i][2];
                        ws8.Cell("E" + (aa2)).Value = data.Rows[i][3];
                        ws8.Cell("F" + (aa2)).Value = data.Rows[i][4];
                        ws8.Cell("G" + (aa2)).Value = data.Rows[i][5];
                        ws8.Cell("H" + (aa2)).Value = data.Rows[i][6];
                        ws8.Cell("I" + (aa2)).Value = data.Rows[i][17];  // parameterName
                        ws8.Cell("J" + (aa2)).Value = data.Rows[i][7];   // min                    
                        ws8.Cell("K" + (aa2)).Value = data.Rows[i][8];    // max                
                        ws8.Cell("L" + (aa2)).Value = data.Rows[i][9];   // part 1 

                        ws8.Cell("M" + (aa2)).Value = checkstatus(data.Rows[i][10].ToString());  // part 1 status
                        if (checkstatus(data.Rows[i][10].ToString()) == "Fail")
                        {
                            ws8.Cell("M" + (aa2)).Style.Font.FontColor = XLColor.DarkRed;
                        }
                        ws8.Cell("N" + (aa2)).Value = data.Rows[i][11];  // part 2

                        ws8.Cell("O" + (aa2)).Value = checkstatus(data.Rows[i][12].ToString()); // part 2 status
                        if (checkstatus(data.Rows[i][12].ToString()) == "Fail")
                        {
                            ws8.Cell("O" + (aa2)).Style.Font.FontColor = XLColor.DarkRed;
                        }

                        ws8.Cell("P" + (aa2)).Value = data.Rows[i][13];     // part 3

                        ws8.Cell("Q" + (aa2)).Value = checkstatus(data.Rows[i][14].ToString()); // part 3 status
                        if (checkstatus(data.Rows[i][14].ToString()) == "Fail")
                        {
                            ws8.Cell("Q" + (aa2)).Style.Font.FontColor = XLColor.DarkRed;
                        }

                        ws8.Cell("R" + (aa2)).Value = data.Rows[i][15];     // part 4

                        ws8.Cell("S" + (aa2)).Value = checkstatus(data.Rows[i][16].ToString()); // part 4 status
                        if (checkstatus(data.Rows[i][16].ToString()) == "Fail")
                        {
                            ws8.Cell("S" + (aa2)).Style.Font.FontColor = XLColor.DarkRed;
                        }

                        aa2++;
                    }

                    //// Adding DataRows.
                    //for (int i = 0; i < data.Rows.Count; i++)
                    //{
                    //    ws8.Cell(columnname(col, 0) + (5)).Value = parameter;

                    //    ws8.Cell(columnname(col, 0) + (6)).Value = "Min";
                    //    ws8.Cell(columnname(col, 0) + (aa3)).Value = data.Rows[i][7];

                    //    ws8.Cell(columnname(col, 1) + (6)).Value = "Max";
                    //    ws8.Cell(columnname(col, 1) + (aa3)).Value = data.Rows[i][8];

                    //    ws8.Cell(columnname(col, 2) + (5)).Value = "Part - 01";
                    //    ws8.Cell(columnname(col, 2) + (aa3)).Value = data.Rows[i][9];

                    //    ws8.Cell(columnname(col, 3) + (5)).Value = "status";
                    //    ws8.Cell(columnname(col, 3) + (aa3)).Value = checkstatus(data.Rows[i][10].ToString());
                    //    if (checkstatus(data.Rows[i][10].ToString()) == "Fail")
                    //    {
                    //        ws8.Cell(columnname(col, 3) + (aa3)).Style.Font.FontColor = XLColor.DarkRed;
                    //    }

                    //    ws8.Cell(columnname(col, 4) + (5)).Value = "Part - 02";
                    //    ws8.Cell(columnname(col, 4) + (aa3)).Value = data.Rows[i][11];


                    //    ws8.Cell(columnname(col, 5) + (5)).Value = "status";
                    //    ws8.Cell(columnname(col, 5) + (aa3)).Value = checkstatus(data.Rows[i][12].ToString());
                    //    if (checkstatus(data.Rows[i][12].ToString()) == "Fail")
                    //    {
                    //        ws8.Cell(columnname(col, 5) + (aa3)).Style.Font.FontColor = XLColor.DarkRed;
                    //    }

                    //    ws8.Cell(columnname(col, 6) + (5)).Value = "Part - 03";
                    //    ws8.Cell(columnname(col, 6) + (aa3)).Value = data.Rows[i][13];

                    //    ws8.Cell(columnname(col, 7) + (5)).Value = "status";
                    //    ws8.Cell(columnname(col, 7) + (aa3)).Value = checkstatus(data.Rows[i][14].ToString());
                    //    if (checkstatus(data.Rows[i][14].ToString()) == "Fail")
                    //    {
                    //        ws8.Cell(columnname(col, 7) + (aa3)).Style.Font.FontColor = XLColor.DarkRed;
                    //    }

                    //    ws8.Cell(columnname(col, 8) + (5)).Value = "Part - 04";
                    //    ws8.Cell(columnname(col, 8) + (aa3)).Value = data.Rows[i][15];

                    //    ws8.Cell(columnname(col, 9) + (5)).Value = "status";
                    //    ws8.Cell(columnname(col, 9) + (aa3)).Value = checkstatus(data.Rows[i][16].ToString());
                    //    if (checkstatus(data.Rows[i][16].ToString()) == "Fail")
                    //    {
                    //        ws8.Cell(columnname(col, 9) + (aa3)).Style.Font.FontColor = XLColor.DarkRed;
                    //    }

                    //    //var aa = dt8.Rows[j][0].ToString();


                    //    //var dd = ws9.CellsUsed(cell => cell.GetString() == aa);


                    //    //////to find the respective cell in a n excel
                    //    ////var dd = ws4.Search(aa, CompareOptions.OrdinalIgnoreCase);

                    //    //search(dd, ws9, dt8.Rows[j][1].ToString(), columnNam);

                    //    aa3++;
                    //}


                    IXLCell firstcelldiag = ws8.Cell("A5");

                    IXLCell lastcelldiag = ws8.LastCellUsed();

                    // the range for which you want to add a table style
                    var rngTable1diag = ws8.Range(firstcelldiag, lastcelldiag);

                    rngTable1diag.Style.Border.BottomBorder = XLBorderStyleValues.Thin;

                    rngTable1diag.Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                    rngTable1diag.Style.Border.TopBorder = XLBorderStyleValues.Thin;

                    rngTable1diag.Style.Border.RightBorder = XLBorderStyleValues.Thin;

                    //rngTable1diag.Rows(1,2).Style.Fill.BackgroundColor = XLColor.LightGray;

                    
                }

                workbook.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                workbook.Style.Font.Bold = true;
                workbook.Save();
                //workbook.SaveAs(filepath);


            }
        }

        public static string checkstatus(string status)
        {
            string status1 = "";

            if (status == "2")
            {
                status1 = "Fail";
            }
            else
            {
                status1 = status;
            }
            
            return status1;
        }

        public static string columnname(string columnName,int add)
        {
            int sum = 0;


            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }

            return calculation(sum+add);
        }

        public static int sum(string columnName)
        {

            int sum = 0;

            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }

            return sum;
        }



        public static String calculation(int a)
        {
            int dividend = a;
            string columnNam = "";
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnNam = Convert.ToChar(65 + modulo).ToString() + columnNam;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnNam;

        }


        public static void search(IXLCells dd1, IXLWorksheet sheet, string valuetoenter, string columnNam44)
        {

            var a = "";

            foreach (IXLCell cell1 in dd1)
            {
                IXLCell coll = sheet.Column(1).LastCellUsed();

                var s = coll.Address.RowNumber.ToString();

                var df = cell1.Value.ToString();

                if (df == null || df == "")
                {

                    Console.WriteLine("not found so created and added");


                }
                else
                {
                    a = cell1.Address.RowNumber.ToString();

                    // var t = ds3.Tables[i].Rows[j][4].ToString();


                    sheet.Cell(columnNam44 + a).Value = valuetoenter;
                    sheet.Cell(columnNam44 + a).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                }


            }

        }



    }
}
