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
    class LoopAllMachines
    {

        public void getData(string connStr, DataSet ds, string machinecode, string path, string filepath, string linecode)
        {
            //DataSet ds = new DataSet();

            DataSet ds1 = new DataSet();
            //DataSet ds2 = new DataSet();

            using (SqlConnection con = new SqlConnection(connStr))
            {
                try
                {
                    con.Open();
                    
                    Console.WriteLine("Data required for excel has been collected ");

                    ///variant list of production qty variant-wise and day-wise - TABLE 6
                    ds1.Tables.Add(ds.Tables[6].Copy());

                    UploadExcelProduction(ds1, machinecode, connStr, path, filepath);

                    //ExportDataSetToExcel(ds);
                    Console.WriteLine("Excel Chart has been generated");


                }
                catch (SqlException ex)
                {
                    ExceptionSetting.SendErrorTomail(ex, connStr);
                    Console.WriteLine("SQL Error: " + ex.Message);
                }
                catch (Exception e)
                {
                    ExceptionSetting.SendErrorTomail(e, connStr);
                    Console.WriteLine("Failed to generate Excel File" + e);
                }

            }



        }

        
        public static void UploadExcelProduction(DataSet ds, string machinecode, String connStr, String path, String filepath)
        {

            string date = DateTime.Today.AddDays(-1).ToString("dd-MM-yyyy");
            string datetoaddinTime = DateTime.Today.AddDays(-1).ToString("yyyy-MM-dd");
            
            DataTable dt10 = new DataTable();

            dt10 = ds.Tables[0];


            //Started reading the Excel file.  
            using (XLWorkbook workbook = new XLWorkbook(filepath))
            {
                
                
                ////variant entering in plative production qty sheet
                IXLWorksheet ws8 = workbook.Worksheet(2);

                if (dt10.Rows.Count > 0)
                {
                    int aa2 = 6;
                    // Adding DataRows.
                    for (int i = 0; i < dt10.Rows.Count; i++)
                    {

                        ws8.Cell("A" + (aa2)).Value = dt10.Rows[i][2];

                        aa2++;
                    }
                }

                DateTime dat = DateTime.Today.AddDays(-1);

                var dates = new List<DateTime>();

                var firstDayOfMonth = new DateTime(dat.Year, dat.Month, 1);
                var lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);

                var NoOfMachine = 9;

                for (var dt33 = dat; dt33 >= firstDayOfMonth; dt33 = dt33.AddDays(-1))
                {
                    dates.Add(dt33);
                }

                //////DAY-WISE production qty
                //IXLWorksheet ws11 = workbook.Worksheet(3);

                //var datecolumnName = "H";

                //for (int i12 = 0; i12 < dates.Count; i12++)
                //{
                //    ws11.Cell(datecolumnName + "4").Value = dates[i12].ToString();
                //    ws11.Cell(datecolumnName + "4").Style.Font.Bold = true;
                //    ws11.Cell(datecolumnName + "4").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                //    int sum111 = sum(datecolumnName);

                //    datecolumnName = calculation(sum111 + 5);


                //}



                workbook.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                workbook.Style.Font.Bold = true;
                workbook.Save();
                //workbook.SaveAs(filepath);

            }

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
