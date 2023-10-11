using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    class Program
    {
        [Obsolete]
        static void Main(string[] args)
        {
            DataTable dt = DT();
            string remoteFileName = ConfigurationManager.AppSettings["InputFile"];
            string FullPathFile = Convert.ToString(AppDomain.CurrentDomain.BaseDirectory).Replace("bin\\Debug\\", "") + "Excel\\" + remoteFileName;

            try
            {

                XSSFWorkbook hssfwb;
                using (FileStream file = new FileStream(FullPathFile, FileMode.Open, FileAccess.Read))
                {
                    hssfwb = new XSSFWorkbook(file);
                }
                int Count = hssfwb.NumberOfSheets;

                for (int i = 0; i < Count; i++)
                {
                    ISheet sheet = hssfwb.GetSheetAt(i);
                    for (int row = 3; row <= sheet.LastRowNum; row++)
                    {
                        if (sheet.GetRow(row) != null) //null is when the row only contains empty cells 
                        {
                            
                                if (Convert.ToString(sheet.GetRow(2).GetCell(0))== "Approve")
                                {
                                if (!string.IsNullOrEmpty(Convert.ToString(sheet.GetRow(row).GetCell(1))))
                                {
                                    DataRow _new = dt.NewRow();
                                    _new["Name"] = Convert.ToString(sheet.GetRow(0).GetCell(12));
                                    _new["Date"] = ADDzero(Convert.ToString(sheet.GetRow(row).GetCell(1)));
                                    _new["Start Time"] = Convert.ToString(sheet.GetRow(row).GetCell(7));
                                    _new["Finish Time"] = Convert.ToString(sheet.GetRow(row).GetCell(8));
                                    _new["Working Hours"] = Convert.ToString(sheet.GetRow(row).GetCell(9));
                                    _new["Work Status"] = Convert.ToString(sheet.GetRow(row).GetCell(3));
                                    _new["Daily Approver"] = Convert.ToString(sheet.GetRow(row).GetCell(55));
                                    dt.Rows.Add(_new);
                                }
                                else
                                {
                                    break;
                                }
                            }
                                else
                                {
                                if (!string.IsNullOrEmpty(Convert.ToString(sheet.GetRow(row).GetCell(0))))
                                {
                                    DataRow _new = dt.NewRow();
                                    _new["Name"] = Convert.ToString(sheet.GetRow(0).GetCell(12));
                                    _new["Date"] = ADDzero(Convert.ToString(sheet.GetRow(row).GetCell(0)));
                                    _new["Start Time"] = Convert.ToString(sheet.GetRow(row).GetCell(6));
                                    _new["Finish Time"] = Convert.ToString(sheet.GetRow(row).GetCell(7));
                                    _new["Working Hours"] = Convert.ToString(sheet.GetRow(row).GetCell(8));
                                    _new["Work Status"] = Convert.ToString(sheet.GetRow(row).GetCell(2));
                                    _new["Daily Approver"] = Convert.ToString(sheet.GetRow(row).GetCell(54));
                                    dt.Rows.Add(_new);
                                }
                                else
                                {
                                    break;
                                }
                            }
                                
                            
                        }
                    }
                }

                WriteExcelWithNPOI(dt, ConfigurationManager.AppSettings["OutputFile"]);
                Console.WriteLine("File Generated Successfully");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error " + ex.Message);
                Console.ReadLine();
            }
        }

        [Obsolete]
        public static void WriteExcelWithNPOI(DataTable dt, string FileName)
        {
            // dll refered NPOI.dll and NPOI.OOXML  

            IWorkbook workbook;
            workbook = new XSSFWorkbook();
            ISheet sheet1 = workbook.CreateSheet("Sheet1");
            var font = workbook.CreateFont();
            font.FontHeightInPoints = (short)11;
            font.FontName = "Calibri";
            font.Color = IndexedColors.Black.Index;
            //cell value
            var fontCell = workbook.CreateFont();
            fontCell.FontHeightInPoints = (short)9;
            fontCell.FontName = "CorpoS";
            var fontStatus = workbook.CreateFont();
            fontStatus.FontHeightInPoints = (short)9;
            fontStatus.FontName = "CorpoS";
            fontStatus.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;
            XSSFCellStyle yourStyleHead = (XSSFCellStyle)workbook.CreateCellStyle();
            yourStyleHead.WrapText = true;
            yourStyleHead.Alignment = HorizontalAlignment.Center;
            yourStyleHead.VerticalAlignment = VerticalAlignment.Center;
            XSSFCellStyle yourStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            yourStyle.WrapText = true;
            yourStyle.Alignment = HorizontalAlignment.Center;
            yourStyle.VerticalAlignment = VerticalAlignment.Center;
            yourStyle.BorderBottom = BorderStyle.Thin;
            yourStyle.BorderRight = BorderStyle.Thin;
            yourStyle.BorderTop = BorderStyle.Thin;
            XSSFCellStyle yourCellTitle = (XSSFCellStyle)workbook.CreateCellStyle();
            yourCellTitle.WrapText = true;
            XSSFCellStyle yourCellDesc = (XSSFCellStyle)workbook.CreateCellStyle();
            yourCellDesc.WrapText = true;
            XSSFCellStyle DateCellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            DateCellStyle.WrapText = true;
            DateCellStyle.Alignment = HorizontalAlignment.Left;
            DateCellStyle.VerticalAlignment = VerticalAlignment.Center;
            XSSFCellStyle yourCellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            yourCellStyle.WrapText = true;
            yourCellStyle.VerticalAlignment = VerticalAlignment.Center;
            XSSFCellStyle SnoCellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            SnoCellStyle.WrapText = true;
            SnoCellStyle.Alignment = HorizontalAlignment.Center;
            SnoCellStyle.VerticalAlignment = VerticalAlignment.Center;
            var Header = workbook.CreateFont();
            Header.FontHeightInPoints = (short)14;
            Header.FontName = "Calibri";
            Header.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;

            XSSFCellStyle yourStyle1 = (XSSFCellStyle)workbook.CreateCellStyle();
            yourStyle1.BorderLeft = NPOI.SS.UserModel.BorderStyle.Medium;
            yourStyle1.BorderTop = NPOI.SS.UserModel.BorderStyle.Medium;
            yourStyle1.BorderRight = NPOI.SS.UserModel.BorderStyle.Medium;
            yourStyle1.BorderBottom = NPOI.SS.UserModel.BorderStyle.Medium;

            ICellStyle testeStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            testeStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Medium;
            testeStyle.FillForegroundColor = IndexedColors.Red.Index;
            testeStyle.FillPattern = FillPattern.SolidForeground;

            //make a header row  
            IRow row1 = sheet1.CreateRow(0);
            NPOI.HSSF.UserModel.HSSFWorkbook wob = new NPOI.HSSF.UserModel.HSSFWorkbook();

            for (int j = 0; j < dt.Columns.Count; j++)
            {
                ICell cell = row1.CreateCell(j);
                cell.CellStyle = SnoCellStyle;
                cell.CellStyle.SetFont(fontCell);
                String columnName = dt.Columns[j].ToString().Replace('_', ' ');
                cell.SetCellValue(columnName);
                cell.CellStyle = workbook.CreateCellStyle();
                cell.CellStyle = yourStyle;
                cell.CellStyle = yourStyle1;
                cell.CellStyle.SetFont(font);
                cell.CellStyle = workbook.CreateCellStyle();
                cell.CellStyle = yourStyle;
                cell.CellStyle.SetFont(font);
                cell.CellStyle.FillForegroundColor = IndexedColors.Grey25Percent.Index;
                cell.CellStyle.FillPattern = FillPattern.SolidForeground;
                sheet1.SetColumnWidth(j, 5000);
            }
            ICellStyle style = workbook.CreateCellStyle();

            IDataFormat dataFormatCustom = workbook.CreateDataFormat();
            //loops through data  
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                IRow row = sheet1.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ICell cell1 = row.CreateCell(j);
                    String columnName = dt.Columns[j].ToString();
                    cell1.SetCellValue(dt.Rows[i][columnName].ToString());
                    cell1.CellStyle = yourStyle1;
                    

                    if (columnName == "Working Hours")
                    {
                        float working = float.Parse(dt.Rows[i]["Working Hours"].ToString().Replace(":", "."));

                        if (working > 10)
                        {

                            cell1.CellStyle = testeStyle;
                            cell1.CellStyle.DataFormat = dataFormatCustom.GetFormat("HH:mm");
                        }
                       
                    }
                    else if(columnName == "Start Time" || columnName == "Finish Time" )
                    {
                        cell1.CellStyle.DataFormat = dataFormatCustom.GetFormat("HH:mm");

                    }


                }
            }
            string FullPathFile = Convert.ToString(AppDomain.CurrentDomain.BaseDirectory).Replace("bin\\Debug\\", "") + "Excel\\" + FileName;
            using (var exportData = new MemoryStream())
            {
                FileStream xfile = new FileStream(FullPathFile, FileMode.Create, System.IO.FileAccess.Write);
                workbook.Write(xfile);
                xfile.Close();
            }
        }

        public static string ADDzero(string text)
        {
            string output = string.Empty;
            if (!string.IsNullOrEmpty(text)) { 
            string value = text.Substring(1, 1);
                if (value == "(")
                {
                    output = "0" + text;
                }
                else
                {
                    output = text;
                }
            }
            return output;


        }
        public static DataTable DT()
        {
            DataTable dt = new DataTable();
            dt.Clear();
            dt.Columns.Add("Name");
            dt.Columns.Add("Date");
            dt.Columns.Add("Start Time");
            dt.Columns.Add("Finish Time");
            dt.Columns.Add("Working Hours");
            dt.Columns.Add("Work Status");
            dt.Columns.Add("Daily Approver");
            return dt;
        }
    }
}
