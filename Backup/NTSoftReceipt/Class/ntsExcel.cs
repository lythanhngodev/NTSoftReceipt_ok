using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.IO;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using System.Drawing.Printing;
using System.Net;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

using System.Globalization;
using ClosedXML.Excel;

namespace NTSoftReceipt.Class
{
    public class ntsExcel
    {
        public ntsExcel()
        {

        }
        public Excel.Application xlApplication;
        public Excel.Workbook xlWorkbook;
        public Excel.Worksheet xlWorksheet;
        public Excel.Range xlRange;
        public object misValue = System.Reflection.Missing.Value;
        Excel.XlFixedFormatType paramExportFormat = Excel.XlFixedFormatType.xlTypePDF;
        Excel.XlFixedFormatQuality paramExportQuality = Excel.XlFixedFormatQuality.xlQualityStandard;
        bool paramOpenAfterPublish = false;
        bool paramIncludeDocProps = true;
        bool paramIgnorePrintAreas = true;
        object FixedFormatExtClassPtr = System.Reflection.Missing.Value;
        object paramFromPage = Type.Missing;
        object paramToPage = Type.Missing;

        #region Thứ tự cột của Excel bắt đầu từ 1
        public static string[] ColumsCell =
            new string[] { 
            "", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
            "AA","AB","AC","AD","AE","AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ",
            "BA","BB","BC","BD","BE","BF","BG","BH","BI","BJ","BK","BL","BM","BN","BO","BP","BQ","BR","BS","BT","BU","BV","BW","BX","BY","BZ",
            "CA","CB","CC","CD","CE","CF","CG","CH","CI","CJ","CK","CL","CM","CN","CO","CP","CQ","CR","CS","CT","CU","CV","CW","CX","CY","CZ",
            "DA","DB","DC","DD","DE","DF","DG","DH","DI","DJ","DK","DL","DM","DN","DO","DP","DQ","DR","DS","DT","DU","DV","DW","DX","DY","DZ",
            "EA","EB","EC","ED","EE","EF","EG","EH","EI","EJ","EK","EL","EM","EN","EO","EP","EQ","ER","ES","ET","EU","EV","EW","EX","EY","EZ",
            "FA","FB","FC","FD","FE","FF","FG","FH","FI","FJ","FK","FL","FM","FN","FO","FP","FQ","FR","FS","FT","FU","FV","FW","FX","FY","FZ",
            "GA","GB","GC","GD","GE","GF","GG","GH","GI","GJ","GK","GL","GM","GN","GO","GP","GQ","GR","GS","GT","GU","GV","GW","GX","GY","GZ",
            "HA","HB","HC","HD","HE","HF","HG","HH","HI","HJ","HK","HL","HM","HN","HO","HP","HQ","HR","HS","HT","HU","HV","HW","HX","HY","HZ"
        };
        #endregion

        private int ColumsMin = -1, ColumsMax = -1, RowsMin = -1, RowsMax = -1, CountMinMax = -1;
        private void GetColumRowMinMax(object cell1, object cell2)
        {
            try
            {
                for (int i = ColumsCell.Length - 1; i > 0; i--)
                {
                    try
                    {
                        if (cell1.ToString().IndexOf(ColumsCell[i]) != -1)
                        {
                            ColumsMin = i;
                            RowsMin = Convert.ToInt32(cell1.ToString().Replace(ColumsCell[i], ""));
                            break;
                        }
                    }
                    catch
                    { }
                }
                for (int i = ColumsCell.Length - 1; i > 0; i--)
                {
                    try
                    {
                        if (cell2.ToString().IndexOf(ColumsCell[i]) != -1)
                        {
                            ColumsMax = i;
                            RowsMax = Convert.ToInt32(cell2.ToString().Replace(ColumsCell[i], ""));
                            break;
                        }
                    }
                    catch
                    { }
                }
                CountMinMax = ((ColumsMax - ColumsMin) + 1) * ((RowsMax - ColumsMin) + 1);
            }
            catch
            { }
        }

        public void OpendFileXLS(string fileOpen, bool visible)
        {
            try
            {
                xlApplication = new Excel.Application();
                xlApplication.Visible = visible;
                System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
                //
                xlWorkbook = xlApplication.Workbooks.Open(fileOpen, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                //xlWorkbook = xlApplication.Workbooks.Open(FileOpen, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);
            }
            catch
            { }
        }

        public void SetActiveWorksheet(int index)
        {
            if (index > 0)
            {
                xlWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkbook.Worksheets.get_Item(index);
                ((Microsoft.Office.Interop.Excel._Worksheet)xlWorksheet).Activate();
            }
        }
        public void CloseFileXLS(bool fileClose)
        {
            try
            {
                if (xlWorkbook != null)
                {
                    xlWorkbook.Close(fileClose, misValue, misValue);
                    xlApplication.Workbooks.Close();
                    xlApplication.Quit();
                    //Finally destroy all the objects.
                    if (xlWorksheet != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet);
                    if (xlWorkbook != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);
                    if (xlApplication != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApplication);

                    xlApplication = null; xlWorkbook = null; xlWorksheet = null;

                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    //System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("Excel");
                    //foreach (System.Diagnostics.Process p in process)
                    //{
                    //    if (!string.IsNullOrEmpty(p.ProcessName))// && p.StartTime.AddSeconds(+5) > DateTime.Now)
                    //    {
                    //        try
                    //        {
                    //            p.Kill();
                    //        }
                    //        catch { }
                    //    }
                    //}
                }
            }
            catch
            {
                xlWorkbook.Close(fileClose, misValue, misValue);
                xlApplication.Workbooks.Close();
                xlApplication.Quit();
                //Finally destroy all the objects.
                if (xlWorksheet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet);
                if (xlWorkbook != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);
                if (xlApplication != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApplication);

                xlApplication = null; xlWorkbook = null; xlWorksheet = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public void SaveFileXLS(string fileSave, bool checkDelete, bool closeFile)
        {
            try
            {
                if (checkDelete)
                    if (System.IO.File.Exists(fileSave))
                        System.IO.File.Delete(fileSave);
                if (xlWorkbook != null)
                {
                    //if (fileSave.ToLower().IndexOf(".xls") != -1)
                    //    xlWorkbook.SaveAs(fileSave, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, System.Reflection.Missing.Value, Missing.Value, false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlUserResolution, true, Missing.Value, Missing.Value, Missing.Value);
                    //else
                    //    xlWorkbook.SaveAs(fileSave, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkbook.SaveAs(fileSave, Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel8,
                    Type.Missing, Type.Missing, false, false,
                    Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing);
                }
                xlWorkbook.Saved = true;
                if (closeFile) //Đóng tập tin lại
                    CloseFileXLS(false);
            }
            catch
            {
                xlWorkbook.Saved = true;
                if (closeFile)
                    CloseFileXLS(false);
            }
        }

        public void SaveFileXlsToPdf(string pathFilePdf, string fileExcel, bool closeFile)
        {
            try
            {
                //OpendFileXLS(fileExcel, false);

                if (xlWorkbook != null)
                {
                    if (fileExcel.ToLower().IndexOf(".xlsx") != -1)
                        xlWorkbook.ExportAsFixedFormat(paramExportFormat, pathFilePdf, paramExportQuality, paramIncludeDocProps, paramIgnorePrintAreas, paramFromPage, paramToPage, paramOpenAfterPublish, misValue);
                    else
                        //xlWorkbook.SaveAs(fileExcel, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                        xlWorkbook.ExportAsFixedFormat(paramExportFormat, pathFilePdf, paramExportQuality, paramIncludeDocProps, paramIgnorePrintAreas, paramFromPage, paramToPage, paramOpenAfterPublish, misValue);
                }
                xlWorkbook.Saved = true;
                if (closeFile) //Đóng tập tin lại
                    CloseFileXLS(false);
            }
            catch
            {
                xlWorkbook.Saved = true;
                if (closeFile)
                    CloseFileXLS(false);
            }
        }


        public void SaveFileHTML(string fileSave, bool checkDelete, bool closeFile)
        {
            try
            {
                if (checkDelete)
                {
                    try
                    {
                        FileInfo file = new FileInfo(fileSave);
                        if (file.Exists)
                            file.Delete();
                    }
                    catch
                    { }
                }

                if (xlWorkbook != null)
                    xlWorkbook.SaveAs(fileSave, Microsoft.Office.Interop.Excel.XlFileFormat.xlHtml, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, misValue, misValue, misValue, misValue, misValue);
                if (closeFile) //Đóng tập tin lại
                    CloseFileXLS(false);
            }
            catch
            {
                if (closeFile)
                    CloseFileXLS(false);
            }
        }
        public void InsertRows(object cell1, object cell2)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Range r = xlWorksheet.get_Range(cell1, cell2);
                r.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown, Missing.Value);
            }
            catch
            { }
        }
        public void DeleteRows(object cell1, object cell2)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Range r = xlWorksheet.get_Range(cell1, cell2);
                r.EntireRow.Delete(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftDown);
            }
            catch
            { }
        }
        public void MergeCells(int CountColumn, object cell1, object cell2)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Range r = xlWorksheet.get_Range(cell1, cell2);
                r.Merge(CountColumn);
            }
            catch
            { }
        }
        public void MergeCells(object cell1, object cell2)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Range r = xlWorksheet.get_Range(cell1, cell2);
                r.Merge(false);
            }
            catch
            { }
        }
        public void UnMergeCells(object cell1, object cell2)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Range r = xlWorksheet.get_Range(cell1, cell2);
                r.UnMerge();
            }
            catch
            { }
        }
        public void SetFontColor(System.Drawing.Color colorCell, object cell1, object cell2)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Range r = xlWorksheet.get_Range(cell1, cell2);
                r.Font.Color = colorCell.ToArgb();
            }
            catch
            { }
        }
        /// <summary>
        /// Format Cells Border xlMedium
        /// </summary>
        public void FormatBorderMedium(object cell1, object cell2)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Range r = xlWorksheet.get_Range(cell1, cell2);
                r.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
            }
            catch
            { }
        }

        public void InteriorColor(System.Drawing.Color colorCell, object cell1, object cell2)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Range r = xlWorksheet.get_Range(cell1, cell2);
                r.Interior.Color = colorCell.ToArgb();
            }
            catch
            { }
        }
        public void SetFontName(string fontName, object cell1, object cell2)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Range r = xlWorksheet.get_Range(cell1, cell2);
                r.Font.Name = fontName;
            }
            catch
            { }
        }
        public void SetFontUnderline(bool value, object cell1, object cell2)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Range r = xlWorksheet.get_Range(cell1, cell2);
                r.Font.Underline = value;
            }
            catch
            { }
        }
        public void SetFontItalic(bool value, object cell1, object cell2)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Range r = xlWorksheet.get_Range(cell1, cell2);
                r.Font.Italic = value;
            }
            catch
            { }
        }
        public void SetFontBold(bool value, object cell1, object cell2)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Range r = xlWorksheet.get_Range(cell1, cell2);
                r.Font.Bold = value;
            }
            catch
            { }
        }
        public void WrapText(bool value, object cell1, object cell2)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Range r = xlWorksheet.get_Range(cell1, cell2);
                r.WrapText = value;
            }
            catch
            { }
        }
        public void SetBordersColor(System.Drawing.Color colorCell, object cell1, object cell2)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Range r = xlWorksheet.get_Range(cell1, cell2);
                r.Borders.Color = colorCell.ToArgb();
            }
            catch
            { }
        }
        public void SetColumnsWidth(int width, object cell1, object cell2)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Range r = xlWorksheet.get_Range(cell1, cell2);
                r.ColumnWidth = width;
            }
            catch
            { }
        }
        public void SetRowsHeight(int height, object cell1, object cell2)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Range r = xlWorksheet.get_Range(cell1, cell2);
                r.RowHeight = height;
            }
            catch
            { }
        }

        public void SetHorizontalAlignment(xLHAlign xLHAlign, object cell1, object cell2)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Range r = xlWorksheet.get_Range(cell1, cell2);
                if (xLHAlign == xLHAlign.xlHAlignCenter)
                    r.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                if (xLHAlign == xLHAlign.xlHAlignCenterAcrossSelection)
                    r.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenterAcrossSelection;
                if (xLHAlign == xLHAlign.xlHAlignDistributed)
                    r.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignDistributed;
                if (xLHAlign == xLHAlign.xlHAlignFill)
                    r.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignFill;
                if (xLHAlign == xLHAlign.xlHAlignGeneral)
                    r.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignGeneral;
                if (xLHAlign == xLHAlign.xlHAlignJustify)
                    r.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignJustify;
                if (xLHAlign == xLHAlign.xlHAlignLeft)
                    r.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                if (xLHAlign == xLHAlign.xlHAlignRight)
                    r.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;
            }
            catch
            { }
        }
        public enum xLHAlign
        {
            xlHAlignCenter,
            xlHAlignCenterAcrossSelection,
            xlHAlignDistributed,
            xlHAlignFill,
            xlHAlignGeneral,
            xlHAlignJustify,
            xlHAlignLeft,
            xlHAlignRight
        }
        public void SetVerticalAlignment(xlVAlign xlVAlign, object cell1, object cell2)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Range r = xlWorksheet.get_Range(cell1, cell2);
                if (xlVAlign == xlVAlign.xlVAlignBottom)
                    r.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignBottom;
                if (xlVAlign == xlVAlign.xlVAlignCenter)
                    r.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
                if (xlVAlign == xlVAlign.xlVAlignDistributed)
                    r.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignDistributed;
                if (xlVAlign == xlVAlign.xlVAlignJustify)
                    r.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignJustify;
                if (xlVAlign == xlVAlign.xlVAlignTop)
                    r.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignTop;
            }
            catch
            { }
        }
        public enum xlVAlign
        {
            xlVAlignBottom,
            xlVAlignCenter,
            xlVAlignDistributed,
            xlVAlignJustify,
            xlVAlignTop
        }

        public void WriteValues(object value, object cell1, object cell2)
        {
            try
            {
                xlRange = xlWorksheet.get_Range(cell1, cell2);
                xlRange.Value2 = value;

                //GetColumRowMinMax(cell1, cell2); //Lấy Dòng và cột min max.
                //for (int i = RowsMin; i <= RowsMax; i++)
                //{
                //    for (int j = ColumsMin; j <= ColumsMax; j++)
                //    {
                //        xlApplication.Cells[i, j] = value;
                //        xlRange = xlWorksheet.get_Range(cell1, cell2);
                //        xlRange.Value2 = value;
                //    }
                //}
            }
            catch
            { }
        }
        public DataTable GetValue(object cell1, object cell2)
        {
            DataTable value = null;
            try
            {
                GetColumRowMinMax(cell1, cell2); //Lấy Dòng và cột min max.
                value = GetDaTaColum((ColumsMax - ColumsMin) + 1);
                int Index1 = 0;
                object[] colum = new object[(ColumsMax - ColumsMin) + 1];
                for (int i = RowsMin; i <= RowsMax; i++)
                {
                    Index1 = 0;
                    for (int j = ColumsMin; j <= ColumsMax; j++)
                    {
                        Microsoft.Office.Interop.Excel.Range r = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[i, j];
                        colum[Index1] = r.Cells.Text;
                        Index1 += 1;
                    }
                    value.Rows.Add(colum);
                }
            }
            catch
            { }
            return value;
        }
        private DataTable GetDaTaColum(int columCount)
        {
            DataTable DataTable = new DataTable();
            if (columCount > 0)
            {
                for (int i = 0; i < columCount; i++)
                {
                    DataTable.Columns.Add("Column" + (i + 1).ToString(), typeof(string));
                }
            }
            return DataTable;
        }
        private DataTable GetDaTaColumPicture(int columCount)
        {
            DataTable DataTable = new DataTable();
            if (columCount > 0)
            {
                for (int i = 0; i < columCount; i++)
                {
                    DataTable.Columns.Add("Column" + (i + 1).ToString(), typeof(Image));
                }
            }
            return DataTable;
        }
        public void PrintOut()
        {
            try
            {
                //Object From, Object To, Object Copies, Object Preview, Object ActivePrinter, Object PrintToFile, Object Collate, Object PrToFileName
                xlWorkbook.PrintOut(Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            }
            catch
            { }
        }
        public void CopyPaseCells(object copycell1, object copycell2, object pasecell1, object pasecell2)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Range sourceRange = xlWorksheet.get_Range(copycell1, copycell2);
                sourceRange.Copy(Type.Missing);
                Microsoft.Office.Interop.Excel.Range destinationRange = xlWorksheet.get_Range(pasecell1, pasecell2);
                destinationRange.PasteSpecial(Microsoft.Office.Interop.Excel.XlPasteType.xlPasteAll, Microsoft.Office.Interop.Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            }
            catch
            { }
        }
        public void AddNewSheet(string sheetName, bool before, int index)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Sheets xlWorksheet1 = xlWorkbook.Sheets as Microsoft.Office.Interop.Excel.Sheets;
                if (before)
                    xlWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorksheet1.Add(xlWorksheet1[index], Type.Missing, Type.Missing, Type.Missing);
                else
                    xlWorksheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorksheet1.Add(Type.Missing, xlWorksheet1[index], Type.Missing, Type.Missing);
                if (sheetName.Trim() != "")
                    xlWorksheet.Name = sheetName;
            }
            catch
            { }
        }
        public void AddCopyNewSheet(string sheetName, bool before, int index)
        {
            try
            {
                SetActiveWorksheet(index);
                int indexNew = index;
                Microsoft.Office.Interop.Excel.Sheets xlWorksheet1 = xlWorkbook.Sheets as Microsoft.Office.Interop.Excel.Sheets;
                if (before)
                    xlWorksheet.Copy(xlWorksheet1[index], Type.Missing);
                else
                {
                    xlWorksheet.Copy(Type.Missing, xlWorksheet1[index]);
                    indexNew = index + 1;
                }
                SetActiveWorksheet(indexNew);
                if (sheetName.Trim() != "")
                    xlWorksheet.Name = sheetName;
            }
            catch
            { }
        }
        public bool PrinterAllFile(string pathToPdf)
        {
            try
            {
                Process process = new Process();
                string printerName = GetDefaultPrinter();
                process.StartInfo.FileName = pathToPdf;
                process.StartInfo.Verb = "printto";
                process.StartInfo.Arguments = "\"" + printerName + "\"";
                process.Start();

                process.WaitForInputIdle();
                process.Kill();
                return true;
            }
            catch
            {
                return false;
            }
        }
        private string GetDefaultPrinter()
        {
            PrinterSettings settings = new PrinterSettings();
            foreach (string printer in PrinterSettings.InstalledPrinters)
            {
                settings.PrinterName = printer;
                if (settings.IsDefaultPrinter)
                    return printer;
            }
            return string.Empty;
        }
        #region In dữ liệu HTML.
        public void _mPrinterFileHTML(string htmlFileName)
        {

        }
        #endregion
        /// <summary>
        /// 
        /// </summary>
        /// <param name="pathOld">Đường dẫn vật lý đến ổ đĩa. Ví dụ: C:\ABC.xls</param>
        /// <param name="urlFolderNew">Đường dẫn vật lý tới thư mục cần copy file vào. Ví dụ: D:\newfolder</param>
        /// <param name="fileNameNew">Tên file mới có phần mở rộng. Ví dụ ABC.xls</param>
        /// <param name="deleteAllFileInFolderNew">True nếu muốn xóa hết file trong thư mục mới</param>
        public void CopyFiles(string pathFileOld, string urlFolderNew, string fileNameNew, bool deleteAllFileInFolderNew)
        {
            string urlFolder = urlFolderNew;
            if (!System.IO.Directory.Exists(urlFolder))
            {
                System.IO.Directory.CreateDirectory(urlFolder);
            }
            if (deleteAllFileInFolderNew == true)
            {
                DirectoryInfo di = new DirectoryInfo(urlFolder);
                FileInfo[] rgFiles = di.GetFiles();
                foreach (FileInfo fi in rgFiles)
                {
                    fi.Delete();
                }
            }
            System.IO.File.Copy(pathFileOld, urlFolderNew + "/" + fileNameNew);
        }
        /// <summary>
        /// Chuyển số kiểu Interger sang kiểu số la mã
        /// </summary>
        /// <param name="num">Tham số kiểu Integer</param>
        /// <returns>string</returns>
        public string IntToRoman(int num)
        {
            string[] thou = { "", "M", "MM", "MMM" };
            string[] hun = { "", "C", "CC", "CCC", "CD", "D", "DC", "DCC", "DCCC", "CM" };
            string[] ten = { "", "X", "XX", "XXX", "XL", "L", "LX", "LXX", "LXXX", "XC" };
            string[] ones = { "", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX" };
            string roman = "";
            roman += thou[(int)(num / 1000) % 10];
            roman += hun[(int)(num / 100) % 10];
            roman += ten[(int)(num / 10) % 10];
            roman += ones[num % 10];

            return roman;
        }
        /// <summary>
        /// Chuyển kiểu số sang kiểu chuỗi
        /// </summary>
        /// <param name="num">Kiểu số Integer</param>
        /// <param name="toUpper">True nếu muốn trả về in hoa. False nếu muốn trả về in thường</param>
        /// <returns>string</returns>
        public string IntToAlphabet(int num, bool toUpper)
        {
            if (toUpper == true)
            {
                return ((char)(num + 64)).ToString();
            }
            else
            {
                return ((char)(num + 96)).ToString();
            }
        }

        /// <summary>
        /// Nối datatable
        /// </summary>
        /// <param name="dt">Danh sách table</param>
        /// <returns></returns>
        public DataTable JoinDataTable(object[] dt)
        {
            DataTable result = new DataTable();
            for (int i = 0; i < dt.Length; i++)
            {
                if (i == 0)
                {
                    result = ((DataTable)dt[i]).Copy();
                }
                else
                {
                    if (((DataTable)dt[i]).Rows.Count > 0)
                    {

                        for (int j = 0; j < ((DataTable)dt[i]).Rows.Count; j++)
                        {
                            for (int k = 0; k < ((DataTable)dt[i]).Columns.Count; k++)
                            {
                                result.Rows.Add();
                                result.Rows[result.Rows.Count - 1][j] = ((DataTable)dt[i]).Rows[j][k];
                            }
                        }
                    }
                }
            }

            return result;
        }

        public void exPortExcel1Group(ref XLWorkbook wb, DataTable tab, bool xuatGroup, int kieuXuat, string[] cotXuat, ref int dongXuat, string rangeStart, string rangeEnd, bool sum, bool sumTop, bool xuatSTT, int kieuXuatSTT, string[] cotTong)
        {
            ntsExcel ntsExcel = new ntsExcel();
            //bien su dung
            int stt = 0;
            var ws = wb.Worksheet(1);
            string nhom1 = "";
            int sttNhom = 0;
            int index = 0;
            //tong dau
            if (sum == true && sumTop == true)
            {
                ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat).InsertRowsBelow(1);
                ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat).CopyTo(ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat));
                ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat).Style.Font.SetBold(true).Alignment.SetWrapText(true).Border.SetTopBorder(XLBorderStyleValues.Thin).Border.SetBottomBorder(XLBorderStyleValues.Thin).Border.SetRightBorder(XLBorderStyleValues.Thin).Border.SetLeftBorder(XLBorderStyleValues.Thin);

                ws.Cell(cotTong[0] + dongXuat).Value = cotTong[1];
                for (int j = 2; j < cotTong.Length; j++)
                {
                    // ws.Cell(cotXuat[int.Parse(cotTong[j])] + dongXuat).Value = tongCot(tab, int.Parse(cotTong[j]), xuatSTT);
                    index = int.Parse(cotTong[j].ToString());
                    ws.Cell(cotXuat[int.Parse(cotTong[j])] + dongXuat).Value = tab.Compute("Sum(" + tab.Columns[index].ColumnName + ")", "");// tongCot(tab, int.Parse(cotTong[j]), false);
                }
                dongXuat += 1;
            }
            for (int i = 0; i < tab.Rows.Count; i++)
            {
                if (xuatGroup == false)
                {//xuat khong co nhom
                    ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat).InsertRowsBelow(1);
                    ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat).CopyTo(ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat));
                    ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat).Style.Alignment.SetWrapText(true).Font.SetBold(false).Border.SetTopBorder(XLBorderStyleValues.Dotted).Border.SetBottomBorder(XLBorderStyleValues.Dotted);
                    if (xuatSTT == true)
                    {
                        stt += 1;
                        ws.Cell(cotXuat[0] + dongXuat).Value = stt.ToString();
                        for (int j = 1; j < cotXuat.Length; j++)
                        {
                            ws.Cell(cotXuat[j] + dongXuat).Value = tab.Rows[i][j - 1].ToString();
                        }
                    }
                    else
                    {
                        for (int j = 0; j < cotXuat.Length; j++)
                        {
                            ws.Cell(cotXuat[j] + dongXuat).Value = tab.Rows[i][j].ToString();
                        }
                    }
                }
                //xuat du lieu co nhom
                else
                {
                    if (xuatSTT == true)
                    {//xuat dong group
                        if (nhom1 != tab.Rows[i][0].ToString())
                        {
                            ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat).InsertRowsBelow(1);
                            ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat).CopyTo(ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat));
                            ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat).Style.Font.SetBold(true).Alignment.SetWrapText(true).Border.SetTopBorder(XLBorderStyleValues.Thin).Border.SetBottomBorder(XLBorderStyleValues.Thin).Border.SetRightBorder(XLBorderStyleValues.Thin).Border.SetLeftBorder(XLBorderStyleValues.Thin);
                            nhom1 = tab.Rows[i][0].ToString();
                            sttNhom += 1; stt = 0;
                            if (kieuXuat == 1) ws.Cell(cotXuat[0] + dongXuat).Value = sttNhom.ToString();
                            else
                                if (kieuXuat == 2) ws.Cell(cotXuat[0] + dongXuat).Value = ntsExcel.IntToRoman(sttNhom).ToString();
                                else
                                    if (kieuXuat == 3) ws.Cell(cotXuat[0] + dongXuat).Value = ntsExcel.IntToAlphabet(sttNhom, false).ToString();
                                    else
                                        if (kieuXuat == 4) ws.Cell(cotXuat[0] + dongXuat).Value = ntsExcel.IntToAlphabet(sttNhom, true).ToString();
                            ws.Cell(cotXuat[1] + dongXuat).Value = nhom1;
                            index = 0;
                            for (int j = 2; j < cotTong.Length; j++)
                            {
                                index = int.Parse(cotTong[j].ToString());
                                ws.Cell(cotXuat[int.Parse(cotTong[j])] + dongXuat).Value = tab.Compute("Sum(" + tab.Columns[index].ColumnName + ")", "tenNhom='" + nhom1 + "'");// tongCot(tab, int.Parse(cotTong[j]), false);
                            }

                            dongXuat += 1;
                        }
                        //xuat stt
                        ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat).InsertRowsBelow(1);
                        ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat).CopyTo(ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat));
                        ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat).Style.Alignment.SetWrapText(true).Border.SetTopBorder(XLBorderStyleValues.Dashed).Border.SetBottomBorder(XLBorderStyleValues.Dashed).Border.SetRightBorder(XLBorderStyleValues.Thin).Border.SetLeftBorder(XLBorderStyleValues.Thin);

                        stt += 1;
                        if (kieuXuatSTT == 1) ws.Cell(cotXuat[0] + dongXuat).Value = stt.ToString();
                        else
                            if (kieuXuatSTT == 2) ws.Cell(cotXuat[0] + dongXuat).Value = "'" + sttNhom.ToString() + "." + stt.ToString();
                        for (int j = 1; j < cotXuat.Length - 1; j++)
                        {
                            ws.Cell(cotXuat[j] + dongXuat).Value = tab.Rows[i][j].ToString();
                        }
                    }
                    else
                    {
                        if (nhom1 != tab.Rows[i][0].ToString())
                        {
                            ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat).InsertRowsBelow(1);
                            ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat).CopyTo(ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat));
                            ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat).Style.Alignment.WrapText = true;
                            ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat).Style.Border.TopBorder = XLBorderStyleValues.Thin;

                            nhom1 = tab.Rows[i][0].ToString();
                            ws.Cell(cotXuat[1] + dongXuat).Value = nhom1;
                            dongXuat += 1;
                        }
                        //xuat stt
                        ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat).InsertRowsBelow(1);
                        ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat).CopyTo(ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat));
                        ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat).Style.Font.Bold = false;
                        ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat).Style.Alignment.WrapText = true;
                        ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat).Style.Border.TopBorder = XLBorderStyleValues.Dotted;
                        for (int j = 2; j < cotXuat.Length; j++)
                        {
                            ws.Cell(cotXuat[j] + dongXuat).Value = tab.Rows[i][j].ToString();
                        }
                    }
                }
                dongXuat += 1;
            }
            //tong chan
            if (sum == true && sumTop == false)
            {
                ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat).InsertRowsBelow(1);
                ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat).CopyTo(ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat));
                ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat).Style.Alignment.WrapText = true;
                ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat).Style.Font.Bold = true;
                ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                ws.Cell(cotTong[0] + dongXuat).Value = cotTong[1];
                index = 0;
                for (int j = 2; j < cotTong.Length; j++)
                {
                    index = int.Parse(cotTong[j].ToString());
                    ws.Cell(cotXuat[int.Parse(cotTong[j])] + dongXuat).Value = tab.Compute("Sum(" + tab.Columns[index].ColumnName + ")", "");// tongCot(tab, int.Parse(cotTong[j]), false);

                }
                dongXuat += 1;
            }
            ws.Range(rangeStart + dongXuat + ":" + rangeEnd + dongXuat).Delete(XLShiftDeletedCells.ShiftCellsUp);
            ws.Range(rangeStart + (dongXuat - 1) + ":" + rangeEnd + (dongXuat - 1)).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
        }


        public void xuatKyTen(ref XLWorkbook wb, string[] mKyTen, int dongXuat, bool hienNgayBC, string ngayIn, string diaDanh)
        {
            var ws = wb.Worksheet(1);
            ws.Cell(mKyTen[0].Split('-')[0].ToString() + dongXuat).Style.Font.SetItalic(true).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right);
            ws.Range(mKyTen[0].Split('-')[0].ToString() + dongXuat, mKyTen[0].Split('-')[1].ToString() + dongXuat).Merge();
            if (hienNgayBC == true)
            {
                ws.Cell(mKyTen[0].Split('-')[0].ToString() + dongXuat).Value = diaDanh
                + ", ngày " + ngayIn.Substring(0, 2)
                + " tháng " + ngayIn.Substring(3, 2)
                + " năm " + ngayIn.Substring(6, 4);
            }
            else
            {
                ws.Cell(mKyTen[0].Split('-')[0].ToString() + dongXuat).Value = "......., ngày.....tháng....năm.....";
            }
            for (int i = 1; i < mKyTen.Length; i++)
            {
                dongXuat += 1;
                ws.Row(dongXuat).Style.Font.SetBold(true).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center).Alignment.SetWrapText(true);
                ws.Range(mKyTen[i].Split('-')[0].ToString() + dongXuat, mKyTen[i].Split('-')[1].ToString() + dongXuat).Merge();
                ws.Cell(mKyTen[i].Split('-')[0].ToString() + dongXuat).Value = mKyTen[i].Split('-')[2].ToString();
                dongXuat += 1;
                ws.Row(dongXuat).Style.Font.SetItalic(true).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center).Alignment.SetWrapText(true);
                ws.Range(mKyTen[i].Split('-')[0].ToString() + dongXuat, mKyTen[i].Split('-')[1].ToString() + dongXuat).Merge();
                ws.Cell(mKyTen[i].Split('-')[0].ToString() + dongXuat).Value = mKyTen[i].Split('-')[3].ToString();
                dongXuat += 4;
                ws.Row(dongXuat).Style.Font.SetBold(true).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center).Alignment.SetWrapText(true);
                ws.Range(mKyTen[i].Split('-')[0].ToString() + dongXuat, mKyTen[i].Split('-')[1].ToString() + dongXuat).Merge();
                ws.Cell(mKyTen[i].Split('-')[0].ToString() + dongXuat).Value = mKyTen[i].Split('-')[4].ToString();

                dongXuat -= 6;
            }
        }

        public int WriteExcel(XLWorkbook wb, DataTable result, int startRows, int startColumns, bool checkInsert)
        {
            int rowsEpx = startRows;
            try
            {
                var ws = wb.Worksheet("Sheet1");
                if (result.Rows.Count > 0)
                {
                    for (int i = 0; i < result.Rows.Count; i++)
                    {
                        if (checkInsert)
                        {
                            ws.Range(rowsEpx, 1, rowsEpx, result.Columns.Count).InsertRowsBelow(1);
                        }
                        for (int j = startColumns; j < result.Columns.Count - 3; j++)//excel tính từ dòng 1 cột 1
                        {
                            ws.Cell(rowsEpx, j + 1).Value = result.Rows[i][j].ToString();
                        }
                        if (result.Rows[i][result.Columns.Count - 3].ToString() == "True")
                        {
                            ws.Range(rowsEpx, 1, rowsEpx, result.Columns.Count - 3).Style.Font.Bold = true;
                        }
                        if (result.Rows[i][result.Columns.Count - 2].ToString() == "True")
                        {
                            ws.Range(rowsEpx, 1, rowsEpx, result.Columns.Count - 3).Style.Font.Italic = true;
                        }
                        if (result.Rows[i][result.Columns.Count - 1].ToString() == "True")
                        {
                            if (result.Rows[i - 1][result.Columns.Count - 1].ToString() == "False")
                            {
                                ws.Range(rowsEpx, 1, rowsEpx, result.Columns.Count - 3).Style.Border.BottomBorder = XLBorderStyleValues.Dotted;
                            }
                            else
                            {
                                ws.Range(rowsEpx, 1, rowsEpx, result.Columns.Count - 3).Style.Border.TopBorder = XLBorderStyleValues.Dotted;
                            }
                        }
                        rowsEpx += 1;
                        //if (i == result.Rows.Count - 1)
                        //{
                        //    ws.Range(i, 1, rowsEpx, result.Columns.Count - 3).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                        //    ws.Range(rowsEpx + 1, 1, rowsEpx + 1, result.Columns.Count - 3).Delete(XLShiftDeletedCells.ShiftCellsUp);
                        //}
                    }
                }
                return rowsEpx;
            }
            catch
            {
                return rowsEpx;
            }
        }
        /// <summary>
        /// Xuất dữ liệu ra excel
        /// </summary>
        /// <param name="wb">Wordkbook cần của file excel</param>
        /// <param name="startRows">Dòng bắt đầu ghi dữ liệu trên file excel (Tính từ 1)</param>
        /// <param name="startColumns">Cột bắt đầu ghi dữ liệu ra file excel (Tính từ 1)</param>
        /// <param name="checkInsert">Trước khi ghi có insert dòng trống không</param>
        /// <param name="dt">Datatable truyền vào</param>
        /// <param name="colSumStart">Cột bắt đầu sum trên datatable (Tính từ 0)</param>
        /// <param name="colSumEnd">Cột kết thúc sum</param>
        /// <param name="idGroup1">STT group 1</param>
        /// <param name="idGroup2">STT group 2</param>
        /// <param name="idGroup3">STT group 3</param>
        /// <param name="idGroup4">STT group 4</param>
        /// <param name="idGroup5">STT group 5</param>
        /// <param name="checkSum">Xuất tổng cộng dòng đầu</param>
        /// <param name="labelSum">Chuỗi dòng tổng cộng</param>
        /// <returns>object[]</returns>
        public object[] WriteExcel5Group(XLWorkbook wb, int startRows, int startColumns, bool checkInsert, DataTable dt, int colSumStart, int colSumEnd, int idGroup1, int idGroup2, int idGroup3, int idGroup4, int idGroup5, bool checkSum, string labelSum)
        {
            object[] arr = new object[2];
            int rowsEpx = startRows;
            int colEpx = startColumns;
            var ws = wb.Worksheet("Sheet1");
            try
            {
                string group1 = "", group2 = "", group3 = "", group4 = "", group5 = "";
                int colContent = 6;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (group1 == dt.Rows[i][1].ToString())//cot 1 chứa nhóm 1
                    {
                        if (group2 == dt.Rows[i][2].ToString())//cot 2 chứa nhóm 2
                        {
                            if (group3 == dt.Rows[i][3].ToString())//cot 3 chứa nhóm 3
                            {
                                if (group4 == dt.Rows[i][4].ToString())//cot 4 chứa nhóm 4
                                {
                                    if (dt.Rows[i][5].ToString() != "")//kiểm tra nếu nhóm 5!=""
                                    {
                                        //    //Them dòng tên dự án
                                        group5 = dt.Rows[i][5].ToString();
                                        idGroup5 += 1;
                                        if (checkInsert)
                                        {
                                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 5).InsertRowsBelow(1);
                                        }
                                        ws.Cell(rowsEpx, colEpx).Value = IntToAlphabet(idGroup5, false);
                                        colEpx += 1;
                                        ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][5];
                                        for (int c = colContent + 1; c <= colSumEnd; c++)
                                        {
                                            if (c < colSumStart)
                                            {
                                                colEpx += 1;
                                                ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][c];//xuất nội dung của các cột kiểu text sau cột nội dung (sau cột 5) cho đến cột kiểu số
                                            }
                                            else if (c > colSumStart || c <= colSumEnd)
                                            {
                                                colEpx += 1;
                                                ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "' AND " + dt.Columns[4].ColumnName + "='" + group4 + "' AND " + dt.Columns[5].ColumnName + "='" + group5 + "'");
                                            }
                                        }
                                        ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 5).Style.Border.BottomBorder = XLBorderStyleValues.Dashed;
                                        rowsEpx += 1;
                                        colEpx = startColumns;
                                        //Kiểm tra xem dự án có 1 hay 2 dòng
                                        if (dt.Rows[i][4].ToString() == dt.Rows[((i + 1) < dt.Rows.Count ? i + 1 : i)][4].ToString() && dt.Rows[i][5].ToString() == dt.Rows[((i + 1) < dt.Rows.Count ? i + 1 : i)][5].ToString() && dt.Rows.Count > 1 && dt.Rows[i][6].ToString() != "" && i < dt.Rows.Count - 1) //tên dự án dòng trên trùng tên dự án dòng dưới
                                        {
                                            //xuat dong nguồn trong nước
                                            if (checkInsert)
                                            {
                                                ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 5).InsertRowsBelow(1);
                                            }
                                            colEpx += 1;
                                            ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][6];
                                            for (int c = colSumStart; c <= colSumEnd; c++)
                                            {
                                                colEpx = c - colContent + 2;
                                                ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][c];
                                            }
                                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 5).Style.Border.TopBorder = XLBorderStyleValues.Dashed;
                                            rowsEpx += 1;
                                            colEpx = startColumns;
                                            //xuất tiếp nguồn ngoài nước
                                            if (checkInsert)
                                            {
                                                ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 5).InsertRowsBelow(1);
                                            }
                                            colEpx += 1;
                                            ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i + 1][6];
                                            for (int c = colSumStart; c <= colSumEnd; c++)
                                            {
                                                colEpx = c - colContent + 2;
                                                ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i + 1][c];
                                            }
                                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 5).Style.Border.TopBorder = XLBorderStyleValues.Dashed;
                                            rowsEpx += 1;
                                            colEpx = startColumns;
                                            i += 1;
                                        }
                                        else if (dt.Rows[i][6].ToString() == "")
                                        {
                                            i += 1;
                                        }
                                        else
                                        {
                                            //tên dự án dòng trên khác tên dự án dòng dưới
                                            if (checkInsert)
                                            {
                                                ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 5).InsertRowsBelow(1);
                                            }
                                            colEpx += 1;
                                            ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][6];
                                            for (int c = colSumStart; c <= colSumEnd; c++)
                                            {
                                                colEpx = c - colContent + 2;
                                                ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][c];
                                            }
                                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 5).Style.Border.TopBorder = XLBorderStyleValues.Dashed;
                                            rowsEpx += 1;
                                            colEpx = startColumns;
                                        }
                                    }
                                    else
                                    {
                                        //tên dự án dòng trên khác tên dự án dòng dưới
                                        if (checkInsert)
                                        {
                                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 5).InsertRowsBelow(1);
                                        }
                                        colEpx += 1;
                                        ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][6];
                                        for (int c = colSumStart; c <= colSumEnd; c++)
                                        {
                                            colEpx = c - colContent + 2;
                                            ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][c];
                                        }
                                        ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 5).Style.Border.TopBorder = XLBorderStyleValues.Dashed;
                                        rowsEpx += 1;
                                        colEpx = startColumns;
                                    }
                                }
                                else
                                {
                                    //Them dòng nhóm 4
                                    group4 = dt.Rows[i][4].ToString();
                                    if (checkInsert)
                                    {
                                        ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 5).InsertRowsBelow(1);
                                    }
                                    idGroup4 += 1;
                                    idGroup5 = 0;
                                    ws.Cell(rowsEpx, colEpx).Value = "'" + idGroup3.ToString() + "." + idGroup4.ToString();
                                    colEpx += 1;
                                    ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][4];
                                    for (int c = colSumStart; c <= colSumEnd; c++)
                                    {
                                        colEpx = c - colContent + 2;
                                        ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "' AND " + dt.Columns[4].ColumnName + "='" + group4 + "'");
                                    }
                                    ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 5).Style.Font.Bold = true;
                                    ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 5).Style.Font.Italic = true;
                                    rowsEpx += 1;
                                    colEpx = startColumns;
                                    i -= 1; //xử lý tiếp dòng dữ liệu hiện hành 
                                }
                            }
                            else
                            {
                                //Them dòng nhóm 3
                                group3 = dt.Rows[i][3].ToString();
                                group4 = "";
                                if (checkInsert)
                                {
                                    ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 5).InsertRowsBelow(1);
                                }
                                idGroup3 += 1;
                                idGroup4 = 0;
                                idGroup5 = 0;
                                ws.Cell(rowsEpx, colEpx).Value = idGroup3.ToString();
                                colEpx += 1;
                                ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][3];
                                for (int c = colSumStart; c <= colSumEnd; c++)
                                {
                                    colEpx = c - colContent + 2;
                                    ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "'");
                                }
                                ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 5).Style.Font.Bold = true;
                                rowsEpx += 1;
                                colEpx = startColumns;
                                i -= 1; //xử lý tiếp dòng dữ liệu hiện hành 
                            }
                        }
                        else
                        {
                            //Them dòng nhóm 2
                            group2 = dt.Rows[i][2].ToString();
                            group3 = "";
                            group4 = "";
                            if (checkInsert)
                            {
                                ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 5).InsertRowsBelow(1);
                            }
                            idGroup2 += 1;
                            idGroup3 = 0;
                            idGroup4 = 0;
                            idGroup5 = 0;
                            ws.Cell(rowsEpx, colEpx).Value = IntToRoman(idGroup2);
                            colEpx += 1;
                            ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][2];
                            for (int c = colSumStart; c <= colSumEnd; c++)
                            {
                                colEpx = c - colContent + 2;
                                ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "'");
                            }
                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 5).Style.Font.Bold = true;
                            rowsEpx += 1;
                            colEpx = startColumns;
                            i -= 1; //xử lý tiếp dòng dữ liệu hiện hành 
                        }
                    }
                    else
                    {
                        //Them dòng nhóm 1
                        group1 = dt.Rows[i][1].ToString();
                        group2 = "";
                        group3 = "";
                        group4 = "";
                        if (checkInsert)
                        {
                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 5).InsertRowsBelow(1);
                        }
                        idGroup1 += 1;
                        idGroup2 = 0;
                        idGroup3 = 0;
                        idGroup4 = 0;
                        idGroup5 = 0;
                        ws.Cell(rowsEpx, colEpx).Value = IntToAlphabet(idGroup1, true);
                        colEpx += 1;
                        ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][1].ToString().ToUpper();
                        for (int c = colSumStart; c <= colSumEnd; c++)
                        {
                            //ghi số tiền vào cột 7 trên file excel: colSumStart=11; colContent=6
                            colEpx = c - colContent + 2;
                            ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "'");
                        }
                        ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 5).Style.Font.Bold = true;
                        rowsEpx += 1;
                        colEpx = startColumns;
                        i -= 1; //xử lý tiếp dòng dữ liệu hiện hành 
                    }
                }
                arr[0] = rowsEpx;
                arr[1] = idGroup1.ToString() + "-" + idGroup2.ToString() + "-" + idGroup3.ToString() + "-" + idGroup4.ToString() + "-" + idGroup5.ToString();
                return arr;
            }
            catch
            {
                return arr;
            }
        }

        /// <summary>
        /// Xuất dữ liệu 6 group
        /// </summary>
        /// <param name="wb">Wordkbook cần của file excel</param>
        /// <param name="startRows">Dòng bắt đầu ghi dữ liệu trên file excel (Tính từ 1)</param>
        /// <param name="startColumns">Cột bắt đầu ghi dữ liệu ra file excel (Tính từ 1)</param>
        /// <param name="checkInsert">Trước khi ghi có insert dòng trống không</param>
        /// <param name="dt">Datatable truyền vào</param>
        /// <param name="colSumStart">Cột bắt đầu sum trên datatable (Tính từ 0)</param>
        /// <param name="colSumEnd">Cột kết thúc sum</param>
        /// <param name="idGroup1">STT group 1</param>
        /// <param name="idGroup2">STT group 2</param>
        /// <param name="idGroup3">STT group 3</param>
        /// <param name="idGroup4">STT group 4</param>
        /// <param name="idGroup5">STT group 5</param>
        /// <param name="idGroup6">STT group 6</param>
        /// <param name="checkSum">Xuất tổng cộng dòng đầu</param>
        /// <param name="labelSum">Chuỗi dòng tổng cộng</param>
        /// <param name="expDetailSum">Xuất nội dung sau dòng tổng cộng</param>
        /// <param name="expDetailGroup1">Xuất nội dung sau dòng nhóm 1</param>
        /// <param name="expDetailGroup2">Xuất nội dung sau dòng nhóm 2</param>
        /// <param name="expDetailGroup3">Xuất nội dung sau dòng nhóm 3</param>
        /// <param name="expDetailGroup4">Xuất nội dung sau dòng nhóm 4</param>
        /// <param name="expDetailGroup5">Xuất nội dung sau dòng nhóm 5</param>
        /// <param name="expDetailGroup6">Xuất nội dung sau dòng nhóm 6</param>
        /// <returns>object[]</returns>
        public object[] WriteExcel6Group(XLWorkbook wb, int startRows, int startColumns, bool checkInsert, DataTable dt, int colSumStart, int colSumEnd, int idGroup1, int idGroup2, int idGroup3, int idGroup4, int idGroup5, int idGroup6, bool checkSum, string labelSum, bool expDetailSum, bool expDetailGroup1, bool expDetailGroup2, bool expDetailGroup3, bool expDetailGroup4, bool expDetailGroup5, bool expDetailGroup6)
        {
            object[] arr = new object[2];
            int rowsEpx = startRows;
            int colEpx = startColumns;
            int flag = 0;
            var ws = wb.Worksheet("Sheet1");
            DataTable detail = null;
            try
            {
                string group1 = "", group2 = "", group3 = "", group4 = "", group5 = "", group6 = "";
                int colContent = 7;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (checkSum == true && flag == 0)
                    {
                        if (checkInsert)
                        {
                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).InsertRowsBelow(1);
                        }
                        ws.Cell(rowsEpx, colEpx + 1).Value = labelSum.ToUpper();
                        for (int c = colSumStart; c <= colSumEnd; c++)
                        {
                            colEpx = c - colContent + 2;
                            ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", "");
                        }
                        ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).Style.Font.Bold = true;
                        rowsEpx += 1;
                        colEpx = startColumns;
                        //kiểm tra xuất detail của nhóm tổng cộng
                        if (expDetailSum == true)
                        {
                            //xuat dong nguồn trong nước
                            if (checkInsert)
                            {
                                ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).InsertRowsBelow(1);
                            }
                            DataView dtView = new DataView(dt);
                            detail = dtView.ToTable(true, dt.Columns[7].ColumnName.ToString());
                            for (int b = 0; b < detail.Rows.Count; b++)
                            {
                                colEpx += 1;
                                ws.Cell(rowsEpx, colEpx).Value = detail.Rows[b][0];
                                for (int c = colSumStart; c <= colSumEnd; c++)
                                {
                                    colEpx = c - colContent + 2;
                                    ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[7].ColumnName + "='" + detail.Rows[b][0] + "'");
                                }
                                ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).Style.Border.TopBorder = XLBorderStyleValues.Dashed;
                                rowsEpx += 1;
                                colEpx = startColumns;
                            }
                        }
                        flag += 1;
                    }
                    if (group1 == dt.Rows[i][1].ToString())//cot 1 chứa nhóm 1
                    {
                        if (group2 == dt.Rows[i][2].ToString())//cot 2 chứa nhóm 2
                        {
                            if (group3 == dt.Rows[i][3].ToString())//cot 3 chứa nhóm 3
                            {
                                if (group4 == dt.Rows[i][4].ToString())//cot 4 chứa nhóm 4
                                {
                                    if (group5 == dt.Rows[i][5].ToString())//cot 5 chứa nhóm 5
                                    {
                                        if (group6 == dt.Rows[i][6].ToString())//cot 6 chứa nhóm 6
                                        {
                                            if (dt.Rows[i][3].ToString() == "Nguồn vốn XDCB tập trung")
                                            {
                                                idGroup6 += 1;
                                                if (checkInsert)
                                                {
                                                    ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).InsertRowsBelow(1);
                                                }
                                                ws.Cell(rowsEpx, colEpx).Value = "";
                                                colEpx += 1;
                                                ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][7];
                                                for (int c = colSumStart; c <= colSumEnd; c++)
                                                {
                                                    colEpx = c - colContent + 2;
                                                    ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][c];
                                                }
                                                ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).Style.Border.BottomBorder = XLBorderStyleValues.Dashed;
                                                rowsEpx += 1;
                                                colEpx = startColumns;
                                            }
                                        }
                                        else
                                        {
                                            //Them dòng nhóm 6, dòng dự án
                                            group6 = dt.Rows[i][6].ToString();
                                            if (checkInsert)
                                            {
                                                ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).InsertRowsBelow(1);
                                            }
                                            idGroup6 += 1;
                                            ws.Cell(rowsEpx, colEpx).Value = "-";
                                            colEpx += 1;
                                            ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][6];
                                            for (int c = colContent + 1; c <= colSumEnd; c++)
                                            {
                                                if (c < colSumStart)
                                                {
                                                    colEpx += 1;
                                                    ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][c];//xuất nội dung của các cột kiểu text sau cột nội dung (sau cột 5) cho đến cột kiểu số
                                                }
                                                else if (c > colSumStart || c <= colSumEnd)
                                                {
                                                    colEpx += 1;
                                                    ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "' AND " + dt.Columns[4].ColumnName + "='" + group4 + "' AND " + dt.Columns[5].ColumnName + "='" + group5 + "' AND " + dt.Columns[6].ColumnName + "='" + group6 + "'");
                                                }
                                            }
                                            rowsEpx += 1;
                                            colEpx = startColumns;
                                            //kiểm tra xuất detail của nhóm tổng cộng
                                            if (expDetailGroup6 == true)
                                            {
                                                //xuat dong nguồn trong nước
                                                if (checkInsert)
                                                {
                                                    ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).InsertRowsBelow(1);
                                                }
                                                DataView dtView = new DataView(dt);
                                                detail = dtView.ToTable(true, dt.Columns[1].ColumnName.ToString(), dt.Columns[7].ColumnName.ToString());
                                                for (int b = 0; b < detail.Rows.Count; b++)
                                                {
                                                    colEpx += 1;
                                                    ws.Cell(rowsEpx, colEpx).Value = detail.Rows[b][1];
                                                    for (int c = colSumStart; c <= colSumEnd; c++)
                                                    {
                                                        colEpx = c - colContent + 2;
                                                        ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "' AND " + dt.Columns[4].ColumnName + "='" + group4 + "' AND " + dt.Columns[5].ColumnName + "='" + group5 + "' AND " + dt.Columns[6].ColumnName + "='" + group6+ "' AND " + dt.Columns[7].ColumnName + "='" + detail.Rows[b][1] + "'");
                                                    }
                                                    ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).Style.Border.TopBorder = XLBorderStyleValues.Dashed;
                                                    rowsEpx += 1;
                                                    colEpx = startColumns;
                                                }
                                            }
                                            i -= 1; //xử lý tiếp dòng dữ liệu hiện hành 
                                        }
                                    }
                                    else
                                    {
                                        //Them dòng nhóm 5
                                        group5 = dt.Rows[i][5].ToString();
                                        group6 = "";
                                        if (checkInsert)
                                        {
                                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).InsertRowsBelow(1);
                                        }
                                        idGroup5 += 1;
                                        idGroup6 = 0;
                                        ws.Cell(rowsEpx, colEpx).Value = "*";
                                        colEpx += 1;
                                        ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][5];
                                        for (int c = colSumStart; c <= colSumEnd; c++)
                                        {
                                            colEpx = c - colContent + 2;
                                            ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "' AND " + dt.Columns[4].ColumnName + "='" + group4 + "' AND " + dt.Columns[5].ColumnName + "='" + group5 + "'");
                                        }
                                        ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).Style.Border.TopBorder = XLBorderStyleValues.Dashed;
                                        rowsEpx += 1;
                                        colEpx = startColumns;
                                        //kiểm tra xuất detail của nhóm tổng cộng
                                        if (expDetailGroup5 == true)
                                        {
                                            //xuat dong nguồn trong nước
                                            if (checkInsert)
                                            {
                                                ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).InsertRowsBelow(1);
                                            }
                                            DataView dtView = new DataView(dt);
                                            detail = dtView.ToTable(true, dt.Columns[1].ColumnName.ToString(), dt.Columns[7].ColumnName.ToString());
                                            for (int b = 0; b < detail.Rows.Count; b++)
                                            {
                                                colEpx += 1;
                                                ws.Cell(rowsEpx, colEpx).Value = detail.Rows[b][1];
                                                for (int c = colSumStart; c <= colSumEnd; c++)
                                                {
                                                    colEpx = c - colContent + 2;
                                                    ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "' AND " + dt.Columns[4].ColumnName + "='" + group4 + "' AND " + dt.Columns[5].ColumnName + "='" + group5 + "' AND " + dt.Columns[7].ColumnName + "='" + detail.Rows[b][1] + "'");
                                                }
                                                ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).Style.Border.TopBorder = XLBorderStyleValues.Dashed;
                                                rowsEpx += 1;
                                                colEpx = startColumns;
                                            }
                                        }
                                        i -= 1; //xử lý tiếp dòng dữ liệu hiện hành 
                                    }
                                }
                                else
                                {
                                    //Them dòng nhóm 4
                                    group4 = dt.Rows[i][4].ToString();
                                    group5 = "";
                                    group6 = "";
                                    if (checkInsert)
                                    {
                                        ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).InsertRowsBelow(1);
                                    }
                                    idGroup4 += 1;
                                    idGroup5 = 0;
                                    idGroup6 = 0;
                                    ws.Cell(rowsEpx, colEpx).Value = "'" + idGroup3.ToString() + "." + idGroup4.ToString();
                                    colEpx += 1;
                                    ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][4];
                                    for (int c = colSumStart; c <= colSumEnd; c++)
                                    {
                                        colEpx = c - colContent + 2;
                                        ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "' AND " + dt.Columns[4].ColumnName + "='" + group4 + "'");
                                    }
                                    ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).Style.Font.Bold = true;
                                    ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).Style.Font.Italic = true;
                                    ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).Style.Border.TopBorder = XLBorderStyleValues.Dashed;
                                    rowsEpx += 1;
                                    colEpx = startColumns;
                                    //kiểm tra xuất detail của nhóm tổng cộng
                                    if (expDetailGroup4 == true)
                                    {
                                        //xuat dong nguồn trong nước
                                        if (checkInsert)
                                        {
                                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).InsertRowsBelow(1);
                                        }
                                        DataView dtView = new DataView(dt);
                                        detail = dtView.ToTable(true, dt.Columns[1].ColumnName.ToString(), dt.Columns[7].ColumnName.ToString());
                                        for (int b = 0; b < detail.Rows.Count; b++)
                                        {
                                            colEpx += 1;
                                            ws.Cell(rowsEpx, colEpx).Value = detail.Rows[b][1];
                                            for (int c = colSumStart; c <= colSumEnd; c++)
                                            {
                                                colEpx = c - colContent + 2;
                                                ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "' AND " + dt.Columns[4].ColumnName + "='" + group4 + "' AND " + dt.Columns[7].ColumnName + "='" + detail.Rows[b][1] + "'");
                                            }
                                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).Style.Border.TopBorder = XLBorderStyleValues.Dashed;
                                            rowsEpx += 1;
                                            colEpx = startColumns;
                                        }
                                    }
                                    i -= 1; //xử lý tiếp dòng dữ liệu hiện hành 
                                }
                            }
                            else
                            {
                                //Them dòng nhóm 3
                                group3 = dt.Rows[i][3].ToString();
                                group4 = "";
                                group5 = "";
                                group6 = "";
                                if (checkInsert)
                                {
                                    ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).InsertRowsBelow(1);
                                }
                                idGroup3 += 1;
                                idGroup4 = 0;
                                idGroup5 = 0;
                                idGroup6 = 0;
                                ws.Cell(rowsEpx, colEpx).Value = idGroup3.ToString();
                                colEpx += 1;
                                ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][3];
                                for (int c = colSumStart; c <= colSumEnd; c++)
                                {
                                    colEpx = c - colContent + 2;
                                    ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "'");
                                }
                                ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).Style.Font.Bold = true;
                                rowsEpx += 1;
                                colEpx = startColumns;
                                //kiểm tra xuất detail của nhóm tổng cộng
                                if (expDetailGroup3 == true)
                                {
                                    //xuat dong nguồn trong nước
                                    if (checkInsert)
                                    {
                                        ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).InsertRowsBelow(1);
                                    }
                                    DataView dtView = new DataView(dt);
                                    detail = dtView.ToTable(true, dt.Columns[1].ColumnName.ToString(), dt.Columns[7].ColumnName.ToString());
                                    for (int b = 0; b < detail.Rows.Count; b++)
                                    {
                                        colEpx += 1;
                                        ws.Cell(rowsEpx, colEpx).Value = detail.Rows[b][1];
                                        for (int c = colSumStart; c <= colSumEnd; c++)
                                        {
                                            colEpx = c - colContent + 2;
                                            ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "' AND " + dt.Columns[7].ColumnName + "='" + detail.Rows[b][1] + "'");
                                        }
                                        ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).Style.Border.TopBorder = XLBorderStyleValues.Dashed;
                                        rowsEpx += 1;
                                        colEpx = startColumns;
                                    }
                                }
                                i -= 1; //xử lý tiếp dòng dữ liệu hiện hành 
                            }
                        }
                        else
                        {
                            //Them dòng nhóm 2
                            group2 = dt.Rows[i][2].ToString();
                            group3 = "";
                            group4 = "";
                            group5 = "";
                            group6 = "";
                            if (checkInsert)
                            {
                                ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).InsertRowsBelow(1);
                            }
                            idGroup2 += 1;
                            idGroup3 = 0;
                            idGroup4 = 0;
                            idGroup5 = 0;
                            idGroup6 = 0;
                            ws.Cell(rowsEpx, colEpx).Value = IntToRoman(idGroup2);
                            colEpx += 1;
                            ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][2];
                            for (int c = colSumStart; c <= colSumEnd; c++)
                            {
                                colEpx = c - colContent + 2;
                                ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "'");
                            }
                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).Style.Font.Bold = true;
                            rowsEpx += 1;
                            colEpx = startColumns;

                            //kiểm tra xuất detail của nhóm tổng cộng
                            if (expDetailGroup2 == true)
                            {
                                //xuat dong nguồn trong nước
                                if (checkInsert)
                                {
                                    ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).InsertRowsBelow(1);
                                }
                                DataView dtView = new DataView(dt);
                                detail = dtView.ToTable(true, dt.Columns[1].ColumnName.ToString(), dt.Columns[7].ColumnName.ToString());
                                for (int b = 0; b < detail.Rows.Count; b++)
                                {
                                    colEpx += 1;
                                    ws.Cell(rowsEpx, colEpx).Value = detail.Rows[b][1];
                                    for (int c = colSumStart; c <= colSumEnd; c++)
                                    {
                                        colEpx = c - colContent + 2;
                                        ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[7].ColumnName + "='" + detail.Rows[b][1] + "'");
                                    }
                                    ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).Style.Border.TopBorder = XLBorderStyleValues.Dashed;
                                    rowsEpx += 1;
                                    colEpx = startColumns;
                                }
                            }
                            i -= 1; //xử lý tiếp dòng dữ liệu hiện hành 
                        }
                    }
                    else
                    {
                        //Them dòng nhóm 1
                        group1 = dt.Rows[i][1].ToString();
                        group2 = "";
                        group3 = "";
                        group4 = "";
                        group5 = "";
                        group6 = "";
                        if (checkInsert)
                        {
                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).InsertRowsBelow(1);
                        }
                        idGroup1 += 1;
                        idGroup2 = 0;
                        idGroup3 = 0;
                        idGroup4 = 0;
                        idGroup5 = 0;
                        idGroup6 = 0;
                        ws.Cell(rowsEpx, colEpx).Value = IntToAlphabet(idGroup1, true);
                        colEpx += 1;
                        ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][1].ToString().ToUpper();
                        for (int c = colSumStart; c <= colSumEnd; c++)
                        {
                            //ghi số tiền vào cột 7 trên file excel: colSumStart=11; colContent=6
                            colEpx = c - colContent + 2;
                            ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "'");
                        }
                        ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).Style.Font.Bold = true;
                        rowsEpx += 1;
                        colEpx = startColumns;

                        //kiểm tra xuất detail của nhóm tổng cộng
                        if (expDetailGroup1 == true)
                        {
                            //xuat dong nguồn trong nước
                            if (checkInsert)
                            {
                                ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).InsertRowsBelow(1);
                            }
                            DataView dtView = new DataView(dt);
                            detail = dtView.ToTable(true, dt.Columns[1].ColumnName.ToString(), dt.Columns[7].ColumnName.ToString());
                            for (int b = 0; b < detail.Rows.Count; b++)
                            {
                                colEpx += 1;
                                ws.Cell(rowsEpx, colEpx).Value = detail.Rows[b][1];
                                for (int c = colSumStart; c <= colSumEnd; c++)
                                {
                                    colEpx = c - colContent + 2;
                                    ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[7].ColumnName + "='" + detail.Rows[b][1] + "'");
                                }
                                ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 6).Style.Border.TopBorder = XLBorderStyleValues.Dashed;
                                rowsEpx += 1;
                                colEpx = startColumns;
                            }
                        }
                        i -= 1; //xử lý tiếp dòng dữ liệu hiện hành 
                    }
                }
                arr[0] = rowsEpx;
                arr[1] = idGroup1.ToString() + "-" + idGroup2.ToString() + "-" + idGroup3.ToString() + "-" + idGroup4.ToString() + "-" + idGroup5.ToString();
                return arr;
            }
            catch
            {
                return arr;
            }
        }

        /// <summary>
        /// Xuất dữ liệu ra excel
        /// </summary>
        /// <param name="wb">Wordkbook cần của file excel</param>
        /// <param name="startRows">Dòng bắt đầu ghi dữ liệu trên file excel (Tính từ 1)</param>
        /// <param name="startColumns">Cột bắt đầu ghi dữ liệu ra file excel (Tính từ 1)</param>
        /// <param name="checkInsert">Trước khi ghi có insert dòng trống không</param>
        /// <param name="dt">Datatable truyền vào</param>
        /// <param name="colSumStart">Cột bắt đầu sum trên datatable (Tính từ 0)</param>
        /// <param name="colSumEnd">Cột kết thúc sum</param>
        /// <param name="idGroup1">STT group 1</param>
        /// <param name="idGroup2">STT group 2</param>
        /// <param name="idGroup3">STT group 3</param>
        /// <param name="idGroup4">STT group 4</param>
        /// <param name="checkSum">Xuất tổng cộng dòng đầu</param>
        /// <param name="labelSum">Chuỗi dòng tổng cộng</param>
        /// <returns>object[]</returns>
        public object[] WriteExcel4Group(XLWorkbook wb, int startRows, int startColumns, bool checkInsert, DataTable dt, int colSumStart, int colSumEnd, int idGroup1, int idGroup2, int idGroup3, int idGroup4, bool checkSum, string labelSum)
        {
            object[] arr = new object[2];
            int rowsEpx = startRows;
            int idSTTDuan = 0;
            int colEpx = startColumns;
            int flag = 0;
            var ws = wb.Worksheet("Sheet1");
            try
            {
                string group1 = "", group2 = "", group3 = "", group4 = "";
                int colContent = 5;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (checkSum == true && flag == 0)
                    {
                        if (checkInsert)
                        {
                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 4).InsertRowsBelow(1);
                        }
                        ws.Cell(rowsEpx, colEpx + 1).Value = labelSum.ToUpper();
                        for (int c = colSumStart; c <= colSumEnd; c++)
                        {
                            colEpx = c - colContent + 2;
                            ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", "");
                        }
                        ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 4).Style.Font.Bold = true;
                        rowsEpx += 1;
                        colEpx = startColumns;
                        flag += 1;
                    }
                    if (group1 == dt.Rows[i][1].ToString())//cot 1 chứa nhóm 1
                    {
                        if (group2 == dt.Rows[i][2].ToString())//cot 2 chứa nhóm 2
                        {
                            if (group3 == dt.Rows[i][3].ToString())//cot 3 chứa nhóm 3
                            {
                                if (group4 == dt.Rows[i][4].ToString())
                                {
                                    //Xuất dòng dự án thứ 1
                                    if (checkInsert)
                                    {
                                        ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 4).InsertRowsBelow(1);
                                    }
                                    idSTTDuan += 1;
                                    ws.Cell(rowsEpx, colEpx).Value = IntToAlphabet(idSTTDuan, false);
                                    colEpx += 1;
                                    ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][5];
                                    for (int c = colSumStart; c <= colSumEnd; c++)
                                    {
                                        colEpx = c - colContent + 2;
                                        ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][c];
                                    }
                                    ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 4).Style.Border.TopBorder = XLBorderStyleValues.Dashed;
                                    rowsEpx += 1;
                                    colEpx = startColumns;
                                }
                                else
                                {
                                    //Them dòng nhóm 4
                                    group4 = dt.Rows[i][4].ToString();
                                    if (checkInsert)
                                    {
                                        ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 4).InsertRowsBelow(1);
                                    }
                                    idGroup4 += 1;
                                    ws.Cell(rowsEpx, colEpx).Value = idGroup4.ToString();
                                    colEpx += 1;
                                    ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][4];
                                    for (int c = colSumStart; c <= colSumEnd; c++)
                                    {
                                        colEpx = c - colContent + 2;
                                        ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "' AND " + dt.Columns[4].ColumnName + "='" + group4 + "'");
                                    }
                                    ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 4).Style.Font.Bold = true;
                                    rowsEpx += 1;
                                    colEpx = startColumns;
                                    i -= 1; //xử lý tiếp dòng dữ liệu hiện hành 
                                }
                            }
                            else
                            {
                                //Them dòng nhóm 3
                                group3 = dt.Rows[i][3].ToString();
                                group4 = "";
                                if (checkInsert)
                                {
                                    ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 4).InsertRowsBelow(1);
                                }
                                idGroup3 += 1;
                                idGroup4 = 0;
                                idSTTDuan = 0;
                                ws.Cell(rowsEpx, colEpx).Value = idGroup3.ToString();
                                colEpx += 1;
                                ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][3];
                                for (int c = colSumStart; c <= colSumEnd; c++)
                                {
                                    colEpx = c - colContent + 2;
                                    ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "'");
                                }
                                ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 4).Style.Font.Bold = true;
                                rowsEpx += 1;
                                colEpx = startColumns;
                                i -= 1; //xử lý tiếp dòng dữ liệu hiện hành 
                            }
                        }
                        else
                        {
                            //Them dòng nhóm 2
                            group2 = dt.Rows[i][2].ToString();
                            group3 = "";
                            group4 = "";
                            if (checkInsert)
                            {
                                ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 4).InsertRowsBelow(1);
                            }
                            idGroup2 += 1;
                            idGroup3 = 0;
                            idGroup4 = 0;
                            idSTTDuan = 0;
                            ws.Cell(rowsEpx, colEpx).Value = IntToRoman(idGroup2);
                            colEpx += 1;
                            ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][2];
                            for (int c = colSumStart; c <= colSumEnd; c++)
                            {
                                colEpx = c - colContent + 2;
                                ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "'");
                            }
                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 4).Style.Font.Bold = true;
                            rowsEpx += 1;
                            colEpx = startColumns;
                            i -= 1; //xử lý tiếp dòng dữ liệu hiện hành 
                        }
                    }
                    else
                    {
                        //Them dòng nhóm 1
                        group1 = dt.Rows[i][1].ToString();
                        group2 = "";
                        group3 = "";
                        group4 = "";
                        if (checkInsert)
                        {
                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 4).InsertRowsBelow(1);
                        }
                        idGroup1 += 1;
                        idGroup2 = 0;
                        idGroup3 = 0;
                        idGroup4 = 0;
                        idSTTDuan = 0;
                        ws.Cell(rowsEpx, colEpx).Value = IntToAlphabet(idGroup1, true);
                        colEpx += 1;
                        ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][1].ToString().ToUpper();
                        for (int c = colSumStart; c <= colSumEnd; c++)
                        {
                            //ghi số tiền vào cột 7 trên file excel: colSumStart=11; colContent=6
                            colEpx = c - colContent + 2;
                            ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "'");
                        }
                        ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 4).Style.Font.Bold = true;
                        rowsEpx += 1;
                        colEpx = startColumns;
                        i -= 1; //xử lý tiếp dòng dữ liệu hiện hành 
                    }
                }
                arr[0] = rowsEpx;
                arr[1] = idGroup1.ToString() + "-" + idGroup2.ToString() + "-" + idGroup3.ToString() + "-" + idGroup4.ToString() ;
                return arr;
            }
            catch
            {
                return arr;
            }
        }

        /// <summary>
        /// Xuất dữ liệu ra excel (3Group, STT [Gr1: A; Gr2: I; Gr3: 1])
        /// </summary>
        /// <param name="wb">Wordkbook cần của file excel</param>
        /// <param name="startRows">Dòng bắt đầu ghi dữ liệu trên file excel (Tính từ 1)</param>
        /// <param name="startColumns">Cột bắt đầu ghi dữ liệu ra file excel (Tính từ 1)</param>
        /// <param name="checkInsert">Trước khi ghi có insert dòng trống không</param>
        /// <param name="dt">Datatable truyền vào</param>
        /// <param name="colSumStart">Cột bắt đầu sum trên datatable (Tính từ 0)</param>
        /// <param name="colSumEnd">Cột kết thúc sum</param>
        /// <param name="idGroup1">STT group 1</param>
        /// <param name="idGroup2">STT group 2</param>
        /// <param name="idGroup3">STT group 3</param>
        /// <param name="checkSum">Xuất tổng cộng dòng đầu</param>
        /// <param name="labelSum">Chuỗi dòng tổng cộng</param>
        /// <returns>object[]</returns>
        public object[] WriteExcel3GroupA(XLWorkbook wb, int startRows, int startColumns, bool checkInsert, DataTable dt, int colSumStart, int colSumEnd, int idGroup1, int idGroup2, int idGroup3, bool checkSum, string labelSum)
        {
            object[] arr = new object[2];
            int rowsEpx = startRows;
            int colEpx = startColumns;
            var ws = wb.Worksheet("Sheet1");
            try
            {
                string group1 = "", group2 = "", group3 = "";
                int colContent = 4;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (group1 == dt.Rows[i][1].ToString())//cot 1 chứa nhóm 1
                    {
                        if (group2 == dt.Rows[i][2].ToString())//cot 2 chứa nhóm 2
                        {

                            //Them dòng tên dự án
                            group3 = dt.Rows[i][3].ToString();
                            idGroup3 += 1;
                            if (checkInsert)
                            {
                                ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 2).InsertRowsBelow(1);
                            }
                            ws.Cell(rowsEpx, colEpx).Value = IntToAlphabet(idGroup3, false);
                            colEpx += 1;
                            ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][3];
                            for (int c = colContent + 1; c <= colSumEnd; c++)
                            {
                                if (c < colSumStart)
                                {
                                    colEpx += 1;
                                    ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][c];//xuất nội dung của các cột kiểu text sau cột nội dung (sau cột 5) cho đến cột kiểu số
                                }
                                else if (c > colSumStart || c <= colSumEnd)
                                {
                                    colEpx += 1;
                                    ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "'");
                                }
                            }
                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 3).Style.Border.BottomBorder = XLBorderStyleValues.Dashed;
                            rowsEpx += 1;
                            colEpx = startColumns;
                            //Kiểm tra xem dự án có 1 hay 2 dòng
                            if (dt.Rows[i][2].ToString() == dt.Rows[((i + 1) < dt.Rows.Count ? i + 1 : i)][2].ToString() && dt.Rows[i][3].ToString() == dt.Rows[((i + 1) < dt.Rows.Count ? i + 1 : i)][3].ToString() && dt.Rows.Count > 1 && i < dt.Rows.Count - 1) //tên dự án dòng trên trùng tên dự án dòng dưới
                            {
                                //xuat dong nguồn trong nước
                                if (checkInsert)
                                {
                                    ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 3).InsertRowsBelow(1);
                                }
                                colEpx += 1;
                                ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][4];
                                for (int c = colSumStart; c <= colSumEnd; c++)
                                {
                                    colEpx = c - colContent + 2;
                                    ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][c];
                                }
                                ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 3).Style.Border.TopBorder = XLBorderStyleValues.Dashed;
                                rowsEpx += 1;
                                colEpx = startColumns;
                                //xuất tiếp nguồn ngoài nước
                                if (checkInsert)
                                {
                                    ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 3).InsertRowsBelow(1);
                                }
                                colEpx += 1;
                                ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i + 1][4];
                                for (int c = colSumStart; c <= colSumEnd; c++)
                                {
                                    colEpx = c - colContent + 2;
                                    ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i + 1][c];
                                }
                                ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 3).Style.Border.TopBorder = XLBorderStyleValues.Dashed;
                                rowsEpx += 1;
                                colEpx = startColumns;
                                i += 1;
                            }
                            else //tên dự án dòng trên khác tên dự án dòng dưới
                            {
                                if (checkInsert)
                                {
                                    ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 3).InsertRowsBelow(1);
                                }
                                colEpx += 1;
                                ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][4];
                                for (int c = colSumStart; c <= colSumEnd; c++)
                                {
                                    colEpx = c - colContent + 2;
                                    ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][c];
                                }
                                ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 3).Style.Border.TopBorder = XLBorderStyleValues.Dashed;
                                rowsEpx += 1;
                                colEpx = startColumns;
                            }
                        }
                        else
                        {
                            //Them dòng nhóm 2
                            group2 = dt.Rows[i][2].ToString();
                            group3 = "";
                            if (checkInsert)
                            {
                                ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 3).InsertRowsBelow(1);
                            }
                            idGroup2 += 1;
                            idGroup3 = 0;
                            ws.Cell(rowsEpx, colEpx).Value = IntToRoman(idGroup2);
                            colEpx += 1;
                            ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][2];
                            for (int c = colSumStart; c <= colSumEnd; c++)
                            {
                                colEpx = c - colContent + 2;
                                ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "'");
                            }
                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 3).Style.Font.Bold = true;
                            rowsEpx += 1;
                            colEpx = startColumns;
                            i -= 1; //xử lý tiếp dòng dữ liệu hiện hành 
                        }
                    }
                    else
                    {
                        //Them dòng nhóm 1
                        group1 = dt.Rows[i][1].ToString();
                        group2 = "";
                        group3 = "";
                        if (checkInsert)
                        {
                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 3).InsertRowsBelow(1);
                        }
                        idGroup1 += 1;
                        idGroup2 = 0;
                        idGroup3 = 0;
                        ws.Cell(rowsEpx, colEpx).Value = IntToAlphabet(idGroup1, true);
                        colEpx += 1;
                        ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][1].ToString().ToUpper();
                        for (int c = colSumStart; c <= colSumEnd; c++)
                        {
                            //ghi số tiền vào cột 7 trên file excel: colSumStart=11; colContent=6
                            colEpx = c - colContent + 2;
                            ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "'");
                        }
                        ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 3).Style.Font.Bold = true;
                        rowsEpx += 1;
                        colEpx = startColumns;
                        i -= 1; //xử lý tiếp dòng dữ liệu hiện hành 
                    }
                }
                arr[0] = rowsEpx;
                arr[1] = idGroup1.ToString() + "-" + idGroup2.ToString() + "-" + idGroup3.ToString();
                return arr;
            }
            catch
            {
                return arr;
            }
        }

        /// <summary>
        /// Xuất dữ liệu ra excel (3Group, STT [Gr1: 1; Gr2: 1.1; Gr3: a])
        /// </summary>
        /// <param name="wb">Wordkbook cần của file excel</param>
        /// <param name="startRows">Dòng bắt đầu ghi dữ liệu trên file excel (Tính từ 1)</param>
        /// <param name="startColumns">Cột bắt đầu ghi dữ liệu ra file excel (Tính từ 1)</param>
        /// <param name="checkInsert">Trước khi ghi có insert dòng trống không</param>
        /// <param name="dt">Datatable truyền vào</param>
        /// <param name="colSumStart">Cột bắt đầu sum trên datatable (Tính từ 0)</param>
        /// <param name="colSumEnd">Cột kết thúc sum</param>
        /// <param name="idGroup1">STT group 1</param>
        /// <param name="idGroup2">STT group 2</param>
        /// <param name="idGroup3">STT group 3</param>
        /// <param name="checkSum">Xuất tổng cộng dòng đầu</param>
        /// <param name="labelSum">Chuỗi dòng tổng cộng</param>
        /// <returns>object[]</returns>
        public object[] WriteExcel3Group1(XLWorkbook wb, int startRows, int startColumns, bool checkInsert, DataTable dt, int colSumStart, int colSumEnd, int idGroup1, int idGroup2, int idGroup3, bool checkSum, string labelSum)
        {
            object[] arr = new object[2];
            int rowsEpx = startRows;
            int colEpx = startColumns;
            int flag = 0;
            var ws = wb.Worksheet("Sheet1");
            try
            {
                string group1 = "", group2 = "", group3 = "";
                int colContent = 4;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (checkSum == true && flag == 0)
                    {
                        if (checkInsert)
                        {
                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 2).InsertRowsBelow(1);
                        }
                        ws.Cell(rowsEpx, colEpx + 1).Value = labelSum.ToUpper();
                        for (int c = colSumStart; c <= colSumEnd; c++)
                        {
                            colEpx = c - colContent + 2;
                            ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", "");
                        }
                        ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 2).Style.Font.Bold = true;
                        rowsEpx += 1;
                        colEpx = startColumns;
                        flag += 1;
                    }
                    if (group1 == dt.Rows[i][1].ToString())//cot 1 chứa nhóm 1
                    {
                        if (group2 == dt.Rows[i][2].ToString())//cot 2 chứa nhóm 2
                        {
                            if (dt.Rows[i][3].ToString() != "")
                            {
                                //Xuất nhóm dự án
                                group3 = dt.Rows[i][3].ToString();
                                idGroup3 += 1;
                                if (checkInsert)
                                {
                                    ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 2).InsertRowsBelow(1);
                                }
                                ws.Cell(rowsEpx, colEpx).Value = IntToAlphabet(idGroup3, false);
                                colEpx += 1;
                                ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][3];
                                for (int c = colSumStart; c <= colSumEnd; c++)
                                {
                                    colEpx = c - colContent + 2;
                                    ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "'");
                                }
                                ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 3).Style.Border.BottomBorder = XLBorderStyleValues.Dashed;
                                rowsEpx += 1;
                                colEpx = startColumns;
                                //Kiểm tra xem dự án có 1 hay 2 dòng
                                if (dt.Rows[i][2].ToString() == dt.Rows[((i + 1) < dt.Rows.Count ? i + 1 : i)][2].ToString() && dt.Rows[i][3].ToString() == dt.Rows[((i + 1) < dt.Rows.Count ? i + 1 : i)][3].ToString() && dt.Rows.Count > 1 && i < dt.Rows.Count - 1) //tên dự án dòng trên trùng tên dự án dòng dưới
                                {
                                    //xuat dong nguồn trong nước
                                    if (checkInsert)
                                    {
                                        ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 3).InsertRowsBelow(1);
                                    }
                                    colEpx += 1;
                                    ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][4];
                                    for (int c = colContent + 1; c <= colSumEnd; c++)
                                    {
                                        colEpx = c - colContent + 2;
                                        ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][c];
                                    }
                                    ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 3).Style.Border.TopBorder = XLBorderStyleValues.Dashed;
                                    rowsEpx += 1;
                                    colEpx = startColumns;
                                    //xuất tiếp nguồn ngoài nước
                                    if (checkInsert)
                                    {
                                        ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 3).InsertRowsBelow(1);
                                    }
                                    colEpx += 1;
                                    ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i + 1][4];
                                    for (int c = colContent + 1; c <= colSumEnd; c++)
                                    {
                                        colEpx = c - colContent + 2;
                                        ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i + 1][c];
                                    }
                                    ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 3).Style.Border.TopBorder = XLBorderStyleValues.Dashed;
                                    rowsEpx += 1;
                                    colEpx = startColumns;
                                    i += 1;
                                }
                                else //tên dự án dòng trên khác tên dự án dòng dưới
                                {
                                    if (checkInsert)
                                    {
                                        ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 3).InsertRowsBelow(1);
                                    }
                                    colEpx += 1;
                                    ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][4];
                                    for (int c = colContent + 1; c <= colSumEnd; c++)
                                    {
                                        colEpx = c - colContent + 2;
                                        ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][c];
                                    }
                                    ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 3).Style.Border.TopBorder = XLBorderStyleValues.Dashed;
                                    rowsEpx += 1;
                                    colEpx = startColumns;
                                }
                            }
                            else
                            {
                                if (checkInsert)
                                {
                                    ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 3).InsertRowsBelow(1);
                                }
                                colEpx += 1;
                                ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][4];
                                for (int c = colContent + 1; c <= colSumEnd; c++)
                                {
                                    colEpx = c - colContent + 2;
                                    ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][c];
                                }
                                ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 3).Style.Border.TopBorder = XLBorderStyleValues.Dashed;
                                rowsEpx += 1;
                                colEpx = startColumns;
                            }
                        }
                        else
                        {
                            //Them dòng nhóm 2
                            group2 = dt.Rows[i][2].ToString();
                            group3 = "";
                            if (checkInsert)
                            {
                                ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 3).InsertRowsBelow(1);
                            }
                            idGroup2 += 1;
                            idGroup3 = 0;
                            ws.Cell(rowsEpx, colEpx).Value = "'" + idGroup1.ToString() + "." + idGroup2.ToString();
                            colEpx += 1;
                            ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][2];
                            for (int c = colSumStart; c <= colSumEnd; c++)
                            {
                                colEpx = c - colContent + 2;
                                ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "'");
                            }
                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 3).Style.Font.Bold = true;
                            rowsEpx += 1;
                            colEpx = startColumns;
                            i -= 1; //xử lý tiếp dòng dữ liệu hiện hành 
                        }
                    }
                    else
                    {
                        //Them dòng nhóm 1
                        group1 = dt.Rows[i][1].ToString();
                        group2 = "";
                        group3 = "";
                        if (checkInsert)
                        {
                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 3).InsertRowsBelow(1);
                        }
                        idGroup1 += 1;
                        idGroup2 = 0;
                        idGroup3 = 0;
                        ws.Cell(rowsEpx, colEpx).Value = idGroup1.ToString();
                        colEpx += 1;
                        ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][1].ToString().ToUpper();
                        for (int c = colSumStart; c <= colSumEnd; c++)
                        {
                            //ghi số tiền vào cột 7 trên file excel: colSumStart=11; colContent=6
                            colEpx = c - colContent + 2;
                            ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "'");
                        }
                        ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 3).Style.Font.Bold = true;
                        rowsEpx += 1;
                        colEpx = startColumns;
                        i -= 1; //xử lý tiếp dòng dữ liệu hiện hành 
                    }
                }
                arr[0] = rowsEpx;
                arr[1] = idGroup1.ToString() + "-" + idGroup2.ToString() + "-" + idGroup3.ToString();
                return arr;
            }
            catch
            {
                return arr;
            }
        }

        /// <summary>
        /// Xuất dữ liệu ra excel
        /// </summary>
        /// <param name="wb">Wordkbook cần của file excel</param>
        /// <param name="startRows">Dòng bắt đầu ghi dữ liệu trên file excel (Tính từ 1)</param>
        /// <param name="startColumns">Cột bắt đầu ghi dữ liệu ra file excel (Tính từ 1)</param>
        /// <param name="checkInsert">Trước khi ghi có insert dòng trống không</param>
        /// <param name="dt">Datatable truyền vào</param>
        /// <param name="colSumStart">Cột bắt đầu sum trên datatable (Tính từ 0)</param>
        /// <param name="colSumEnd">Cột kết thúc sum</param>
        /// <param name="idGroup1">STT group 1</param>
        /// <param name="idGroup2">STT group 2</param>
        /// <param name="checkSum">Xuất tổng cộng dòng đầu</param>
        /// <param name="labelSum">Chuỗi dòng tổng cộng</param>
        /// <returns>object[]</returns>
        public object[] WriteExcel2Group(XLWorkbook wb, int startRows, int startColumns, bool checkInsert, DataTable dt, int colSumStart, int colSumEnd, int idGroup1, int idGroup2, bool checkSum, string labelSum)
        {
            object[] arr = new object[2];
            int rowsEpx = startRows;
            int colEpx = startColumns;
            var ws = wb.Worksheet("Sheet1");
            try
            {
                string group1 = "", group2 = "";
                int colContent = 3;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (group1 == dt.Rows[i][1].ToString())//cot 1 chứa nhóm 1
                    {
                        //Them dòng tên dự án
                        group2 = dt.Rows[i][2].ToString();
                        idGroup2 += 1;
                        if (checkInsert)
                        {
                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count).InsertRowsBelow(1);
                        }
                        ws.Cell(rowsEpx, colEpx).Value = IntToAlphabet(idGroup2, false);
                        colEpx += 1;
                        ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][2];
                        for (int c = colContent + 1; c <= colSumEnd; c++)
                        {
                            if (c < colSumStart)
                            {
                                colEpx += 1;
                                ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][c];//xuất nội dung của các cột kiểu text sau cột nội dung (sau cột 5) cho đến cột kiểu số
                            }
                            else if (c > colSumStart || c <= colSumEnd)
                            {
                                colEpx += 1;
                                ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' ");
                            }
                        }
                        ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 2).Style.Border.BottomBorder = XLBorderStyleValues.Dashed;
                        rowsEpx += 1;
                        colEpx = startColumns;
                        //Kiểm tra xem dự án có 1 hay 2 dòng
                        if (dt.Rows[i][1].ToString() == dt.Rows[((i + 1) < dt.Rows.Count ? i + 1 : i)][1].ToString() && dt.Rows[i][2].ToString() == dt.Rows[((i + 1) < dt.Rows.Count ? i + 1 : i)][2].ToString() && dt.Rows.Count > 1 && i< dt.Rows.Count-1) //tên dự án dòng trên trùng tên dự án dòng dưới
                        {
                            //xuat dong nguồn trong nước
                            if (checkInsert)
                            {
                                ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 2).InsertRowsBelow(1);
                            }
                            colEpx += 1;
                            ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][3];
                            for (int c = colSumStart; c <= colSumEnd; c++)
                            {
                                colEpx = c - colContent + 2;
                                ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][c];
                            }
                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 2).Style.Border.TopBorder = XLBorderStyleValues.Dashed;
                            rowsEpx += 1;
                            colEpx = startColumns;
                            //xuất tiếp nguồn ngoài nước
                            if (checkInsert)
                            {
                                ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 2).InsertRowsBelow(1);
                            }
                            colEpx += 1;
                            ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i + 1][3];
                            for (int c = colSumStart; c <= colSumEnd; c++)
                            {
                                colEpx = c - colContent + 2;
                                ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i + 1][c];
                            }
                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 2).Style.Border.TopBorder = XLBorderStyleValues.Dashed;
                            rowsEpx += 1;
                            colEpx = startColumns;
                            i += 1;
                        }
                        else //tên dự án dòng trên khác tên dự án dòng dưới
                        {
                            if (checkInsert)
                            {
                                ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 2).InsertRowsBelow(1);
                            }
                            colEpx += 1;
                            ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][3];
                            for (int c = colSumStart; c <= colSumEnd; c++)
                            {
                                colEpx = c - colContent + 2;
                                ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][c];
                            }
                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 2).Style.Border.TopBorder = XLBorderStyleValues.Dashed;
                            rowsEpx += 1;
                            colEpx = startColumns;
                        }
                    }
                    else
                    {
                        //Them dòng nhóm 1
                        group1 = dt.Rows[i][1].ToString();
                        group2 = "";
                        if (checkInsert)
                        {
                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 2).InsertRowsBelow(1);
                        }
                        idGroup1 += 1;
                        idGroup2 = 0;
                        ws.Cell(rowsEpx, colEpx).Value = IntToAlphabet(idGroup1, true);
                        colEpx += 1;
                        ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][1].ToString().ToUpper();
                        for (int c = colSumStart; c <= colSumEnd; c++)
                        {
                            //ghi số tiền vào cột 7 trên file excel: colSumStart=11; colContent=6
                            colEpx = c - colContent + 2;
                            ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "'");
                        }
                        ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 2).Style.Font.Bold = true;
                        rowsEpx += 1;
                        colEpx = startColumns;
                        i -= 1; //xử lý tiếp dòng dữ liệu hiện hành 
                    }
                }
                arr[0] = rowsEpx;
                arr[1] = idGroup1.ToString() + "-" + idGroup2.ToString();
                return arr;
            }
            catch
            {
                return arr;
            }
        }

        /// <summary>
        /// Xuất dữ liệu 1 group, STT đánh số 1,2,3
        /// </summary>
        /// <param name="wb">Wordkbook cần của file excel</param>
        /// <param name="startRows">Dòng bắt đầu ghi dữ liệu trên file excel (Tính từ 1)</param>
        /// <param name="startColumns">Cột bắt đầu ghi dữ liệu ra file excel (Tính từ 1)</param>
        /// <param name="checkInsert">Trước khi ghi có insert dòng trống không</param>
        /// <param name="dt">Datatable truyền vào</param>
        /// <param name="colSumStart">Cột bắt đầu sum trên datatable (Tính từ 0)</param>
        /// <param name="colSumEnd">Cột kết thúc sum</param>
        /// <param name="idGroup1">STT group 1</param>
        /// <param name="checkSum">Xuất tổng cộng dòng đầu</param>
        /// <param name="labelSum">Chuỗi dòng tổng cộng</param>
        /// <param name="expDetailSum">Xuất chi tiết dưới dòng tổng cộng</param>
        /// <returns>object[]</returns>
        public object[] WriteExcel1GroupID1(XLWorkbook wb, int startRows, int startColumns, bool checkInsert, DataTable dt, int colSumStart, int colSumEnd, int idGroup1, bool checkSum, string labelSum, bool expDetailSum)
        {
            object[] arr = new object[2];
            int rowsEpx = startRows;
            int colEpx = startColumns;
            int flag = 0;
            int colContent = 2;
            var ws = wb.Worksheet("Sheet1");
            try
            {
                string group1 = "";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (checkSum == true && flag == 0)
                    {
                        if (checkInsert)
                        {
                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 1).InsertRowsBelow(1);
                        }
                        ws.Cell(rowsEpx, colEpx + 1).Value = labelSum.ToUpper();
                        for (int c = colSumStart; c <= colSumEnd; c++)
                        {
                            colEpx = c - colContent + 2;
                            ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", "");
                        }
                        ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 1).Style.Font.Bold = true;
                        rowsEpx += 1;
                        colEpx = startColumns;
                        //kiểm tra xuất detail của nhóm tổng cộng
                        if (expDetailSum == true)
                        {
                            //xuat dong nguồn trong nước
                            if (checkInsert)
                            {
                                ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 1).InsertRowsBelow(1);
                            }
                            colEpx += 1;
                            ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][2];
                            for (int c = colSumStart; c <= colSumEnd; c++)
                            {
                                colEpx = c - colContent + 2;
                                ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[2].ColumnName + "='" + dt.Rows[i][2] + "' ");
                            }
                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 1).Style.Border.TopBorder = XLBorderStyleValues.Dashed;
                            rowsEpx += 1;
                            colEpx = startColumns;
                        }
                        flag += 1;
                    }
                    if (group1 == dt.Rows[i][1].ToString())
                    {
                        //xuat dong nguồn trong nước
                        if (checkInsert)
                        {
                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 1).InsertRowsBelow(1);
                        }
                        colEpx += 1;
                        ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][2];
                        for (int c = colContent + 1; c <= colSumEnd; c++)
                        {
                            colEpx = c;
                            ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][c];
                        }
                        ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 1).Style.Border.TopBorder = XLBorderStyleValues.Dashed;
                        rowsEpx += 1;
                        colEpx = startColumns;
                    }
                    else
                    {
                        //Xuất dòng group
                        group1 = dt.Rows[i][1].ToString();
                        idGroup1 += 1;
                        if (checkInsert)
                        {
                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 1).InsertRowsBelow(1);
                        }
                        ws.Cell(rowsEpx, colEpx).Value = idGroup1.ToString();
                        colEpx += 1;
                        ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][1];
                        for (int c = colContent + 1; c <= colSumEnd; c++)
                        {
                            if (c < colSumStart)
                            {
                                colEpx = c;
                                ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][c];//xuất nội dung của các cột kiểu text sau cột nội dung (sau cột 5) cho đến cột kiểu số
                            }
                            else if (c > colSumStart || c <= colSumEnd)
                            {
                                colEpx = c;
                                ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' ");
                            }
                        }
                        ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 1).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        rowsEpx += 1;
                        colEpx = startColumns;
                        i -= 1; 
                    }
                }
                arr[0] = rowsEpx;
                arr[1] = idGroup1.ToString();
                return arr;
            }
            catch
            {
                return arr;
            }
        }

        /// <summary>
        /// Xuất dữ liệu 1 group, STT đánh số 1,2,3
        /// </summary>
        /// <param name="wb">Wordkbook cần của file excel</param>
        /// <param name="startRows">Dòng bắt đầu ghi dữ liệu trên file excel (Tính từ 1)</param>
        /// <param name="startColumns">Cột bắt đầu ghi dữ liệu ra file excel (Tính từ 1)</param>
        /// <param name="checkInsert">Trước khi ghi có insert dòng trống không</param>
        /// <param name="dt">Datatable truyền vào</param>
        /// <param name="colSumStart">Cột bắt đầu sum trên datatable (Tính từ 0)</param>
        /// <param name="colSumEnd">Cột kết thúc sum</param>
        /// <param name="idGroup1">STT group 1</param>
        /// <param name="checkSum">Xuất tổng cộng dòng đầu</param>
        /// <param name="labelSum">Chuỗi dòng tổng cộng</param>
        /// <param name="expDetailSum">Xuất chi tiết dưới dòng tổng cộng</param>
        /// <returns>object[]</returns>
        public object[] WriteExcel1GroupNoneID(XLWorkbook wb, int startRows, int startColumns, bool checkInsert, DataTable dt, int colSumStart, int colSumEnd, string idGroup1, bool checkSum, string labelSum,bool expDetailSum)
        {
            object[] arr = new object[2];
            int rowsEpx = startRows;
            int colEpx = startColumns;
            int flag = 0;
            int colContent = 2;
            var ws = wb.Worksheet("Sheet1");
            try
            {
                string group1 = "";
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (checkSum == true && flag == 0)
                    {
                        if (checkInsert)
                        {
                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 1).InsertRowsBelow(1);
                        }
                        ws.Cell(rowsEpx, colEpx + 1).Value = labelSum.ToUpper();
                        for (int c = colSumStart; c <= colSumEnd; c++)
                        {
                            colEpx = c - colContent + 2;
                            ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", "");
                        }
                        ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 1).Style.Font.Bold = true;
                        rowsEpx += 1;
                        colEpx = startColumns;
                        //kiểm tra xuất detail của nhóm tổng cộng
                        if (expDetailSum == true)
                        {
                            //xuat dong nguồn trong nước
                            if (checkInsert)
                            {
                                ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 1).InsertRowsBelow(1);
                            }
                            colEpx += 1;
                            ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][2];
                            for (int c = colSumStart; c <= colSumEnd; c++)
                            {
                                colEpx = c - colContent + 2;
                                ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[2].ColumnName + "='" + dt.Rows[i][2] + "' ");
                            }
                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 1).Style.Border.TopBorder = XLBorderStyleValues.Dashed;
                            rowsEpx += 1;
                            colEpx = startColumns;
                        }
                        flag += 1;
                    }
                    if (group1 == dt.Rows[i][1].ToString())
                    {
                        //xuat dong nguồn trong nước
                        if (checkInsert)
                        {
                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 1).InsertRowsBelow(1);
                        }
                        colEpx += 1;
                        ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][2];
                        for (int c = colContent + 1; c <= colSumEnd; c++)
                        {
                            colEpx = c;
                            ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][c];
                        }
                        ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 1).Style.Border.TopBorder = XLBorderStyleValues.Dashed;
                        rowsEpx += 1;
                        colEpx = startColumns;
                    }
                    else
                    {
                        //Xuất dòng group
                        group1 = dt.Rows[i][1].ToString();
                        if (checkInsert)
                        {
                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 1).InsertRowsBelow(1);
                        }
                        ws.Cell(rowsEpx, colEpx).Value = idGroup1.ToString();
                        colEpx += 1;
                        ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][1];
                        for (int c = colContent + 1; c <= colSumEnd; c++)
                        {
                            if (c < colSumStart)
                            {
                                colEpx = c;
                                ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][c];//xuất nội dung của các cột kiểu text sau cột nội dung (sau cột 5) cho đến cột kiểu số
                            }
                            else if (c > colSumStart || c <= colSumEnd)
                            {
                                colEpx = c;
                                ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' ");
                            }
                        }
                        ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count - 1).Style.Border.TopBorder = XLBorderStyleValues.Thin;
                        rowsEpx += 1;
                        colEpx = startColumns;
                        i -= 1;
                    }
                }
                arr[0] = rowsEpx;
                arr[1] = idGroup1.ToString();
                return arr;
            }
            catch
            {
                return arr;
            }
        }

        /// <summary>
        /// Xuất dữ liệu ra excel
        /// </summary>
        /// <param name="wb">Wordkbook cần của file excel</param>
        /// <param name="startRows">Dòng bắt đầu ghi dữ liệu trên file excel (Tính từ 1)</param>
        /// <param name="startColumns">Cột bắt đầu ghi dữ liệu ra file excel (Tính từ 1)</param>
        /// <param name="checkInsert">Trước khi ghi có insert dòng trống không</param>
        /// <param name="dt">Datatable truyền vào</param>
        /// <param name="colSumStart">Cột bắt đầu sum trên datatable (Tính từ 0)</param>
        /// <param name="colSumEnd">Cột kết thúc sum trên datatable</param>
        /// <param name="idGroup1">STT group 1</param>
        /// <param name="checkSum">Xuất tổng cộng dòng đầu</param>
        /// <param name="labelSum">Chuỗi dòng tổng cộng</param>
        /// <returns>object[]</returns>
        public object[] WriteExcel0Group(XLWorkbook wb, int startRows, int startColumns, bool checkInsert, DataTable dt, int colSumStart, int colSumEnd, int id, bool checkSum, string labelSum, bool epxID)
        {
            object[] arr = new object[2];
            int rowsEpx = startRows;
            int colEpx = startColumns;
            var ws = wb.Worksheet("Sheet1");
            int flag = 0;
            try
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (checkSum == true && flag == 0)
                    {
                        if (checkInsert)
                        {
                            ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count).InsertRowsBelow(1);
                        }
                        ws.Cell(rowsEpx, colEpx).Value = labelSum.ToUpper();
                        for (int c = colSumStart; c <= colSumEnd; c++)
                        {
                            colEpx = c + 1;
                            ws.Cell(rowsEpx, colEpx).Value = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", "");
                        }
                        ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count).Style.Font.Bold = true;
                        rowsEpx += 1;
                        colEpx = startColumns;
                        flag += 1;
                    }
                    //Them dòng tên dự án
                    if (checkInsert)
                    {
                        ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count).InsertRowsBelow(1);
                    }
                    //kiểm tra xuất STT
                    if (epxID == true)
                    {
                        id += 1;
                        ws.Cell(rowsEpx, colEpx).Value = id.ToString();
                        colEpx += 1;
                    }
                    ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][1];
                    for (int c = colSumStart; c <= colSumEnd; c++)
                    {
                        colEpx = c + 1;
                        ws.Cell(rowsEpx, colEpx).Value = dt.Rows[i][c];//xuất nội dung của các cột kiểu text sau cột nội dung (sau cột 5) cho đến cột kiểu số
                    }
                    ws.Range(rowsEpx, 1, rowsEpx, dt.Columns.Count).Style.Border.TopBorder = XLBorderStyleValues.Dashed;
                    rowsEpx += 1;
                    colEpx = startColumns;
                }
                arr[0] = rowsEpx;
                arr[1] = id.ToString();
                return arr;
            }
            catch
            {
                return arr;
            }
        }

        /// <summary>
        /// Xuất DataTable 5 nhóm
        /// </summary>
        /// <param name="dt">Datatable truyền vào</param>
        /// <param name="colSumStart">Cột bắt đầu sum</param>
        /// <param name="colSumEnd">Cột kết thúc sum</param>
        /// <param name="idGroup1">STT group 1</param>
        /// <param name="idGroup2">STT group 2</param>
        /// <param name="idGroup3">STT group 3</param>
        /// <param name="idGroup4">STT group 4</param>
        /// <param name="idGroup5">STT group 5</param>
        /// <param name="checkSum">Xuất tổng cộng dòng đầu</param>
        /// <param name="labelSum">Chuỗi dòng tổng cộng</param>
        /// <returns>object[]</returns>
        public object[] getTable5Group(DataTable dt, int colSumStart, int colSumEnd, int idGroup1, int idGroup2, int idGroup3, int idGroup4, int idGroup5,bool checkSum,string labelSum)
        {
            object[] arr = new object[2];

            DataTable tblDes = dt.Copy();
            DataColumn colAddID = tblDes.Columns.Add("ID", System.Type.GetType("System.String"));
            DataColumn colAddBold = tblDes.Columns.Add("Bold", System.Type.GetType("System.Boolean"));
            DataColumn colAddItalic = tblDes.Columns.Add("Italic", System.Type.GetType("System.Boolean"));
            DataColumn colAddBorder = tblDes.Columns.Add("Border", System.Type.GetType("System.Boolean"));
            try
            {
                tblDes.Columns.Remove(tblDes.Columns[6]);//bỏ cột loaiNguon
                tblDes.Columns.Remove(tblDes.Columns[0]);//bỏ cột sttSort
                tblDes.Columns.Remove(tblDes.Columns[0]);//bỏ cột nhóm 1
                tblDes.Columns.Remove(tblDes.Columns[0]);//bỏ cột nhóm 2
                tblDes.Columns.Remove(tblDes.Columns[0]);//bỏ cột nhóm 3
                tblDes.Columns.Remove(tblDes.Columns[0]);//bỏ cột nhóm 4
                tblDes.Rows.Clear();
                string group1 = "", group2 = "", group3 = "", group4 = "", group5 = "";
                int colContent = 6;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (group1 == dt.Rows[i][1].ToString())//cot 1 chứa nhóm 1
                    {
                        if (group2 == dt.Rows[i][2].ToString())//cot 2 chứa nhóm 2
                        {
                            if (group3 == dt.Rows[i][3].ToString())//cot 3 chứa nhóm 3
                            {
                                if (group4 == dt.Rows[i][4].ToString())//cot 4 chứa nhóm 4
                                {
                                    //Them dòng tên dự án
                                    tblDes.Rows.Add();
                                    tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][5];//cột 5 chứa nhóm 5
                                    group5 = dt.Rows[i][5].ToString();
                                    idGroup5 += 1;
                                    tblDes.Rows[tblDes.Rows.Count - 1]["ID"] = IntToAlphabet(idGroup5, false);
                                    tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "False";
                                    tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                                    tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "True";

                                    for (int c = colContent + 1; c <= colSumEnd; c++)
                                    {
                                        if (c < colSumStart)
                                        {
                                            tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Rows[i][c];//xuất nội dung của các cột kiểu text sau cột nội dung (sau cột 5) cho đến cột kiểu số
                                        }
                                        else if (c > colSumStart || c <= colSumEnd)
                                        {
                                            tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "' AND " + dt.Columns[4].ColumnName + "='" + group4 + "' AND " + dt.Columns[5].ColumnName + "='" + group5 + "'");
                                        }
                                    }
                                    //Kiểm tra xem dự án có 1 hay 2 dòng
                                    if (dt.Rows[i][4].ToString() == dt.Rows[((i + 1) < dt.Rows.Count ? i + 1 : i)][4].ToString() && dt.Rows[i][5].ToString() == dt.Rows[((i + 1) < dt.Rows.Count ? i + 1 : i)][5].ToString() && dt.Rows.Count > 1) //tên dự án dòng trên trùng tên dự án dòng dưới
                                    {
                                        tblDes.Rows.Add();
                                        tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][6];//cột 6 chứa nội dung
                                        tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "False";
                                        tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                                        tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "True";
                                        for (int c = colSumStart; c <= colSumEnd; c++)
                                        {
                                            tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Rows[i][c];
                                        }
                                        tblDes.Rows.Add();
                                        tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i + 1][6];//cột 6 chứa nội dung
                                        tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "False";
                                        tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                                        tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "True";
                                        for (int c = colSumStart; c <= colSumEnd; c++)
                                        {
                                            tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Rows[i + 1][c];
                                        }
                                        i += 1;
                                    }
                                    else //tên dự án dòng trên khác tên dự án dòng dưới
                                    {
                                        //xảy ra lỗi khi xét dòng cuối cùng nên phải kiểm tra (i+1) với rowcount
                                        tblDes.Rows.Add();
                                        tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][6];//cột 6 chứa nội dung
                                        tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "False";
                                        tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                                        tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "True";
                                        for (int c = colSumStart; c <= colSumEnd; c++)
                                        {
                                            tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Rows[i][c];
                                        }
                                    }
                                }
                                else
                                {//Them dòng nhóm 4
                                    group4 = dt.Rows[i][4].ToString();
                                    tblDes.Rows.Add();
                                    tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][4];
                                    for (int c = colSumStart; c <= colSumEnd; c++)
                                    {
                                        tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "' AND " + dt.Columns[4].ColumnName + "='" + group4 + "'");
                                    }
                                    idGroup4 += 1;
                                    idGroup5 = 0;
                                    tblDes.Rows[tblDes.Rows.Count - 1]["ID"] = "'" + idGroup3.ToString() + "." + idGroup4.ToString();
                                    tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "True";
                                    tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "True";
                                    tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "False";
                                    i -= 1; //xử lý tiếp dòng dữ liệu hiện hành
                                }
                            }
                            else
                            {//Them dòng nhóm 3
                                group3 = dt.Rows[i][3].ToString();
                                group4 = "";
                                tblDes.Rows.Add();
                                tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][3];
                                for (int c = colSumStart; c <= colSumEnd; c++)
                                {
                                    tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "'");
                                }
                                idGroup3 += 1;
                                idGroup4 = 0;
                                idGroup5 = 0;
                                tblDes.Rows[tblDes.Rows.Count - 1]["ID"] = idGroup3.ToString();
                                tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "True";
                                tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                                tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "False";
                                i -= 1; //xử lý tiếp dòng dữ liệu hiện hành
                            }
                        }
                        else
                        {
                            //Them dòng nhóm 2
                            group2 = dt.Rows[i][2].ToString();
                            group3 = "";
                            group4 = "";
                            tblDes.Rows.Add();
                            tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][2];
                            for (int c = colSumStart; c <= colSumEnd; c++)
                            {
                                tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "'");
                            }
                            idGroup2 += 1;
                            idGroup3 = 0;
                            idGroup4 = 0;
                            idGroup5 = 0;
                            tblDes.Rows[tblDes.Rows.Count - 1]["ID"] = IntToRoman(idGroup2);
                            tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "True";
                            tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                            tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "False";
                            i -= 1; //xử lý tiếp dòng dữ liệu hiện hành1
                        }
                    }
                    else
                    {
                        //Them dòng nhóm 1
                        group1 = dt.Rows[i][1].ToString();
                        group2 = "";
                        group3 = "";
                        group4 = "";
                        tblDes.Rows.Add();
                        tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][1].ToString().ToUpper();
                        for (int c = colSumStart; c <= colSumEnd; c++)
                        {
                            tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "'");
                        }
                        idGroup1 += 1;
                        idGroup2 = 0;
                        idGroup3 = 0;
                        idGroup4 = 0;
                        idGroup5 = 0;
                        tblDes.Rows[tblDes.Rows.Count - 1]["ID"] = IntToAlphabet(idGroup1, true);
                        tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "True";
                        tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                        tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "False";
                        i -= 1; //xử lý tiếp dòng dữ liệu hiện hành 
                    }


                    //if (checkSum == true && i == 0)
                    //{
                    //    tblDes.Rows.Add();
                    //    tblDes.Rows[tblDes.Rows.Count - 1][0] = labelSum.ToUpper();
                    //    for (int c = colSumStart; c <= colSumEnd; c++)
                    //    {
                    //        tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", "");
                    //    }
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "True";
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "False";
                    //}
                    //#region "nếu nhóm 1 khác nhom 1 thì xuất nhóm 1,2,3,4,5"
                    //if (group1 != dt.Rows[i][1].ToString() && dt.Rows[i][1].ToString()!="")//nếu nhóm 1 khác nhom 1 thì xuất nhóm 1,2,3,4,5
                    //{
                    //    //Them dòng nhóm 1
                    //    group1 = dt.Rows[i][1].ToString();
                    //    tblDes.Rows.Add();
                    //    tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][1].ToString().ToUpper();
                    //    for (int c = colSumStart; c <= colSumEnd; c++)
                    //    {
                    //        tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "'");
                    //    }
                    //    idGroup1 += 1;
                    //    idGroup2 = 0;
                    //    idGroup3 = 0;
                    //    idGroup4 = 0;
                    //    idGroup5 = 0;
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["ID"] = IntToAlphabet(idGroup1, true);
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "True";
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "False";

                    //    //Them dòng nhóm 2
                    //    if (dt.Rows[i][2].ToString() != "")
                    //    {
                    //        group2 = dt.Rows[i][2].ToString();
                    //        tblDes.Rows.Add();
                    //        tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][2];
                    //        for (int c = colSumStart; c <= colSumEnd; c++)
                    //        {
                    //            tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "'");
                    //        }
                    //        idGroup2 += 1;
                    //        tblDes.Rows[tblDes.Rows.Count - 1]["ID"] = IntToRoman(idGroup2);
                    //        tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "True";
                    //        tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                    //        tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "False";
                    //    }

                    //    //Them dòng nhóm 3
                    //    if (dt.Rows[i][3].ToString() != "")
                    //    {
                    //        group3 = dt.Rows[i][3].ToString();
                    //        tblDes.Rows.Add();
                    //        tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][3];
                    //        for (int c = colSumStart; c <= colSumEnd; c++)
                    //        {
                    //            tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "'");
                    //        }
                    //        idGroup3 += 1;
                    //        tblDes.Rows[tblDes.Rows.Count - 1]["ID"] = idGroup3.ToString();
                    //        tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "True";
                    //        tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                    //        tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "False";
                    //    }
                    //    //Them dòng nhóm 4
                    //    if (dt.Rows[i][4].ToString() != "")
                    //    {
                    //        group4 = dt.Rows[i][4].ToString();
                    //        tblDes.Rows.Add();
                    //        tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][4];
                    //        for (int c = colSumStart; c <= colSumEnd; c++)
                    //        {
                    //            tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "' AND " + dt.Columns[4].ColumnName + "='" + group4 + "'");
                    //        }
                    //        idGroup4 += 1;
                    //        tblDes.Rows[tblDes.Rows.Count - 1]["ID"] = "'" + idGroup3.ToString() + "." + idGroup4.ToString();
                    //        tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "True";
                    //        tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "True";
                    //        tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "False";
                    //    }
                    //    //Them dòng tên dự án
                    //    tblDes.Rows.Add();
                    //    tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][5];//cột 5 chứa nhóm 5
                    //    group5 = dt.Rows[i][5].ToString();

                    //    for (int c = colContent + 1; c <= colSumEnd; c++)
                    //    {
                    //        if (c < colSumStart)
                    //        {
                    //            tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Rows[i][c];//xuất nội dung của các cột kiểu text sau cột nội dung (sau cột 5) cho đến cột kiểu số
                    //        }
                    //        else if (c > colSumStart || c <= colSumEnd)
                    //        {
                    //            tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "' AND " + dt.Columns[4].ColumnName + "='" + group4 + "' AND " + dt.Columns[5].ColumnName + "='" + group5 + "'");
                    //        }
                    //    }
                    //    idGroup5 += 1;
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["ID"] = IntToAlphabet(idGroup5, false);
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "False";
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "True";
                    //}
                    //#endregion

                    //#region "nếu nhóm 2 khác nhom 2 thì xuất nhóm 2,3,4,5"
                    //if (group2 != dt.Rows[i][2].ToString() && dt.Rows[i][2].ToString() != "")//nếu nhóm 1 khác nhom 1 thì xuất nhóm 1,2,3,4,5
                    //{
                    //    //Them dòng nhóm 2
                    //    group2 = dt.Rows[i][2].ToString();
                    //    tblDes.Rows.Add();
                    //    tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][2];
                    //    for (int c = colSumStart; c <= colSumEnd; c++)
                    //    {
                    //        tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "'");
                    //    }
                    //    idGroup2 += 1;
                    //    idGroup3 = 0;
                    //    idGroup4 = 0;
                    //    idGroup5 = 0;
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["ID"] = IntToRoman(idGroup2);
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "True";
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "False";

                    //    //Them dòng nhóm 3
                    //    if (dt.Rows[i][3].ToString() != "")
                    //    {
                    //        group3 = dt.Rows[i][3].ToString();
                    //        tblDes.Rows.Add();
                    //        tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][3];
                    //        for (int c = colSumStart; c <= colSumEnd; c++)
                    //        {
                    //            tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "'");
                    //        }
                    //        idGroup3 += 1;
                    //        tblDes.Rows[tblDes.Rows.Count - 1]["ID"] = idGroup3.ToString();
                    //        tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "True";
                    //        tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                    //        tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "False";
                    //    }

                    //    //Them dòng nhóm 4
                    //    if (dt.Rows[i][4].ToString() != "")
                    //    {
                    //        group4 = dt.Rows[i][4].ToString();
                    //        tblDes.Rows.Add();
                    //        tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][4];
                    //        for (int c = colSumStart; c <= colSumEnd; c++)
                    //        {
                    //            tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "' AND " + dt.Columns[4].ColumnName + "='" + group4 + "'");
                    //        }
                    //        idGroup4 += 1;
                    //        tblDes.Rows[tblDes.Rows.Count - 1]["ID"] = "'" + idGroup3.ToString() + "." + idGroup4.ToString();
                    //        tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "True";
                    //        tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "True";
                    //        tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "False";
                    //    }
                    //    //Them dòng tên dự án
                    //    tblDes.Rows.Add();
                    //    tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][5];//cột 5 chứa nhóm 5
                    //    group5 = dt.Rows[i][5].ToString();

                    //    for (int c = colContent + 1; c <= colSumEnd; c++)
                    //    {
                    //        if (c < colSumStart)
                    //        {
                    //            tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Rows[i][c];//xuất nội dung của các cột kiểu text sau cột nội dung (sau cột 5) cho đến cột kiểu số
                    //        }
                    //        else if (c > colSumStart || c <= colSumEnd)
                    //        {
                    //            tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "' AND " + dt.Columns[4].ColumnName + "='" + group4 + "' AND " + dt.Columns[5].ColumnName + "='" + group5 + "'");
                    //        }
                    //    }
                    //    idGroup5 += 1;
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["ID"] = IntToAlphabet(idGroup5, false);
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "False";
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "True";
                    //}
                    //#endregion

                    //#region "nếu nhóm 3 khác nhom 3 thì xuất nhóm 3,4,5"
                    //if (group3 != dt.Rows[i][3].ToString() && dt.Rows[i][3].ToString()!="")//nếu nhóm 1 khác nhom 1 thì xuất nhóm 1,2,3,4,5
                    //{
                    //    //Them dòng nhóm 3
                    //    group3 = dt.Rows[i][3].ToString();
                    //    tblDes.Rows.Add();
                    //    tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][3];
                    //    for (int c = colSumStart; c <= colSumEnd; c++)
                    //    {
                    //        tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "'");
                    //    }
                    //    idGroup3 += 1;
                    //    idGroup4 = 0;
                    //    idGroup5 = 0;
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["ID"] = idGroup3.ToString();
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "True";
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "False";

                    //    //Them dòng nhóm 4
                    //    if (dt.Rows[i][4].ToString() != "")
                    //    {
                    //        group4 = dt.Rows[i][4].ToString();
                    //        tblDes.Rows.Add();
                    //        tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][4];
                    //        for (int c = colSumStart; c <= colSumEnd; c++)
                    //        {
                    //            tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "' AND " + dt.Columns[4].ColumnName + "='" + group4 + "'");
                    //        }
                    //        idGroup4 += 1;
                    //        tblDes.Rows[tblDes.Rows.Count - 1]["ID"] = "'" + idGroup3.ToString() + "." + idGroup4.ToString();
                    //        tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "True";
                    //        tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "True";
                    //        tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "False";
                    //    }
                    //    //Them dòng tên dự án
                    //    tblDes.Rows.Add();
                    //    tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][5];//cột 5 chứa nhóm 5
                    //    group5 = dt.Rows[i][5].ToString();

                    //    for (int c = colContent + 1; c <= colSumEnd; c++)
                    //    {
                    //        if (c < colSumStart)
                    //        {
                    //            tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Rows[i][c];//xuất nội dung của các cột kiểu text sau cột nội dung (sau cột 5) cho đến cột kiểu số
                    //        }
                    //        else if (c > colSumStart || c <= colSumEnd)
                    //        {
                    //            tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "' AND " + dt.Columns[4].ColumnName + "='" + group4 + "' AND " + dt.Columns[5].ColumnName + "='" + group5 + "'");
                    //        }
                    //    }
                    //    idGroup5 += 1;
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["ID"] = IntToAlphabet(idGroup5, false);
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "False";
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "True";
                    //}
                    //#endregion

                    //#region "nếu nhóm 4 khác nhom 5 thì xuất nhóm 4,5"
                    //if (group4 != dt.Rows[i][4].ToString() && dt.Rows[i][4].ToString() != "")//nếu nhóm 1 khác nhom 1 thì xuất nhóm 1,2,3,4,5
                    //{
                    //    //Them dòng nhóm 4
                    //    group4 = dt.Rows[i][4].ToString();
                    //    tblDes.Rows.Add();
                    //    tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][4];
                    //    for (int c = colSumStart; c <= colSumEnd; c++)
                    //    {
                    //        tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "' AND " + dt.Columns[4].ColumnName + "='" + group4 + "'");
                    //    }
                    //    idGroup4 += 1;
                    //    idGroup5 = 0;
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["ID"] = "'" + idGroup3.ToString() + "." + idGroup4.ToString();
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "True";
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "True";
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "False";

                    //    //Them dòng tên dự án
                    //    tblDes.Rows.Add();
                    //    tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][5];//cột 5 chứa nhóm 5
                    //    group5 = dt.Rows[i][5].ToString();

                    //    for (int c = colContent + 1; c <= colSumEnd; c++)
                    //    {
                    //        if (c < colSumStart)
                    //        {
                    //            tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Rows[i][c];//xuất nội dung của các cột kiểu text sau cột nội dung (sau cột 5) cho đến cột kiểu số
                    //        }
                    //        else if (c > colSumStart || c <= colSumEnd)
                    //        {
                    //            tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "' AND " + dt.Columns[4].ColumnName + "='" + group4 + "' AND " + dt.Columns[5].ColumnName + "='" + group5 + "'");
                    //        }
                    //    }
                    //    idGroup5 += 1;
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["ID"] = IntToAlphabet(idGroup5, false);
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "False";
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "True";
                    //}
                    //#endregion

                    //#region "nếu nhóm 5 khác nhom 6 thì xuất nhóm 5"
                    //if (group5 != dt.Rows[i][5].ToString() && dt.Rows[i][5].ToString()!="")//nếu nhóm 1 khác nhom 1 thì xuất nhóm 1,2,3,4,5
                    //{
                    //    //Them dòng tên dự án
                    //    tblDes.Rows.Add();
                    //    tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][5];//cột 5 chứa nhóm 5
                    //    group5 = dt.Rows[i][5].ToString();

                    //    for (int c = colContent + 1; c <= colSumEnd; c++)
                    //    {
                    //        if (c < colSumStart)
                    //        {
                    //            tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Rows[i][c];//xuất nội dung của các cột kiểu text sau cột nội dung (sau cột 5) cho đến cột kiểu số
                    //        }
                    //        else if (c > colSumStart || c <= colSumEnd)
                    //        {
                    //            tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "' AND " + dt.Columns[4].ColumnName + "='" + group4 + "' AND " + dt.Columns[5].ColumnName + "='" + group5 + "'");
                    //        }
                    //    }
                    //    idGroup5 += 1;
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["ID"] = IntToAlphabet(idGroup5, false);
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "False";
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                    //    tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "True";
                    //}
                    //#endregion

                    //#region "Cuối cùng xuất nội dung"
                    ////Them dòng tên dự án
                    //tblDes.Rows.Add();
                    //tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][6];//cột 6 chứa nội dung
                    //for (int c = colContent + 5; c < dt.Columns.Count; c++)
                    //{
                    //    tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Rows[i][c];
                    //}
                    //tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "False";
                    //tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                    //tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "True";
                    //#endregion

                }
                colAddID.SetOrdinal(0);//set cột STT về vị trí đầu tiên
                arr[0] = tblDes;
                arr[1] = idGroup1.ToString() + "-" + idGroup2.ToString() + "-" + idGroup3.ToString() + "-" + idGroup4.ToString() + "-" + idGroup5.ToString();
                return arr;
            }
            catch
            {
                return arr;
            }
        }

        /// <summary>
        /// Xuất DataTable 4 nhóm
        /// </summary>
        /// <param name="dt">Datatable truyền vào</param>
        /// <param name="colSumStart">Cột bắt đầu sum</param>
        /// <param name="colSumEnd">Cột kết thúc sum</param>
        /// <param name="idGroup1">STT group 1</param>
        /// <param name="idGroup2">STT group 2</param>
        /// <param name="idGroup3">STT group 3</param>
        /// <param name="checkSum">Xuất tổng cộng dòng đầu</param>
        /// <param name="labelSum">Chuỗi dòng tổng cộng</param>
        /// <returns>object[]</returns>
        public object[] getTable3Group(DataTable dt, int colSumStart, int colSumEnd, int idGroup1, int idGroup2, int idGroup3, bool checkSum, string labelSum)
        {
            object[] arr = new object[2];

            DataTable tblDes = dt.Copy();
            DataColumn colAddID = tblDes.Columns.Add("ID", System.Type.GetType("System.String"));
            DataColumn colAddBold = tblDes.Columns.Add("Bold", System.Type.GetType("System.Boolean"));
            DataColumn colAddItalic = tblDes.Columns.Add("Italic", System.Type.GetType("System.Boolean"));
            DataColumn colAddBorder = tblDes.Columns.Add("Border", System.Type.GetType("System.Boolean"));
            try
            {
                tblDes.Columns.Remove(tblDes.Columns[4]);//bỏ cột loaiNguon
                tblDes.Columns.Remove(tblDes.Columns[0]);//bỏ cột sttSort
                tblDes.Columns.Remove(tblDes.Columns[0]);//bỏ cột nhom1
                tblDes.Columns.Remove(tblDes.Columns[0]);//bỏ cột nhom2
                tblDes.Rows.Clear();
                string group1 = "", group2 = "", group3 = "";
                int colContent = 4;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (checkSum == true && i == 0)
                    {
                        tblDes.Rows.Add();
                        tblDes.Rows[tblDes.Rows.Count - 1][0] = labelSum.ToUpper();
                        for (int c = colSumStart; c <= colSumEnd; c++)
                        {
                            tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", "");
                        }
                        tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "True";
                        tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                        tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "False";
                    }
                    if (group1 == dt.Rows[i][1].ToString())//cot 1 chứa nhóm 1
                    {
                        if (group2 == dt.Rows[i][2].ToString())//cot 2 chứa nhóm 2
                        {
                            //Them dòng tên dự án
                            tblDes.Rows.Add();
                            tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][3];//cột 5 chứa nhóm 5
                            group3 = dt.Rows[i][3].ToString();
                            idGroup3 += 1;
                            tblDes.Rows[tblDes.Rows.Count - 1]["ID"] = IntToAlphabet(idGroup3, false);
                            tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "False";
                            tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                            tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "True";

                            for (int c = colContent + 1; c <= colSumEnd; c++)
                            {
                                if (c < colSumStart)
                                {
                                    tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Rows[i][c];//xuất nội dung của các cột kiểu text sau cột nội dung (sau cột 5) cho đến cột kiểu số
                                }
                                else if (c > colSumStart || c <= colSumEnd)
                                {
                                    tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "'");
                                }
                            }
                            //Kiểm tra xem dự án có 1 hay 2 dòng
                            if (dt.Rows[i][2].ToString() == dt.Rows[((i + 1) < dt.Rows.Count ? i + 1 : i)][2].ToString() && dt.Rows[i][3].ToString() == dt.Rows[((i + 1) < dt.Rows.Count ? i + 1 : i)][3].ToString() && dt.Rows.Count > 1) //tên dự án dòng trên trùng tên dự án dòng dưới
                            {
                                tblDes.Rows.Add();
                                tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][4];//cột 6 chứa nội dung
                                tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "False";
                                tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                                tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "True";
                                for (int c = colSumStart; c <= colSumEnd; c++)
                                {
                                    tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Rows[i][c];
                                }
                                tblDes.Rows.Add();
                                tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i + 1][4];//cột 6 chứa nội dung
                                tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "False";
                                tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                                tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "True";
                                for (int c = colSumStart; c <= colSumEnd; c++)
                                {
                                    tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Rows[i + 1][c];
                                }
                                i += 1;
                            }
                            else //tên dự án dòng trên khác tên dự án dòng dưới
                            {
                                //xảy ra lỗi khi xét dòng cuối cùng nên phải kiểm tra (i+1) với rowcount
                                tblDes.Rows.Add();
                                tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][4];//cột 6 chứa nội dung
                                tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "False";
                                tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                                tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "True";
                                for (int c = colSumStart; c <= colSumEnd; c++)
                                {
                                    tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Rows[i][c];
                                }
                            }
                        }
                        else
                        {
                            //Them dòng nhóm 2
                            group2 = dt.Rows[i][2].ToString();
                            group3 = "";
                            tblDes.Rows.Add();
                            tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][2];
                            for (int c = colSumStart; c <= colSumEnd; c++)
                            {
                                tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "'");
                            }
                            idGroup2 += 1;
                            idGroup3 = 0;
                            tblDes.Rows[tblDes.Rows.Count - 1]["ID"] = IntToRoman(idGroup2);
                            tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "True";
                            tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                            tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "False";
                            i -= 1; //xử lý tiếp dòng dữ liệu hiện hành1
                        }
                    }
                    else
                    {
                        //Them dòng nhóm 1
                        group1 = dt.Rows[i][1].ToString();
                        group2 = "";
                        group3 = "";
                        tblDes.Rows.Add();
                        tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][1].ToString().ToUpper();
                        for (int c = colContent+5; c <= colSumEnd; c++)
                        {
                            tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "'");
                        }
                        idGroup1 += 1;
                        idGroup2 = 0;
                        idGroup3 = 0;
                        tblDes.Rows[tblDes.Rows.Count - 1]["ID"] = IntToAlphabet(idGroup1, true);
                        tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "True";
                        tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                        tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "False";
                        i -= 1; //xử lý tiếp dòng dữ liệu hiện hành 
                    }
                }
                colAddID.SetOrdinal(0);//set cột STT về vị trí đầu tiên
                arr[0] = tblDes;
                arr[1] = idGroup1.ToString() + "-" + idGroup2.ToString() + "-" + idGroup3.ToString();
                return arr;
            }
            catch
            {
                return arr;
            }
        }

        /// <summary>
        /// Xuất DataTable 2 nhóm
        /// </summary>
        /// <param name="dt">Datatable truyền vào</param>
        /// <param name="colSumStart">Cột bắt đầu sum</param>
        /// <param name="colSumEnd">Cột kết thúc sum</param>
        /// <param name="idGroup1">STT group 1</param>
        /// <param name="idGroup2">STT group 2</param>
        /// <param name="checkSum">Xuất tổng cộng dòng đầu</param>
        /// <param name="labelSum">Chuỗi dòng tổng cộng</param>
        /// <returns>object[]</returns>
        public object[] getTable2Group(DataTable dt, int colSumStart, int colSumEnd, int idGroup1, int idGroup2, bool checkSum, string labelSum)
        {
            object[] arr = new object[2];

            DataTable tblDes = dt.Copy();
            DataColumn colAddID = tblDes.Columns.Add("ID", System.Type.GetType("System.String"));
            DataColumn colAddBold = tblDes.Columns.Add("Bold", System.Type.GetType("System.Boolean"));
            DataColumn colAddItalic = tblDes.Columns.Add("Italic", System.Type.GetType("System.Boolean"));
            DataColumn colAddBorder = tblDes.Columns.Add("Border", System.Type.GetType("System.Boolean"));
            try
            {
                tblDes.Columns.Remove(tblDes.Columns[0]);
                tblDes.Columns.Remove(tblDes.Columns[0]);
                tblDes.Rows.Clear();
                string group1 = "", group2 = "";
                int colContent = 3;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (checkSum == true && i == 0)
                    {
                        tblDes.Rows.Add();
                        tblDes.Rows[tblDes.Rows.Count - 1][0] = labelSum.ToUpper();
                        for (int c = colSumStart; c <= colSumEnd; c++)
                        {
                            tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", "");
                        }
                        tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "True";
                        tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                        tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "False";
                    }
                    if (group1 == dt.Rows[i][1].ToString())//cot 1 chứa nhóm 1
                    {

                        //Them dòng tên dự án
                        tblDes.Rows.Add();
                        tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][3];//cột 5 chứa nhóm 5
                        tblDes.Rows[tblDes.Rows.Count - 1]["ID"] = IntToAlphabet(idGroup2, false);
                        tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "False";
                        tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                        tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "True";

                        for (int c = colContent + 1; c <= colSumEnd; c++)
                        {
                            if (c < colSumStart)
                            {
                                tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Rows[i][c];//xuất nội dung của các cột kiểu text sau cột nội dung (sau cột 5) cho đến cột kiểu số
                            }
                            else if (c > colSumStart || c <= colSumEnd)
                            {
                                tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' ");
                            }
                        }
                        //Kiểm tra xem dự án có 1 hay 2 dòng
                        if (dt.Rows[i][1].ToString() == dt.Rows[((i + 1) < dt.Rows.Count ? i + 1 : i)][1].ToString() && dt.Rows[i][2].ToString() == dt.Rows[((i + 1) < dt.Rows.Count ? i + 1 : i)][2].ToString() && dt.Rows.Count > 1) //tên dự án dòng trên trùng tên dự án dòng dưới
                        {
                            tblDes.Rows.Add();
                            tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][3];//cột 3 chứa nội dung
                            tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "False";
                            tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                            tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "True";
                            for (int c = colSumStart; c <= colSumEnd; c++)
                            {
                                tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Rows[i][c];
                            }
                            tblDes.Rows.Add();
                            tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i + 1][3];//cột 3 chứa nội dung
                            tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "False";
                            tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                            tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "True";
                            for (int c = colSumStart; c <= colSumEnd; c++)
                            {
                                tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Rows[i + 1][c];
                            }
                            i += 1;
                        }
                        else //tên dự án dòng trên khác tên dự án dòng dưới
                        {
                            //xảy ra lỗi khi xét dòng cuối cùng nên phải kiểm tra (i+1) với rowcount
                            tblDes.Rows.Add();
                            tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][3];//cột 3 chứa nội dung
                            tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "False";
                            tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                            tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "True";
                            for (int c = colSumStart; c <= colSumEnd; c++)
                            {
                                tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Rows[i][c];
                            }
                        }
                    }
                    else
                    {
                        //Them dòng nhóm 1
                        group1 = dt.Rows[i][1].ToString();
                        group2 = "";
                        tblDes.Rows.Add();
                        tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][1].ToString().ToUpper();
                        for (int c = colContent + 5; c <= colSumEnd; c++)
                        {
                            tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "'");
                        }
                        idGroup1 += 1;
                        idGroup2 = 0;
                        tblDes.Rows[tblDes.Rows.Count - 1]["ID"] = IntToAlphabet(idGroup1, true);
                        tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "True";
                        tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
                        tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "False";
                        i -= 1; //xử lý tiếp dòng dữ liệu hiện hành 
                    }
                }
                colAddID.SetOrdinal(0);//set cột STT về vị trí đầu tiên
                arr[0] = tblDes;
                arr[1] = idGroup1.ToString() + "-" + idGroup2.ToString();
                return arr;
            }
            catch
            {
                return arr;
            }
        }

        //public object[] getDataExportExcel(DataTable dt, int colSumStart, int colSumEnd, int groupCount, int idGroup1, int idGroup2, int idGroup3, int idGroup4, int idGroup5)
        //{
        //    object[] arr = new object[2];

        //    DataTable tblDes = dt.Copy();
        //    DataColumn colAddID = tblDes.Columns.Add("ID", System.Type.GetType("System.String"));
        //    DataColumn colAddBold = tblDes.Columns.Add("Bold", System.Type.GetType("System.Boolean"));
        //    DataColumn colAddItalic = tblDes.Columns.Add("Italic", System.Type.GetType("System.Boolean"));
        //    DataColumn colAddBorder = tblDes.Columns.Add("Border", System.Type.GetType("System.Boolean"));
        //    try
        //    {
        //        tblDes.Columns.Remove(tblDes.Columns[0]);
        //        tblDes.Columns.Remove(tblDes.Columns[0]);
        //        tblDes.Columns.Remove(tblDes.Columns[0]);
        //        tblDes.Columns.Remove(tblDes.Columns[0]);
        //        tblDes.Columns.Remove(tblDes.Columns[0]);
        //        tblDes.Rows.Clear();
        //        string group1 = "", group2 = "", group3 = "", group4 = "", group5 = "";
        //        int colContent = 6;
        //        for (int i = 0; i < dt.Rows.Count; i++)
        //        {
        //            if (group1 == dt.Rows[i][1].ToString())//cot 1 chứa nhóm 1
        //            {
        //                if (group2 == dt.Rows[i][2].ToString())//cot 2 chứa nhóm 2
        //                {
        //                    if (group3 == dt.Rows[i][3].ToString())//cot 3 chứa nhóm 3
        //                    {
        //                        if (group4 == dt.Rows[i][4].ToString())//cot 4 chứa nhóm 4
        //                        {
        //                            //group4 = "";
        //                            //Them dòng tên dự án
        //                            tblDes.Rows.Add();
        //                            tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][5];//cột 5 chứa nhóm 5
        //                            group5 = dt.Rows[i][5].ToString();
        //                            idGroup5 += 1;
        //                            tblDes.Rows[tblDes.Rows.Count - 1]["ID"] = IntToAlphabet(idGroup5, false);
        //                            tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "False";
        //                            tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
        //                            tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "True";

        //                            for (int c = colContent + 1; c <= colSumEnd; c++)
        //                            {
        //                                if (c < colSumStart)
        //                                {
        //                                    tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Rows[i][c];//xuất nội dung của các cột kiểu text sau cột nội dung (sau cột 5) cho đến cột kiểu số
        //                                }
        //                                else if (c > colSumStart || c <= colSumEnd)
        //                                {
        //                                    tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "' AND " + dt.Columns[4].ColumnName + "='" + group4 + "' AND " + dt.Columns[5].ColumnName + "='" + group5 + "'");
        //                                }
        //                            }
        //                            //Kiểm tra xem dự án có 1 hay 2 dòng
        //                            if (dt.Rows[i][4].ToString() == dt.Rows[((i + 1) < dt.Rows.Count ? i + 1 : i)][4].ToString() && dt.Rows[i][5].ToString() == dt.Rows[((i + 1) < dt.Rows.Count ? i + 1 : i)][5].ToString()) //tên dự án dòng trên khác tên dự án dòng dưới
        //                            {
        //                                tblDes.Rows.Add();
        //                                tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][6];//cột 6 chứa nội dung
        //                                tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "False";
        //                                tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
        //                                tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "True";
        //                                for (int c = colContent + 5; c <= colSumEnd; c++)
        //                                {
        //                                    tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Rows[i][c];
        //                                }
        //                                tblDes.Rows.Add();
        //                                tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i + 1][6];//cột 6 chứa nội dung
        //                                tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "False";
        //                                tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
        //                                tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "True";
        //                                for (int c = colContent + 5; c <= colSumEnd; c++)
        //                                {
        //                                    tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Rows[i + 1][c];
        //                                }
        //                                i += 1;

        //                            }
        //                            else //tên dự án dòng trên cùng tên dự án dòng dưới
        //                            {
        //                                //xảy ra lỗi khi xét dòng cuối cùng nên phải kiểm tra (i+1) với rowcount
        //                                tblDes.Rows.Add();
        //                                tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][6];//cột 6 chứa nội dung
        //                                tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "False";
        //                                tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
        //                                tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "True";
        //                                for (int c = colContent + 5; c <= colSumEnd; c++)
        //                                {
        //                                    tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Rows[i][c];
        //                                }
        //                            }

        //                            ////Kiểm tra xem dự án có 1 hay 2 dòng
        //                            //if (dt.Rows[i][4].ToString() == dt.Rows[((i + 1) < dt.Rows.Count ? i + 1 : i)][4].ToString() && dt.Rows[i][5].ToString() != dt.Rows[((i + 1) < dt.Rows.Count ? i + 1 : i)][5].ToString() || (i + 1) == dt.Rows.Count) //tên dự án dòng trên khác tên dự án dòng dưới
        //                            //{
        //                            //    //xảy ra lỗi khi xét dòng cuối cùng nên phải kiểm tra (i+1) với rowcount
        //                            //    tblDes.Rows.Add();
        //                            //    tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][6];//cột 6 chứa nội dung
        //                            //    tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "False";
        //                            //    tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
        //                            //    tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "True";
        //                            //    for (int c = colContent + 5; c <= colSumEnd; c++)
        //                            //    {
        //                            //        tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Rows[i][c];
        //                            //    }
        //                            //}
        //                            //else //tên dự án dòng trên cùng tên dự án dòng dưới
        //                            //{
        //                            //    tblDes.Rows.Add();
        //                            //    tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][6];//cột 6 chứa nội dung
        //                            //    tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "False";
        //                            //    tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
        //                            //    tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "True";
        //                            //    for (int c = colContent + 5; c <= colSumEnd; c++)
        //                            //    {
        //                            //        tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Rows[i][c];
        //                            //    }
        //                            //    tblDes.Rows.Add();
        //                            //    tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i + 1][6];//cột 6 chứa nội dung
        //                            //    tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "False";
        //                            //    tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
        //                            //    tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "True";
        //                            //    for (int c = colContent + 5; c <= colSumEnd; c++)
        //                            //    {
        //                            //        tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Rows[i + 1][c];
        //                            //    }
        //                            //    i += 1;
        //                            //}
        //                        }
        //                        else
        //                        {//Them dòng nhóm 4
        //                            group4 = dt.Rows[i][4].ToString();
        //                            tblDes.Rows.Add();
        //                            tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][4];
        //                            for (int c = colContent + 5; c <= colSumEnd; c++)
        //                            {
        //                                tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "' AND " + dt.Columns[4].ColumnName + "='" + group4 + "'");
        //                            }
        //                            idGroup4 += 1;
        //                            tblDes.Rows[tblDes.Rows.Count - 1]["ID"] = "'" + idGroup3.ToString() + "." + idGroup4.ToString();
        //                            tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "True";
        //                            tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "True";
        //                            tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "False";
        //                            i -= 1; //xử lý tiếp dòng dữ liệu hiện hành
        //                        }
        //                    }
        //                    else
        //                    {//Them dòng nhóm 3
        //                        group3 = dt.Rows[i][3].ToString();
        //                        tblDes.Rows.Add();
        //                        tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][3];
        //                        for (int c = colContent + 5; c <= colSumEnd; c++)
        //                        {
        //                            tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "' AND " + dt.Columns[3].ColumnName + "='" + group3 + "'");
        //                        }
        //                        idGroup3 += 1;
        //                        tblDes.Rows[tblDes.Rows.Count - 1]["ID"] = idGroup3.ToString();
        //                        tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "True";
        //                        tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
        //                        tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "False";
        //                        i -= 1; //xử lý tiếp dòng dữ liệu hiện hành
        //                    }
        //                }
        //                else
        //                {
        //                    //Them dòng nhóm 2
        //                    group2 = dt.Rows[i][2].ToString();
        //                    tblDes.Rows.Add();
        //                    tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][2];
        //                    for (int c = colContent + 5; c <= colSumEnd; c++)
        //                    {
        //                        tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "' AND " + dt.Columns[2].ColumnName + "='" + group2 + "'");
        //                    }
        //                    idGroup2 += 1;
        //                    tblDes.Rows[tblDes.Rows.Count - 1]["ID"] = IntToRoman(idGroup2);
        //                    tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "True";
        //                    tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
        //                    tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "False";
        //                    i -= 1; //xử lý tiếp dòng dữ liệu hiện hành1
        //                }
        //            }
        //            else
        //            {
        //                //Them dòng nhóm 1
        //                group1 = dt.Rows[i][1].ToString();
        //                tblDes.Rows.Add();
        //                tblDes.Rows[tblDes.Rows.Count - 1][0] = dt.Rows[i][1].ToString().ToUpper();
        //                for (int c = colContent + 5; c <= colSumEnd; c++)
        //                {
        //                    tblDes.Rows[tblDes.Rows.Count - 1][c - colContent] = dt.Compute("SUM(" + dt.Columns[c].ColumnName + ")", dt.Columns[1].ColumnName + "='" + group1 + "'");
        //                }
        //                idGroup1 += 1;
        //                tblDes.Rows[tblDes.Rows.Count - 1]["ID"] = IntToAlphabet(idGroup1, true);
        //                tblDes.Rows[tblDes.Rows.Count - 1]["Bold"] = "True";
        //                tblDes.Rows[tblDes.Rows.Count - 1]["Italic"] = "False";
        //                tblDes.Rows[tblDes.Rows.Count - 1]["Border"] = "False";
        //                i -= 1; //xử lý tiếp dòng dữ liệu hiện hành 
        //            }
        //        }
        //        colAddID.SetOrdinal(0);//set cột STT về vị trí đầu tiên
        //        arr[0] = tblDes;
        //        arr[1] = idGroup1.ToString() + "-" + idGroup2.ToString() + "-" + idGroup3.ToString() + "-" + idGroup4.ToString() + "-" + idGroup5.ToString();
        //        return arr;
        //    }
        //    catch
        //    {
        //        return arr;
        //    }
        //}
    }
}
