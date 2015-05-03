using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using System.Data;
using System.Globalization;
using System.Threading;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace SerqUtil.ExcelUtility
{
    public class ExcelWriter
    {
        public static void Write(List<DataTable> workSheets, string filePath)
        {
            Process[] ExcelsBefore =  Process.GetProcessesByName("Excel");
            Application xlApp = new Application();
            Process[] ExcelsAfter = Process.GetProcessesByName("Excel");
            // store the process to kill in final block
            Process excel = ExcelsAfter.Except(ExcelsBefore).Cast<Process>().First();

            xlApp.ScreenUpdating = false;
            CultureInfo CurrentCI = Thread.CurrentThread.CurrentCulture;
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            Workbooks workbooks = xlApp.Workbooks;
            Workbook workbook = workbooks.Add();
            Sheets _workSheets = (Sheets)workbook.Worksheets; 
          
            try
            {
                int workSheetCount = _workSheets.Count;
                int workSheetToBeAddedCount = workSheets.Count - workSheetCount;
                if (workSheetToBeAddedCount > 0)
                {
                    for (int z = 0; z < workSheetToBeAddedCount; z++)
                    {
                        Worksheet worksheet = (Worksheet)_workSheets.Add();
                    }
                }

                for (int a = 0; a < workSheets.Count; a++)
                {
                    Worksheet worksheet = (Worksheet)workbook.Worksheets[a + 1];
                    worksheet.EnableCalculation = false;
                    if (!string.IsNullOrEmpty(workSheets[a].TableName))
                        worksheet.Name = workSheets[a].TableName;

                    for (int i = 0; i < workSheets[a].Columns.Count; i++)
                    {
                        worksheet.Cells[1, i + 1] = workSheets[a].Columns[i].ColumnName;
                        Range range = (Range)worksheet.Cells[1, i + 1];
                        range.Interior.ColorIndex = 15;
                        range.Font.Bold = true;
                        range.EntireColumn.AutoFit();
                    }

                    //Get dimensions of the 2-d array 
                    int rowCount = workSheets[a].Rows.Count;
                    if (rowCount > 0)
                    {
                        int columnCount = workSheets[a].Columns.Count;
                        string[,] array = new string[rowCount, columnCount];

                        for (int s = 0; s < rowCount; s++)
                        {
                            for (int s2 = 0; s2 < columnCount; s2++)
                            {
                                array[s, s2] = workSheets[a].Rows[s][s2].ToString();
                            }
                        }

                        // Get an Excel Range of the same dimensions 
                        Range range2 = (Range)worksheet.Cells[2, 1];
                        range2 = range2.get_Resize(rowCount, columnCount);
                        // Assign the 2-d array to the Excel Range 
                        range2.set_Value(XlRangeValueDataType.xlRangeValueDefault, array);
                    }

                }
                workbook.SaveCopyAs(filePath);
            }
            finally
            {
                try
                {
                    xlApp.ScreenUpdating = true;
                    xlApp.Quit();
                }
                catch
                {
                    //supress
                }
                excel.Kill();
            }

        }

    }
}
