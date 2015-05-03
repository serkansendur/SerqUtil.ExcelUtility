using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Collections.Specialized;

namespace SerqUtil.ExcelUtility
{
    public class ExcelReader
    {
        public StringCollection GetWorkSheetList(string filePath)
        {
            return ExcelHelper.GetWorkSheetList(filePath);
        }

        private static string NonEmptyRowsFilter(string filePath, string workSheet, bool headers)
        {
            List<string> columnNames = new List<string>();
            foreach(string s in ExcelHelper.GetColumnNames(filePath, workSheet, headers,20))
            {
                columnNames.Add("[" + s + "]");
            }
            if (columnNames.Count == 1)
                return columnNames[0] + " is not null";
          return "(" + string.Join(" is not null or ",columnNames.ToArray()) + " is not null)";
        }

        public DataTable Read(string workSheetName, string filePath, bool headers)
        {
            if (ExcelHelper.GetWorkSheetList(filePath).Contains(workSheetName))
            {
                DataTable dt = new DataTable();
                using (OleDbConnection excelConnection = new OleDbConnection(ExcelHelper.GetConnectionString(filePath,headers)))
                {
                    excelConnection.Open();
                    using (OleDbDataAdapter excelAdapter = new OleDbDataAdapter("SELECT * FROM [" +
                       workSheetName + "] where " + NonEmptyRowsFilter(filePath, workSheetName, headers)
                    , excelConnection))
                    {
                        excelAdapter.Fill(dt);
                    }
                    excelConnection.Close();
                }
                return dt;
            }
            else
                throw new Exception("The specified worksheet was not found!");
        }

    }
}
