using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections.Specialized;
using System.Data.OleDb;
using System.Data;

namespace SerqUtil.ExcelUtility
{
    public class ExcelHelper
    {
        public static string GetConnectionString(string filePath,bool headers)
        {
            string connectionString = string.Empty;
            switch (System.IO.Path.GetExtension(filePath))
            {
                case ".xls":
                    connectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source= " +
                        filePath + ";Extended Properties=\"Excel 8.0;IMEX=1;HDR=" + 
                        (headers ? "YES" : "NO")
                        + "\"";
                    break;
                case ".xlsx":
                    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source= " +
                        filePath + ";Extended Properties=\"Excel 12.0;IMEX=1;HDR=" +
                         (headers ? "YES" : "NO")
                        + "\"";
                    break;
            }
            return connectionString;
        }

        public static StringCollection GetWorkSheetList(string filePath)
        {
            using (OleDbConnection excelConnection = new OleDbConnection(GetConnectionString(filePath,true)))
            {
                StringCollection sheetNames = new StringCollection();
                excelConnection.Open();
                DataTable worksheets = excelConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (worksheets == null)
                    throw new Exception("no worksheets found!");
                else
                    foreach (DataRow row in worksheets.Rows)
                        sheetNames.Add(row["TABLE_NAME"].ToString());

                excelConnection.Close();
                return sheetNames;
            }
        }

        public static StringCollection GetColumnNames(string filePath, string workSheet, bool headers, int? count = null)
        {
            using (OleDbConnection excelConnection = new OleDbConnection(GetConnectionString(filePath, headers)))
            {
                StringCollection columnNames = new StringCollection();
                excelConnection.Open();
                String[] restriction = { null, null, workSheet, null };
                DataTable dtColumns = excelConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Columns, restriction);

                if (dtColumns == null)
                    throw new Exception("no columns found!");
                else
                {
                    if(count == null || count >= dtColumns.Rows.Count)
                     foreach (DataRow row in dtColumns.Rows)
                        columnNames.Add(row["COLUMN_NAME"].ToString());
                    else
                    {
                        for (int i = 0; i < count; i++)
                            columnNames.Add(dtColumns.Rows[i]["COLUMN_NAME"].ToString());
                    }
                       
                }

                excelConnection.Close();
                return columnNames;
            }
        }
    }
}
