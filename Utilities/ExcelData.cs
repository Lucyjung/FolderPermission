using ExcelDataReader;
using System.Data;
using System.IO;

namespace FolderPermission.Utilities
{
    class ExcelData
    {
        public static DataTable readData(string filePath, int sheetIndex = 0)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    // Choose one of either 1 or 2:

                    // 1. Use the reader methods
                    do
                    {
                        while (reader.Read())
                        {
                            // reader.GetDouble(0);
                        }
                    } while (reader.NextResult());

                    // 2. Use the AsDataSet extension method
                    var result = reader.AsDataSet();

                    // The result of each spreadsheet is in result.Tables
                    var dt = result.Tables[sheetIndex];

                    foreach (DataColumn column in dt.Columns)
                    {
                        string cName = dt.Rows[0][column.ColumnName].ToString();
                        if (!dt.Columns.Contains(cName) && cName != "")
                        {
                            column.ColumnName = cName;
                        }

                    }
                    dt.Rows.Remove(dt.Rows[0]);
                    return dt;
                }
            }
        }
    }
}
