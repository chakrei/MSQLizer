using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI.WebControls;

namespace FileToDB.Models
{
    public class ExcelHelper
    {
        private static ExcelHelper _ExcelHelper;

        public static ExcelHelper GetExcelHelperInstance()
        {
            if (_ExcelHelper == null)
                _ExcelHelper = new ExcelHelper();
            return _ExcelHelper;
        }

        public List<string> GetSheets(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                throw new FileNotFoundException();

            List<string> sheets = new List<string>();

            using (var fs = File.OpenRead(filePath))
            using (DataSet fileDS = ExcelReaderFactory.CreateReader(fs).AsDataSet())
                foreach (DataTable table in fileDS.Tables)
                    sheets.Add(table.TableName);

            return sheets;
        }
        public List<string> GetColumns(string filePath, string sheetName, bool firstrow)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                throw new FileNotFoundException();

            List<string> columns = new List<string>();

            using (var fs = File.OpenRead(filePath))
            using (DataSet fileDS = ExcelReaderFactory.CreateReader(fs).AsDataSet())
                if (!firstrow)
                    foreach (DataColumn column in fileDS.Tables[sheetName].Columns)
                        columns.Add(column.ColumnName);
                else
                {
                    var sheetColumns = fileDS.Tables[sheetName]?.Columns;
                    DataRow firstRow = null;
                    if (fileDS.Tables[sheetName].Rows.Count > 0)
                        firstRow = fileDS.Tables[sheetName].Rows[0];

                    foreach (DataColumn column in sheetColumns)
                        columns.Add(firstRow[column.ColumnName].ToString());
                }

            return columns;
        }
        private string PrepareInsertStatement(ScriptModel script)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("\nGO\n");
            sb.AppendLine($"INSERT INTO {script.DBDetails.TableName}");
            sb.Append("(");

            sb.Append(string.Join(",", script.Rows.Select(x => "'" + x.ColumnName + "'")));
            sb.AppendLine(")");
            sb.AppendLine("VALUES");
            return sb.ToString();
        }
        public string GenerateScript(string filePath, ScriptModel script)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                throw new FileNotFoundException();

            StringBuilder sb = new StringBuilder();

            if (!string.IsNullOrEmpty(script.DBDetails.InitialCatalog))
                sb.AppendLine($"USE {script.DBDetails.InitialCatalog} \nGO");


            if (string.IsNullOrEmpty(script.DBDetails.TableName))
                script.DBDetails.TableName = "PlaceHolderTable";

            sb.AppendLine($"IF NOT EXISTS(SELECT 1 FROM SYS.OBJECTS WHERE NAME = '{script.DBDetails.TableName}')");
            sb.AppendLine($"CREATE TABLE {script.DBDetails.TableName} (");
            foreach (Row col in script.Rows)
                sb.AppendLine($"[{col.ColumnName}] { (col.DataType.ToLower() == "string" ? "VARCHAR(MAX)" : col.DataType) }");
            sb.AppendLine(")");
            sb.AppendLine("GO");

            using (var fs = File.OpenRead(filePath))
            using (DataSet fileDS = ExcelReaderFactory.CreateReader(fs).AsDataSet())
            {
                int rowIndex = 0;
                foreach (DataRow row in fileDS.Tables[script.SheetName].Rows)
                {
                    StringBuilder tableValues = new StringBuilder();
                    if (rowIndex == 0)
                        sb.AppendLine(PrepareInsertStatement(script));

                    rowIndex++;
                    foreach (var col in script.Rows)
                    {
                        if (col.DataType.ToLower() == "datetime")
                            tableValues.Append($"'{ DateTime.ParseExact(Convert.ToString(row[col.ColumnName]), col.ExtraInfo, CultureInfo.InvariantCulture)}',");
                        else
                        {
                            string cellValue = Convert.ToString(row[col.ColumnName]);
                            col.ExtraInfo = col.ExtraInfo ?? "";

                            foreach (var ch in col.ExtraInfo.ToCharArray())
                                cellValue = cellValue.Replace(ch.ToString(), "\\" + ch);

                            tableValues.Append($"'{cellValue }',");
                        }
                    }
                    sb.AppendLine("(");
                    sb.AppendLine(tableValues.ToString().TrimEnd(','));
                    sb.Append("),");

                    if (rowIndex == 500)
                        rowIndex = 0;
                }
            }

            var targetFilePath = Path.Combine(Directory.GetParent(filePath).FullName, Path.GetFileNameWithoutExtension(filePath) + "_Output.sql");
            File.WriteAllText(targetFilePath, sb.ToString().TrimEnd(','));
            return Path.GetFileName(targetFilePath);
        }
    }
}