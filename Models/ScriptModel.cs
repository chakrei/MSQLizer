using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace FileToDB.Models
{
    public class ScriptModel
    {
        public string SheetName { get; set; }
        public bool Firstrow { get; set; }
        public List<Row> Rows { get; set; }
        public DBModel DBDetails { get; set; }
    }

    public class Row
    {
        public string ColumnName { get; set; }
        public string DataType { get; set; }
        public string ExtraInfo { get; set; }
    }

    public class DBModel
    {
        public string TableName { get; set; }
        public string InstanceName { get; set; }
        public string InitialCatalog { get; set; }
        public string Password { get; set; }
    }
}