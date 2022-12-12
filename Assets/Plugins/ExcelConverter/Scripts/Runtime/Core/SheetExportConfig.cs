using Newtonsoft.Json.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using UnityEngine;

namespace ExcelConverter
{
    public class SheetExportConfig
    {
        public List<DataTable> sheets { get; private set; }
        public JsonConfig jsonConfig { get; private set; }

        public SheetExportConfig(JsonConfig jsonConfig)
        {
            this.jsonConfig = jsonConfig;
            this.sheets = new List<DataTable>();
        }

        public void AddSheet(DataTable sheet)
        {
            this.sheets.Add(sheet);
        }

        public bool Match(string tableName)
        {
            return this.jsonConfig.dataList.Contains(tableName);
        }
    }
}
