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
        public DataTable sheet { get; private set; }
        public JsonConfig jsonConfig { get; private set; }

        public SheetExportConfig(DataTable sheet, JsonConfig jsonConfig)
        {
            this.sheet = sheet;
            this.jsonConfig = jsonConfig;
        }
    }
}
