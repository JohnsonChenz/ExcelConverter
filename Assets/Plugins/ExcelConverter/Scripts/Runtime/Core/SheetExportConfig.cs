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
        public bool ableToExport { get; private set; }                // 是否可被輸出
        public JsonConfig.KeyType mainKeyType { get; private set; }   // 主Key類型, None = 無, Lower = 轉小寫, Upper = 轉大寫
        public JsonConfig.KeyType subKeyType { get; private set; }    // 副Key類型, None = 無, Lower = 轉小寫, Upper = 轉大寫
        public int mainKeyColumn { get; private set; }                // 主Key欄(可以有多個 用逗號相隔 會在初始時就做剔除)
        public int subKeyRow { get; private set; }                    // 副Key列
        public int firstDataRow { get; private set; }                 // 數據資料起始列

        public SheetExportConfig(DataTable sheet, bool ableToExport, JsonConfig.KeyType mainKeyType = JsonConfig.KeyType.None, JsonConfig.KeyType subKeyType = JsonConfig.KeyType.None, int mainKeyColumn = 0, int subKeyRow = -1, int firstDataRow = -1)
        {
            this.sheet = sheet;
            this.ableToExport = ableToExport;
            this.mainKeyType = mainKeyType;
            this.subKeyType = subKeyType;
            this.mainKeyColumn = mainKeyColumn;
            this.subKeyRow = subKeyRow;
            this.firstDataRow = firstDataRow;
        }
    }
}
