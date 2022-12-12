using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Threading.Tasks;
using UnityEngine.UIElements;
using System.Linq;

namespace ExcelConverter
{
    public class FileExporter
    {
        class ExportDataDefine
        {
            public const string keyExportType = "export_type";
            public const string keyExportData = "data";
            public const string exportTypeJObject = "JObject";
            public const string exportTypeJArray = "JArray";
        }

        private readonly List<SheetExportConfig> listSheetExportConfig;
        private readonly string savePath;
        private readonly bool formatJson;
        private readonly ProgressBar progressBar;

        public FileExporter(List<SheetExportConfig> listSheetExportConfig, string savePath, bool formatJson, ProgressBar progressBar)
        {
            this.listSheetExportConfig = listSheetExportConfig;
            this.savePath = savePath;
            this.progressBar = progressBar;
            this.formatJson = formatJson;
        }

        public async Task<int> SaveJsonFiles()
        {
            return await this._ReadAllSheets(new Func<SheetExportConfig, Task<int>>[] { this._SaveSheetJson });
        }

        public async Task<int> SaveBsonFiles()
        {
            return await this._ReadAllSheets(new Func<SheetExportConfig, Task<int>>[] { this._SaveSheetBson });
        }

        public async Task<int> SaveBothFiles()
        {
            return await this._ReadAllSheets(new Func<SheetExportConfig, Task<int>>[] { this._SaveSheetJson, this._SaveSheetBson });
        }

        private async Task<int> _ReadAllSheets(Func<SheetExportConfig, Task<int>>[] exportFuncs)
        {
            if (this.listSheetExportConfig.Count <= 0)
            {
                LogAdder.AddLogError("沒有讀取到任何Excel!!");
                return 0;
            }

            this.progressBar.lowValue = 0;
            this.progressBar.value = 0;
            this.progressBar.highValue = this.listSheetExportConfig.Sum(x => x.sheets.Count) * exportFuncs.Length;

            int result = 0;

            foreach (var exportFunc in exportFuncs)
            {
                foreach (var sheetExportConfig in this.listSheetExportConfig)
                {
                    result += await exportFunc(sheetExportConfig);
                }
            }

            return result;
        }

        private async Task<int> _SaveSheetJson(SheetExportConfig sheetExportConfig)
        {
            int result = 0;

            foreach (var sheet in sheetExportConfig.sheets)
            {
                // 檢測Sheet名稱是否有重名
                string adjustedName = this._VerifyDuplicatedName(sheet);

                // 更新進度條顯示
                this.progressBar.value++;
                this.progressBar.title = $" ({this.progressBar.value}/{this.progressBar.highValue}) 【{adjustedName}.json】";

                if (sheet.Rows.Count <= 0)
                {
                    LogAdder.AddLogError(string.Format("內容為空 文件名:{0}", sheet.TableName));
                    continue;
                }

                // 將Sheet的資料序列化成Json字串
                string json = await this._SerializeSheetToJson(sheet, sheetExportConfig.jsonConfig, this.formatJson);

                // 儲存檔案
                string dstFolder = this.savePath + "/json/";

                // 如果沒有資料夾就創一個
                if (!Directory.Exists(dstFolder))
                {
                    Directory.CreateDirectory(dstFolder);
                }

                string path = dstFolder + adjustedName + ".json";

                // 輸出檔案至目標路徑
                File.WriteAllText(path, json);

                result++;

                await Task.Yield();
            }

            return result;
        }

        private async Task<int> _SaveSheetBson(SheetExportConfig sheetExportConfig)
        {
            int result = 0;

            foreach (var sheet in sheetExportConfig.sheets)
            {
                // 檢測Sheet名稱是否有重名
                string adjustedName = this._VerifyDuplicatedName(sheet);

                // 更新進度條顯示
                this.progressBar.value++;
                this.progressBar.title = $"({this.progressBar.value}/{this.progressBar.highValue}) 【{adjustedName}.bytes】";

                if (sheet.Rows.Count <= 0)
                {
                    LogAdder.AddLogError(string.Format("內容為空 文件名:{0}", sheet.TableName));
                    continue;
                }

                // 將Sheet的資料序列化成Json字串
                string json = await this._SerializeSheetToJson(sheet, sheetExportConfig.jsonConfig, this.formatJson);

                // 儲存檔案
                string dstFolder = this.savePath + "/bson/";

                // 如果沒有資料夾就創一個
                if (!Directory.Exists(dstFolder))
                {
                    Directory.CreateDirectory(dstFolder);
                }

                // 將Json字串轉為JObject
                JObject jObj = JObject.Parse(json);

                byte[] bson = jObj.ToBson();

                string path = dstFolder + adjustedName + ".bytes";

                // 輸出檔案至目標路徑
                File.WriteAllBytes(path, bson);

                result++;

                await Task.Yield();
            }

            return result;
        }

        private async Task<string> _SerializeSheetToJson(DataTable sheet, JsonConfig jsonConfig, bool formatJson)
        {
            int columns = sheet.Columns.Count;
            int rows = sheet.Rows.Count;
            string json = "";

            // 如果有主Key
            if (jsonConfig.enableMainKey)
            {
                // 宣告存放MainKey的Column
                List<int> mainKeyColumns = new List<int>();
                switch (jsonConfig.mainKeySelectType)
                {
                    case JsonConfig.MainKeySelectType.FirstNColumn:
                        for (int i = 0; i < jsonConfig.columnOfMainKey; i++)
                        {
                            mainKeyColumns.Add(i);
                        }
                        break;
                    case JsonConfig.MainKeySelectType.SpecificColumn:
                        string specificColumns = jsonConfig.columnOfMainKey.ToString();
                        for (int i = 0; i < specificColumns.Length; i++)
                        {
                            int column = Convert.ToInt32(specificColumns[i].ToString()) - 1;
                            mainKeyColumns.Add(column);
                        }
                        break;
                }

                // 如果有副Key
                if (jsonConfig.enableSubKey)
                {
                    // 使用Map包Map處理
                    Dictionary<string, Dictionary<string, object>> columnData = new Dictionary<string, Dictionary<string, object>>();

                    // 由於Excel內Index(或是說外人眼中的Index)跟程式中不同，故減1
                    // 以"列"做讀取
                    for (int i = jsonConfig.rowOfFirstData - 1; i < rows; i++)
                    {
                        // 存放列資料
                        Dictionary<string, object> rowData = new Dictionary<string, object>();

                        // 以"欄"去設置副Key及資料
                        for (int j = 0; j < columns; j++)
                        {
                            // 副Key將不包含主Key欄之資料
                            if (mainKeyColumns.Contains(j)) continue;

                            string subKey = sheet.Rows[jsonConfig.rowOfSubKey - 1][j].ToString().TrimEnd();  // 取第subKeyRow列的j欄當副Key，Index同樣做減1處理                                                                         

                            // 刪除副Key中的@#字元                                                             
                            if (subKey.Contains("@") || subKey.Contains("#")) subKey = subKey.Replace("@", string.Empty).Replace("#", string.Empty);

                            if (subKey.ToUpper() == "EMPTY" || subKey.Contains("*") || subKey == "") continue; // 假設副Key資料為empty或*或空字元 跳過

                            // 副Key大小寫設定
                            switch (jsonConfig.subKeyType)
                            {
                                case JsonConfig.KeyType.None:
                                    break;
                                case JsonConfig.KeyType.Lowercase:
                                    subKey = subKey.ToLower();
                                    break;
                                case JsonConfig.KeyType.Uppercase:
                                    subKey = subKey.ToUpper();
                                    break;
                            }

                            // 如果資料可以被轉換成int
                            if (int.TryParse(sheet.Rows[i][j].ToString(), out int result))
                            {
                                rowData[subKey] = result; // 從第i列取j欄當副Key的資料並轉成int
                            }
                            else
                            {
                                string strValue = sheet.Rows[i][j].ToString(); // 從第i列取j欄當副Key的資料
                                try
                                {
                                    rowData[subKey] = (strValue.IndexOfAny(new char[] { '[', ']' }, 0) != -1) ? JsonConvert.DeserializeObject(strValue) : strValue;
                                }
                                catch (Exception ex)
                                {
                                    UnityEngine.Debug.LogError(ex);
                                    UnityEngine.Debug.Log(string.Format("This value is <color=#FFBD00>{0}</color> in TableName: <color=#00FF37>{1}</color>, <color=#FF0000>Cannot convert to array, please to check your value of field format.</color>", strValue, sheet.TableName));
                                }
                            }
                        }

                        // 開始設置主Key
                        string mainKey = "";

                        for (int k = 0; k < mainKeyColumns.Count; k++)
                        {
                            string accMainKey = "";

                            int column = mainKeyColumns[k];
                            // 防止IndexOutOfRange
                            if (column < 0 || column >= sheet.Rows[i].ItemArray.Length)
                            {
                                UnityEngine.Debug.LogError($"IndexOutOfRange!! 無法新增主Key。請檢查資料表 : {sheet.TableName}的第{i + 1}列，第{column + 1}欄是否有資料");
                                LogAdder.AddLogError($"IndexOutOfRange!! 無法新增主Key。請檢查資料表 : {sheet.TableName}的第{i + 1}列，第{column + 1}欄是否有資料");
                                continue;
                            }

                            accMainKey = sheet.Rows[i][column].ToString();
                            if (string.IsNullOrEmpty(accMainKey))
                            {
                                UnityEngine.Debug.LogError($"無主Key資料!! 請檢查資料表 : {sheet.TableName}的第{column + 1}欄，第{i + 1}列是否有資料");
                                LogAdder.AddLogError($"無主Key資料!! 請檢查資料表 : {sheet.TableName}的第{column + 1}欄，第{i + 1}列是否有資料");
                            }
                            else
                            {
                                mainKey += accMainKey; // 把第i列的第mainKeyColumn個(欄)數據當主key (因為是多Key 所以這裡String會做累加)
                            }
                        }

                        if (string.IsNullOrEmpty(mainKey))
                        {
                            UnityEngine.Debug.LogError($"組成的主Key資料為空!! SheetName : {sheet.TableName}. 請檢查對應資料表或是設定檔");
                            LogAdder.AddLogError($"組成的主Key資料為空!! SheetName : {sheet.TableName}. 請檢查對應資料表或是設定檔");
                        }
                        else
                        {
                            try
                            {
                                // 主Key大小寫設定
                                switch (jsonConfig.mainKeyType)
                                {
                                    case JsonConfig.KeyType.None:
                                        break;
                                    case JsonConfig.KeyType.Lowercase:
                                        mainKey = mainKey.ToLower();
                                        break;
                                    case JsonConfig.KeyType.Uppercase:
                                        mainKey = mainKey.ToUpper();
                                        break;
                                }

                                columnData.Add(mainKey, rowData); // 加入一列主Key及副Key含子資料
                            }
                            catch
                            {
                                UnityEngine.Debug.LogError(string.Format("【{0}】 with the same key has already been added. Key: {1}", sheet.TableName, mainKey));
                                LogAdder.AddLogError(string.Format("【{0}】 with the same key has already been added. Key: {1}", sheet.TableName, mainKey));
                            }
                        }
                    }

                    Dictionary<string, object> tData = new Dictionary<string, object>()
                    {
                        {ExportDataDefine.keyExportType,ExportDataDefine.exportTypeJObject},
                        {ExportDataDefine.keyExportData,columnData}
                    };

                    // 將序列化後的Json轉為字串
                    json = JsonConvert.SerializeObject(tData, formatJson ? Formatting.Indented : Formatting.None);
                    return json;
                }
                // 如果沒副Key
                else
                {
                    // 使用Map包Array處理
                    Dictionary<string, List<object>> columnData = new Dictionary<string, List<object>>();

                    // 由於Excel內Index(或是說外人眼中的Index)跟程式中不同，故減1
                    // 以"列"做讀取
                    for (int i = jsonConfig.rowOfFirstData - 1; i < rows; i++)
                    {
                        List<object> rowData = new List<object>();

                        // 以"欄"去設置資料
                        for (int j = 0; j < columns; j++)
                        {
                            if (mainKeyColumns.Contains(j)) continue;

                            // 從主key欄位之後的欄位為起點開始放副資料
                            if (int.TryParse(sheet.Rows[i][j].ToString(), out int result)) // 如果資料可以被轉換成int
                            {
                                rowData.Add(result); // 從第i列取j欄當資料並轉成int
                            }
                            else
                            {
                                string strValue = sheet.Rows[i][j].ToString(); // 從第i列取j欄當資料
                                try
                                {
                                    rowData.Add((strValue.IndexOfAny(new char[] { '[', ']' }, 0) != -1) ? JsonConvert.DeserializeObject(strValue) : strValue);
                                }
                                catch (Exception ex)
                                {
                                    UnityEngine.Debug.LogError(ex);
                                    UnityEngine.Debug.Log(string.Format("This value is <color=#FFBD00>{0}</color> in TableName: <color=#00FF37>{1}</color>, <color=#FF0000>Cannot convert to array, please to check your value of field format.</color>", strValue, sheet.TableName));
                                }
                            }
                        }

                        // 開始設置主Key
                        string mainKey = "";

                        for (int k = 0; k < mainKeyColumns.Count; k++)
                        {
                            string accMainKey = "";

                            int column = mainKeyColumns[k];
                            // 防止IndexOutOfRange
                            if (column < 0 || column >= sheet.Rows[i].ItemArray.Length)
                            {
                                UnityEngine.Debug.LogError($"IndexOutOfRange!! 無法新增主Key。請檢查資料表 : {sheet.TableName}的第{i + 1}列，第{column + 1}欄是否有資料");
                                LogAdder.AddLogError($"IndexOutOfRange!! 無法新增主Key。請檢查資料表 : {sheet.TableName}的第{i + 1}列，第{column + 1}欄是否有資料");
                                continue;
                            }

                            accMainKey = sheet.Rows[i][column].ToString();
                            if (string.IsNullOrEmpty(accMainKey))
                            {
                                UnityEngine.Debug.LogError($"無主Key資料!! 請檢查資料表 : {sheet.TableName}的第{column + 1}欄，第{i + 1}列是否有資料");
                                LogAdder.AddLogError($"無主Key資料!! 請檢查資料表 : {sheet.TableName}的第{column + 1}欄，第{i + 1}列是否有資料");
                            }
                            else
                            {
                                mainKey += accMainKey; // 把第i列的第mainKeyColumn個(欄)數據當主key (因為是多Key 所以這裡String會做累加)
                            }
                        }

                        if (string.IsNullOrEmpty(mainKey))
                        {
                            UnityEngine.Debug.LogError($"組成的主Key資料為空!! SheetName : {sheet.TableName}. 請檢查對應資料表或是設定檔");
                            LogAdder.AddLogError($"組成的主Key資料為空!! SheetName : {sheet.TableName}. 請檢查對應資料表或是設定檔");
                        }
                        else
                        {
                            try
                            {
                                // 主Key大小寫設定
                                switch (jsonConfig.mainKeyType)
                                {
                                    case JsonConfig.KeyType.None:
                                        break;
                                    case JsonConfig.KeyType.Lowercase:
                                        mainKey = mainKey.ToLower();
                                        break;
                                    case JsonConfig.KeyType.Uppercase:
                                        mainKey = mainKey.ToUpper();
                                        break;
                                }

                                columnData.Add(mainKey, rowData); // 加入一列主Key及副Key含子資料
                            }
                            catch
                            {
                                UnityEngine.Debug.LogError(string.Format("【{0}】 with the same key has already been added. Key: {1}", sheet.TableName, mainKey));
                                LogAdder.AddLogError(string.Format("【{0}】 with the same key has already been added. Key: {1}", sheet.TableName, mainKey));
                            }
                        }
                    }

                    Dictionary<string, object> tData = new Dictionary<string, object>()
                    {
                        {ExportDataDefine.keyExportType,ExportDataDefine.exportTypeJObject},
                        {ExportDataDefine.keyExportData,columnData}
                    };

                    // 將序列化後的Json轉為字串
                    json = JsonConvert.SerializeObject(tData, formatJson ? Formatting.Indented : Formatting.None);
                    return json;
                }
            }
            // 如果沒有主Key
            else
            {
                // 使用List(Array)處理
                List<object> columnData = new List<object>();

                // 由於Excel內Index(或是說外人眼中的Index)跟程式中不同，故減1
                // 以"列"做讀取
                for (int i = jsonConfig.rowOfFirstData - 1; i < rows; i++)
                {
                    // 如果有副Key
                    if (jsonConfig.enableSubKey)
                    {
                        // 存放列資料
                        Dictionary<string, object> rowData = new Dictionary<string, object>();

                        // 以"欄"去設置副Key及資料
                        for (int j = 0; j < columns; j++)
                        {
                            string subKey = sheet.Rows[jsonConfig.rowOfSubKey - 1][j].ToString().TrimEnd(); // 取第subKeyRow列的j欄當副Key，Index同樣做減1處理                                                                         

                            // 刪除副Key中的@#字元                                                             
                            if (subKey.Contains("@") || subKey.Contains("#")) subKey = subKey.Replace("@", string.Empty).Replace("#", string.Empty);

                            if (subKey.ToUpper() == "EMPTY" || subKey.Contains("*") || subKey == "") continue; // 假設副Key資料為empty或*或空字元 跳過

                            // 副Key大小寫設定                                                                 
                            switch (jsonConfig.subKeyType)
                            {
                                case JsonConfig.KeyType.None:
                                    break;
                                case JsonConfig.KeyType.Lowercase:
                                    subKey = subKey.ToLower();
                                    break;
                                case JsonConfig.KeyType.Uppercase:
                                    subKey = subKey.ToUpper();
                                    break;
                            }

                            if (int.TryParse(sheet.Rows[i][j].ToString(), out int result)) //如果資料可以被轉換成int
                            {
                                rowData[subKey] = result; //從第i列取j欄當副Key的資料並轉成int
                            }
                            else
                            {
                                string strValue = sheet.Rows[i][j].ToString(); // 從第i列取j欄當副Key的資料
                                try
                                {
                                    rowData[subKey] = (strValue.IndexOfAny(new char[] { '[', ']' }, 0) != -1) ? JsonConvert.DeserializeObject(strValue) : strValue;
                                }
                                catch (Exception ex)
                                {
                                    UnityEngine.Debug.LogError(ex);
                                    UnityEngine.Debug.Log(string.Format("This value is <color=#FFBD00>{0}</color> in TableName: <color=#00FF37>{1}</color>, <color=#FF0000>Cannot convert to array, please to check your value of field format.</color>", strValue, sheet.TableName));
                                }
                            }
                        }

                        columnData.Add(rowData);   // 加入副Key含子資料
                    }
                    // 如果沒有副Key
                    else
                    {
                        List<object> rowData = new List<object>();

                        // 以"欄"去設置資料
                        for (int j = 0; j < columns; j++)
                        {
                            if (int.TryParse(sheet.Rows[i][j].ToString(), out int result)) //如果資料可以被轉換成int
                            {
                                rowData.Add(result); //從第i列取j欄當資料並轉成int
                            }
                            else
                            {
                                string strValue = sheet.Rows[i][j].ToString(); // 從第i列取j欄當副Key的資料
                                try
                                {
                                    rowData.Add((strValue.IndexOfAny(new char[] { '[', ']' }, 0) != -1) ? JsonConvert.DeserializeObject(strValue) : strValue);
                                }
                                catch (Exception ex)
                                {
                                    UnityEngine.Debug.LogError(ex);
                                    UnityEngine.Debug.Log(string.Format("This value is <color=#FFBD00>{0}</color> in TableName: <color=#00FF37>{1}</color>, <color=#FF0000>Cannot convert to array, please to check your value of field format.</color>", strValue, sheet.TableName));
                                }

                            }
                        }

                        columnData.Add(rowData);
                    }
                }

                await Task.Yield();

                Dictionary<string, object> tData = new Dictionary<string, object>()
                {
                    {ExportDataDefine.keyExportType, ExportDataDefine.exportTypeJArray},
                    {ExportDataDefine.keyExportData, columnData}
                };

                // 將序列化後的Json轉為字串
                json = JsonConvert.SerializeObject(tData, formatJson ? Formatting.Indented : Formatting.None);
                return json;
            }
        }

        private string _VerifyDuplicatedName(DataTable sheet)
        {
            for (int i = 0; i < this.listSheetExportConfig.Count; i++)
            {
                for (int j = 0; j < this.listSheetExportConfig[i].sheets.Count; j++)
                {
                    // 如果兩個Sheet隸屬的Excel檔案不同，但名稱相同，就算重名
                    if (this.listSheetExportConfig[i].sheets[j].Namespace != sheet.Namespace && this.listSheetExportConfig[i].sheets[j].TableName == sheet.TableName)
                    {
                        string adjustedName = "[" + sheet.Namespace + "]" + sheet.TableName;
                        LogAdder.AddLogError(string.Format("檢測到Sheet有重名 Sheet名稱:{0}，輸出檔名將會更改為:{1}!", sheet.TableName, adjustedName));
                        return adjustedName;
                    }
                }
            }
            return sheet.TableName;
        }
    }
}
