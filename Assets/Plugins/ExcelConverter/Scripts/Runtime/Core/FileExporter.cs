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
        public class ExportDataDefine
        {
            public const string keyExportType = "export_type";
            public const string keyExportData = "data";
            public const string exportTypeJObject = "JObject";
            public const string exportTypeJArray = "JArray";
        }

        private readonly List<SheetExportConfig> exportSettingList;
        private readonly string savePath;
        private readonly int formattingOption;
        private readonly ProgressBar progressBar;

        public FileExporter(List<SheetExportConfig> exportSettingList, string savePath, bool isJsonFormatting, ProgressBar progressBar)
        {
            this.exportSettingList = exportSettingList;
            this.savePath = savePath;
            this.progressBar = progressBar;

            if (isJsonFormatting)
            {
                this.formattingOption = 1;
            }
            else
            {
                this.formattingOption = 0;
            }
        }

        public async Task<int> SaveJsonFiles()
        {
            return await this._ReadAllSheets(new Func<SheetExportConfig, string, Task<int>>[] { this._SaveSheetJson });
        }

        public async Task<int> SaveBsonFiles()
        {
            return await this._ReadAllSheets(new Func<SheetExportConfig, string, Task<int>>[] { this._SaveSheetBson });
        }

        public async Task<int> SaveBothFiles()
        {
            return await this._ReadAllSheets(new Func<SheetExportConfig, string, Task<int>>[] { this._SaveSheetJson, this._SaveSheetBson });
        }

        private async Task<int> _ReadAllSheets(Func<SheetExportConfig, string, Task<int>>[] exportFuncs)
        {
            if (this.exportSettingList.Count <= 0)
            {
                LogAdder.AddLogError("沒有讀取到任何Excel!!");
                return 0;
            }

            this.progressBar.lowValue = 0;
            this.progressBar.highValue = this.exportSettingList.Count(x => x.ableToExport) * exportFuncs.Length;

            int result = 0;

            foreach (var exportFunc in exportFuncs)
            {
                for (var i = 0; i < this.exportSettingList.Count; i++)
                {
                    if (this.exportSettingList[i].ableToExport)
                    {
                        // 檢測Sheet名稱是否有重名
                        string adjustedName = this._VerifyDuplicatedName(this.exportSettingList[i].sheet);

                        this.progressBar.value = result += await exportFunc(this.exportSettingList[i], adjustedName);

                        await Task.Yield();
                    }
                    else
                    {
                        //Debug.LogError("此文件沒有在設定檔資料中或設定檔有誤 判定為無法輸出:" + this.exportSettings[i].sheet.TableName);
                    }
                }
            }

            return result;
        }

        private async Task<int> _SaveSheetJson(SheetExportConfig exportSettings, string fileName)
        {
            if (exportSettings.sheet.Rows.Count <= 0)
            {
                LogAdder.AddLogError(string.Format("內容為空 文件名:{0}", exportSettings.sheet.TableName));
                return 0;
            }

            // 將Sheet的資料序列化成Json字串
            string json = await this._SerializeSheetToJson(exportSettings.sheet, exportSettings, this.formattingOption);

            // 儲存檔案
            string dstFolder = this.savePath + "/json/";

            // 如果沒有資料夾就創一個
            if (!Directory.Exists(dstFolder))
            {
                Directory.CreateDirectory(dstFolder);
            }

            string path = dstFolder + fileName + ".json";

            // 更新進度調顯示
            this.progressBar.title = $"({this.progressBar.value}/{this.progressBar.highValue}) 【{fileName}.json】";

            // 輸出檔案至目標路徑
            File.WriteAllText(path, json);

            return 1;
        }

        private async Task<int> _SaveSheetBson(SheetExportConfig exportSettings, string fileName)
        {
            if (exportSettings.sheet.Rows.Count <= 0)
            {
                LogAdder.AddLogError(string.Format("內容為空 文件名:{0}", exportSettings.sheet.TableName));
                return 0;
            }

            // 將Sheet的資料序列化成Json字串
            string json = await this._SerializeSheetToJson(exportSettings.sheet, exportSettings, this.formattingOption);

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

            string path = dstFolder + fileName + ".bytes";

            // 更新進度調顯示
            this.progressBar.title = $"({this.progressBar.value}/{this.progressBar.highValue}) 【{fileName}.bytes】";

            // 輸出檔案至目標路徑
            File.WriteAllBytes(path, bson);

            return 1;
        }

        private async Task<string> _SerializeSheetToJson(DataTable sheet, SheetExportConfig exportSettings, int FormattingOption = 0)
        {
            int columns = sheet.Columns.Count;
            int rows = sheet.Rows.Count;
            string json = "";

            // 如果有主Key
            if (exportSettings.mainKeyColumn > 0)
            {
                // 如果有副Key
                if (exportSettings.subKeyRow > 0)
                {
                    // 使用Map包Map處理
                    Dictionary<string, Dictionary<string, object>> columnData = new Dictionary<string, Dictionary<string, object>>();

                    // 由於Excel內Index(或是說外人眼中的Index)跟程式中不同，故減1
                    // 以"列"做讀取
                    for (int i = exportSettings.firstDataRow - 1; i < rows; i++)
                    {
                        // 存放列資料
                        Dictionary<string, object> rowData = new Dictionary<string, object>();

                        // 以"欄"去設置副Key及資料
                        for (int j = 0; j < columns; j++)
                        {
                            // 副Key將不包含主Key欄之資料
                            if (j < exportSettings.mainKeyColumn) continue;

                            string subKey = sheet.Rows[exportSettings.subKeyRow - 1][j].ToString().TrimEnd();  // 取第subKeyRow列的j欄當副Key，Index同樣做減1處理                                                                         

                            // 刪除副Key中的@#字元                                                             
                            if (subKey.Contains("@") || subKey.Contains("#")) subKey = subKey.Replace("@", string.Empty).Replace("#", string.Empty);

                            if (subKey.ToUpper() == "EMPTY" || subKey.Contains("*") || subKey == "") continue; // 假設副Key資料為empty或*或空字元 跳過

                            // 副Key大小寫設定
                            switch (exportSettings.subKeyType)
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

                        if (exportSettings.mainKeyColumn > 0)
                        {
                            for (int k = 0; k < exportSettings.mainKeyColumn; k++)
                            {
                                mainKey += sheet.Rows[i][k].ToString(); // 把第i列的第mainKeyColumn個(欄)數據當主key (因為是多Key 所以這裡String會做累加)
                            }
                        }

                        // 主Key大小寫設定
                        switch (exportSettings.mainKeyType)
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

                        try
                        {
                            columnData.Add(mainKey, rowData); // 加入一列主Key及副Key含子資料
                        }
                        catch
                        {
                            UnityEngine.Debug.LogError(string.Format("【{0}】 with the same key has already been added. Key: {1}", sheet.TableName, mainKey));
                            LogAdder.AddLogError(string.Format("【{0}】 with the same key has already been added. Key: {1}", sheet.TableName, mainKey));
                        }
                    }

                    Dictionary<string, object> tData = new Dictionary<string, object>()
                    {
                        {ExportDataDefine.keyExportType,ExportDataDefine.exportTypeJObject},
                        {ExportDataDefine.keyExportData,columnData}
                    };

                    // 將序列化後的Json轉為字串
                    json = JsonConvert.SerializeObject(tData, (Formatting)FormattingOption);
                    return json;
                }
                // 如果沒副Key
                else
                {
                    // 使用Map包Array處理
                    Dictionary<string, List<object>> columnData = new Dictionary<string, List<object>>();

                    // 由於Excel內Index(或是說外人眼中的Index)跟程式中不同，故減1
                    // 以"列"做讀取
                    for (int i = exportSettings.firstDataRow - 1; i < rows; i++)
                    {
                        List<object> rowData = new List<object>();

                        // 以"欄"去設置資料
                        for (int j = 0; j < columns; j++)
                        {
                            if (j < exportSettings.mainKeyColumn) continue;

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

                        if (exportSettings.mainKeyColumn > 0)
                        {
                            for (int k = 0; k < exportSettings.mainKeyColumn; k++)
                            {
                                mainKey += sheet.Rows[i][k].ToString(); // 把第i列的第mainKeyColumn個(欄)數據當主key (因為是多Key 所以這裡String會做累加)
                            }
                        }

                        // 主Key大小寫設定         
                        switch (exportSettings.mainKeyType)
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

                        try
                        {
                            columnData.Add(mainKey, rowData); // 加入一列主Key及副Key含子資料
                        }
                        catch
                        {
                            UnityEngine.Debug.LogError(string.Format("【{0}】 with the same key has already been added. Key: {1}", sheet.TableName, mainKey));
                            LogAdder.AddLogError(string.Format("【{0}】 with the same key has already been added. Key: {1}", sheet.TableName, mainKey));
                        }
                    }

                    Dictionary<string, object> tData = new Dictionary<string, object>()
                    {
                        {ExportDataDefine.keyExportType,ExportDataDefine.exportTypeJObject},
                        {ExportDataDefine.keyExportData,columnData}
                    };

                    // 將序列化後的Json轉為字串
                    json = JsonConvert.SerializeObject(tData, (Formatting)FormattingOption);
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
                for (int i = exportSettings.firstDataRow - 1; i < rows; i++)
                {
                    // 如果有副Key
                    if (exportSettings.subKeyRow > 0)
                    {
                        // 存放列資料
                        Dictionary<string, object> rowData = new Dictionary<string, object>();

                        // 以"欄"去設置副Key及資料
                        for (int j = 0; j < columns; j++)
                        {
                            string subKey = sheet.Rows[exportSettings.subKeyRow - 1][j].ToString().TrimEnd(); // 取第subKeyRow列的j欄當副Key，Index同樣做減1處理                                                                         

                            // 刪除副Key中的@#字元                                                             
                            if (subKey.Contains("@") || subKey.Contains("#")) subKey = subKey.Replace("@", string.Empty).Replace("#", string.Empty);

                            if (subKey.ToUpper() == "EMPTY" || subKey.Contains("*") || subKey == "") continue; // 假設副Key資料為empty或*或空字元 跳過

                            // 副Key大小寫設定                                                                 
                            switch (exportSettings.subKeyType)
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
                json = JsonConvert.SerializeObject(tData, (Formatting)FormattingOption);
                return json;
            }
        }

        private string _VerifyDuplicatedName(DataTable sheet)
        {
            for (int i = 0; i < exportSettingList.Count; i++)
            {
                // 如果用來比對的檔案沒有要輸出就跳過
                if (exportSettingList[i].ableToExport == false) continue;
                // 如果兩個Sheet隸屬的Excel檔案不同，但名稱相同，就算重名
                if (exportSettingList[i].sheet.Namespace != sheet.Namespace && exportSettingList[i].sheet.TableName == sheet.TableName)
                {
                    string adjustedName = "[" + sheet.Namespace + "]" + sheet.TableName;
                    LogAdder.AddLogError(string.Format("檢測到Sheet有重名 Sheet名稱:{0}，輸出檔名將會更改為:{1}!", sheet.TableName, adjustedName));
                    return adjustedName;
                }
            }
            return sheet.TableName;
        }
    }
}
