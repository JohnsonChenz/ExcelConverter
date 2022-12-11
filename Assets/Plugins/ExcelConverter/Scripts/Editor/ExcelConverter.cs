using System.Collections;
using System.Collections.Generic;
using System.IO;
using System;
using System.Linq;
using UnityEditor;
using UnityEngine;
using UnityEngine.UIElements;
using ExcelDataReader;
using System.Data;
using Newtonsoft.Json.Linq;
using UnityEditor.UIElements;
using System.Threading;
using System.Threading.Tasks;
using System.Text;

namespace ExcelConverter.Editor
{
    public class ExcelConverterEditor : EditorWindow
    {
        public class ExportType
        {
            public const string Json = "Json";
            public const string Bson = "Bson";
            public const string Both = "Both";
        }

        // --- Editor相關
        private bool isEditorRunning = false;

        // --- EditorPrefs相關Key
        private const string keyFilePath = "Excel_FilePath";
        private const string keyOutPath = "Excel_OutPath";
        public const string keyJsonPath = "Json_FilePath";
        private const string keyExportType = "File_ExportType";
        private const string keyJsonFormatting = "Json_Formatting";
        private const string keyAutoScan = "Auto_Scan";

        // --- 檔案相關設定                       
        private List<bool> isUseFiles = null;            // 判斷是否有被勾選
        private List<FileInfo> excelFiles = null;        // 存放Excel
        private List<DataTable> excelsWithSheet = null;  // 存放Excel中的Sheet

        // --- Json設定檔相關
        private List<JsonConfig> jsonConfigList;

        // --- Root
        private VisualElement root;

        // --- 輸出相關
        private int genResult;                           // 輸出結果數量
        private bool formatJson = false;            // Json是否要格式化

        // --- UI相關組件
        private TextField outPathText = null;
        private TextField filePathText = null;
        private TextField jsonPathText = null;
        private Toggle chooseAllTgl = null;
        private Toggle jsonFormattingTgl = null;
        private Toggle autoScanTgl = null;
        private Button exportBtn = null;
        private Button browseOutPathBtn = null;
        private Button browseFilePathBtn = null;
        private Button browseJsonPathBtn = null;
        private Button scanBtn = null;
        private Button clearLogBtn = null;
        private Button jsonGeneratorBtn = null;
        private Button resetBtn = null;
        private Button openFilePathBtn = null;
        private Button openOutPathBtn = null;
        private ScrollView scrollViewExcelFiles = null;
        private PopupField<string> exportOptionPop = null;
        private ProgressBar excelLoadingProgress = null;

        [MenuItem("Plugins/Excel Converter")]
        static void Init()
        {
            var window = GetWindow<ExcelConverterEditor>();
            window.titleContent = new GUIContent("ExcelConverter Tool");
            window.minSize = new Vector2(540, 770);
        }

        private async void OnEnable()
        {
            this.root = rootVisualElement;

            var vta = Resources.Load<VisualTreeAsset>("ExcelConverter");
            vta.CloneTree(this.root);

            // 初始化各組件，路徑以及按鈕
            this._InitComponents();
            this._InitPaths();
            this._InitEvents();
            this._InitSettings();

            this.isEditorRunning = true;
            await Task.Yield();

            if (!string.IsNullOrEmpty(this.filePathText.value))
            {
                if (this.autoScanTgl.value == true)
                {
                    await this._LoadExcelFiles(this.filePathText.value);
                }
                else
                {
                    LogAdder.AddLog("並無開啟AutoScan，請手動點選Scan按鈕掃描所設置路徑的Excel檔");
                }
            }
            else LogAdder.AddLogError("請選擇Excel檔案路徑!!");

            this._SetExportFileControl(this.excelsWithSheet != null && this.excelsWithSheet.Count > 0);
        }

        private void OnGUI()
        {
            LogAdder.OnUpdate();
        }

        private void OnDisable()
        {
            this.isEditorRunning = false;
            GC.Collect();
        }

        private void _InitComponents()
        {
            //this.testBoolBtn = this.root.Q<Button>("TestBool");
            this.exportBtn = this.root.Q<Button>("Export");
            this.chooseAllTgl = this.root.Q<Toggle>("ChooseAll");

            this.outPathText = this.root.Q<TextField>("OutPath");
            this.browseOutPathBtn = this.root.Q<Button>("Browse_OutPath");

            this.filePathText = this.root.Q<TextField>("FilePath");
            this.browseFilePathBtn = this.root.Q<Button>("Browse_FilePath");

            this.jsonPathText = root.Q<TextField>("JsonPath");
            this.browseJsonPathBtn = root.Q<Button>("Browse_JsonPath");

            this.scanBtn = this.root.Q<Button>("Scan");

            this.scrollViewExcelFiles = root.Q<ScrollView>("ScrollViewExcelFiles");

            this.clearLogBtn = root.Q<Button>("ClearLog");

            this.jsonFormattingTgl = root.Q<Toggle>("JsonFormatting");

            this.exportOptionPop = this._InitCustomPopupField(this.exportOptionPop);

            LogAdder.SetScrollView(root.Q<ScrollView>("ScrollViewLog"));

            this.jsonGeneratorBtn = root.Q<Button>("JsonGenerator");

            this.resetBtn = root.Q<Button>("Reset");

            this.openFilePathBtn = root.Q<Button>("Open_FilePath");

            this.openOutPathBtn = root.Q<Button>("Open_OutPath");

            this.autoScanTgl = root.Q<Toggle>("AutoScan");

            this.excelLoadingProgress = root.Q<ProgressBar>("ProgressBar");
            this.SetProgressBarVisble(false);
        }

        private void _InitPaths()
        {
            this.outPathText.value = EditorPrefs.GetString(keyOutPath, Application.dataPath);
            this.filePathText.value = EditorPrefs.GetString(keyFilePath, Application.dataPath);
            this.jsonPathText.value = EditorPrefs.GetString(keyJsonPath, Application.dataPath);
            LogAdder.AddLog(string.Format("檔案路徑:{0}", this.filePathText.value));
            LogAdder.AddLog(string.Format("輸出路徑:{0}", this.outPathText.value));
            LogAdder.AddLog(string.Format("Json設定檔路徑:{0}", this.jsonPathText.value));
        }

        private void _InitSettings()
        {
            this.exportOptionPop.value = EditorPrefs.GetString(keyExportType, ExportType.Json);
            this.jsonFormattingTgl.value = EditorPrefs.GetBool(keyJsonFormatting, false);
            this.autoScanTgl.value = EditorPrefs.GetBool(keyAutoScan, false);

            this._RefreshJsonFormattingTgl();
        }

        private void _InitEvents()
        {
            this.browseOutPathBtn.clickable.clicked += () =>
            {
                // 讀取先前紀錄的路徑
                string tempPath = EditorPrefs.GetString(keyOutPath, Application.dataPath);

                string path = EditorUtility.OpenFolderPanel("Choose Excel Output Files Folder", !string.IsNullOrEmpty(tempPath) ? tempPath : "", "");

                if (!string.IsNullOrEmpty(path))
                {
                    this.outPathText.value = path;
                    EditorPrefs.SetString(keyOutPath, path);
                    LogAdder.AddLog(string.Format("路徑顯示:{0}", path));
                }
            };

            this.browseFilePathBtn.clickable.clicked += async () =>
            {
                // 讀取先前紀錄的路徑
                string tempPath = EditorPrefs.GetString(keyFilePath, Application.dataPath);

                string path = EditorUtility.OpenFolderPanel("Choose Excel Files Folder", !string.IsNullOrEmpty(tempPath) ? tempPath : "", "");

                if (!string.IsNullOrEmpty(path))
                {
                    this.filePathText.value = path;
                    await this._LoadExcelFiles(this.filePathText.value);
                    EditorPrefs.SetString(keyFilePath, path);
                    LogAdder.AddLog(string.Format("路徑顯示:{0}", path));
                }
            };

            this.browseJsonPathBtn.clickable.clicked += async () =>
            {
                // 讀取先前紀錄的路徑
                string tempPath = EditorPrefs.GetString(keyJsonPath, Application.dataPath);
                string fileName = Path.GetFileName(tempPath);
                tempPath = tempPath.Replace(fileName, "");

                string path = EditorUtility.OpenFilePanel("Choose Json Settings File", !string.IsNullOrEmpty(tempPath) ? tempPath : "", "json");

                if (!string.IsNullOrEmpty(path))
                {
                    this.jsonPathText.value = path;
                    EditorPrefs.SetString(keyJsonPath, path);
                    LogAdder.AddLog(string.Format("路徑顯示:{0}", path));
                }
            };

            this.scanBtn.clickable.clicked += async () =>
            {
                if (!string.IsNullOrEmpty(this.filePathText.value))
                {
                    await this._LoadExcelFiles(this.filePathText.value);
                }

                this._SetExportFileControl(this.excelsWithSheet != null && this.excelsWithSheet.Count > 0);
            };

            this.exportBtn.clickable.clicked += async () =>
            {
                if (!string.IsNullOrEmpty(this.outPathText.value))
                {
                    if (this.excelsWithSheet != null && this.excelsWithSheet.Count > 0)
                    {
                        await this._ExportExcelFiles(this.outPathText.value);
                    }
                    else LogAdder.AddLogError("無任何可輸出之檔案!!");
                }
                else LogAdder.AddLogError("請選擇輸出路徑!!");
            };

            this.chooseAllTgl.RegisterValueChangedCallback(evt =>
            {
                if (this.excelsWithSheet != null && this.excelsWithSheet.Count > 0)
                {
                    for (int i = 0; i < this.excelsWithSheet.Count; i++)
                    {
                        Toggle tog = this.root.Q<Toggle>(this.excelsWithSheet[i].Namespace + this.excelsWithSheet[i].TableName);
                        tog.SetValueWithoutNotify(evt.newValue);
                    }
                    this._RefreshUseFileStatus();
                }
            });

            this.clearLogBtn.clickable.clicked += () =>
            {
                LogAdder.ClearLog();
            };

            this.jsonFormattingTgl.RegisterValueChangedCallback(evt =>
            {
                this.formatJson = evt.newValue;
                EditorPrefs.SetBool(keyJsonFormatting, evt.newValue);
            });


            this.exportOptionPop.RegisterCallback<ChangeEvent<string>>((evt) =>
            {
                this.exportOptionPop.value = evt.newValue;

                EditorPrefs.SetString(keyExportType, this.exportOptionPop.value);

                this._RefreshJsonFormattingTgl();

                LogAdder.AddLog(string.Format("目前輸出檔案類型為:{0}", this.exportOptionPop.value));
            });

            this.jsonGeneratorBtn.clickable.clicked += () =>
            {
                JsonConfigGenerator.OpenJsonConfigGenWindow();
            };

            this.resetBtn.clickable.clicked += async () =>
            {
                await this._Reset();
            };

            this.openFilePathBtn.clickable.clicked += () =>
            {
                this._OpenFolder(this.filePathText.value);
            };

            this.openOutPathBtn.clickable.clicked += () =>
            {
                this._OpenFolder(this.outPathText.value);
            };

            this.autoScanTgl.RegisterValueChangedCallback((value) =>
            {
                EditorPrefs.SetBool(keyAutoScan, value.newValue);
            });
        }

        private async Task _ExportExcelFiles(string outPath)
        {
            if (!this._ParseJsonConfigFile(this.jsonPathText.value) || this.jsonConfigList == null)
            {
                LogAdder.AddLogError("Json設定檔出錯或格式有誤，無法進行輸出!!");
                return;
            }

            // 關閉編輯器控制
            this._SetEditorControl(false);

            // 開啟進度調顯示
            this.SetProgressBarVisble(true);

            // 確定Json設定檔案轉換完成後，驗證所有檔案，名稱是否有被寫在Json設定檔內，有的話就部署相關設定
            List<SheetExportConfig> listSheetExportConfig = this._GetSheetExportConfig();

            FileExporter fileExporter = new FileExporter(listSheetExportConfig, outPath, this.formatJson, this.excelLoadingProgress);

            switch (this.exportOptionPop.value)
            {
                case ExportType.Json:
                    this.genResult = await fileExporter.SaveJsonFiles();
                    break;
                case ExportType.Bson:
                    this.genResult = await fileExporter.SaveBsonFiles();
                    break;
                case ExportType.Both:
                    this.genResult = await fileExporter.SaveBothFiles();
                    break;
                default:
                    return;
            }

            fileExporter = null;

            LogAdder.AddLog(string.Format("成功輸出的檔案數量:{0}", this.genResult));

            AssetDatabase.Refresh();

            // 重新開啟編輯器控制
            this._SetEditorControl(true);

            // 關閉進度調顯示
            this.SetProgressBarVisble(false);
        }

        private bool _ParseJsonConfigFile(string path)
        {
            if (path == null || !path.Contains(".json"))
            {
                LogAdder.AddLogError("請設定Json設定檔案路徑!!");
                return false;
            }

            // 儲存設定檔Json字串
            string tempJsonConfigData = File.ReadAllText(path);

            this.jsonConfigList = JsonConfigGenerator.GetParsedJsonConfig(tempJsonConfigData);

            if (this.jsonConfigList == null || this.jsonConfigList.Count == 0)
            {
                LogAdder.AddLogError("Json設定檔案轉換失敗!!");
                return false;
            }

            LogAdder.AddLog("Json設定檔案轉換成功");
            return true;
        }

        private async Task _LoadExcelFiles(string path)
        {
            if (path == null)
            {
                LogAdder.AddLogError("請設定Excel檔案路徑!!");
                return;
            }

            // 關閉Editor控制
            this._SetEditorControl(false);

            this.excelFiles = new List<FileInfo>();

            // 取得所有Excel
            await this._GetFolderFiles(this.excelFiles, path);

            // 取得所有Excel中的Sheet
            this.excelsWithSheet = await this._GetExcelSheets(this.excelFiles);

            // isUseFiles在此全預設為True
            this.isUseFiles = new List<bool>(Enumerable.Repeat(true, this.excelsWithSheet.Count));

            // 將全選的Toggle也預設為True
            this.chooseAllTgl.SetValueWithoutNotify(true);
            this.chooseAllTgl.label = string.Format("Sheets:{0}", this.excelsWithSheet.Count);
            this.chooseAllTgl.text = string.Format("Excels:{0}", this.excelFiles.Count);

            LogAdder.AddLog($"讀取完畢, Excel數量 : {this.excelFiles.Count}, Sheet數量 : {this.excelsWithSheet.Count}");

            // 重新開啟Editor控制
            this._SetEditorControl(true);
        }

        private async Task _GetFolderFiles(List<FileInfo> excelFiles, string path)
        {
            DirectoryInfo folder = new DirectoryInfo(path);
            if (!folder.Exists) return;

            // 將獲取到的所有Excel檔放進List中
            excelFiles.AddRange(this._GetFilesByExtensions(folder, ".xlsx", ".xls"));

            // 迴圈搜尋子資料夾
            foreach (var info in folder.GetDirectories())
            {
                await this._GetFolderFiles(excelFiles, info.FullName);
            }
        }

        private async Task<List<DataTable>> _GetExcelSheets(List<FileInfo> excelFiles)
        {
            // 清空存放Toggle的ScrollView
            this.scrollViewExcelFiles.Clear();

            // 接下來透過上面的Excel List開始獲取Excel內的Sheet
            List<DataTable> tempDataTable = new List<DataTable>();

            // 如果有Excel
            if (excelFiles.Count > 0)
            {
                StringBuilder strBuilder = new StringBuilder();
                strBuilder.Append("Excel讀取完畢，讀取失敗的檔案:");

                this.SetProgressBarVisble(true);

                this.excelLoadingProgress.lowValue = 0;
                this.excelLoadingProgress.highValue = excelFiles.Count;

                LogAdder.AddLog("開始讀取Excel檔案....");

                for (int i = 0; i < excelFiles.Count; i++)
                {
                    if (!this.isEditorRunning) break;

                    this.excelLoadingProgress.value = i + 1;
                    this.excelLoadingProgress.title = $"({this.excelLoadingProgress.value}/{excelFiles.Count}) 【{excelFiles[i].Name}】";

                    await Task.Yield();

                    // 讀取該Excel檔案路徑
                    FileStream stream;

                    do
                    {
                        try
                        {
                            // 讀取該Excel檔案路徑 
                            stream = File.Open(excelFiles[i].FullName, FileMode.Open, FileAccess.Read, FileShare.Read);
                        }
                        catch (IOException)
                        {
                            stream = null;
                            bool readAgain = EditorUtility.DisplayDialog("File Opened Detected", $"File is detected opened while ExcelConverter is trying to read it, Path :\n{excelFiles[i].FullName}\nWould you close it and try again?", "Yes", "No");
                            if (readAgain == false)
                            {
                                strBuilder.Append($"\n{excelFiles[i].FullName}");
                                break;
                            }
                        }
                    }
                    while (stream == null);

                    if (stream == null)
                    {
                        Debug.Log(string.Format($"<color=#FFBC00>{excelFiles[i].FullName}</color> <color=#FF007D>Passed reading!!</color>"));
                        continue;
                    }

                    DataSet dataSet = null;

                    using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        dataSet = reader.AsDataSet();
                    }

                    // 如果有Sheet
                    if (dataSet.Tables.Count > 0)
                    {
                        for (int j = 0; j < dataSet.Tables.Count; j++)
                        {
                            if (!this.isEditorRunning) break;

                            // 將Excel實際名稱加入到Table的NameSpace，這裡是為了防止以名稱尋找Toggle時發現重名，導致相關Bug產生
                            dataSet.Tables[j].Namespace = excelFiles[i].Name.Replace(".xlsx", "").Replace(".xls", "");

                            // 將獲得的Sheet新增到暫存List裡頭
                            tempDataTable.Add(dataSet.Tables[j]);

                            // 新增Toggle
                            Toggle tog = new Toggle()
                            {
                                value = true,
                                text = dataSet.Tables[j].Namespace,
                                label = dataSet.Tables[j].TableName,
                                // Toggle名稱設定為 "Namespace(Sheet隸屬的Excel名稱)" + "Tablename(Sheet實際名稱)"，防止以名稱尋找Toggle時重名
                                name = dataSet.Tables[j].Namespace + dataSet.Tables[j].TableName,
                                focusable = false
                            };

                            // 新增Class到Toggle
                            tog.AddToClassList("unity-custom-toggle");

                            // 添加Toggle事件 點擊時更新isUseFiles List狀態
                            tog.RegisterValueChangedCallback(evt =>
                            {
                                this._RefreshUseFileStatus();
                            });

                            // 將Toggle加到ScrollView裡頭
                            this.scrollViewExcelFiles.Add(tog);

                            await Task.Yield();
                        }
                    }

                    stream = null;
                }

                LogAdder.AddLog($"{strBuilder}");

                this.SetProgressBarVisble(false);
            }
            return tempDataTable;
        }

        IEnumerable<FileInfo> _GetFilesByExtensions(DirectoryInfo dir, params string[] extensions)
        {
            if (extensions == null) return null;

            IEnumerable<FileInfo> files = dir.EnumerateFiles();
            return files.Where(f => extensions.Contains(f.Extension.ToLower()) && !f.FullName.Contains("~$"));
        }

        private List<SheetExportConfig> _GetSheetExportConfig()
        {
            List<SheetExportConfig> listSheetExportConfig = new List<SheetExportConfig>();

            // 將Sheet一個一個做核對
            for (int i = 0; i < this.excelsWithSheet.Count; i++)
            {
                DataTable sheet = this.excelsWithSheet[i];

                SheetExportConfig sheetExportConfig = this._GenerateSheetExportConfig(sheet, this.isUseFiles[i]);
                if (sheetExportConfig != null)
                {
                    listSheetExportConfig.Add(sheetExportConfig);
                    LogAdder.AddLog($"sheet >> 【{sheet.TableName}】, 是否可以被輸出: ●");
                }
                else
                {
                    LogAdder.AddLog($"sheet >> 【{sheet.TableName}】, 是否可以被輸出: ○");
                }
            }

            return listSheetExportConfig;
        }

        private SheetExportConfig _GenerateSheetExportConfig(DataTable sheet, bool isUseFiles)
        {
            // 如果檔案沒勾選 不做核對
            if (!isUseFiles)
            {
                LogAdder.AddLogError(string.Format("-------------\nSheet無勾選使用 名稱:{0}\n-------------", sheet.TableName));
                return null;
            }
            else
            {
                for (int i = 0; i < this.jsonConfigList.Count; i++)
                {
                    JsonConfig jsonConfig = this.jsonConfigList[i];

                    for (int j = 0; j < jsonConfig.dataList.Count; j++)
                    {
                        if (sheet.TableName == jsonConfig.dataList[j])
                        {
                            LogAdder.AddLog(string.Format("---------------\nSheet核對有資料 檔案名稱:{0}", sheet.TableName));
                            LogAdder.AddLog(string.Format("數據資料起始列:{0}", jsonConfig.rowOfFirstData));
                            LogAdder.AddLog(string.Format("是否啟用主Key:{0}", jsonConfig.enableMainKey));
                            if (jsonConfig.enableMainKey)
                            {
                                LogAdder.AddLog(string.Format("主Key欄:{0}", jsonConfig.columnOfMainKey));
                                LogAdder.AddLog(string.Format("主Key大小寫類型:{0}", jsonConfig.mainKeyType));
                            }
                            LogAdder.AddLog(string.Format("是否啟用副Key:{0}", jsonConfig.enableSubKey));
                            if (jsonConfig.enableSubKey)
                            {
                                LogAdder.AddLog(string.Format("副Key列:{0}", jsonConfig.rowOfSubKey));
                                LogAdder.AddLog(string.Format("副Key大小寫類型:{0}\n-------------", jsonConfig.subKeyType));
                            }

                            return new SheetExportConfig(sheet, jsonConfig);
                        }
                    }
                }
            }

            LogAdder.AddLogError(string.Format("-------------\nSheet核對無資料 名稱:{0}\n-------------", sheet.TableName));

            return null;
        }

        private void _RefreshUseFileStatus()
        {
            if (this.excelsWithSheet.Count > 0)
            {
                for (int i = 0; i < this.excelsWithSheet.Count; i++)
                {
                    Toggle tog = root.Q<Toggle>(this.excelsWithSheet[i].Namespace + this.excelsWithSheet[i].TableName);
                    this.isUseFiles[i] = tog.value;

                    LogAdder.AddLog(string.Format("--------------\n Index值:{0} \n isUseFile目前狀態:{1} \n Sheet實際名稱:{2} \n Sheet歸屬之Excel為:{3}", i, isUseFiles[i], this.excelsWithSheet[i].TableName, this.excelsWithSheet[i].Namespace));
                }
            }
        }

        private PopupField<string> _InitCustomPopupField(PopupField<string> popupField)
        {
            // 找到存放選單的VisualElememt
            VisualElement container = root.Q<VisualElement>("LYT_PopupField");

            // 清除VisualElememt子內容
            container.Clear();

            // 設定選單內含的選項
            List<string> choices = new List<string> { ExportType.Json, ExportType.Bson, ExportType.Both };

            // 初始化下拉式選單
            popupField = new PopupField<string>("Option:", choices, 0);

            // 將unity-custom-popup這個Selector加到下拉式選單的Style Class List
            popupField.AddToClassList("unity-custom-popup");

            // 設定下拉式選單預設選項
            popupField.value = ExportType.Json;

            // 這邊是將下拉式選單的Label文字用Style Sheet作微調
            var text = popupField.Q<Label>();
            text.AddToClassList("unity-custom-text");

            // 將下拉式選單加到VisualElememt中
            container.Add(popupField);

            return popupField;
        }

        private async Task _Reset()
        {
            this._ResetPaths();

            if (!string.IsNullOrEmpty(this.filePathText.value))
            {
                await this._LoadExcelFiles(this.filePathText.value);
            }
        }

        private void _ResetPaths()
        {
            EditorPrefs.SetString(keyOutPath, Application.dataPath);
            EditorPrefs.SetString(keyFilePath, Application.dataPath);
            EditorPrefs.SetString(keyJsonPath, Application.dataPath);

            this.outPathText.value = EditorPrefs.GetString(keyOutPath, Application.dataPath);
            this.filePathText.value = EditorPrefs.GetString(keyFilePath, Application.dataPath);
            this.jsonPathText.value = EditorPrefs.GetString(keyJsonPath, Application.dataPath);
        }

        private void _OpenFolder(string path)
        {
            if (string.IsNullOrEmpty(path)) return;

            if (!Directory.Exists(path))
            {
                LogAdder.AddLogError(string.Format("此路徑沒有資料夾!! {0}", path));
            }

            System.Diagnostics.Process.Start("explorer.exe", path.Replace("/", "\\"));
        }

        private void _SetEditorControl(bool isOn)
        {
            this.outPathText.SetEnabled(isOn);
            this.filePathText.SetEnabled(isOn);
            this.jsonPathText.SetEnabled(isOn);
            this.chooseAllTgl.SetEnabled(isOn);
            this.jsonFormattingTgl.SetEnabled(isOn);
            this.exportBtn.SetEnabled(isOn);
            this.browseOutPathBtn.SetEnabled(isOn);
            this.browseFilePathBtn.SetEnabled(isOn);
            this.browseJsonPathBtn.SetEnabled(isOn);
            this.scanBtn.SetEnabled(isOn);
            this.clearLogBtn.SetEnabled(isOn);
            this.jsonGeneratorBtn.SetEnabled(isOn);
            this.resetBtn.SetEnabled(isOn);
            this.openFilePathBtn.SetEnabled(isOn);
            this.openOutPathBtn.SetEnabled(isOn);
            this.scrollViewExcelFiles.SetEnabled(isOn);
            this.exportOptionPop.SetEnabled(isOn);
            this.autoScanTgl.SetEnabled(isOn);
        }

        private void _SetExportFileControl(bool isOn)
        {
            this.exportBtn.SetEnabled(isOn);
            this.chooseAllTgl.SetEnabled(isOn);
            this.scrollViewExcelFiles.SetEnabled(isOn);
        }

        private void _RefreshJsonFormattingTgl()
        {
            switch (this.exportOptionPop.value)
            {
                case ExportType.Json:
                    this.jsonFormattingTgl.SetEnabled(true);
                    break;
                case ExportType.Both:
                    this.jsonFormattingTgl.SetEnabled(true);
                    break;
                case ExportType.Bson:
                    this.jsonFormattingTgl.SetEnabled(false);
                    break;
            }
        }

        public void SetProgressBarVisble(bool isOn)
        {
            this.excelLoadingProgress.visible = isOn;
        }
    }
}