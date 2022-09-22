using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using UnityEngine.UIElements;
using UnityEditor;
using System;
using UnityEditor.UIElements;
using Newtonsoft.Json;
using System.IO;
using Newtonsoft.Json.Linq;
using ExcelConverter;
using System.Threading.Tasks;

namespace ExcelConverter.Editor
{
    public class JsonConfigGenerator : EditorWindow
    {
        [SerializeField, Header("【Json設定檔列表】")]
        public List<JsonConfig> jsonConfigList = new List<JsonConfig>();

        private SerializedObject serializedObject;                           // 序列化對象

        private SerializedProperty assetListProperty;                        // 序列化屬性

        private Vector2 scrollPosition;                                      // 捲動座標暫存

        public static void OpenJsonConfigGenWindow()
        {
            var win = GetWindow<JsonConfigGenerator>("Json Generator");
            win.Show();
            win.minSize = new Vector2(465, 700);
            win.maxSize = new Vector2(465, 700);
        }

        private void OnEnable()
        {
            // 使用當前類別初始化
            this.serializedObject = new SerializedObject(this);

            // 找到當前類別可序列化的屬性
            this.assetListProperty = this.serializedObject.FindProperty("jsonConfigList");
        }

        private async void OnGUI()
        {
            // 更新
            this.serializedObject.Update();

            // 開始檢查是否有修改
            EditorGUI.BeginChangeCheck();

            // 暫存ScrollView滑動位置的同時，也設置ScrollView的滑動位置
            this.scrollPosition = GUILayout.BeginScrollView(this.scrollPosition, GUILayout.MaxHeight(500));

            // 顯示List內容，第二個參數要為true否則無法顯示子節點即List內容
            EditorGUILayout.PropertyField(this.assetListProperty, true);

            // 結束檢查是否有修改
            if (EditorGUI.EndChangeCheck())
            {
                // 提交修改
                this.serializedObject.ApplyModifiedProperties();
            }

            // 結束ScrollView區域
            GUILayout.EndScrollView();

            // 設置間隔
            GUILayout.Space(30);

            // 開始水平排版區域
            GUILayout.BeginHorizontal();

            if (GUILayout.Button("Load Json File", GUILayout.Width(100)))
            {
                // 讀取先前紀錄的路徑
                string tempPath = EditorPrefs.GetString(ExcelConverterEditor.keyJsonPath, Application.dataPath);
                string fileName = Path.GetFileName(tempPath);
                tempPath = tempPath.Replace(fileName, "");

                string path = EditorUtility.OpenFilePanel("Select Json File", !string.IsNullOrEmpty(tempPath) ? tempPath : "", "json");

                if (!string.IsNullOrEmpty(path))
                {
                    string json = File.ReadAllText(path);
                    this.jsonConfigList = GetParsedJsonConfig(json);
                }
            }

            GUILayout.Space(75);

            if (GUILayout.Button("Save Json File", GUILayout.Width(100)))
            {
                // 讀取先前紀錄的路徑
                string tempPath = EditorPrefs.GetString(ExcelConverterEditor.keyJsonPath, Application.dataPath);
                string fileName = Path.GetFileName(tempPath);
                tempPath = tempPath.Replace(fileName, "");

                string path = EditorUtility.SaveFilePanel("Save Json File", !string.IsNullOrEmpty(tempPath) ? tempPath : "", "JsonConfig", "json");

                if (!string.IsNullOrEmpty(path))
                {
                    string json = JsonConvert.SerializeObject(this.jsonConfigList, Formatting.Indented);
                    File.WriteAllText(path, json);
                }
            }

            GUILayout.Space(75);

            if (GUILayout.Button("Reset", GUILayout.Width(100)))
            {
                this.jsonConfigList = new List<JsonConfig>();
            }

            // 停止水平排版區域
            GUILayout.EndHorizontal();
        }

        public static List<JsonConfig> GetParsedJsonConfig(string json)
        {
            JArray jArray;
            List<JsonConfig> listJsonConfig = new List<JsonConfig>();
            try
            {
                jArray = JArray.Parse(json);
            }
            catch
            {
                EditorUtility.DisplayDialog("錯誤", "Json載入失敗!!，Json資料無法轉為JArray!!", "了解");
                return null;
            }

            if (jArray.Count > 0)
            {
                for (int i = 0; i < jArray.Count; i++)
                {
                    for (int j = 0; j < JsonConfig.configKeyChecker.Length; j++)
                    {
                        if (jArray[i][JsonConfig.configKeyChecker[j]] == null)
                        {
                            EditorUtility.DisplayDialog("錯誤", "Json載入失敗!!，Json資料與比對的Key不相符.", "了解");
                            return null;
                        }
                    }

                    JsonConfig jsonConfig = null;

                    try
                    {
                        jsonConfig = jArray[i].ToObject<JsonConfig>();
                    }
                    catch
                    {
                        EditorUtility.DisplayDialog("錯誤", "Json載入失敗!!，Json資料序列化失敗.", "了解");
                        return null;
                    }

                    if (jsonConfig != null)
                    {
                        listJsonConfig.Add(jsonConfig);
                    }
                }
            }

            return listJsonConfig;
        }
    }
}
