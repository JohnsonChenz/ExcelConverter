using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.Serialization;
using UnityEngine;

namespace ExcelConverter
{
    [System.Serializable]
    public class JsonConfig
    {
        // 用來檢驗Json格式是否正確的字串
        [NonSerialized]
        public static readonly string[] configKeyChecker = new string[]
        {
            "enableMainKey",
            "mainKeyType",
            "mainKeySelectType",
            "columnOfMainKey",
            "enableSubKey",
            "subKeyType",
            "rowOfSubKey",
            "rowOfFirstData",
            "dataList"
        };

        public enum KeyType
        {
            None,
            Lowercase,
            Uppercase,
        }

        public enum MainKeySelectType
        {
            FirstNColumn,
            SpecificColumn
        }
        [Tooltip("【是否啟用Mainkey】")]
        public bool enableMainKey;
        [Tooltip("【主Key字母類型】"), JsonConverter(typeof(StringEnumConverter)), DrawIf("enableMainKey", true)]
        public KeyType mainKeyType;
        [Tooltip("【Mainkey挑選模式】"), JsonConverter(typeof(StringEnumConverter)), DrawIf("enableMainKey", true)]
        public MainKeySelectType mainKeySelectType;
        [Tooltip("【主Key欄】"), DrawIf("enableMainKey", true)]
        public int columnOfMainKey;
        [Tooltip("【是否啟用Subkey】")]
        public bool enableSubKey;
        [Tooltip("【副Key字母類型】"), JsonConverter(typeof(StringEnumConverter)), DrawIf("enableSubKey", true)]
        public KeyType subKeyType;
        [Tooltip("【副Key列】"), DrawIf("enableSubKey", true)]
        public int rowOfSubKey;
        [Tooltip("【數據資料起始列】")]
        public int rowOfFirstData;

        [Tooltip("【文件名稱列表】")]
        public List<string> dataList;

        public JsonConfig()
        {
            this.dataList = new List<string>();
        }
    }
}
