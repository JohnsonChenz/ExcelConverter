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
        public static readonly string[] configKeyChecker = new string[6]
        {
            "firstDataRow",
            "mainKeyColumn",
            "subKeyRow",
            "mainKeyType",
            "subKeyType",
            "dataList"
        };

        public enum KeyType
        {
            None,
            Lowercase,
            Uppercase,
        }

        [Header("【主Key字母類型】"), JsonConverter(typeof(StringEnumConverter))]
        public KeyType mainKeyType;
        [Header("【副Key字母類型】"), JsonConverter(typeof(StringEnumConverter))]
        public KeyType subKeyType;
        [Header("【主Key欄】")]
        public int mainKeyColumn;
        [Header("【副Key列】")]
        public int subKeyRow;
        [Header("【數據資料起始列】")]
        public int firstDataRow;

        [Header("【文件名稱列表】")]
        public List<string> dataList;

        public JsonConfig()
        {
            this.dataList = new List<string>();
        }
    }
}
