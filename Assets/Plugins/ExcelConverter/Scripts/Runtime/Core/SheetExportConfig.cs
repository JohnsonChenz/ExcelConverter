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
        public bool ableToExport { get; private set; }                // �O�_�i�Q��X
        public JsonConfig.KeyType mainKeyType { get; private set; }   // �DKey����, None = �L, Lower = ��p�g, Upper = ��j�g
        public JsonConfig.KeyType subKeyType { get; private set; }    // ��Key����, None = �L, Lower = ��p�g, Upper = ��j�g
        public int mainKeyColumn { get; private set; }                // �DKey��(�i�H���h�� �γr���۹j �|�b��l�ɴN���簣)
        public int subKeyRow { get; private set; }                    // ��Key�C
        public int firstDataRow { get; private set; }                 // �ƾڸ�ư_�l�C

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
