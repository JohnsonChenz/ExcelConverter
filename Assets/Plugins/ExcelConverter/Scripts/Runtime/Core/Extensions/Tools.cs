using Newtonsoft.Json;
using Newtonsoft.Json.Bson;
using Newtonsoft.Json.Linq;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using UnityEngine;

namespace ExcelConverter
{
    public static partial class Tools
    {
        public static byte[] ToBson<T>(this T value)
        {
            MemoryStream ms = new MemoryStream();
            using (BsonWriter writer = new BsonWriter(ms))
            {
                JsonSerializer serializer = new JsonSerializer();
                serializer.Serialize(writer, value);
            }

            byte[] data = ms.ToArray();

            return data;
        }
    }
}
