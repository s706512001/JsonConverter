using CsvHelper;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;

namespace JsonConverter
{
    class CsvHelper
    {
        #region Load
        public static Task<List<Dictionary<string, string>>> ReadCsvDataAsync(string filePath)
            => Task.Run(() => ReadCsvData(filePath));

        private static List<Dictionary<string, string>> ReadCsvData(string filePath)
        {
            var result = new List<Dictionary<string, string>>();

            using (var reader = new StreamReader(filePath))
            {
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    csv.Read();
                    csv.ReadHeader();
                    // 取得全部的標題
                    var headers = csv.HeaderRecord;

                    while (csv.Read())
                    {
                        var data = new Dictionary<string, string>(headers.Length);
                        for (var i = 0; i < headers.Length; ++i)
                        {
                            if (csv.TryGetField(headers[i], out string value))
                                data[headers[i]] = value;
                            else
                                data[headers[i]] = string.Empty;
                        }
                        result.Add(data);
                    }
                }
            }
            return result;
        }
        #endregion Load

        #region Write
        public static Task WriteCsvDataAsync(List<Dictionary<string, string>> jsonData, string savePath)
            => Task.Run(() => WriteCsvData(jsonData, savePath));

        private static void WriteCsvData(List<Dictionary<string, string>> jsonData, string savePath)
        {
            var titleList = ForTitleList(jsonData);
            var valueList = ForValueList(jsonData, titleList);

            using (var write = new StreamWriter(savePath))
            {
                // 因為讀取進來的jsonData是不定欄位數量，因此這裡採用手動寫入csv內容，而沒有樣版class
                using (var csv = new CsvWriter(write, CultureInfo.InvariantCulture))
                {
                    // 先寫入標題資料
                    for (var i = 0; i < titleList.Count; ++i)
                        csv.WriteField(titleList[i]);

                    csv.NextRecord();
                    // 再寫入內容資料
                    for (var i = 0; i < valueList.Count; ++i)
                    {
                        var value = valueList[i];
                        for (var j = 0; j < value.Count; ++j)
                            csv.WriteField(value[j]);

                        csv.NextRecord();
                    }
                }
            }
        }

        private static List<string> ForTitleList(List<Dictionary<string, string>> jsonData)
        {
            var result = new List<string>();

            if (null != jsonData)
            {
                // 彙整全部資料裡面所有存在的key
                for (var i = 0; i < jsonData.Count; ++i)
                {
                    var iterator = jsonData[i].Keys.GetEnumerator();
                    while (iterator.MoveNext())
                    {
                        var key = iterator.Current;
                        if (false == result.Contains(key))
                            result.Add(key);
                    }
                }
            }

            return result;
        }

        private static List<List<string>> ForValueList(List<Dictionary<string, string>> jsonData, List<string> titleList)
        {
            var result = new List<List<string>>();

            for (var i = 0; i < jsonData.Count; ++i)
            {
                var data = jsonData[i];
                var values = new List<string>(titleList.Count);
                // 依照欄位的順序，將資料放在一樣順序的位置上
                for (var j = 0; j < titleList.Count; ++j)
                {
                    if (data.TryGetValue(titleList[j], out var value))
                        values.Add(value);
                }
                result.Add(values);
            }

            return result;
        }
        #endregion Write
    }
}
