using CsvHelper;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JsonConverter
{
    class CsvProcessor : Processor
    {
        public override Task<List<Dictionary<string, string>>> ReadFileAsync(string filePath)
            => Task.Run(() => ReadCsvData(filePath));

        public override Task WriteFileAsync(List<Dictionary<string, string>> jsonData, string savePath)
            => Task.Run(() => WriteCsvData(jsonData, savePath));

        private List<Dictionary<string, string>> ReadCsvData(string filePath)
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

        private void WriteCsvData(List<Dictionary<string, string>> jsonData, string savePath)
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
    }
}
