using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JsonConverter
{
    class Main
    {
        private static Main mInstance = null;
        public static Main instance
        {
            get
            {
                if (null == mInstance)
                    mInstance = new Main();

                return mInstance;
            }
        }

        public void StartConvert(string filePath)
        {
            var extension = Path.GetExtension(filePath);
            if (extension == ".json")
                JsonConvertToCsv(filePath);
            else if (extension == ".csv")
                CsvConvertToJson(filePath);
        }

        #region Json
        private async void JsonConvertToCsv(string filePath)
        {
            Console.WriteLine("Start Load Json");

            var jsonDataList = await JsonHelper.LoadJsonDataAsync(filePath);

            if (null != jsonDataList)
            {
                Console.WriteLine("Load Json Complete");

                RemoveEmptyData(jsonDataList);

                Console.WriteLine("Convert JsonData To Csv");

                var directoreName = Path.GetDirectoryName(filePath);
                var fileName = Path.GetFileNameWithoutExtension(filePath);
                var savePath = $"{directoreName}\\{fileName}.csv";

                await CsvHelper.WriteCsvDataAsync(jsonDataList, savePath);

                Console.WriteLine("Convert End");

                EventDispatcher.instance.OnUpdateInformation("轉換完成");
            }
            else
            {
                Console.WriteLine("Has something error, json data list is null");
            }
        }

        /// <summary>移除資料裡面，id為空的資料</summary>
        private void RemoveEmptyData(List<Dictionary<string, string>> list)
        {
            if (null == list || list.Count <= 0)
                return;

            for (var i = list.Count - 1; i >= 0; --i)
            {
                if (list[i].TryGetValue("id", out var value) && string.IsNullOrEmpty(value))
                    list.RemoveAt(i);
            }
        }
        #endregion Json

        #region Csv
        private async void CsvConvertToJson(string filePath)
        {
            Console.WriteLine("Start Load Csv");

            // 從csv讀進來的資料，不過在讀取期間就轉成jsonData格式
            var csvData = await CsvHelper.ReadCsvDataAsync(filePath);

            if (null != csvData)
            {
                Console.WriteLine("Load Csv Complete");

                Console.WriteLine("Write JsonData");

                var directoreName = Path.GetDirectoryName(filePath);
                var fileName = Path.GetFileNameWithoutExtension(filePath);
                var savePath = $"{directoreName}\\{fileName}_new.json";

                await JsonHelper.WriteJsonFileAsync(csvData, savePath);

                Console.WriteLine("Convert End");

                EventDispatcher.instance.OnUpdateInformation("轉換完成");
            }
        }
        #endregion Csv
    }
}
