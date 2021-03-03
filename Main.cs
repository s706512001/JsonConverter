using System;
using System.Collections.Generic;
using System.IO;
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

        public async void CheckCommandLineInput()
        {
            var args = Environment.GetCommandLineArgs();
            if (args.Length >= 2 && CheckFilePathValid(args[1]))
            {
                await StartConvertAsync(args[1]);
                Environment.Exit(0);
            }
        }

        public async Task StartConvertAsync(string filePath)
        {
            switch (ForFileType(filePath))
            {
                case FileType.Json:
                    await JsonConvertToCsv(filePath);
                    break;
                case FileType.Csv:
                    await CsvConvertToJson(filePath);
                    break;
                default:
                    EventDispatcher.instance.OnUpdateInformation(Message.ERROR_INVALID_FILE_TYPE);
                    break;
            }
        }

        public FileType ForFileType(string filePath)
        {
            var extension = Path.GetExtension(filePath);
            if (".json" == extension)
                return FileType.Json;
            else if (".csv" == extension)
                return FileType.Csv;
            else
                return FileType.None;
        }

        private string ForSavePath(string oldPath, FileType saveType)
        {
            var result = $"{Path.GetDirectoryName(oldPath)}\\{Path.GetFileNameWithoutExtension(oldPath)}";
            switch (saveType)
            {
                case FileType.Json:
                    result = string.Concat(result, ".json");
                    break;
                case FileType.Csv:
                    result = string.Concat(result, ".csv");
                    break;
            }
            return result;
        }

        private bool CheckFilePathValid(string filePath)
        {
            if (FileType.None == ForFileType(filePath))
            {
                EventDispatcher.instance.OnUpdateInformationWithFilePath(filePath, Message.ERROR_INVALID_FILE_TYPE);
                return false;
            }
            else if (false == File.Exists(filePath))
            {
                EventDispatcher.instance.OnUpdateInformationWithFilePath(filePath, Message.ERROR_FILE_NOT_EXIST);
                return false;
            }
            return true;
        }

        #region Json
        private async Task JsonConvertToCsv(string filePath)
        {
            Console.WriteLine("Start Load Json");

            var jsonDataList = await JsonHelper.LoadJsonDataAsync(filePath);

            if (null != jsonDataList)
            {
                Console.WriteLine("Load Json Complete");

                RemoveEmptyData(jsonDataList);

                Console.WriteLine("Convert JsonData To Csv");

                await CsvHelper.WriteCsvDataAsync(jsonDataList, ForSavePath(filePath, FileType.Csv));

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
        private async Task CsvConvertToJson(string filePath)
        {
            Console.WriteLine("Start Load Csv");

            // 從csv讀進來的資料，不過在讀取期間就轉成jsonData格式
            var csvData = await CsvHelper.ReadCsvDataAsync(filePath);

            if (null != csvData)
            {
                Console.WriteLine("Load Csv Complete");

                Console.WriteLine("Write JsonData");

                await JsonHelper.WriteJsonFileAsync(csvData, ForSavePath(filePath, FileType.Json));

                Console.WriteLine("Convert End");

                EventDispatcher.instance.OnUpdateInformation("轉換完成");
            }
        }
        #endregion Csv
    }
}
