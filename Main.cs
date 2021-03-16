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

        private IProcess mReadProcessor = null;
        private IProcess mWriteProcessor = null;

        public async void CheckCommandLineInput()
        {
            var args = Environment.GetCommandLineArgs();
            if (args.Length >= 3 && CheckFilePathValid(args[1]) && CheckFileTypeValid(args[2]))
            {
                await StartConvertAsync(args[1], ForStringToFileType(args[2]));
                Environment.Exit(0);
            }
        }

        public async Task StartConvertAsync(string filePath, FileType convertType)
        {
            var wirteType = FileType.none;
            switch (ForConvertStrategy(filePath, convertType))
            {
                case ConvertStrategy.JsonToCsv:
                    wirteType = FileType.csv;
                    mReadProcessor = new JsonProcessor();
                    mWriteProcessor = new CsvProcessor();
                    break;
                case ConvertStrategy.JsonToXlsx:
                    wirteType = FileType.xlsx;
                    mReadProcessor = new JsonProcessor();
                    mWriteProcessor = new ExcelProcessor();
                    break;
                case ConvertStrategy.CsvToJson:
                    wirteType = FileType.json;
                    mReadProcessor = new CsvProcessor();
                    mWriteProcessor = new JsonProcessor();
                    break;
                case ConvertStrategy.XlsxToJson:
                    wirteType = FileType.json;
                    mReadProcessor = new ExcelProcessor();
                    mWriteProcessor = new JsonProcessor();
                    break;
                default:
                    mReadProcessor = null;
                    mWriteProcessor = null;
                    EventDispatcher.instance.OnUpdateInformation(Message.ERROR_INVALID_FILE_TYPE);
                    break;
            }

            if (null != mReadProcessor && null != mWriteProcessor)
            {
                var jsonData = await mReadProcessor.ReadFileAsync(filePath);

                RemoveEmptyData(jsonData);

                await mWriteProcessor.WriteFileAsync(jsonData, ForSavePath(filePath, wirteType));

                EventDispatcher.instance.OnUpdateInformation("轉換完成");
            }
        }

        public FileType ForPathFileType(string filePath)
        {
            var extension = Path.GetExtension(filePath);
            if (".json" == extension)
                return FileType.json;
            else if (".csv" == extension)
                return FileType.csv;
            else if (".xlsx" == extension)
                return FileType.xlsx;
            else
                return FileType.none;
        }

        public FileType ForStringToFileType(string typeString)
            => (FileType)Enum.Parse(typeof(FileType), typeString?.ToLower());

        private string ForSavePath(string oldPath, FileType saveType)
        {
            var result = $"{Path.GetDirectoryName(oldPath)}\\{Path.GetFileNameWithoutExtension(oldPath)}";
            switch (saveType)
            {
                case FileType.json:
                    result = string.Concat(result, ".json");
                    break;
                case FileType.csv:
                    result = string.Concat(result, ".csv");
                    break;
                case FileType.xlsx:
                    result = string.Concat(result, ".xlsx");
                    break;
            }
            return result;
        }

        public ConvertStrategy ForConvertStrategy(string filePath, FileType convertTo)
        {
            switch (ForPathFileType(filePath))
            {
                case FileType.json:
                    switch (convertTo)
                    {
                        case FileType.csv:
                            return ConvertStrategy.JsonToCsv;
                        case FileType.xlsx:
                            return ConvertStrategy.JsonToXlsx;
                        default:
                            return ConvertStrategy.None;
                    }    
                case FileType.csv:
                    return ConvertStrategy.CsvToJson;
                case FileType.xlsx:
                    return ConvertStrategy.XlsxToJson;
                default:
                    return ConvertStrategy.None;
            }
        }

        private bool CheckFilePathValid(string filePath)
        {
            if (FileType.none == ForPathFileType(filePath))
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

        private bool CheckFileTypeValid(string typeString)
            => Enum.IsDefined(typeof(FileType), typeString);

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
    }
}
