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

        /// <summary>檢查是否由 command line 執行</summary>
        public async void CheckCommandLineInput()
        {
            var args = Environment.GetCommandLineArgs();
            if (args.Length >= 3 && CheckFilePathValid(args[1]) && CheckFileTypeValid(args[2]))
            {
                await StartConvertAsync(args[1], ForStringToFileType(args[2]));
                Environment.Exit(0);
            }
        }

        /// <summary>
        /// 開始轉檔
        /// </summary>
        /// <param name="filePath">輸入檔案路徑</param>
        /// <param name="convertType">要轉換的檔案類型</param>
        public async Task StartConvertAsync(string filePath, FileType convertType)
        {
            switch (ForConvertStrategy(filePath, convertType))
            {
                case ConvertStrategy.JsonToCsv:
                    SetupProcessor(FileType.json, FileType.csv);
                    break;
                case ConvertStrategy.JsonToXlsx:
                    SetupProcessor(FileType.json, FileType.xlsx);
                    break;
                case ConvertStrategy.CsvToJson:
                    SetupProcessor(FileType.csv, FileType.json);
                    break;
                case ConvertStrategy.XlsxToJson:
                    SetupProcessor(FileType.xlsx, FileType.json);
                    break;
                default:
                    SetupProcessor(FileType.none, FileType.none);
                    EventDispatcher.instance.OnUpdateInformation(Message.ERROR_INVALID_FILE_TYPE);
                    break;
            }

            if (null != mReadProcessor && null != mWriteProcessor)
            {
                var jsonData = await mReadProcessor.ReadFileAsync(filePath);

                RemoveEmptyData(jsonData);

                await mWriteProcessor.WriteFileAsync(jsonData, ForSavePath(filePath, ForConvertStrategy(filePath, convertType)));

                EventDispatcher.instance.OnUpdateInformation("轉換完成");
            }
        }

        /// <summary>輸入路徑的檔案類型</summary>
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

        /// <summary>
        /// string 轉 FileType
        /// </summary>
        /// <param name="typeString">json, csv, xlsx...</param>
        private FileType ForStringToFileType(string typeString)
            => (FileType)Enum.Parse(typeof(FileType), typeString?.ToLower());

        /// <summary>設置讀取和寫入的處理器</summary>
        private void SetupProcessor(FileType readType, FileType writeType)
        {
            mReadProcessor = Processor.CreateProcessor(readType);
            mWriteProcessor = Processor.CreateProcessor(writeType);
        }

        /// <summary>
        /// 產生寫入檔案的路徑
        /// </summary>
        /// <param name="oldPath">原輸入檔案路徑</param>
        /// <param name="strategy">轉檔策略</param>
        private string ForSavePath(string oldPath, ConvertStrategy strategy)
        {
            var result = $"{Path.GetDirectoryName(oldPath)}\\{Path.GetFileNameWithoutExtension(oldPath)}";
            switch (strategy)
            {
                case ConvertStrategy.CsvToJson:
                case ConvertStrategy.XlsxToJson:
                    result = string.Concat(result, ".json");
                    break;
                case ConvertStrategy.JsonToCsv:
                    result = string.Concat(result, ".csv");
                    break;
                case ConvertStrategy.JsonToXlsx:
                    result = string.Concat(result, ".xlsx");
                    break;
            }
            return result;
        }

        /// <summary>
        /// 偵測轉檔策略
        /// </summary>
        /// <param name="filePath">輸入檔案路徑</param>
        /// <param name="convertTo">轉換後的檔案類型</param>
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

        /// <summary>偵測檔案路徑是否合法</summary>
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

        /// <summary>偵測 command line 輸入的檔案類型是否合法</summary>
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
