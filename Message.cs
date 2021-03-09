namespace JsonConverter
{
    static class Message
    {
        public const string INFO_INPUT_FILE = "選擇一個檔案，或將檔案拖曳進來";
        public const string INFO_JSON_TO_CSV = "Json -> Csv";
        public const string INFO_CSV_TO_JSON = "Csv -> Json";
        public const string INFO_CONVERT_TO_TITLE_JSON = "Json ->";
        public const string INFO_CONVERT_TO_TITLE_CSV = "Csv ->";
        public const string INFO_CONVERT_TO_TITLE_XLSX = "Xlsx ->";
        public const string INFO_SELECT_A_CONVERT_TO_TYPE = "請選擇要轉換的檔案類型";

        public const string ERROR_INVALID_FILE_TYPE = "錯誤的檔案格式\n請放入 json 或 csv 檔案";
        public const string ERROR_FILE_NOT_EXIST = "檔案不存在";
    }
}
