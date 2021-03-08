namespace JsonConverter
{
    public enum FileType
    {
        none,
        /// <summary>Json</summary>
        json,
        /// <summary>Csv</summary>
        csv,
        /// <summary>Excel</summary>
        xlsx,
    }

    public enum ConvertStrategy
    {
        None,
        JsonToCsv,
        JsonToXlsx,
        CsvToJson,
        XlsxToJson,
    }
}
