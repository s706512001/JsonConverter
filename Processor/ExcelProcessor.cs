using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JsonConverter
{
    class ExcelProcessor : Processor
    {
        /// <summary>
        /// 提供者名稱  Microsoft.Jet.OLEDB.4.0適用於2003以前版本，Microsoft.ACE.OLEDB.12.0 適用於2007以後的版本處理 xlsx 檔案
        /// </summary>
        protected virtual string providerName { private set; get; } = "Microsoft.ACE.OLEDB.12.0;";
        /// <summary>
        /// Excel版本，Excel 8.0 針對Excel2000及以上版本，Excel5.0 針對Excel97
        /// </summary>
        protected virtual string extendedString { private set; get; } = "Excel 8.0";

        public override Task<List<Dictionary<string, string>>> ReadFileAsync(string filePath)
            => Task.Run(() => ReadXlsxData(filePath));

        public override Task WriteFileAsync(List<Dictionary<string, string>> jsonData, string savePath)
            => Task.Run(() => WriteXlsxData(jsonData, savePath));

        /// <summary>
        /// 取得連結Excel檔案的字串
        /// </summary>
        /// <param name="filePath">檔案開啟路徑</param>
        /// <param name="HDR">第一行是否為標題</param>
        /// <param name="IMEX">IMEX=1 通知驅動程序始終將「互混」數據列作為文本讀取</param>
        private string ForOledbConnectString(string filePath, string HDR = "", string IMEX = "")
        {
            // 每個參數用 ";" 結尾區隔
            var extendedProperties = string.Concat(extendedString, ";");
            if (!string.IsNullOrEmpty(HDR))
                extendedProperties = string.Concat(extendedProperties, $"HDR={HDR};");
            if (!string.IsNullOrEmpty(IMEX))
                extendedProperties = string.Concat(extendedProperties, $"IMEX={IMEX};");

            var result = string.Format("Data Source={0};Provider={1}Extended Properties='{2}';",
                filePath,
                providerName,
                extendedProperties.Substring(0, extendedProperties.Length - 1));    // 最後一個參數不用 ";"

            return result;
        }

        #region Read
        private List<Dictionary<string, string>> ReadXlsxData(string filePath)
        {
            var dataTable = new DataTable();
            using (OleDbConnection connect = new OleDbConnection(ForOledbConnectString(filePath, "No", "1")))
                ForDataTable(connect, ForFirstSheetTableName(connect), dataTable);

            return ForJsonData(dataTable);
        }

        private DataTable ForDataTable(OleDbConnection connect, string sheetName, DataTable result)
        {
            if (null == result)
                result = new DataTable();

            var queryString = $"select * from [{sheetName}$]";
            try
            {
                using (OleDbDataAdapter dr = new OleDbDataAdapter(queryString, connect))
                    dr.Fill(result);
            }
            catch (Exception ex)
            {
                EventDispatcher.instance.OnShowMessageBox(ex.Message);
            }
            return result;
        }

        private string ForFirstSheetTableName(OleDbConnection connect)
        {
            var result = string.Empty;
            if ((null == connect) || (ConnectionState.Broken == connect.State))
                return result;

            connect.Open();

            var tables = connect.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            if (tables.Rows.Count > 0)
                result = tables.Rows[0]["TABLE_NAME"].ToString();

            connect.Close();

            return result;
        }

        private List<Dictionary<string, string>> ForJsonData(DataTable dataTable)
        {
            var result = new List<Dictionary<string, string>>();
            var keys = ForDataTableColumnsName(dataTable);

            for (var i = 1; i < dataTable.Rows.Count; ++i)
            {
                var rowData = dataTable.Rows[i].ItemArray;
                var dict = new Dictionary<string, string>(rowData.Length);
                for (var j = 0; j < rowData.Length; ++j)
                    dict.Add(keys[j], rowData[j].ToString());

                result.Add(dict);
            }

            return result;
        }

        private string[] ForDataTableColumnsName(DataTable dataTable)
        {
            return Array.ConvertAll(dataTable.Rows[0].ItemArray, ConvertToString);

            string ConvertToString(object obj)
                => obj.ToString();
        }
        #endregion Read

        #region Write
        private async void WriteXlsxData(List<Dictionary<string, string>> jsonData, string savePath)
        {
            if (File.Exists(savePath))
                File.Delete(savePath);

            using (var connect = new OleDbConnection(ForOledbConnectString(savePath, "Yes")))
            {
                try
                {
                    connect.Open();
                    var cmd = connect.CreateCommand();
                    var fileName = Path.GetFileNameWithoutExtension(savePath);
                    var titleList = ForTitleList(jsonData);
                    var valueList = ForValueList(jsonData, titleList);

                    cmd.CommandText = $"CREATE TABLE {fileName} ({string.Format(ForTitleString(titleList), " NTEXT")})";
                    cmd.ExecuteNonQuery();

                    cmd.Parameters.AddRange(ForParameters(titleList));
                    cmd.CommandText =
                        $"INSERT INTO {fileName} ({string.Format(ForTitleString(titleList), "")}) VALUES({ForValueString(titleList)})";

                    for (var i = 0; i < valueList.Count; ++i)
                    {
                        for (var j = 0; j < titleList.Count; ++j)
                            cmd.Parameters[$"@{titleList[j]}"].Value = valueList[i][j];

                        await cmd.ExecuteNonQueryAsync();
                    }
                }
                finally
                {
                    if (ConnectionState.Open == connect.State)
                        connect.Close();
                }
            }
        }

        private OleDbParameter[] ForParameters(List<string> list)
        {
            var result = new List<OleDbParameter>(list.Count);
            for (var i = 0; i < list.Count; ++i)
                result.Add(new OleDbParameter($"@{list[i]}", string.Empty));

            return result.ToArray();
        }

        private string ForTitleString(List<string> list)
        {
            var result = new StringBuilder();
            for (var i = 0; i < list.Count; ++i)
                result.Append(string.Concat(list[i], "{0},"));

            return result.ToString().Substring(0, result.Length - 1);
        }

        private string ForValueString(List<string> list)
        {
            var result = new StringBuilder();
            for (var i = 0; i < list.Count; ++i)
                result.Append("?,");

            return result.ToString().Substring(0, result.Length - 1);
        }
        #endregion Write
    }
}
