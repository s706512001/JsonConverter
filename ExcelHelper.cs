using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using System.Windows.Forms;
using System.IO;

namespace JsonConverter
{
    class ExcelHelper
    {
        #region Load
        public static Task<DataTable> ReadXlsxDataAsync(string filePath)
            => Task.Run(() => ReadXlsxData(filePath));

        private static DataTable ReadXlsxData(string filePath)
        {
            // 提供者名稱  Microsoft.Jet.OLEDB.4.0適用於2003以前版本，Microsoft.ACE.OLEDB.12.0 適用於2007以後的版本處理 xlsx 檔案
            var providerName = "Microsoft.ACE.OLEDB.12.0;";
            // Excel版本，Excel 8.0 針對Excel2000及以上版本，Excel5.0 針對Excel97。
            var extendedString = "'Excel 8.0;";
            // 第一行是否為標題(;結尾區隔)
            var HDR = "No;";
            // IMEX=1 通知驅動程序始終將「互混」數據列作為文本讀取(;結尾區隔,'文字結尾)
            var IMEX = "1';";
            var connectString =
                $"Data Source={filePath};Provider={providerName}Extended Properties={extendedString}HDR={HDR}IMEX={IMEX}";

            var result = new DataTable();
            using (OleDbConnection connect = new OleDbConnection(connectString))
                ForDataTable(connect, ForFirstSheetTableName(connect), result);

            return result;
        }

        public static List<Dictionary<string, string>> ConvertToJsonData(DataTable dataTable)
        {
            var result = new List<Dictionary<string, string>>();
            var keys = ForTableKeys(dataTable);

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

        private static DataTable ForDataTable(OleDbConnection connect, string sheetName, DataTable result)
        {
            if (null == result)
                result = new DataTable();

            var queryString = $"select * from [{sheetName}]";
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

        private static string ForFirstSheetTableName(OleDbConnection connect)
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

        private static string[] ForTableKeys(DataTable dataTable)
        {
            return Array.ConvertAll(dataTable.Rows[0].ItemArray, ConvertToString);

            string ConvertToString(object obj)
                => obj.ToString();
        }
        #endregion Load

        #region Save
        public static Task WriteXlsxDataAsync(List<Dictionary<string, string>> jsonData, string savePath)
            => Task.Run(() => WriteXlsxData(jsonData, savePath));

        private static async void WriteXlsxData(List<Dictionary<string, string>> jsonData, string savePath)
        {
            // 提供者名稱  Microsoft.Jet.OLEDB.4.0適用於2003以前版本，Microsoft.ACE.OLEDB.12.0 適用於2007以後的版本處理 xlsx 檔案
            var providerName = "Microsoft.ACE.OLEDB.12.0;";
            // Excel版本，Excel 8.0 針對Excel2000及以上版本，Excel5.0 針對Excel97。
            var extendedString = "'Excel 8.0;";
            // 第一行是否為標題(;結尾區隔)
            var HDR = "Yes';";

            var connectString =
                $"Data Source={savePath};Provider={providerName}Extended Properties={extendedString}HDR={HDR}";

            using (var connect = new OleDbConnection(connectString))
            {
                try
                {
                    connect.Open();
                    var cmd = connect.CreateCommand();
                    var fileName = Path.GetFileNameWithoutExtension(savePath);
                    var titleList = ForTitleList(jsonData);

                    cmd.CommandText = $"CREATE TABLE {fileName} ({string.Format(ForTitleString(titleList), " NTEXT")})";
                    cmd.ExecuteNonQuery();

                    var paramList = new List<OleDbParameter>();
                    for (var i = 0; i < titleList.Count; ++i)
                    {
                        var param = new OleDbParameter();
                        param.ParameterName = "@" + titleList[i];
                        paramList.Add(param);
                    }
                    cmd.Parameters.AddRange(paramList.ToArray());

                    var valueList = ForValueList(jsonData, titleList);
                    var titleString = ForTitleString(titleList);
                    
                    for (var i = 0; i < valueList.Count; ++i)
                    {
                        cmd.CommandText = $"INSERT INTO {fileName} ({string.Format(titleString, "")}) VALUES(?,?,?,?,?,?)";
                        for (var j = 0; j < titleList.Count; ++j)
                            cmd.Parameters["@" + titleList[j]].Value = valueList[i][j];

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

        private static string ForTitleString(List<string> list)
        {
            var result = string.Empty;
            for (var i = 0; i < list.Count; ++i)
                result += string.Concat(list[i], "{0},");
            result = result.Substring(0, result.Length - 1);

            return result;
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
        #endregion Save
    }
}
