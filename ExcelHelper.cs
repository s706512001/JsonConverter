using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using System.Windows.Forms;

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

        private static void WriteXlsxData(List<Dictionary<string, string>> jsonData, string savePath)
        {

        }
        #endregion Save
    }
}
