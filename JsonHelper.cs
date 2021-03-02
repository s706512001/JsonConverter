using Newtonsoft.Json;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace JsonConverter
{
    class JsonHelper
    {
        #region Load
        /// <summary>載入JSON檔案</summary>
        public static Task<List<Dictionary<string, string>>> LoadJsonDataAsync(string jsonFileData)
            => Task.Run(() => LoadJsonData(jsonFileData));

        private static List<Dictionary<string, string>> LoadJsonData(string jsonFilePath)
        {
            if (false == File.Exists(jsonFilePath))
            {
                // TODO show something in dialog
                return null;
            }

            var jsonString = File.ReadAllText(jsonFilePath, Encoding.UTF8);
            var list = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(jsonString);

            return list;
        }
        #endregion Load

        #region Save
        public static Task WriteJsonFileAsync(List<Dictionary<string, string>> jsonData, string savePath)
            => Task.Run(() => WriteJsonFile(jsonData, savePath));

        private static void WriteJsonFile(List<Dictionary<string, string>> jsonData, string savePath)
        {
            var jsonString = JsonConvert.SerializeObject(jsonData, Formatting.Indented);

            File.WriteAllText(savePath, jsonString);
        }
        #endregion Save
    }
}
