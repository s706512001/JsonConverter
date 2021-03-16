using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JsonConverter
{
    class JsonProcessor : Processor
    {
        public override Task<List<Dictionary<string, string>>> ReadFileAsync(string filePath)
            => Task.Run(() => LoadJsonData(filePath));

        public override Task WriteFileAsync(List<Dictionary<string, string>> jsonData, string savePath)
            => Task.Run(() => WriteJsonFile(jsonData, savePath));

        private List<Dictionary<string, string>> LoadJsonData(string jsonFilePath)
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

        private static void WriteJsonFile(List<Dictionary<string, string>> jsonData, string savePath)
        {
            var jsonString = JsonConvert.SerializeObject(jsonData, Formatting.Indented);

            File.WriteAllText(savePath, jsonString);
        }
    }
}
