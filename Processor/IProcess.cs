using System.Collections.Generic;
using System.Threading.Tasks;

namespace JsonConverter
{
    interface IProcess
    {
        Task<List<Dictionary<string, string>>> ReadFileAsync(string filePath);
        Task WriteFileAsync(List<Dictionary<string, string>> jsonData, string savePath);
    }
}
