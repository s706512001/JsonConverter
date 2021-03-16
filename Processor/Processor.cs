using System.Collections.Generic;
using System.Threading.Tasks;

namespace JsonConverter
{
    abstract class Processor : IProcess
    {
        public abstract Task<List<Dictionary<string, string>>> ReadFileAsync(string filePath);
        public abstract Task WriteFileAsync(List<Dictionary<string, string>> jsonData, string savePath);

        /// <summary>取出所有 Column 名稱</summary>
        protected virtual List<string> ForTitleList(List<Dictionary<string, string>> jsonData)
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

        /// <summary>取出所有 Row 的資料</summary>
        protected virtual List<List<string>> ForValueList(List<Dictionary<string, string>> jsonData, List<string> titleList)
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
    }
}
