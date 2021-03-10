# JsonConverter
JsonConverter

### 前置作業

1. 要先安裝 Microsoft Access Database Engine 2010 可轉散發套件

    https://www.microsoft.com/zh-tw/download/details.aspx?id=13255
  
    建議是 32bit 和 64bit 兩個版本都安裝
  
    安裝完其中一個後，要另外使用 command 的方式安裝第二個
    ```
     AccessDatabaseEngine.exe /passive
    ```
### Command Line 使用方式

./JsonConverter.exe [Source] [ConverTo Type]

```
ConvertTo Type (大小寫不拘)
  json
  csv
  xlsx
  
目前只提供以下轉換方案
  json -> csv
  json -> xlsx
  csv -> json
  xlsx -> json
  
不在此範圍內的轉換方案，理應會在 information 位置顯示錯誤訊息
```
