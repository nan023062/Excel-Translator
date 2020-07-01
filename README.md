# Excel-Translator
A simple and easy-to-use configuration table generation tool

一 核心类 TranslatorTable

该类负责缓存一张表格数据内容(byte格式)并作提供各种数据转换接口。

1.通过Excel、Lua、Json、Bytes等多种数据构造

TranslatorTable(ExcelWorksheet excelSheet)

TranslatorTable(string json)

TranslatorTable(string lua)

TranslatorTable(byte[] bytes)

2.转换成多种数据格式内容

string ToJson(),

string ToLuaTable() ,

byte[] ToBytes() 等等

3.可以生成C#或其他语言的读取类

string ToDataEntryClass()

......

二 工具类ExcelTranslatorUtility

1 负责读取Excel表，获取ExcelWorksheet 格式内容。

2 负责讲目标格式写到文件xx.lua，xx.json，xx.cs，xx.byte等等。

三 工具类ExcelTranslatorBuffer提供字节操作缓冲区

1 用于将数据转换的byte buffer

2 作为C#类的序列化和反序列化工具

当前是用在Unity游戏开发项目中，非常简易而且避免了使用反射来序列化C#对象。
后面我会继续维护并作升级。
目标1是将配置表FieldType(字段类型)的转换方式、格式文件的生成方式，目标语言Data对象的序列化和反序列化方式抽象成对象，可以便于新增语言和数据类型， 
目标2是将工具实现的更简易，于Unity编辑起解耦。
