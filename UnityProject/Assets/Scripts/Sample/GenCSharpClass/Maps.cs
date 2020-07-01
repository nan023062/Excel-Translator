//File Generate By ExcelTranslator, Don't Modify It!
using System;
using Engine.Core.ExcelTranslator;

namespace data
{
	public class DataEntry_Maps : DataEntryBase
	{
		public static string sheetName = "Maps";
		public string ID ;
		public int ItemID;
		public string MapNameString;
		public string Route;
		public string ModelName;
		public int MapType;

		public override void DeSerialized(ExcelTranslatorBuffer buffer)
		{
			buffer.Out(out ID );
			buffer.Out(out ItemID);
			buffer.Out(out MapNameString);
			buffer.Out(out Route);
			buffer.Out(out ModelName);
			buffer.Out(out MapType);
			KEY = ID .ToString();
		}
	}
}