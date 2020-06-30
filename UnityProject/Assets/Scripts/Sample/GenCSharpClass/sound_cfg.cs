//File Generate By ExcelTranslator, Don't Modify It!
using System;
using Engine.Core.ExcelTranslator;

namespace data
{
	public class DataEntry_sound_cfg : DataEntryBase
	{
		public static string sheetName = "sound_cfg";
		public string Key;
		public string zh_cn;
		public string en;
		public string zh_tw;

		public override void DeSerialized(ExcelTranslatorBuffer buffer)
		{
			buffer.Out(out Key);
			buffer.Out(out zh_cn);
			buffer.Out(out en);
			buffer.Out(out zh_tw);
			KEY = Key.ToString();
		}
	}
}