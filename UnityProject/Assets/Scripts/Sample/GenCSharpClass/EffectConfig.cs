//File Generate By ExcelTranslator, Don't Modify It!
using System;
using Engine.Core.ExcelTranslator;

namespace data
{
	public class DataEntry_EffectConfig : DataEntryBase
	{
		public static string sheetName = "EffectConfig";
		public string id;
		public int type;
		public string path;
		public string bind;
		public int ReceiveLight;
		public string[] music;

		public override void DeSerialized(ExcelTranslatorBuffer buffer)
		{
			buffer.Out(out id);
			buffer.Out(out type);
			buffer.Out(out path);
			buffer.Out(out bind);
			buffer.Out(out ReceiveLight);
			buffer.Out(out music);
			KEY = id.ToString();
		}
	}
}