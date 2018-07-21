using XlsxHandling.Enums;
using XlsxHandling.Interfaces.Layer;

namespace XlsxHandling.Layer
{
	public class XlsxFont : IXlsxFont
	{
		public XlsxFont(
			int fontSize = 11, 
			FontType fontType = FontType.Calibri, 
			bool bold = false, 
			bool italic = false, 
			UnderlineType underlineType = UnderlineType.None)
		{
			Size = fontSize;
			FontType = fontType;
			Bold = bold;
			Italic = italic;
			UnderlineType = underlineType;
		}

		public const int FontFamilyId = 2;
		public int Size { get; set; }
		public FontType FontType { get; set; }
		public bool Bold { get; set; }
		public bool Italic { get; set; }
		public UnderlineType UnderlineType { get; set; }
		public string ColorArgb { get; set; }
	}
}
