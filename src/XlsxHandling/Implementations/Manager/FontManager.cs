using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxHandling.Enums;
using XlsxHandling.Interfaces.Layer;
using XlsxHandling.Interfaces.Manager;
using XlsxHandling.Layer;

namespace XlsxHandling.Implementations.Manager
{
	public class FontManager : IFontManager
	{
		public FontManager()
		{
			ActualFonts = new Fonts();
			SetDefaultFont();
		}

		public uint GetFont(IXlsxFont fontLayer)
		{
			XlsxFont layer = fontLayer as XlsxFont;
			if(layer == null) { throw new Exception("Unexpected Font Type"); }

			return GetFont(layer.Size, layer.FontType.ToString(), XlsxFont.FontFamilyId, layer.Bold, layer.Italic, layer.UnderlineType, layer.ColorArgb);
		}

		public uint GetDefaultFont() => 0;

		public Fonts ActualFonts { get; set; }

		/// <summary>
		/// maybe later FontSchemeValues fontScheme = FontSchemeValues.None,
		/// and uint colorSchemeId = 0
		/// </summary>
		/// <param name="size"></param>
		/// <param name="fontName"></param>
		/// <param name="fontFamilyId"></param>
		/// <param name="bold"></param>
		/// <param name="italic"></param>
		/// <param name="underlineValue"></param>
		/// <returns></returns>
		private uint GetFont(double size, string fontName, int fontFamilyId, bool bold, bool italic, UnderlineType underlineType, string fontColor)
		{
			UnderlineValues underlineValue;
			switch(underlineType) {
				case UnderlineType.SingleLine:
					underlineValue = UnderlineValues.Single;
					break;
				case UnderlineType.DoubleLine:
					underlineValue = UnderlineValues.Double;
					break;
				case UnderlineType.None:
					underlineValue = UnderlineValues.None;
					break;
				default:
					throw new Exception("Unknown Underline Type");
			}

			//Check font type exists and return index
			uint fontIndex = 0;
			foreach(var fnt in ActualFonts.ChildElements.OfType<Font>()) {
				if(fnt.FontSize.Val == size &&
					fnt.FontName.Val == fontName &&
					fnt.FontFamilyNumbering.Val == fontFamilyId &&
					fnt.Bold.Val == bold &&
					fnt.Italic.Val == italic &&
					fnt.Underline.Val.Value == underlineValue &&
					(string.IsNullOrEmpty(fontColor) || fnt.Color == null || fnt.Color.Rgb == fontColor)) {
					return fontIndex;
				}
				fontIndex++;
			}

			//if not exists create new font and append to actual fonts
			SetFont(size, fontName, fontFamilyId, bold, italic, underlineValue, fontColor);
			return fontIndex;
		}

		private void SetFont(double size, string fontName, int fontFamilyId, bool bold, bool italic, UnderlineValues underlineValue, string fontColor = null)
		{
			Font font = new Font();
			font.FontSize = new FontSize() { Val = DoubleValue.FromDouble(size) };
			font.FontName = new FontName() { Val = StringValue.FromString(fontName) };
			font.FontFamilyNumbering = new FontFamilyNumbering() { Val = Int32Value.FromInt32(fontFamilyId) };
			font.Bold = new Bold() { Val = BooleanValue.FromBoolean(bold) };
			font.Italic = new Italic() { Val = BooleanValue.FromBoolean(italic) };
			font.Underline = new Underline() { Val = new EnumValue<UnderlineValues>(underlineValue) };
			if(!string.IsNullOrEmpty(fontColor)) { font.Color = new Color() { Rgb = fontColor }; }
			
			//font.FontScheme = new FontScheme() { Val = new EnumValue<FontSchemeValues>(fontScheme) };
			//font.Color = new Color() { Theme = UInt32Value.FromUInt32(colorSchemeId) };

			ActualFonts.Append(font);
			SetFontsCount();
		}

		private void SetDefaultFont()
		{
			SetFont(11, "Calibri", 2, false, false, UnderlineValues.None);
		}

		private void SetFontsCount()
		{
			ActualFonts.Count = (uint)ActualFonts.ChildElements.Count;
		}


	}
}
