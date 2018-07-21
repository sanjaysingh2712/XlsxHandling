using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxHandling.Interfaces.Layer;
using XlsxHandling.Interfaces.Manager;
using XlsxHandling.Layer;

namespace XlsxHandling.Implementations.Manager
{
	public class NumberingFormatManager : INumberingFormatManager
	{
		public NumberingFormatManager()
		{
			ActualNumberingFormats = new NumberingFormats();
			SetDefaultNumberingFormats();
		}

		public NumberingFormats ActualNumberingFormats { get; set; }

		public uint GetNumberingFormat(IXlsxNumberingFormat numberingFormatLayer)
		{
			XlsxNumberingFormat layer = numberingFormatLayer as XlsxNumberingFormat;
			if(layer == null) { throw new Exception("Unexpected Numbering Format Type"); }

			return GetNumberingFormat(layer.NumberingFormatId, layer.FormatCode);
		}

		public uint GetDefaultNumberingFormat() => 0;

		private uint GetNumberingFormat(uint numberingFormatId, string formatCode)
		{
			//Check numbering format exists and return index
			uint numberingFormatIndex = 0;
			foreach(var numberingFormat in ActualNumberingFormats.ChildElements.OfType<NumberingFormat>()) {
				if(numberingFormat.NumberFormatId == numberingFormatId && numberingFormat.FormatCode == formatCode) {
					return numberingFormatIndex;
				}
				numberingFormatIndex++;
			}

			//if not exists create new numbering format and append to actual numbering formats
			SetNumberingFormat(numberingFormatId, formatCode);
			return numberingFormatIndex;
		}

		private void SetNumberingFormat(uint numberingFormatId, string formatCode = null)
		{
			NumberingFormat numberingFormat = new NumberingFormat();
			numberingFormat.NumberFormatId = UInt32Value.FromUInt32(numberingFormatId);
			if(formatCode != null) { numberingFormat.FormatCode = StringValue.FromString(formatCode); }

			ActualNumberingFormats.Append(numberingFormat);
			SetNumberingFormatsCount();
		}

		private void SetNumberingFormatsCount()
		{
			ActualNumberingFormats.Count = (uint)ActualNumberingFormats.ChildElements.Count;
		}

		private void SetDefaultNumberingFormats()
		{
			SetNumberingFormat(164, @"[$/*-*/F800]dddd\,\ mmmm\ dd\,\ yyyy");
			SetNumberingFormat(169, @"d/m/yyyy\ h:mm;@");
			SetNumberingFormat(171, @"dd/mm/yyyy\ hh:mm:ss;@");
		}
	}
}
