using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XlsxHandling.Helper
{
	public class XlsxHelper
	{
		public static bool DoubleIsInteger(double value)
		{
			return Math.Abs(value % 1) <= (Double.Epsilon * 100);
		}

		private static CultureInfo originalCultureInfo = null;
		public static void SetContextEnglish()
		{
			originalCultureInfo = Thread.CurrentThread.CurrentCulture;
			Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
		}

		public static void ResetContext()
		{
			if(originalCultureInfo != null) { Thread.CurrentThread.CurrentCulture = originalCultureInfo; }
		}

		public Dictionary<uint, String> BuildFormatMappingsFromXlsx(string fileName)
		{
			Dictionary<uint, String> formatMappings = new Dictionary<uint, String>();

			using(SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, true)) {
				var stylePart = document.WorkbookPart.WorkbookStylesPart;

				var numFormatsParentNodes = stylePart.Stylesheet.ChildElements.OfType<NumberingFormats>();

				foreach(var numFormatParentNode in numFormatsParentNodes) {
					var formatNodes = numFormatParentNode.ChildElements.OfType<NumberingFormat>();
					foreach(var formatNode in formatNodes) {
						formatMappings.Add(formatNode.NumberFormatId.Value, formatNode.FormatCode);
					}
				}
			}

			return formatMappings;
		}
	}
}
