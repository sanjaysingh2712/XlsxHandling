using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxHandling.Layer;

namespace XlsxHandling.Interfaces.Manager
{
	public interface IStylesheetManager
	{
		Stylesheet GetStyleSheet();
		UInt32Value GetStyleIndex(string valueType, XlsxCell xlsxCell);
	}
}
