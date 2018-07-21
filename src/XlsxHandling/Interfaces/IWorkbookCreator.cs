using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxHandling.Layer;

namespace XlsxHandling.Interfaces
{
	public interface IWorkbookCreator
	{
		void CreateWorkbook(SpreadsheetDocument ssDoc);
		void CreateWorksheet(uint index, XlsxSheet xlsxSheet);	
		Cell GetCell(XlsxCell xlsxCell);
		void Save();
		void Append(Row row);
	}
}
