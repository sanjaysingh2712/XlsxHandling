using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxHandling.Implementations.Manager;
using XlsxHandling.Interfaces;
using XlsxHandling.Interfaces.Layer;
using XlsxHandling.Interfaces.Manager;
using XlsxHandling.Layer;
using XlsxHandling.Resources;

namespace XlsxHandling.Implementations
{
	public class XlsxCreator : IXlsxCreator
	{
		IXlsxFile xlsxFile;
		IWorkbookCreator workbookCreator;

		public XlsxCreator(IWorkbookCreator wbCreator)
		{
			workbookCreator = wbCreator;
		}

		public bool Create(IXlsxFile file)
		{
			xlsxFile = file as XlsxFile;
			if(xlsxFile == null) { throw new ArgumentException(XlsxRes.NoFileProvided); }

			XlsxSheet[] xlsxSheets = xlsxFile.Sheets.Cast<XlsxSheet>().ToArray();

			using(SpreadsheetDocument ssDoc = SpreadsheetDocument.Create(xlsxFile.PathToStoreAt, SpreadsheetDocumentType.Workbook)) {

				workbookCreator.CreateWorkbook(ssDoc);

				uint sheetId = 1;
				foreach(XlsxSheet xlsxSheet in xlsxSheets) {

					workbookCreator.CreateWorksheet(sheetId, xlsxSheet);

					IXlsxCell[][] xlsxCells = xlsxSheet.Cells;
					for(uint i = 0; i < xlsxCells.Length; i++) {

						IXlsxCell[] xlsxCellRow = xlsxCells[i];
						Row row = new Row() { RowIndex = i + 1 };

						for(uint j = 0; j < xlsxCellRow.Length; j++) {
							XlsxCell xlsxCell = xlsxCellRow[j] as XlsxCell;
							if(xlsxCell == null) { throw new Exception(XlsxRes.UnknownCellType); }
							Cell cell = workbookCreator.GetCell(xlsxCell);
							row.Append(cell);
						}

						workbookCreator.Append(row);
						sheetId++;
					}
				}
				workbookCreator.Save();
			}

			return true;
		}
	}
}
