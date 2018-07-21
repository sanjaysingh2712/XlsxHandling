using System;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxHandling.Helper;
using XlsxHandling.Interfaces;
using XlsxHandling.Interfaces.Manager;
using XlsxHandling.Layer;

namespace XlsxHandling.Implementations
{
	public class WorkbookCreator : IWorkbookCreator
	{
		private IStylesheetManager StylesheetManager;
		private ISharedStringManager SharedStringManager;

		public WorkbookCreator(IStylesheetManager stylesheetManager, ISharedStringManager sharedStringManager)
		{
			StylesheetManager = stylesheetManager;
			SharedStringManager = sharedStringManager;
		}

		private Sheets Sheets { get; set; }
		private SheetData SheetData { get; set; }
		private SpreadsheetDocument SsDoc { get; set; }
		private WorkbookPart WorkbookPart { get; set; }

		public void CreateWorkbook(SpreadsheetDocument ssDoc)
		{
			SsDoc = ssDoc;
			// Add a WorkbookPart to the document.
			WorkbookPart = SsDoc.AddWorkbookPart();
			WorkbookPart.Workbook = new Workbook();

			// Add SharedStringTablePart to the Workbook.
			SharedStringManager.SstPart = WorkbookPart.AddNewPart<SharedStringTablePart>();

			// Add StyleSheetPart to the Workbook..
			WorkbookStylesPart stylesPart = WorkbookPart.AddNewPart<WorkbookStylesPart>();
			stylesPart.Stylesheet = StylesheetManager.GetStyleSheet();
			stylesPart.Stylesheet.Save();

			// Add Sheets to the Workbook.
			Sheets = ssDoc.WorkbookPart.Workbook.AppendChild(new Sheets());
		}

		public void CreateWorksheet(uint index, XlsxSheet xlsxSheet)
		{
			// Append a new worksheet and associate  it with the workbook.
			// Add a WorksheetPart to the WorkbookPart.
			WorksheetPart worksheetPart = WorkbookPart.AddNewPart<WorksheetPart>();
			SheetData = new SheetData();
			worksheetPart.Worksheet = new Worksheet(SheetData);

			Sheets.Append(new Sheet {
				Id = SsDoc.WorkbookPart.GetIdOfPart(worksheetPart),
				SheetId = index,
				Name = xlsxSheet.SheetName
			});
		}

		public Cell GetCell(XlsxCell xlsxCell)
		{
			Cell cell = new Cell();
			object value = xlsxCell.Value;
			string valueType = value.GetType().Name;
			CellValue cellValue = new CellValue();
			cell.StyleIndex = StylesheetManager.GetStyleIndex(xlsxCell.IsDateWithoutTime ? "Date" : valueType, xlsxCell);

			switch(valueType) {
				case "Int32":		
					cellValue.Text = xlsxCell.Value.ToString();
					break;
				case "Int64":
				case "Single":
				case "Double":
					XlsxHelper.SetContextEnglish();
					cellValue.Text = xlsxCell.Value.ToString();
					XlsxHelper.ResetContext();
					break;
				case "String":
					cellValue.Text = SharedStringManager.GetIdByValue(value.ToString());
					cell.DataType = CellValues.SharedString;
					break;
				case "DateTime":
					double date = ((DateTime)value).ToOADate();
					XlsxHelper.SetContextEnglish();
					cellValue.Text = date.ToString();
					XlsxHelper.ResetContext();
					break;
				case "Boolean":
					//todo handle boolean
					break;
				default:
					break;
			}
			cell.CellValue = cellValue;
			return cell;
		}

		public void Append(Row row)
		{
			SheetData.Append(row);
		}

		public void Save()
		{
			WorkbookPart.Workbook.Save();
		}
	}
}
