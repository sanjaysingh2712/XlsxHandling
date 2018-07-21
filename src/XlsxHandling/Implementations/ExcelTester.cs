using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;

namespace TestExcel
{
    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public class ExcelTester
	{


		public static Stylesheet GenerateStylesheet2()
		{
			Stylesheet ss = new Stylesheet();

			Fonts fts = new Fonts();
			DocumentFormat.OpenXml.Spreadsheet.Font ft = new DocumentFormat.OpenXml.Spreadsheet.Font();
			FontName ftn = new FontName();
			ftn.Val = "Calibri";
			FontSize ftsz = new FontSize();
			ftsz.Val = 11;
			ft.FontName = ftn;
			ft.FontSize = ftsz;
			fts.Append(ft);
			fts.Count = (uint)fts.ChildElements.Count;

			Fills fills = new Fills();
			Fill fill;
			PatternFill patternFill;
			fill = new Fill();
			patternFill = new PatternFill();
			patternFill.PatternType = PatternValues.None;
			fill.PatternFill = patternFill;
			fills.Append(fill);
			fill = new Fill();
			patternFill = new PatternFill();
			patternFill.PatternType = PatternValues.Gray125;
			fill.PatternFill = patternFill;
			fills.Append(fill);
			fills.Count = (uint)fills.ChildElements.Count;

			Borders borders = new Borders();
			Border border = new Border();
			border.LeftBorder = new LeftBorder();
			border.RightBorder = new RightBorder();
			border.TopBorder = new TopBorder();
			border.BottomBorder = new BottomBorder();
			border.DiagonalBorder = new DiagonalBorder();
			borders.Append(border);
			borders.Count = (uint)borders.ChildElements.Count;

			CellStyleFormats csfs = new CellStyleFormats();
			CellFormat cf = new CellFormat();
			cf.NumberFormatId = 0;
			cf.FontId = 0;
			cf.FillId = 0;
			cf.BorderId = 0;
			csfs.Append(cf);
			csfs.Count = (uint)csfs.ChildElements.Count;

			uint iExcelIndex = 164;
			NumberingFormats nfs = new NumberingFormats();
			CellFormats cfs = new CellFormats();

			cf = new CellFormat();
			cf.NumberFormatId = 0;
			cf.FontId = 0;
			cf.FillId = 0;
			cf.BorderId = 0;
			cf.FormatId = 0;
			cfs.Append(cf);

			NumberingFormat nf;
			nf = new NumberingFormat();
			nf.NumberFormatId = iExcelIndex++;
			nf.FormatCode = "dd/mm/yyyy hh:mm:ss";
			nfs.Append(nf);
			cf = new CellFormat();
			cf.NumberFormatId = nf.NumberFormatId;
			cf.FontId = 0;
			cf.FillId = 0;
			cf.BorderId = 0;
			cf.FormatId = 0;
			cf.ApplyNumberFormat = true;
			cfs.Append(cf);

			nf = new NumberingFormat();
			nf.NumberFormatId = iExcelIndex++;
			nf.FormatCode = "#,##0.0000";
			nfs.Append(nf);
			cf = new CellFormat();
			cf.NumberFormatId = nf.NumberFormatId;
			cf.FontId = 0;
			cf.FillId = 0;
			cf.BorderId = 0;
			cf.FormatId = 0;
			cf.ApplyNumberFormat = true;
			cfs.Append(cf);

			// #,##0.00 is also Excel style index 4
			nf = new NumberingFormat();
			nf.NumberFormatId = iExcelIndex++;
			nf.FormatCode = "#,##0.00";
			nfs.Append(nf);
			cf = new CellFormat();
			cf.NumberFormatId = nf.NumberFormatId;
			cf.FontId = 0;
			cf.FillId = 0;
			cf.BorderId = 0;
			cf.FormatId = 0;
			cf.ApplyNumberFormat = true;
			cfs.Append(cf);

			// @ is also Excel style index 49
			nf = new NumberingFormat();
			nf.NumberFormatId = iExcelIndex++;
			nf.FormatCode = "@";
			nfs.Append(nf);
			cf = new CellFormat();
			cf.NumberFormatId = nf.NumberFormatId;
			cf.FontId = 0;
			cf.FillId = 0;
			cf.BorderId = 0;
			cf.FormatId = 0;
			cf.ApplyNumberFormat = true;
			cfs.Append(cf);

			nfs.Count = (uint)nfs.ChildElements.Count;
			cfs.Count = (uint)cfs.ChildElements.Count;

			ss.Append(nfs);
			ss.Append(fts);
			ss.Append(fills);
			ss.Append(borders);
			ss.Append(csfs);
			ss.Append(cfs);

			CellStyles css = new CellStyles();
			CellStyle cs = new CellStyle();
			cs.Name = "Normal";
			cs.FormatId = 0;
			cs.BuiltinId = 0;
			css.Append(cs);
			css.Count = (uint)css.ChildElements.Count;
			ss.Append(css);

			DifferentialFormats dfs = new DifferentialFormats();
			dfs.Count = 0;
			ss.Append(dfs);

			TableStyles tss = new TableStyles();
			tss.Count = 0;
			tss.DefaultTableStyle = "TableStyleMedium9";
			tss.DefaultPivotStyle = "PivotStyleLight16";
			ss.Append(tss);

			return ss;
		}
		//var numberCell = new Cell {
		//	DataType = CellValues.Number,
		//	CellReference = header + index,
		//	CellValue = new CellValue(text),
		//	StyleIndex = 3
		//};
		public void Execute()
		{
			List<Person> personen = GetData();
			String pathToTemp = @"..\..\testFolder";
			if(!Directory.Exists(pathToTemp)) {
				Directory.CreateDirectory(pathToTemp);
			}

			String tempFile = pathToTemp + "\\" + DateTime.Now.Ticks / TimeSpan.TicksPerMillisecond + ".xlsx";

			using(SpreadsheetDocument spreadSheet = SpreadsheetDocument.Create(tempFile, SpreadsheetDocumentType.Workbook)) {
				// create the workbook

				var dateFormat = new NumberingFormat() {
					NumberFormatId = (UInt32Value)0,
					FormatCode = StringValue.FromString("dd.MM.yyyy")
				};
				WorkbookPart part = spreadSheet.AddWorkbookPart();
				part.Workbook = new Workbook();
				part.AddNewPart<WorksheetPart>();
				part.WorksheetParts.First().Worksheet = new Worksheet();
				WorkbookStylesPart sp = spreadSheet.WorkbookPart.AddNewPart<WorkbookStylesPart>();
				sp.Stylesheet = new Stylesheet();
				sp.Stylesheet.NumberingFormats = new NumberingFormats();
				sp.Stylesheet.NumberingFormats.Append(dateFormat);

				CellFormat cellFormat = new CellFormat();
				cellFormat.NumberFormatId = dateFormat.NumberFormatId;
				cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(true);


				sp.Stylesheet.CellFormats = new CellFormats();
				sp.Stylesheet.CellFormats.AppendChild<CellFormat>(cellFormat);

				sp.Stylesheet.CellFormats.Count = UInt32Value.FromUInt32((uint)sp.Stylesheet.CellFormats.ChildElements.Count);


				sp.Stylesheet.Save();
				Sheets sheets = spreadSheet.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());


				// create sheet data

				foreach(Person p in personen) {
					WorksheetPart tabPart = part.AddNewPart<WorksheetPart>();
					Worksheet workSheet1 = new Worksheet();
					SheetData sheetData1 = new SheetData();

					Sheet sheet1 = new Sheet() {
						Id = spreadSheet.WorkbookPart.GetIdOfPart(tabPart),
						SheetId = 1,
						Name = p.Name
					};

					sheets.Append(sheet1);

					Row r = new Row();

					r.AppendChild(new Cell() { CellValue = new CellValue(p.Name), DataType = CellValues.String });
					r.AppendChild(new Cell() { CellValue = new CellValue(p.BirthDay.ToOADate().ToString()), StyleIndex = 0 });
					r.AppendChild(new Cell() { CellValue = new CellValue(p.HeightInCm.ToString(CultureInfo.InvariantCulture)), DataType = CellValues.Number });
					r.AppendChild(new Cell() { CellValue = new CellValue(p.Weight.ToString(CultureInfo.InvariantCulture)), DataType = CellValues.Number });


					sheetData1.AppendChild(r);


					workSheet1.AppendChild(sheetData1);
					tabPart.Worksheet = workSheet1;
				}
				part.Workbook.Save();

			}

			Process.Start(@tempFile);
			Environment.Exit(0);
		}
        public List<Person> GetData()
        {
            List<Person> result = new List<Person>();
            result.Add(new Person
            {
                Name = "ASDF",
                Weight = 192.8,
                BirthDay = DateTime.Now,
                HeightInCm = 198
            });


            result.Add(new Person
            {
                Name = "ASe2",
                Weight = 23.8,
                BirthDay = DateTime.Now,
                HeightInCm = 57
            });

            return result;
        }
    }

	public class Person
	{
		public string Name { get; internal set; }
		public double Weight { get; internal set; }
		public DateTime BirthDay { get; internal set; }
		public int HeightInCm { get; internal set; }
	}
}
