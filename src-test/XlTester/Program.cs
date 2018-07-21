using System;
using System.Collections.Generic;
using Autofac;
using Autofac.Core;
//using Autofac.Core;
using XlsxHandling.Enums;
using XlsxHandling.Implementations;
using XlsxHandling.Implementations.Manager;
using XlsxHandling.Interfaces;
using XlsxHandling.Interfaces.Layer;
using XlsxHandling.Interfaces.Manager;
using XlsxHandling.Layer;

namespace XlTester
{
	public class Program
	{
		public static void Main(string[] args)
		{
			string path = @"..\..\testfile.xlsx";
			IList<IXlsxSheet> sheets = SetSheets();

			IXlsxFile file = new XlsxFile() {
				PathToStoreAt = path,
				Sheets = sheets
			};

			IContainer container = BuildIocContainer();

			IXlsxCreator creator = container.Resolve<IXlsxCreator>();

			//IXlsxCreator creator = new XlsxCreator(new WorkbookCreator(new StylesheetManager(new FontManager(), new BorderManager(), new FillManager(), new NumberingFormatManager(), new CellFormatManager()), new SharedStringManager()));
			try {
				creator.Create(file);
				Console.WriteLine("{0} created", path);
			} catch(Exception ex) {
				Console.WriteLine(ex.Message);
			}
			Console.ReadLine();
		}

		private static IContainer BuildIocContainer()
		{
			//Create builder
			var builder = new ContainerBuilder();

			//Register types
			builder.RegisterType<XlsxCreator>().As<IXlsxCreator>();
			builder.RegisterType<WorkbookCreator>().As<IWorkbookCreator>();
			builder.RegisterType<StylesheetManager>().As<IStylesheetManager>();
			builder.RegisterType<FontManager>().As<IFontManager>();
			builder.RegisterType<BorderManager>().As<IBorderManager>();
			builder.RegisterType<FillManager>().As<IFillManager>();
			builder.RegisterType<NumberingFormatManager>().As<INumberingFormatManager>();
			builder.RegisterType<CellFormatManager>().As<ICellFormatManager>();
			builder.RegisterType<SharedStringManager>().As<ISharedStringManager>();

			//return container
			return builder.Build();
		}

		private static IList<IXlsxSheet> SetSheets()
		{
			IList<IXlsxSheet> sheets = new List<IXlsxSheet>();
			XlsxSheet sheet = new XlsxSheet();
			sheet.SheetName = "Just to test";

			IXlsxFill fill1 = new XlsxFill {
				PatternType = PatternType.DarkTrellis
			};

			IXlsxFill fill2 = new XlsxFill {
				PatternType = PatternType.LightTrellis,
				BackgroundColorArgb = "FF0000FF"
			};

			IXlsxBorder border1 = new XlsxBorder {
				TopBorder = BorderType.Dashed,
				RightBorder = BorderType.Double
			};

			IXlsxBorder border2 = new XlsxBorder {
				DiagonalBorder = BorderType.MediumDashDot
			};

			XlsxCell[] header = new XlsxCell[] {
				new XlsxCell("Name") { Fill = fill1, Border = border1 },
				new XlsxCell("Zahl") { Fill = fill1, Border = border1 },
				new XlsxCell("Datum") { Fill = fill1, Border = border1 }
			};
			XlsxCell[] row1 = new XlsxCell[] {
				new XlsxCell("Sanjay", new XlsxFont{
					FontType = FontType.Calibri,
					Bold = true,
					Italic = true,
					Size = 13,
					UnderlineType = UnderlineType.DoubleLine,
					ColorArgb = "FF00FF00"
				}){
					Fill = fill1
				},
				new XlsxCell(5.2765) {
					Fill = fill2
				},
				new XlsxCell(new DateTime(2018,1,1,12,12,12)){ Border = border2 }
			};

			IXlsxCell[][] cells = new XlsxCell[][] {
				header,
				row1
			};

			sheet.Cells = cells;
			sheets.Add(sheet);
			return sheets;
		}
	}
}
