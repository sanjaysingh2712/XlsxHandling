using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxHandling.Interfaces.Layer;
using XlsxHandling.Interfaces.Manager;
using XlsxHandling.Layer;

namespace XlsxHandling.Implementations.Manager
{
	public class StylesheetManager : IStylesheetManager
	{
		public StylesheetManager(IFontManager fontManager, IBorderManager borderManager, IFillManager fillManager, INumberingFormatManager numberingFormatManager, ICellFormatManager cellFormatManager)
		{
			FontManager = fontManager as FontManager;
			BorderManager = borderManager as BorderManager;
			FillManager = fillManager as FillManager;
			NumberingFormatManager = numberingFormatManager as NumberingFormatManager;
			CellFormatManager = cellFormatManager as CellFormatManager;

			ActualStyleSheet = new Stylesheet {
				Fonts = FontManager.ActualFonts,
				Borders = BorderManager.ActualBorders,
				Fills = FillManager.ActualFills,
				NumberingFormats = NumberingFormatManager.ActualNumberingFormats,
				CellFormats = CellFormatManager.ActualCellFormats
			};
		}

		public Stylesheet GetStyleSheet()
		{
			return ActualStyleSheet;
		}

		private FontManager FontManager { get; set; }
		private BorderManager BorderManager { get; set; }
		private FillManager FillManager { get; set; }
		private NumberingFormatManager NumberingFormatManager { get; set; }
		private CellFormatManager CellFormatManager { get; set; }

		private Stylesheet ActualStyleSheet { get; set; }

		public UInt32Value GetStyleIndex(string valueType, XlsxCell xlsxCell)
		{
			uint fontId = xlsxCell.Font != null
				? FontManager.GetFont(xlsxCell.Font)
				: FontManager.GetDefaultFont();
			uint borderId = xlsxCell.Border != null
				? BorderManager.GetBorder(xlsxCell.Border)
				: BorderManager.GetDefaultBorder();
			uint fillId = xlsxCell.Fill != null
				? FillManager.GetFill(xlsxCell.Fill)
				: FillManager.GetDefaultFill();
			uint numberingFormatId = xlsxCell.NumberingFormat != null
				? NumberingFormatManager.GetNumberingFormat(xlsxCell.NumberingFormat)
				: GetIndexByDataType(valueType);

			IXlsxCellFormat cellFormat = new XlsxCellFormat {
				FontId = fontId,
				BorderId = borderId,
				FillId = fillId,
				NumberFormatId = numberingFormatId,
				ApplyNumberFormat = xlsxCell.ApplyNumberFormat,
				ApplyFont = xlsxCell.ApplyFont,
				ApplyFill = xlsxCell.ApplyFill
			};

			return CellFormatManager.GetCellFormat(cellFormat);
		}

		private uint GetIndexByDataType(string valueType)
		{
			switch(valueType) {
				case "Int32":
					return CellFormatManager.NumberStyleIndex;
				case "Int64":
				case "Single":
				case "Double":
					return CellFormatManager.FloatingPointStyleIndex;
				case "String":
					return CellFormatManager.DefaultStyleIndex;
				case "DateTime":
					return CellFormatManager.DateTimeStyleIndex;
				case "Date":
					return CellFormatManager.DateStyleIndex;
				case "Boolean":
					return CellFormatManager.NumberStyleIndex;
				default:
					throw new Exception("Unexpected Data Type");
			}

		}

		private void CreateStyleSheetOld()
		{
			//needed to create xlsx with stylesheet
			//ActualStyleSheet.Borders = BorderManager.ActualBorders;
			//ActualStyleSheet.Fills = FillManager.ActualFills;
			//ActualStyleSheet.Fonts = FontManager.ActualFonts;
			//ActualStyleSheet.NumberingFormats = NumberingFormatManager.ActualNumberingFormats;
			//ActualStyleSheet.CellFormats = CellFormatManager.ActualCellFormats;

			//not needed to create xlsx with stylesheet
			//ActualStyleSheet.CellStyleFormats = GetCellStyleFormats();
			//ActualStyleSheet.CellStyles = GetCellStyles();
			//ActualStyleSheet.TableStyles = GetTableStyles();
			//ActualStyleSheet.DifferentialFormats = GetDifferentialFormats();
			//ActualStyleSheet.StylesheetExtensionList = GetStylesheetExtensionList();
		}

		#region Not needed to create xlsx

		private StylesheetExtensionList GetStylesheetExtensionList()
		{
			StylesheetExtensionList stylesheetExtensionList = new StylesheetExtensionList();
			return stylesheetExtensionList;
		}

		private DifferentialFormats GetDifferentialFormats()
		{
			DifferentialFormats differentialFormats = new DifferentialFormats();
			differentialFormats.Count = 0;
			return differentialFormats;
		}

		private TableStyles GetTableStyles()
		{
			TableStyles tableStyles = new TableStyles();
			tableStyles.DefaultTableStyle = "TableStyleMedium9";
			tableStyles.DefaultPivotStyle = "PivotStyleLight16";
			tableStyles.Count = 0;
			return tableStyles;
		}

		private CellStyleFormats GetCellStyleFormats()
		{
			CellStyleFormats cellStyleFormats = new CellStyleFormats();
			CellFormat cellStyleFormat = CellFormatManager.ActualCellFormats.Elements<CellFormat>().FirstOrDefault();
			cellStyleFormats.Append(cellStyleFormat);
			SetCellStyleFormatsCount(cellStyleFormats);
			return cellStyleFormats;
		}

		private void SetCellStyleFormatsCount(CellStyleFormats cellStyleFormats)
		{
			cellStyleFormats.Count = (uint)cellStyleFormats.ChildElements.Count;
		}

		private CellStyles GetCellStyles()
		{
			CellStyles cellStyles = new CellStyles();
			CellStyle cellStyle = GetCellStyle("Standard");
			cellStyles.Append(cellStyle);
			SetCellStylesCount(cellStyles);
			return cellStyles;
		}

		private CellStyle GetCellStyle(string cellStyleName, uint cellStyleFormatId = 0, uint builtInId = 0)
		{
			CellStyle cellStyle = new CellStyle();
			cellStyle.Name = cellStyleName;
			cellStyle.FormatId = cellStyleFormatId;
			cellStyle.BuiltinId = builtInId;
			return cellStyle;
		}

		private void SetCellStylesCount(CellStyles cellStyles)
		{
			cellStyles.Count = (uint)cellStyles.ChildElements.Count;
		}

		#endregion
	}
}
