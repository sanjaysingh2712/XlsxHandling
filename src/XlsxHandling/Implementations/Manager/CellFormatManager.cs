using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxHandling.Interfaces.Layer;
using XlsxHandling.Interfaces.Manager;
using XlsxHandling.Layer;

namespace XlsxHandling.Implementations.Manager
{
	public class CellFormatManager : ICellFormatManager
	{
		public CellFormatManager()
		{
			ActualCellFormats = new CellFormats();
			SetDefaultCellFormat();
		}

		public CellFormats ActualCellFormats { get; set; }

		public uint GetCellFormat(IXlsxCellFormat fillLayer)
		{
			XlsxCellFormat layer = fillLayer as XlsxCellFormat;
			if(layer == null) { throw new Exception("Unexpected Cell Format Type"); }

			return GetCellFormat(layer.NumberFormatId, layer.ApplyNumberFormat, layer.ApplyFont, layer.ApplyFill, layer.FontId, layer.FillId, layer.BorderId);
		}

		private uint GetCellFormat(uint numberFormatId, bool applyNumberFormat, bool applyFont, bool applyFill, uint fontId, uint fillId, uint borderId)
		{
			//Check numbering format exists and return index
			uint cellFormatIndex = 0;
			foreach(var cellFormat in ActualCellFormats.ChildElements.OfType<CellFormat>().ToArray()) {
				if(cellFormat.NumberFormatId == numberFormatId &&
					(cellFormat.ApplyNumberFormat ?? false) == applyNumberFormat &&
					(cellFormat.ApplyFont ?? false) == applyFont &&
					(cellFormat.ApplyFill ?? false) == applyFill &&
					cellFormat.FontId == fontId &&
					cellFormat.FillId == fillId &&
					cellFormat.BorderId == borderId) {
					return cellFormatIndex;
				}
				cellFormatIndex++;
			}

			//if not exists create new cell format and append to actual cell formats
			SetCellFormat(numberFormatId, applyNumberFormat, applyFont, applyFill, fontId, fillId, borderId);
			return cellFormatIndex;
		}

		private void SetCellFormat(uint numberFormatId, bool applyNumberFormat, bool applyFont, bool applyFill, uint fontId, uint fillId, uint borderId)
		{
			CellFormat cellFormat = new CellFormat();
			cellFormat.NumberFormatId = UInt32Value.FromUInt32(numberFormatId);

			if(applyNumberFormat) {
				cellFormat.ApplyNumberFormat = BooleanValue.FromBoolean(applyNumberFormat);
			}
			if(applyFont) {
				cellFormat.ApplyFont = BooleanValue.FromBoolean(applyFont);
			}
			if(applyFill) {
				cellFormat.ApplyFill = BooleanValue.FromBoolean(applyFill);
			}
			cellFormat.FontId = UInt32Value.FromUInt32(fontId);
			cellFormat.FillId = UInt32Value.FromUInt32(fillId);
			cellFormat.BorderId = UInt32Value.FromUInt32(borderId);
			//cellFormat.FormatId = UInt32Value.FromUInt32(cellStyleFormatId);

			ActualCellFormats.Append(cellFormat);
			SetCellFormatsCount();
		}

		internal uint DefaultStyleIndex { get; private set; }
		internal uint NumberStyleIndex { get; private set; }
		internal uint DateStyleIndex { get; private set; }
		internal uint FloatingPointStyleIndex { get; private set; }
		internal uint DateTimeStyleIndex { get; private set; }

		private void SetCellFormatsCount()
		{
			ActualCellFormats.Count = (uint)ActualCellFormats.ChildElements.Count;
		}

		private void SetDefaultCellFormat()
		{
			SetCellFormat(0, false, false, false, 0, 0, 0); //standard
			DefaultStyleIndex = 0;
			SetCellFormat(2, true, false, false, 0, 0, 0); //floating point
			FloatingPointStyleIndex = 2;
			SetCellFormat(1, true, false, false, 0, 0, 0); //integer
			NumberStyleIndex = 1;
			SetCellFormat(14, true, false, false, 0, 0, 0); //date
			DateStyleIndex = 14;
			SetCellFormat(169, true, false, false, 0, 0, 0); //datetime
			DateTimeStyleIndex = 169;
			//CellFormat cellFormat5 = GetCellFormat(164, true); //date
			//CellFormat cellFormat6 = GetCellFormat(169, true); //date
		}
	}
}
