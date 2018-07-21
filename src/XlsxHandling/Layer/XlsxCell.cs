using XlsxHandling.Interfaces.Layer;

namespace XlsxHandling.Layer
{
	public class XlsxCell : IXlsxCell
	{
		public XlsxCell(object value, IXlsxFont font = null)
		{
			Value = value;
			Font = font;
		}

		public object Value { get; set; }
		public IXlsxFont Font { get; set; }
		public IXlsxBorder Border { get; set; }
		public IXlsxFill Fill { get; set; }
		public IXlsxNumberingFormat NumberingFormat { get; set; }
		public bool ApplyNumberFormat
		{
			get { return NumberingFormat != null && !string.IsNullOrEmpty(NumberingFormat.FormatCode); }
		}
		public bool ApplyFont
		{
			get { return Font != null; }
		}
		public bool ApplyFill
		{
			get { return Fill != null && !string.IsNullOrEmpty(Fill.BackgroundColorArgb); }
		}
		public bool IsDateWithoutTime { get; set; }
	}
}
