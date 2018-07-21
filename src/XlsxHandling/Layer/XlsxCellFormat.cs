using XlsxHandling.Interfaces.Layer;

namespace XlsxHandling.Layer
{
	public class XlsxCellFormat : IXlsxCellFormat
	{
		public uint NumberFormatId { get; set; }
		public bool ApplyNumberFormat { get; set; }
		public bool ApplyFont { get; set; }
		public bool ApplyFill { get; set; }
		public uint FontId { get; set; }
		public uint FillId { get; set; }
		public uint BorderId { get; set; }
	}
}
