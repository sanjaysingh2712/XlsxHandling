namespace XlsxHandling.Interfaces.Layer
{
	public interface IXlsxCellFormat
	{
		uint NumberFormatId { get; set; }
		bool ApplyNumberFormat { get; set; }
		bool ApplyFont { get; set; }
		bool ApplyFill { get; set; }
		uint FontId { get; set; }
		uint FillId { get; set; }
		uint BorderId { get; set; }
	}
}
