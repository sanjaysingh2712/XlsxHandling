namespace XlsxHandling.Interfaces.Layer
{
	public interface IXlsxCell
	{
		object Value { get; set; }
		IXlsxFont Font { get; set; }
		IXlsxBorder Border { get; set; }
		IXlsxFill Fill { get; set; }
		IXlsxNumberingFormat NumberingFormat { get; set; }
		bool ApplyNumberFormat { get; }
		bool ApplyFont { get; }
		bool ApplyFill { get; }
		bool IsDateWithoutTime { get; set; }
	}
}