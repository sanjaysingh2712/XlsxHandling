namespace XlsxHandling.Interfaces.Layer
{
	public interface IXlsxNumberingFormat
	{
		uint NumberingFormatId { get; set; }
		string FormatCode { get; set; }
	}
}
