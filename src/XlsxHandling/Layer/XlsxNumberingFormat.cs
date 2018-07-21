using XlsxHandling.Interfaces.Layer;

namespace XlsxHandling.Layer
{
	public class XlsxNumberingFormat : IXlsxNumberingFormat
	{
		public uint NumberingFormatId { get; set; }
		public string FormatCode { get; set; }
	}
}
