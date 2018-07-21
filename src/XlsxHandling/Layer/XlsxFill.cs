using XlsxHandling.Enums;
using XlsxHandling.Interfaces.Layer;

namespace XlsxHandling.Layer
{
	public class XlsxFill : IXlsxFill
	{
		public PatternType PatternType { get; set; }
		public string BackgroundColorArgb { get; set; }
	}
}
