using XlsxHandling.Enums;

namespace XlsxHandling.Interfaces.Layer
{
	public interface IXlsxFill
	{
		PatternType PatternType { get; set; }
		string BackgroundColorArgb { get; set; }
	}
}
