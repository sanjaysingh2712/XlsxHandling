using XlsxHandling.Interfaces.Layer;

namespace XlsxHandling.Interfaces.Manager
{
	public interface INumberingFormatManager
	{
		uint GetNumberingFormat(IXlsxNumberingFormat numberingFormatLayer);
		uint GetDefaultNumberingFormat();
	}
}
