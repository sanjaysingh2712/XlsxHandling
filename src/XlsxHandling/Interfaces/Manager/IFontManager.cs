using XlsxHandling.Interfaces.Layer;

namespace XlsxHandling.Interfaces.Manager
{
	public interface IFontManager
	{
		uint GetFont(IXlsxFont fontLayer);
		uint GetDefaultFont();
	}
}
