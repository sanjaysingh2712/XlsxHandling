using XlsxHandling.Interfaces.Layer;

namespace XlsxHandling.Interfaces.Manager
{
	public interface IBorderManager
	{
		uint GetBorder(IXlsxBorder borderLayer);
		uint GetDefaultBorder();
	}
}
