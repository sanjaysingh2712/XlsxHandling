using XlsxHandling.Interfaces.Layer;

namespace XlsxHandling.Interfaces.Manager
{
	public interface IFillManager
	{
		uint GetFill(IXlsxFill fillLayer);
		uint GetDefaultFill();
	}
}
