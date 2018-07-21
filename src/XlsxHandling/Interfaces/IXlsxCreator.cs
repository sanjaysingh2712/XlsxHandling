using XlsxHandling.Interfaces.Layer;

namespace XlsxHandling.Interfaces
{
	public interface IXlsxCreator
	{
		bool Create(IXlsxFile file);
	}
}
