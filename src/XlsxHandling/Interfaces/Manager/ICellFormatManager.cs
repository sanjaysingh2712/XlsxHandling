using XlsxHandling.Interfaces.Layer;

namespace XlsxHandling.Interfaces.Manager
{
	public interface ICellFormatManager
	{
		uint GetCellFormat(IXlsxCellFormat fillLayer);
	}
}
