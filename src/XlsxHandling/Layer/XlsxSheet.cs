using XlsxHandling.Interfaces.Layer;

namespace XlsxHandling.Layer
{
	public class XlsxSheet : IXlsxSheet
	{
		public string SheetName { get; set; }
		public IXlsxCell[][] Cells { get; set; }
		public XlsxFont GlobalFont { get; set; }
	}
}
