using XlsxHandling.Enums;
using XlsxHandling.Interfaces.Layer;

namespace XlsxHandling.Layer
{
	public class XlsxBorder : IXlsxBorder
	{
		public BorderType TopBorder { get; set; }
		public BorderType RightBorder { get; set; }
		public BorderType LeftBorder { get; set; }
		public BorderType BottomBorder { get; set; }
		public BorderType DiagonalBorder { get; set; }
	}
}
