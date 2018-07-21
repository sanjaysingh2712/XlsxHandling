using XlsxHandling.Enums;

namespace XlsxHandling.Interfaces.Layer
{
	public interface IXlsxBorder
	{
		BorderType TopBorder { get; set; }
		BorderType RightBorder { get; set; }
		BorderType LeftBorder { get; set; }
		BorderType BottomBorder { get; set; }
		BorderType DiagonalBorder { get; set; }
	}
}
