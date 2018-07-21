using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxHandling.Enums;
using XlsxHandling.Interfaces.Layer;
using XlsxHandling.Interfaces.Manager;
using XlsxHandling.Layer;

namespace XlsxHandling.Implementations.Manager
{
	public class BorderManager : IBorderManager
	{
		public BorderManager()
		{
			ActualBorders = new Borders();
			SetDefaultBorder();
		}

		public Borders ActualBorders { get; set; }

		public uint GetBorder(IXlsxBorder borderLayer)
		{
			XlsxBorder layer = borderLayer as XlsxBorder;
			if(layer == null) { throw new Exception("Unexpected Border Type"); }

			return GetBorder(layer.TopBorder, layer.RightBorder, layer.BottomBorder, layer.LeftBorder, layer.DiagonalBorder);
		}

		public uint GetDefaultBorder() => 0;

		private uint GetBorder(BorderType topBorder, BorderType rightBorder, BorderType bottomBorder, BorderType leftBorder, BorderType diagonalBorder)
		{
			BorderStyleValues topBorderValue = GetBorderStyle(topBorder);
			BorderStyleValues rightBorderValue = GetBorderStyle(rightBorder);
			BorderStyleValues bottomBorderValue = GetBorderStyle(bottomBorder);
			BorderStyleValues leftBorderValue = GetBorderStyle(leftBorder);
			BorderStyleValues diagonalBorderValue = GetBorderStyle(diagonalBorder);

			//Check font type exists and return index
			uint borderIndex = 0;
			foreach(var border in ActualBorders.ChildElements.OfType<Border>()) {
				if(border.TopBorder.Style == topBorderValue &&
					border.RightBorder.Style == rightBorderValue &&
					border.BottomBorder.Style == bottomBorderValue &&
					border.LeftBorder.Style == leftBorderValue &&
					border.DiagonalBorder.Style == diagonalBorderValue) {
					return borderIndex;
				}
				borderIndex++;
			}

			//if not exists create new font and append to actual fonts
			SetBorder(topBorderValue, rightBorderValue, bottomBorderValue, leftBorderValue, diagonalBorderValue);
			return borderIndex;
		}

		private BorderStyleValues GetBorderStyle(BorderType borderType)
		{
			BorderStyleValues borderStyleValue;

			switch(borderType) {
				case BorderType.None:
					borderStyleValue = BorderStyleValues.None;
					break;
				case BorderType.Dashed:
					borderStyleValue = BorderStyleValues.Dashed;
					break;
				case BorderType.Dotted:
					borderStyleValue = BorderStyleValues.Dotted;
					break;
				case BorderType.Thick:
					borderStyleValue = BorderStyleValues.Thick;
					break;
				case BorderType.Double:
					borderStyleValue = BorderStyleValues.Double;
					break;
				case BorderType.Hair:
					borderStyleValue = BorderStyleValues.Hair;
					break;
				case BorderType.MediumDashed:
					borderStyleValue = BorderStyleValues.MediumDashed;
					break;
				case BorderType.DashDot:
					borderStyleValue = BorderStyleValues.DashDot;
					break;
				case BorderType.MediumDashDot:
					borderStyleValue = BorderStyleValues.MediumDashDot;
					break;
				case BorderType.DashDotDot:
					borderStyleValue = BorderStyleValues.DashDotDot;
					break;
				case BorderType.MediumDashDotDot:
					borderStyleValue = BorderStyleValues.MediumDashDotDot;
					break;
				case BorderType.SlantDashDot:
					borderStyleValue = BorderStyleValues.SlantDashDot;
					break;
				default:
					throw new Exception("Unknown Border Type");
			}

			return borderStyleValue;
		}

		private void SetBorder(BorderStyleValues topBorderValue, BorderStyleValues rightBorderValue, BorderStyleValues bottomBorderValue, BorderStyleValues leftBorderValue, BorderStyleValues diagonalBorderValue)
		{
			Border border = new Border();
			border.TopBorder = new TopBorder() { Style = new EnumValue<BorderStyleValues>(topBorderValue) };
			border.RightBorder = new RightBorder() { Style = new EnumValue<BorderStyleValues>(rightBorderValue) };
			border.BottomBorder = new BottomBorder() { Style = new EnumValue<BorderStyleValues>(bottomBorderValue) };
			border.LeftBorder = new LeftBorder() { Style = new EnumValue<BorderStyleValues>(leftBorderValue) };
			border.DiagonalBorder = new DiagonalBorder() { Style = new EnumValue<BorderStyleValues>(diagonalBorderValue) };

			ActualBorders.Append(border);
			SetBordersCount();
		}

		private void SetBordersCount()
		{
			ActualBorders.Count = (uint)ActualBorders.ChildElements.Count;
		}

		private void SetDefaultBorder()
		{
			SetBorder(BorderStyleValues.None, BorderStyleValues.None, BorderStyleValues.None, BorderStyleValues.None, BorderStyleValues.None);
		}
	}
}
