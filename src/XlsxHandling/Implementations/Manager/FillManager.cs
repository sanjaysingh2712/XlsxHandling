using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxHandling.Enums;
using XlsxHandling.Interfaces.Layer;
using XlsxHandling.Interfaces.Manager;
using XlsxHandling.Layer;

namespace XlsxHandling.Implementations.Manager
{
	public class FillManager : IFillManager
	{
		public FillManager()
		{
			ActualFills = new Fills();
			SetDefaultFill();
		}

		public Fills ActualFills { get; set; }

		public uint GetFill(IXlsxFill fillLayer)
		{
			XlsxFill layer = fillLayer as XlsxFill;
			if(layer == null) { throw new Exception("Unexpected Fill Type"); }

			return GetFill(layer.PatternType, layer.BackgroundColorArgb);
		}

		public uint GetDefaultFill() => 0;

		private uint GetFill(PatternType patternType, string backgroundColor)
		{
			PatternValues patternValue = GetFillStyle(patternType);

			//Check font type exists and return index
			uint fillIndex = 0;
			foreach(var fill in ActualFills.ChildElements.OfType<Fill>()) {
				if(fill.PatternFill.PatternType == patternValue &&
					(
						string.IsNullOrEmpty(backgroundColor) ||
						fill.PatternFill.ForegroundColor == null ||
						fill.PatternFill.ForegroundColor.Rgb == backgroundColor
					)) {
					return fillIndex;
				}
				fillIndex++;
			}

			//if not exists create new fill and append to actual fills
			SetFill(patternValue, backgroundColor);
			return fillIndex;
		}

		private PatternValues GetFillStyle(PatternType patternType)
		{
			PatternValues patternValue;

			switch(patternType) {
				case PatternType.None:
					patternValue = PatternValues.None;
					break;
				case PatternType.Solid:
					patternValue = PatternValues.Solid;
					break;
				case PatternType.MediumGray:
					patternValue = PatternValues.MediumGray;
					break;
				case PatternType.DarkGray:
					patternValue = PatternValues.DarkGray;
					break;
				case PatternType.LightGray:
					patternValue = PatternValues.LightGray;
					break;
				case PatternType.DarkHorizontal:
					patternValue = PatternValues.DarkHorizontal;
					break;
				case PatternType.DarkVertical:
					patternValue = PatternValues.DarkVertical;
					break;
				case PatternType.DarkDown:
					patternValue = PatternValues.DarkDown;
					break;
				case PatternType.DarkUp:
					patternValue = PatternValues.DarkUp;
					break;
				case PatternType.DarkGrid:
					patternValue = PatternValues.DarkGrid;
					break;
				case PatternType.DarkTrellis:
					patternValue = PatternValues.DarkTrellis;
					break;
				case PatternType.LightHorizontal:
					patternValue = PatternValues.LightHorizontal;
					break;
				case PatternType.LightVertical:
					patternValue = PatternValues.LightVertical;
					break;
				case PatternType.LightDown:
					patternValue = PatternValues.LightDown;
					break;
				case PatternType.LightUp:
					patternValue = PatternValues.LightUp;
					break;
				case PatternType.LightGrid:
					patternValue = PatternValues.LightGrid;
					break;
				case PatternType.LightTrellis:
					patternValue = PatternValues.LightTrellis;
					break;
				case PatternType.Gray125:
					patternValue = PatternValues.Gray125;
					break;
				case PatternType.Gray0625:
					patternValue = PatternValues.Gray0625;
					break;
				default:
					throw new Exception("Unknown Fill Pattern Type");
			}

			return patternValue;
		}

		private void SetFill(PatternValues patternValue, string backgroundColor = null)
		{
			Fill fill = new Fill();
			PatternFill patternFill = new PatternFill();
			patternFill.PatternType = new EnumValue<PatternValues>(patternValue);
			if(!string.IsNullOrEmpty(backgroundColor)) {
				patternFill.ForegroundColor = new ForegroundColor { Rgb = backgroundColor };
			}
			fill.PatternFill = patternFill;

			ActualFills.Append(fill);
			SetFillsCount();
		}

		private void SetFillsCount()
		{
			ActualFills.Count = (uint)ActualFills.ChildElements.Count;
		}

		private void SetDefaultFill()
		{
			SetFill(PatternValues.None);
			//SetFill(PatternValues.Gray125);
		}
	}
}
