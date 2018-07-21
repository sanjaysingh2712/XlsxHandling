using System.Collections.Generic;
using XlsxHandling.Interfaces.Layer;

namespace XlsxHandling.Layer
{
	public class XlsxFile : IXlsxFile
	{
		public string PathToStoreAt { get; set; }
		public IList<IXlsxSheet> Sheets { get; set; }
	}
}
