using System.Collections.Generic;

namespace XlsxHandling.Interfaces.Layer
{
    public interface IXlsxFile
    {
		string PathToStoreAt { get; set; }
		IList<IXlsxSheet> Sheets { get; set; }
    }
}