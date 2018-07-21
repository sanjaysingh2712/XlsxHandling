using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XlsxHandling.Interfaces.Manager;

namespace XlsxHandling.Implementations.Manager
{
	public class SharedStringManager : ISharedStringManager
	{
		public SharedStringTablePart SstPart { get; set; }

		public string GetIdByValue(string value)
		{
			if(SstPart.SharedStringTable == null) {
				SstPart.SharedStringTable = new SharedStringTable();
			}

			int i = 0;
			foreach(SharedStringItem ssi in SstPart.SharedStringTable.Elements<SharedStringItem>()) {
				if(ssi.InnerText == value) { return i.ToString(); }
				i++;
			}

			// The text does not exist in the part. Create the SharedStringItem and return its index.
			SstPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(value)));
			SstPart.SharedStringTable.Save();

			return i.ToString();
		}
	}
}
