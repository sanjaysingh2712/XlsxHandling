using DocumentFormat.OpenXml.Packaging;

namespace XlsxHandling.Interfaces.Manager
{
	public interface ISharedStringManager
	{
		SharedStringTablePart SstPart { get; set; }
		string GetIdByValue(string value);
	}
}
