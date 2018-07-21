using System;
using Moq;
using NUnit.Framework;
using XlsxHandling.Implementations;
using XlsxHandling.Implementations.Manager;
using XlsxHandling.Interfaces;
using XlsxHandling.Interfaces.Layer;
using XlsxHandling.Layer;

namespace XlsxHandling.Test
{
	[TestFixture]
	public class XlsxCreatorTests
	{
		[Test]
		public void Create_XlsxCreated_IsTrue()
		{
			//Arrange
			IXlsxCreator creator = new XlsxCreator(new WorkbookCreator(new StylesheetManager(new FontManager(), new BorderManager(), new FillManager(), new NumberingFormatManager(), new CellFormatManager()), new SharedStringManager()));
			string path = @"C:\Users\sanjay\source\repos\Libraries\XlsxHandling\src-test\XlsxHandling.Test\output.xlsx";
			IXlsxFile file = new XlsxFile() { PathToStoreAt = path };

			//Actual
			bool created = creator.Create(file);

			//Assert
			Assert.IsTrue(created);
		}

		[Test]
		public void Create_NoFileProvided_ArgumentException()
		{
			//Arrange
			XlsxCreator creator = new XlsxCreator(new WorkbookCreator(new StylesheetManager(new FontManager(), new BorderManager(), new FillManager(), new NumberingFormatManager(), new CellFormatManager()), new SharedStringManager()));
			Mock<IXlsxFile> fileMock = new Mock<IXlsxFile>();
			Type expectedException = typeof(ArgumentException);

			//Actual
			try {
				creator.Create(fileMock.Object);
			} catch(Exception ex) {
				Assert.AreEqual(expectedException, ex.GetType(), "Expected an ArgumentException");
			}
		}
	}
}
