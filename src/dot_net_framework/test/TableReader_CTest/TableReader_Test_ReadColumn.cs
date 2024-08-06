using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using TableReader.Excel;
using TableReader.TableData;

namespace TableReader_CTest
{
	public partial class TableReader_Test
	{
		[TestMethod]
		[Description("ReadColumn(Range range)")]
		public void ReadColumn_test_001()
		{
			var testDataPath = @"..\..\..\TestData\ReadColumn_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "ReadColum_test_001";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				Range range = new Range()
				{
					StartRow = 1,
					StartColumn = 1,
				};
				IEnumerable<string> items = reader.ReadColumn(range);
				Assert.AreEqual(1, items.Count());
				Assert.AreEqual("Item_001_001", items.ElementAt(0));
			}
		}

		[TestMethod]
		[Description("ReadColumn(Range range)")]
		public void ReadColumn_test_002()
		{
			var testDataPath = @"..\..\..\TestData\ReadColumn_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "ReadColum_test_001";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				Range range = new Range()
				{
					StartRow = 1,
					StartColumn = 2,
				};
				IEnumerable<string> items = reader.ReadColumn(range);
				Assert.AreEqual(2, items.Count());
				Assert.AreEqual("Item_001_002", items.ElementAt(0));
				Assert.AreEqual("Item_002_002", items.ElementAt(1));
			}
		}

		[TestMethod]
		[Description("ReadColumn(Range range)")]
		public void ReadColumn_test_003()
		{
			var testDataPath = @"..\..\..\TestData\ReadColumn_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "ReadColum_test_001";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				Range range = new Range()
				{
					StartRow = 1,
					StartColumn = 3,
				};
				IEnumerable<string> items = reader.ReadColumn(range);
				Assert.AreEqual(4, items.Count());
				Assert.AreEqual("Item_001_003", items.ElementAt(0));
				Assert.AreEqual("Item_002_003", items.ElementAt(1));
				Assert.AreEqual("", items.ElementAt(2));
				Assert.AreEqual("Item_004_003", items.ElementAt(3));
			}
		}

		[TestMethod]
		[Description("ReadColumn(Range range)")]
		public void ReadColumn_test_004()
		{
			var testDataPath = @"..\..\..\TestData\ReadColumn_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "ReadColum_test_001";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				Range range = new Range()
				{
					StartRow = 1,
					StartColumn = 4,
				};
				IEnumerable<string> items = reader.ReadColumn(range);
				Assert.AreEqual(5, items.Count());
				Assert.AreEqual("", items.ElementAt(0));
				Assert.AreEqual("Item_002_004", items.ElementAt(1));
				Assert.AreEqual("", items.ElementAt(2));
				Assert.AreEqual("", items.ElementAt(3));
				Assert.AreEqual("Item_005_004", items.ElementAt(4));
			}
		}

		[TestMethod]
		[Description("ReadColumn(Range range)")]
		public void ReadColumn_test_005()
		{
			var testDataPath = @"..\..\..\TestData\ReadColumn_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "ReadColum_test_001";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				Range range = new Range()
				{
					StartRow = 2,
					StartColumn = 3,
				};
				IEnumerable<string> items = reader.ReadColumn(range);
				Assert.AreEqual(3, items.Count());
				Assert.AreEqual("Item_002_003", items.ElementAt(0));
				Assert.AreEqual("", items.ElementAt(1));
				Assert.AreEqual("Item_004_003", items.ElementAt(2));
			}
		}
	}
}
