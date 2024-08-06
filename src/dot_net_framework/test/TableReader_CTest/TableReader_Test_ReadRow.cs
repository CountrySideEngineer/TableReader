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
		[Description("ReadRow(Range range)")]
		public void ReadRow_test_001()
		{
			var testDataPath = @"..\..\..\TestData\ReadRow_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "ReadRow_test_001";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				Range range = new Range()
				{
					StartRow = 1,
					StartColumn = 1,
				};
				IEnumerable<string> items = reader.ReadRow(range);
				Assert.AreEqual(1, items.Count());
				Assert.AreEqual("Item_001_001", items.ElementAt(0));
			}
		}

		[TestMethod]
		[Description("ReadRow(Range range)")]
		public void ReadRow_test_002()
		{
			var testDataPath = @"..\..\..\TestData\ReadRow_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "ReadRow_test_001";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				Range range = new Range()
				{
					StartRow = 2,
					StartColumn = 1,
				};
				IEnumerable<string> items = reader.ReadRow(range);
				Assert.AreEqual(2, items.Count());
				Assert.AreEqual("Item_002_001", items.ElementAt(0));
				Assert.AreEqual("Item_002_002", items.ElementAt(1));
			}
		}

		[TestMethod]
		[Description("ReadRow(Range range)")]
		public void ReadRow_test_003()
		{
			var testDataPath = @"..\..\..\TestData\ReadRow_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "ReadRow_test_001";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				Range range = new Range()
				{
					StartRow = 3,
					StartColumn = 1,
				};
				IEnumerable<string> items = reader.ReadRow(range);
				Assert.AreEqual(4, items.Count());
				Assert.AreEqual("Item_003_001", items.ElementAt(0));
				Assert.AreEqual("Item_003_002", items.ElementAt(1));
				Assert.AreEqual("", items.ElementAt(2));
				Assert.AreEqual("Item_003_004", items.ElementAt(3));
			}
		}

		[TestMethod]
		[Description("ReadRow(Range range)")]
		public void ReadRow_test_004()
		{
			var testDataPath = @"..\..\..\TestData\ReadRow_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "ReadRow_test_001";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				Range range = new Range()
				{
					StartRow = 4,
					StartColumn = 1,
				};
				IEnumerable<string> items = reader.ReadRow(range);
				Assert.AreEqual(5, items.Count());
				Assert.AreEqual("", items.ElementAt(0));
				Assert.AreEqual("Item_004_002", items.ElementAt(1));
				Assert.AreEqual("", items.ElementAt(2));
				Assert.AreEqual("", items.ElementAt(3));
				Assert.AreEqual("Item_004_005", items.ElementAt(4));
			}
		}

		[TestMethod]
		[Description("ReadRow(Range range)")]
		public void ReadRow_test_005()
		{
			var testDataPath = @"..\..\..\TestData\ReadRow_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "ReadRow_test_001";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				Range range = new Range()
				{
					StartRow = 4,
					StartColumn = 2,
				};
				IEnumerable<string> items = reader.ReadRow(range);
				Assert.AreEqual(4, items.Count());
				Assert.AreEqual("Item_004_002", items.ElementAt(0));
				Assert.AreEqual("", items.ElementAt(1));
				Assert.AreEqual("", items.ElementAt(2));
				Assert.AreEqual("Item_004_005", items.ElementAt(3));
			}
		}
	}
}
