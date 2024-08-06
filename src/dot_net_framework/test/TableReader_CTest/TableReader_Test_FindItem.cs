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
		[Description("FindItem(string item)")]
		public void FindItem_test_001()
		{
			var testDataPath = @"..\..\..\TestData\FindItem_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindItem_test_001";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = "Item_001";
				IEnumerable<Range> ranges = reader.FindItem(item);
				Assert.AreEqual(1, ranges.Count());
				Assert.AreEqual(1, ranges.ElementAt(0).StartRow);
				Assert.AreEqual(1, ranges.ElementAt(0).StartColumn);
				Assert.AreEqual(1, ranges.ElementAt(0).RowCount);
				Assert.AreEqual(1, ranges.ElementAt(0).ColumnCount);
			}
		}

		[TestMethod]
		[Description("FindItem(string item)")]
		public void FindItem_test_002()
		{
			var testDataPath = @"..\..\..\TestData\FindItem_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindItem_test_002";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = "Item_001";
				IEnumerable<Range> ranges = reader.FindItem(item);
				Assert.AreEqual(2, ranges.Count());
				Assert.AreEqual(1, ranges.ElementAt(0).StartRow);
				Assert.AreEqual(1, ranges.ElementAt(0).StartColumn);
				Assert.AreEqual(1, ranges.ElementAt(0).RowCount);
				Assert.AreEqual(1, ranges.ElementAt(0).ColumnCount);
				Assert.AreEqual(1, ranges.ElementAt(1).StartRow);
				Assert.AreEqual(2, ranges.ElementAt(1).StartColumn);
				Assert.AreEqual(1, ranges.ElementAt(1).RowCount);
				Assert.AreEqual(1, ranges.ElementAt(1).ColumnCount);
			}
		}

		[TestMethod]
		[Description("FindItem(string item)")]
		public void FindItem_test_003()
		{
			var testDataPath = @"..\..\..\TestData\FindItem_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindItem_test_003";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = "Item_001";
				IEnumerable<Range> ranges = reader.FindItem(item);
				Assert.AreEqual(2, ranges.Count());
				Assert.AreEqual(1, ranges.ElementAt(0).StartRow);
				Assert.AreEqual(1, ranges.ElementAt(0).StartColumn);
				Assert.AreEqual(1, ranges.ElementAt(0).RowCount);
				Assert.AreEqual(1, ranges.ElementAt(0).ColumnCount);
				Assert.AreEqual(2, ranges.ElementAt(1).StartRow);
				Assert.AreEqual(2, ranges.ElementAt(1).StartColumn);
				Assert.AreEqual(1, ranges.ElementAt(1).RowCount);
				Assert.AreEqual(1, ranges.ElementAt(1).ColumnCount);
			}
		}

		[TestMethod]
		[Description("FindItem(string item)")]
		public void FindItem_test_004()
		{
			var testDataPath = @"..\..\..\TestData\FindItem_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindItem_test_004";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = "Item_002";
				IEnumerable<Range> ranges = reader.FindItem(item);
				Assert.AreEqual(1, ranges.Count());
				Assert.AreEqual(2, ranges.ElementAt(0).StartRow);
				Assert.AreEqual(1, ranges.ElementAt(0).StartColumn);
				Assert.AreEqual(1, ranges.ElementAt(0).RowCount);
				Assert.AreEqual(1, ranges.ElementAt(0).ColumnCount);
			}
		}

		[TestMethod]
		[Description("FindItem(string item)")]
		public void FindItem_test_005()
		{
			var testDataPath = @"..\..\..\TestData\FindItem_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindItem_test_004";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = "Item_003";
				IEnumerable<Range> ranges = reader.FindItem(item);
				Assert.AreEqual(3, ranges.Count());
				Assert.AreEqual(1, ranges.ElementAt(0).StartRow);
				Assert.AreEqual(2, ranges.ElementAt(0).StartColumn);
				Assert.AreEqual(1, ranges.ElementAt(0).RowCount);
				Assert.AreEqual(1, ranges.ElementAt(0).ColumnCount);
				Assert.AreEqual(3, ranges.ElementAt(1).StartRow);
				Assert.AreEqual(1, ranges.ElementAt(1).StartColumn);
				Assert.AreEqual(1, ranges.ElementAt(1).RowCount);
				Assert.AreEqual(1, ranges.ElementAt(1).ColumnCount);
				Assert.AreEqual(3, ranges.ElementAt(2).StartRow);
				Assert.AreEqual(3, ranges.ElementAt(2).StartColumn);
				Assert.AreEqual(1, ranges.ElementAt(2).RowCount);
				Assert.AreEqual(1, ranges.ElementAt(2).ColumnCount);
			}
		}

		[TestMethod]
		[Description("FindItem(string item)")]
		public void FindItem_test_006()
		{
			var testDataPath = @"..\..\..\TestData\FindItem_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindItem_test_001";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = "Item_NoData";
				var ex = Assert.ThrowsException<ArgumentException>(() =>
				{
					IEnumerable<Range> ranges = reader.FindItem(item);
				});
				Assert.AreEqual("No cell contains \"Item_NoData\" in FindItem_test_001.", ex.Message);
			}
		}

		[TestMethod]
		[Description("FindItem(string item)")]
		public void FindItem_test_007()
		{
			var testDataPath = @"..\..\..\TestData\FindItem_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindItem_test_001";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = string.Empty;
				var ex = Assert.ThrowsException<ArgumentException>(() =>
				{
					IEnumerable<Range> ranges = reader.FindItem(item);
				});
				Assert.AreEqual("Target string must not be empty.", ex.Message);
			}
		}
	}
}
