using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using TableReader.Excel;
using TableReader.TableData;

namespace TableReader_CTest
{
	public partial class TableReader_Test
	{
		[TestMethod]
		[Description("FindFirstItem(string item)")]
		public void FindFirstItem_test_001_001_001()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItem_test_001_001.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItem_test_001";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = "FirstItem_01_01";
				Range range = reader.FindFirstItem(item);

				Assert.AreEqual(1, range.StartRow);
				Assert.AreEqual(1, range.StartColumn);
				Assert.AreEqual(1, range.RowCount);
				Assert.AreEqual(1, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("FindFirstItem(string item)")]
		public void FindFirstItem_test_001_001_002()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItem_test_001_001.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItem_test_001";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = "FirstItem_02_01";
				Range range = reader.FindFirstItem(item);

				Assert.AreEqual(2, range.StartRow);
				Assert.AreEqual(1, range.StartColumn);
				Assert.AreEqual(1, range.RowCount);
				Assert.AreEqual(1, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("FindFirstItem(string item)")]
		public void FindFirstItem_test_001_001_003()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItem_test_001_001.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItem_test_001";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = "FirstItem_01_02";
				Range range = reader.FindFirstItem(item);

				Assert.AreEqual(1, range.StartRow);
				Assert.AreEqual(2, range.StartColumn);
				Assert.AreEqual(1, range.RowCount);
				Assert.AreEqual(1, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("FindFirstItem(string item)")]
		public void FindFirstItem_test_001_001_004()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItem_test_001_001.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItem_test_001";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = "FirstItem_02_02";
				Range range = reader.FindFirstItem(item);

				Assert.AreEqual(2, range.StartRow);
				Assert.AreEqual(2, range.StartColumn);
				Assert.AreEqual(1, range.RowCount);
				Assert.AreEqual(1, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("FindFirstItem(string item)")]
		public void FindFirstItem_test_001_001_005()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItem_test_001_001.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItem_test_001";
				var ex = Assert.ThrowsException<ArgumentException>(() =>
				{
					string item = "FirstItem_NoData";
					var reader = new ExcelTableReader(testDataStream, sheetName);
					Range range = reader.FindFirstItem(item);
				});
				Assert.AreEqual("No cell contains \"FirstItem_NoData\" in FindFirstItem_test_001.", ex.Message);
			}
		}

		[TestMethod]
		[Description("FindFirstItem(string item)")]
		public void FindFirstItem_test_001_001_006()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItem_test_001_001.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItem_test_001";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = "FirstItem_02_02";
				Range range = reader.FindFirstItem(item);

				Assert.AreEqual(2, range.StartRow);
				Assert.AreEqual(2, range.StartColumn);
				Assert.AreEqual(1, range.RowCount);
				Assert.AreEqual(1, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("FindFirstItem(string item)")]
		public void FindFirstItem_test_001_001_007()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItem_test_001_001.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItem_test_001";
				var ex = Assert.ThrowsException<ArgumentException>(() =>
				{
					string item = string.Empty;
					var reader = new ExcelTableReader(testDataStream, sheetName);
					Range range = reader.FindFirstItem(item);
				});
				Assert.AreEqual("The string to be searched must have value set.", ex.Message);
			}
		}

		[TestMethod]
		[Description("FindFirstItem(string item)")]
		public void FindFirstItem_test_001_001_008()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItem_test_001_001.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItem_test_001";
				var ex = Assert.ThrowsException<ArgumentException>(() =>
				{
					string item = string.Empty;
					var reader = new ExcelTableReader(testDataStream, sheetName);
					Range range = reader.FindFirstItem(item);
				});
				Assert.AreEqual("The string to be searched must have value set.", ex.Message);
			}
		}

		[TestMethod]
		[Description("FindFirstItem(string item)")]
		public void FindFirstItem_test_001_001_009()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItem_test_001_001.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItem_test_001";
				var ex = Assert.ThrowsException<ArgumentException>(() =>
				{
					string item = " ";
					var reader = new ExcelTableReader(testDataStream, sheetName);
					Range range = reader.FindFirstItem(item);
				});
				Assert.AreEqual("No cell contains \" \" in FindFirstItem_test_001.", ex.Message);
			}
		}

		[TestMethod]
		[Description("FindFirstItem(string item)")]
		public void FindFirstItem_test_001_001_010()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItem_test_001_001.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItem_test_001";
				var ex = Assert.ThrowsException<ArgumentException>(() =>
				{
					string item = "\t";
					var reader = new ExcelTableReader(testDataStream, sheetName);
					Range range = reader.FindFirstItem(item);
				});
				Assert.AreEqual("No cell contains \"\t\" in FindFirstItem_test_001.", ex.Message);
			}
		}

		[TestMethod]
		[Description("FindFirstItem(string item)")]
		public void FindFirstItem_test_001_001_011()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItem_test_001_001.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItem_test_001";
				var ex = Assert.ThrowsException<ArgumentException>(() =>
				{
					string item = "　";
					var reader = new ExcelTableReader(testDataStream, sheetName);
					Range range = reader.FindFirstItem(item);
				});
				Assert.AreEqual("No cell contains \"　\" in FindFirstItem_test_001.", ex.Message);
			}
		}
	}
}
