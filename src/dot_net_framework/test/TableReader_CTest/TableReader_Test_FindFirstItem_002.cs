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
		[Description("FindFirstItem(string item, Range range)")]
		public void FindFirstItem_Test_002_001()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItem_Test_001_001.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItem_test_002";
				Range tableRange = new Range()
				{
					StartRow = 1,
					StartColumn = 1,
					RowCount = 1,
					ColumnCount = 1
				};
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = "FirstItem_01_01";
				Range range = reader.FindFirstItem(item, tableRange);

				Assert.AreEqual(1, range.StartRow);
				Assert.AreEqual(1, range.StartColumn);
				Assert.AreEqual(1, range.RowCount);
				Assert.AreEqual(1, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("FindFirstItem(string item, Range range)")]
		public void FindFirstItem_Test_002_002()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItem_Test_001_001.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItem_test_002";
				Range tableRange = new Range()
				{
					StartRow = 1,
					StartColumn = 1,
					RowCount = 2,
					ColumnCount = 2
				};
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = "FirstItem_01_01";
				Range range = reader.FindFirstItem(item, tableRange);

				Assert.AreEqual(1, range.StartRow);
				Assert.AreEqual(1, range.StartColumn);
				Assert.AreEqual(1, range.RowCount);
				Assert.AreEqual(1, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("FindFirstItem(string item, Range range)")]
		public void FindFirstItem_Test_002_003()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItem_Test_001_001.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItem_test_002";
				Range tableRange = new Range()
				{
					StartRow = 1,
					StartColumn = 1,
					RowCount = 2,
					ColumnCount = 2
				};
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = "FirstItem_02_01";
				Range range = reader.FindFirstItem(item, tableRange);

				Assert.AreEqual(2, range.StartRow);
				Assert.AreEqual(1, range.StartColumn);
				Assert.AreEqual(1, range.RowCount);
				Assert.AreEqual(1, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("FindFirstItem(string item, Range range)")]
		public void FindFirstItem_Test_002_004()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItem_Test_001_001.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItem_test_002";
				Range tableRange = new Range()
				{
					StartRow = 1,
					StartColumn = 1,
					RowCount = 2,
					ColumnCount = 2
				};
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = "FirstItem_01_02";
				Range range = reader.FindFirstItem(item, tableRange);

				Assert.AreEqual(1, range.StartRow);
				Assert.AreEqual(2, range.StartColumn);
				Assert.AreEqual(1, range.RowCount);
				Assert.AreEqual(1, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("FindFirstItem(string item, Range range)")]
		public void FindFirstItem_Test_002_005()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItem_Test_001_001.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItem_test_002";
				Range tableRange = new Range()
				{
					StartRow = 1,
					StartColumn = 1,
					RowCount = 2,
					ColumnCount = 2
				};
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = "FirstItem_02_02";
				Range range = reader.FindFirstItem(item, tableRange);

				Assert.AreEqual(2, range.StartRow);
				Assert.AreEqual(2, range.StartColumn);
				Assert.AreEqual(1, range.RowCount);
				Assert.AreEqual(1, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("FindFirstItem(string item, Range range)")]
		public void FindFirstItem_Test_002_006()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItem_Test_001_001.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItem_test_002";
				Range tableRange = new Range()
				{
					StartRow = 3,
					StartColumn = 2,
					RowCount = 4,
					ColumnCount = 1
				};
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = "FirstItem_02_02";
				Range range = reader.FindFirstItem(item, tableRange);

				Assert.AreEqual(3, range.StartRow);
				Assert.AreEqual(2, range.StartColumn);
				Assert.AreEqual(1, range.RowCount);
				Assert.AreEqual(1, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("FindFirstItem(string item, Range range)")]
		public void FindFirstItem_Test_002_007()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItem_Test_001_001.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItem_test_002";
				Range tableRange = new Range()
				{
					StartRow = 1,
					StartColumn = 1,
					RowCount = 9,
					ColumnCount = 3,
				};
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = "FirstItem_08_03";
				Range range = reader.FindFirstItem(item, tableRange);

				Assert.AreEqual(8, range.StartRow);
				Assert.AreEqual(3, range.StartColumn);
				Assert.AreEqual(1, range.RowCount);
				Assert.AreEqual(1, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("FindFirstItem(string item, Range range)")]
		public void FindFirstItem_Test_002_008()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItem_test_001_001.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItem_test_002";
				var ex = Assert.ThrowsException<ArgumentException>(() =>
				{
					string item = "FirstItem_01_01";
					Range tableRange = new Range()
					{
						StartRow = 2,
						StartColumn = 1,
						RowCount = 1,
						ColumnCount = 1
					};
					var reader = new ExcelTableReader(testDataStream, sheetName);
					Range range = reader.FindFirstItem(item, tableRange);
				});
				Assert.AreEqual("No cell contains \"FirstItem_01_01\" in FindFirstItem_test_002.", ex.Message);
			}
		}

		[TestMethod]
		[Description("FindFirstItem(string item, Range range)")]
		public void FindFirstItem_Test_002_009()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItem_test_001_001.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItem_test_002";
				var ex = Assert.ThrowsException<ArgumentException>(() =>
				{
					string item = "FirstItem_01_01";
					Range tableRange = new Range()
					{
						StartRow = 2,
						StartColumn = 1,
						RowCount = 1,
						ColumnCount = 1
					};
					var reader = new ExcelTableReader(testDataStream, sheetName);
					Range range = reader.FindFirstItem(item, tableRange);
				});
				Assert.AreEqual("No cell contains \"FirstItem_01_01\" in FindFirstItem_test_002.", ex.Message);
			}
		}
	}
}
