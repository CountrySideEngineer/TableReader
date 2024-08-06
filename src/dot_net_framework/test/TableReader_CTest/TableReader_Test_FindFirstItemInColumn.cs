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
		[Description("FindFirstItemInColumn(string item, Range range)")]
		public void FindFirstItemInColumn_Test_001()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItemInColumn_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItemInColumn_001";
				Range tableRange = new Range()
				{
					StartRow = 1,
					StartColumn = 1,
					RowCount = 1,
					ColumnCount = 0
				};
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = "FirstItem_01_01";
				Range range = reader.FindFirstItemInColumn(item, tableRange);

				Assert.AreEqual(1, range.StartRow);
				Assert.AreEqual(1, range.StartColumn);
				Assert.AreEqual(1, range.RowCount);
				Assert.AreEqual(1, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("FindFirstItemInColumn(string item, Range range)")]
		public void FindFirstItemInColumn_Test_002()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItemInColumn_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItemInColumn_001";
				Range tableRange = new Range()
				{
					StartRow = 1,
					StartColumn = 1,
					RowCount = 2,
					ColumnCount = 1
				};
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = "FirstItem_01_01";
				Range range = reader.FindFirstItemInColumn(item, tableRange);

				Assert.AreEqual(1, range.StartRow);
				Assert.AreEqual(1, range.StartColumn);
				Assert.AreEqual(1, range.RowCount);
				Assert.AreEqual(1, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("FindFirstItemInColumn(string item, Range range)")]
		public void FindFirstItemInColumn_Test_003()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItemInColumn_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItemInColumn_001";
				Range tableRange = new Range()
				{
					StartRow = 2,
					StartColumn = 1,
					RowCount = 2,
					ColumnCount = 1
				};
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = "FirstItem_02_01";
				Range range = reader.FindFirstItemInColumn(item, tableRange);

				Assert.AreEqual(2, range.StartRow);
				Assert.AreEqual(1, range.StartColumn);
				Assert.AreEqual(1, range.RowCount);
				Assert.AreEqual(1, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("FindFirstItemInColumn(string item, Range range)")]
		public void FindFirstItemInColumn_Test_004()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItemInColumn_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItemInColumn_001";
				Range tableRange = new Range()
				{
					StartRow = 1,
					StartColumn = 2,
					RowCount = 8,
					ColumnCount = 0
				};
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = "FirstItem_01_02";
				Range range = reader.FindFirstItemInColumn(item, tableRange);

				Assert.AreEqual(1, range.StartRow);
				Assert.AreEqual(2, range.StartColumn);
				Assert.AreEqual(1, range.RowCount);
				Assert.AreEqual(1, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("FindFirstItemInColumn(string item, Range range)")]
		public void FindFirstItemInColumn_Test_005()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItemInColumn_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItemInColumn_001";
				Range tableRange = new Range()
				{
					StartRow = 1,
					StartColumn = 2,
					RowCount = 8,
					ColumnCount = 0
				};
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = "FirstItem_06_02";
				Range range = reader.FindFirstItemInColumn(item, tableRange);

				Assert.AreEqual(6, range.StartRow);
				Assert.AreEqual(2, range.StartColumn);
				Assert.AreEqual(1, range.RowCount);
				Assert.AreEqual(1, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("FindFirstItemInColumn(string item, Range range)")]
		public void FindFirstItemInColumn_Test_006()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItemInColumn_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItemInColumn_001";
				Range tableRange = new Range()
				{
					StartRow = 1,
					StartColumn = 2,
					RowCount = 8,
					ColumnCount = 0
				};
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = "FirstItem_08_02";
				Range range = reader.FindFirstItemInColumn(item, tableRange);

				Assert.AreEqual(8, range.StartRow);
				Assert.AreEqual(2, range.StartColumn);
				Assert.AreEqual(1, range.RowCount);
				Assert.AreEqual(1, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("FindFirstItemInColumn(string item, Range range)")]
		public void FindFirstItemInColumn_Test_007()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItemInColumn_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItemInColumn_001";
				string item = "FirstItem_09_02";
				Range tableRange = new Range()
				{
					StartRow = 1,
					StartColumn = 2,
					RowCount = 8,
					ColumnCount = 0
				};
				var reader = new ExcelTableReader(testDataStream, sheetName);
				var ex = Assert.ThrowsException<ArgumentException>(() =>
				{
					Range range = reader.FindFirstItemInColumn(item, tableRange);
				});
				Assert.AreEqual("No cell contains \"FirstItem_09_02\" in FindFirstItemInColumn_001.", ex.Message);
			}
		}

		[TestMethod]
		[Description("FindFirstItemInColumn(string item, Range range)")]
		public void FindFirstItemInColumn_Test_008()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItemInColumn_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItemInColumn_001";
				string item = "FirstItem_09_02";
				Range tableRange = new Range()
				{
					StartRow = 1,
					StartColumn = 2,
					RowCount = 8,
					ColumnCount = 0
				};
				var reader = new ExcelTableReader(testDataStream, sheetName);
				var ex = Assert.ThrowsException<ArgumentNullException>(() =>
				{
					Range range = reader.FindFirstItemInColumn(item, null);
				});
				Assert.AreEqual("range", ex.ParamName);
				Assert.AreEqual("Range to read has not been set.\r\nパラメーター名:range", ex.Message);
			}
		}

		[TestMethod]
		[Description("FindFirstItemInColumn(string item, Range range)")]
		public void FindFirstItemInColumn_Test_009()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItemInColumn_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItemInColumn_001";
				string item = string.Empty;
				Range tableRange = new Range()
				{
					StartRow = 1,
					StartColumn = 2,
					RowCount = 8,
					ColumnCount = 0
				};
				var reader = new ExcelTableReader(testDataStream, sheetName);
				var ex = Assert.ThrowsException<ArgumentException>(() =>
				{
					Range range = reader.FindFirstItemInColumn(item, tableRange);
				});
				Assert.AreEqual("Targe item to scan shoul have any value, not empty.", ex.Message);
			}
		}

		[TestMethod]
		[Description("FindFirstItemInColumn(string item, Range range)")]
		public void FindFirstItemInColumn_Test_010()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItemInColumn_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = string.Empty;
				string item = string.Empty;
				Range tableRange = new Range()
				{
					StartRow = 1,
					StartColumn = 2,
					RowCount = 8,
					ColumnCount = 0
				};
				var reader = new ExcelTableReader(testDataStream, sheetName);
				var ex = Assert.ThrowsException<InvalidDataException>(() =>
				{
					Range range = reader.FindFirstItemInColumn(item, tableRange);
				});
				Assert.AreEqual("Sheet Name to scan is invalid.", ex.Message);
			}
		}

		[TestMethod]
		[Description("FindFirstItemInColumn(string item, Range range)")]
		public void FindFirstItemInColumn_Test_011()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItemInColumn_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItemInColumn_001";
				string item = string.Empty;
				Range tableRange = new Range()
				{
					StartRow = 1,
					StartColumn = 2,
					RowCount = 8,
					ColumnCount = 0
				};
				var reader = new ExcelTableReader(null, sheetName);
				var ex = Assert.ThrowsException<NullReferenceException>(() =>
				{
					Range range = reader.FindFirstItemInColumn(item, tableRange);
				});
				Assert.AreEqual("Stream data to read has not been set.", ex.Message);
			}
		}
	}
}
