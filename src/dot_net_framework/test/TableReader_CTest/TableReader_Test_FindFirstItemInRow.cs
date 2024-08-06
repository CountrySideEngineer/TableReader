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
		[Description("FindFirstItemInRow(string item, Range range)")]
		public void FindFirstItemInRow_Test_001()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItemInRow_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItemInRow_001";
				Range tableRange = new Range()
				{
					StartRow = 1,
					StartColumn = 1,
					RowCount = 0,
					ColumnCount = 1
				};
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = "FirstItem_01_01";
				Range range = reader.FindFirstItemInRow(item, tableRange);

				Assert.AreEqual(1, range.StartRow);
				Assert.AreEqual(1, range.StartColumn);
				Assert.AreEqual(1, range.RowCount);
				Assert.AreEqual(1, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("FindFirstItemInRow(string item, Range range)")]
		public void FindFirstItemInRow_Test_002()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItemInRow_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItemInRow_001";
				Range tableRange = new Range()
				{
					StartRow = 1,
					StartColumn = 1,
					RowCount = 0,
					ColumnCount = 2
				};
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = "FirstItem_01_01";
				Range range = reader.FindFirstItemInRow(item, tableRange);

				Assert.AreEqual(1, range.StartRow);
				Assert.AreEqual(1, range.StartColumn);
				Assert.AreEqual(1, range.RowCount);
				Assert.AreEqual(1, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("FindFirstItemInRow(string item, Range range)")]
		public void FindFirstItemInRow_Test_003()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItemInRow_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItemInRow_001";
				Range tableRange = new Range()
				{
					StartRow = 1,
					StartColumn = 2,
					RowCount = 1,
					ColumnCount = 2
				};
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = "FirstItem_01_02";
				Range range = reader.FindFirstItemInRow(item, tableRange);

				Assert.AreEqual(1, range.StartRow);
				Assert.AreEqual(2, range.StartColumn);
				Assert.AreEqual(1, range.RowCount);
				Assert.AreEqual(1, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("FindFirstItemInRow(string item, Range range)")]
		public void FindFirstItemInRow_Test_004()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItemInRow_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItemInRow_001";
				Range tableRange = new Range()
				{
					StartRow = 2,
					StartColumn = 1,
					RowCount = 1,
					ColumnCount = 8
				};
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = "FirstItem_02_01";
				Range range = reader.FindFirstItemInRow(item, tableRange);

				Assert.AreEqual(2, range.StartRow);
				Assert.AreEqual(1, range.StartColumn);
				Assert.AreEqual(1, range.RowCount);
				Assert.AreEqual(1, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("FindFirstItemInRow(string item, Range range)")]
		public void FindFirstItemInRow_Test_005()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItemInRow_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItemInRow_001";
				Range tableRange = new Range()
				{
					StartRow = 2,
					StartColumn = 1,
					RowCount = 1,
					ColumnCount = 8
				};
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = "FirstItem_02_05";
				Range range = reader.FindFirstItemInRow(item, tableRange);

				Assert.AreEqual(2, range.StartRow);
				Assert.AreEqual(5, range.StartColumn);
				Assert.AreEqual(1, range.RowCount);
				Assert.AreEqual(1, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("FindFirstItemInRow(string item, Range range)")]
		public void FindFirstItemInRow_Test_006()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItemInRow_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItemInRow_001";
				Range tableRange = new Range()
				{
					StartRow = 2,
					StartColumn = 1,
					RowCount = 1,
					ColumnCount = 8
				};
				var reader = new ExcelTableReader(testDataStream, sheetName);

				string item = "FirstItem_02_07";
				Range range = reader.FindFirstItemInRow(item, tableRange);

				Assert.AreEqual(2, range.StartRow);
				Assert.AreEqual(7, range.StartColumn);
				Assert.AreEqual(1, range.RowCount);
				Assert.AreEqual(1, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("FindFirstItemInRow(string item, Range range)")]
		public void FindFirstItemInRow_Test_007()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItemInRow_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItemInRow_001";
				Range tableRange = new Range()
				{
					StartRow = 2,
					StartColumn = 1,
					RowCount = 1,
					ColumnCount = 8
				};
				string item = "FirstItem_02_09";
				var reader = new ExcelTableReader(testDataStream, sheetName);
				var ex = Assert.ThrowsException<ArgumentException>(() =>
				{
					Range range = reader.FindFirstItemInRow(item, tableRange);
				});
				Assert.AreEqual("No cell contains \"FirstItem_02_09\" in FindFirstItemInRow_001.", ex.Message);
			}
		}

		[TestMethod]
		[Description("FindFirstItemInRow(string item, Range range)")]
		public void FindFirstItemInRow_Test_008()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItemInRow_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItemInRow_001";
				Range tableRange = new Range()
				{
					StartRow = 2,
					StartColumn = 1,
					RowCount = 1,
					ColumnCount = 8
				};
				string item = "FirstItem_02_09";
				var reader = new ExcelTableReader(testDataStream, sheetName);
				var ex = Assert.ThrowsException<ArgumentNullException>(() =>
				{
					Range range = reader.FindFirstItemInRow(item, null);
				});
				Assert.AreEqual("range", ex.ParamName);
				Assert.AreEqual("Range to read has not been set.\r\nパラメーター名:range", ex.Message);
			}
		}

		[TestMethod]
		[Description("FindFirstItemInRow(string item, Range range)")]
		public void FindFirstItemInRow_Test_009()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItemInRow_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItemInRow_001";
				Range tableRange = new Range()
				{
					StartRow = 2,
					StartColumn = 1,
					RowCount = 1,
					ColumnCount = 8
				};
				string item = string.Empty;
				var reader = new ExcelTableReader(testDataStream, sheetName);
				var ex = Assert.ThrowsException<ArgumentException>(() =>
				{
					Range range = reader.FindFirstItemInRow(item, tableRange);
				});
				Assert.AreEqual("Targe item to scan shoul have any value, not empty.", ex.Message);
			}
		}

		[TestMethod]
		[Description("FindFirstItemInRow(string item, Range range)")]
		public void FindFirstItemInRow_Test_010()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItemInRow_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = string.Empty;
				Range tableRange = new Range()
				{
					StartRow = 2,
					StartColumn = 1,
					RowCount = 1,
					ColumnCount = 8
				};
				string item = string.Empty;
				var reader = new ExcelTableReader(testDataStream, sheetName);
				var ex = Assert.ThrowsException<InvalidDataException>(() =>
				{
					Range range = reader.FindFirstItemInRow(item, tableRange);
				});
				Assert.AreEqual("Sheet Name to scan is invalid.", ex.Message);
			}
		}

		[TestMethod]
		[Description("FindFirstItemInRow(string item, Range range)")]
		public void FindFirstItemInRow_Test_011()
		{
			var testDataPath = @"..\..\..\TestData\FindFirstItemInRow_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "FindFirstItemInRow_001";
				Range tableRange = new Range()
				{
					StartRow = 2,
					StartColumn = 1,
					RowCount = 1,
					ColumnCount = 8
				};
				string item = "FirstItem_02_09";
				var reader = new ExcelTableReader(null, sheetName);
				var ex = Assert.ThrowsException<NullReferenceException>(() =>
				{
					Range range = reader.FindFirstItemInRow(item, tableRange);
				});
				Assert.AreEqual("Stream data to read has not been set.", ex.Message);
			}
		}
	}
}
