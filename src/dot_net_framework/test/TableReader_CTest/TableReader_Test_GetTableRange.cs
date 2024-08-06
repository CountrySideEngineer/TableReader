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
		[Description("GetTableRange(ref Range range)")]
		public void GetTableRange_test_001()
		{
			var testDataPath = @"..\..\..\TestData\GetTableRange_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetTableRange_test_001";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				Range range = new Range()
				{
					StartRow = -1,
					RowCount = -1,
					StartColumn = -1,
					ColumnCount = -1,

				};
				reader.GetTableRange(ref range);
				Assert.AreEqual(1, range.StartRow);
				Assert.AreEqual(1, range.RowCount);
				Assert.AreEqual(1, range.StartColumn);
				Assert.AreEqual(1, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("GetTableRange(ref Range range)")]
		public void GetTableRange_test_002()
		{
			var testDataPath = @"..\..\..\TestData\GetTableRange_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetTableRange_test_002";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				Range range = new Range()
				{
					StartRow = -1,
					RowCount = -1,
					StartColumn = -1,
					ColumnCount = -1,

				};
				reader.GetTableRange(ref range);
				Assert.AreEqual(1, range.StartRow);
				Assert.AreEqual(2, range.RowCount);
				Assert.AreEqual(1, range.StartColumn);
				Assert.AreEqual(1, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("GetTableRange(ref Range range)")]
		public void GetTableRange_test_003()
		{
			var testDataPath = @"..\..\..\TestData\GetTableRange_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetTableRange_test_003";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				Range range = new Range()
				{
					StartRow = -1,
					RowCount = -1,
					StartColumn = -1,
					ColumnCount = -1,

				};
				reader.GetTableRange(ref range);
				Assert.AreEqual(2, range.StartRow);
				Assert.AreEqual(2, range.RowCount);
				Assert.AreEqual(1, range.StartColumn);
				Assert.AreEqual(1, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("GetTableRange(ref Range range)")]
		public void GetTableRange_test_004()
		{
			var testDataPath = @"..\..\..\TestData\GetTableRange_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetTableRange_test_004";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				Range range = new Range()
				{
					StartRow = -1,
					RowCount = -1,
					StartColumn = -1,
					ColumnCount = -1,

				};
				reader.GetTableRange(ref range);
				Assert.AreEqual(1, range.StartRow);
				Assert.AreEqual(1, range.RowCount);
				Assert.AreEqual(2, range.StartColumn);
				Assert.AreEqual(2, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("GetTableRange(ref Range range)")]
		public void GetTableRange_test_005()
		{
			var testDataPath = @"..\..\..\TestData\GetTableRange_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetTableRange_test_005";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				Range range = new Range()
				{
					StartRow = -1,
					RowCount = -1,
					StartColumn = -1,
					ColumnCount = -1,

				};
				reader.GetTableRange(ref range);
				Assert.AreEqual(2, range.StartRow);
				Assert.AreEqual(2, range.RowCount);
				Assert.AreEqual(2, range.StartColumn);
				Assert.AreEqual(1, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("GetTableRange(ref Range range)")]
		public void GetTableRange_test_006()
		{
			var testDataPath = @"..\..\..\TestData\GetTableRange_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetTableRange_test_006";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				Range range = new Range()
				{
					StartRow = -1,
					RowCount = -1,
					StartColumn = -1,
					ColumnCount = -1,

				};
				reader.GetTableRange(ref range);
				Assert.AreEqual(2, range.StartRow);
				Assert.AreEqual(1, range.RowCount);
				Assert.AreEqual(2, range.StartColumn);
				Assert.AreEqual(2, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("GetTableRange(ref Range range)")]
		public void GetTableRange_test_007()
		{
			var testDataPath = @"..\..\..\TestData\GetTableRange_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetTableRange_test_007";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				Range range = new Range()
				{
					StartRow = -1,
					RowCount = -1,
					StartColumn = -1,
					ColumnCount = -1,

				};
				reader.GetTableRange(ref range);
				Assert.AreEqual(2, range.StartRow);
				Assert.AreEqual(4, range.RowCount);
				Assert.AreEqual(2, range.StartColumn);
				Assert.AreEqual(3, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("GetTableRange(ref Range range)")]
		public void GetTableRange_test_008()
		{
			var testDataPath = @"..\..\..\TestData\GetTableRange_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetTableRange_test_008";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				Range range = new Range()
				{
					StartRow = -1,
					RowCount = -1,
					StartColumn = -1,
					ColumnCount = -1,

				};
				reader.GetTableRange(ref range);
				Assert.AreEqual(2, range.StartRow);
				Assert.AreEqual(6, range.RowCount);
				Assert.AreEqual(3, range.StartColumn);
				Assert.AreEqual(4, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("GetTableRange(ref Range range)")]
		public void GetTableRange_test_009()
		{
			var testDataPath = @"..\..\..\TestData\GetTableRange_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetTableRange_test_009";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				Range range = new Range()
				{
					StartRow = -1,
					RowCount = -1,
					StartColumn = -1,
					ColumnCount = -1,

				};
				reader.GetTableRange(ref range);
				Assert.AreEqual(0, range.StartRow);
				Assert.AreEqual(0, range.RowCount);
				Assert.AreEqual(0, range.StartColumn);
				Assert.AreEqual(0, range.ColumnCount);
			}
		}
	}
}
