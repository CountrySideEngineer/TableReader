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
		[Description("GetRowRange(ref Range range)")]
		public void GetRowRange_test_001()
		{
			var testDataPath = @"..\..\..\TestData\GetRowRange_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetRowRange_test_001";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				Range range = new Range()
				{
					StartRow = -1,
					RowCount = -1,
				};
				reader.GetRowRange(ref range);
				Assert.AreEqual(1, range.StartRow);
				Assert.AreEqual(1, range.RowCount);
			}
		}

		[TestMethod]
		[Description("GetRowRange(ref Range range)")]
		public void GetRowRange_test_002()
		{
			var testDataPath = @"..\..\..\TestData\GetRowRange_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetRowRange_test_002";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				Range range = new Range()
				{
					StartRow = -1,
					RowCount = -1,
				};
				reader.GetRowRange(ref range);
				Assert.AreEqual(1, range.StartRow);
				Assert.AreEqual(2, range.RowCount);
			}
		}

		[TestMethod]
		[Description("GetRowRange(ref Range range)")]
		public void GetRowRange_test_003()
		{
			var testDataPath = @"..\..\..\TestData\GetRowRange_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetRowRange_test_003";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				Range range = new Range()
				{
					StartRow = -1,
					RowCount = -1,
				};
				reader.GetRowRange(ref range);
				Assert.AreEqual(2, range.StartRow);
				Assert.AreEqual(2, range.RowCount);
			}
		}

		[TestMethod]
		[Description("GetRowRange(ref Range range)")]
		public void GetRowRange_test_004()
		{
			var testDataPath = @"..\..\..\TestData\GetRowRange_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetRowRange_test_004";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				Range range = new Range()
				{
					StartRow = -1,
					RowCount = -1,
				};
				reader.GetRowRange(ref range);
				Assert.AreEqual(2, range.StartRow);
				Assert.AreEqual(3, range.RowCount);
			}
		}

		[TestMethod]
		[Description("GetRowRange(ref Range range)")]
		public void GetRowRange_test_005()
		{
			var testDataPath = @"..\..\..\TestData\GetRowRange_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetRowRange_test_005";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				Range range = new Range()
				{
					StartRow = -1,
					RowCount = -1,
				};
				reader.GetRowRange(ref range);
				Assert.AreEqual(2, range.StartRow);
				Assert.AreEqual(3, range.RowCount);
			}
		}
	}
}
