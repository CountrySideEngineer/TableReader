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
		[Description("GetColumnRange(ref Range range)")]
		public void GetColumnRange_test_001()
		{
			var testDataPath = @"..\..\..\TestData\GetColumnRange_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetColumnRange_test_001";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				Range range = new Range()
				{
					StartColumn = -1,
					ColumnCount = -1,
				};
				reader.GetColumnRange(ref range);
				Assert.AreEqual(1, range.StartColumn);
				Assert.AreEqual(1, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("GetColumnRange(ref Range range)")]
		public void GetColumnRange_test_002()
		{
			var testDataPath = @"..\..\..\TestData\GetColumnRange_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetColumnRange_test_002";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				Range range = new Range()
				{
					StartColumn = -1,
					ColumnCount = -1,
				};
				reader.GetColumnRange(ref range);
				Assert.AreEqual(1, range.StartColumn);
				Assert.AreEqual(2, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("GetColumnRange(ref Range range)")]
		public void GetColumnRange_test_003()
		{
			var testDataPath = @"..\..\..\TestData\GetColumnRange_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetColumnRange_test_003";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				Range range = new Range()
				{
					StartColumn = -1,
					ColumnCount = -1,
				};
				reader.GetColumnRange(ref range);
				Assert.AreEqual(2, range.StartColumn);
				Assert.AreEqual(2, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("GetColumnRange(ref Range range)")]
		public void GetColumnRange_test_004()
		{
			var testDataPath = @"..\..\..\TestData\GetColumnRange_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetColumnRange_test_004";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				Range range = new Range()
				{
					StartColumn = -1,
					ColumnCount = -1,
				};
				reader.GetColumnRange(ref range);
				Assert.AreEqual(2, range.StartColumn);
				Assert.AreEqual(3, range.ColumnCount);
			}
		}

		[TestMethod]
		[Description("GetColumnRange(ref Range range)")]
		public void GetColumnRange_test_005()
		{
			var testDataPath = @"..\..\..\TestData\GetColumnRange_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetColumnRange_test_005";
				var reader = new ExcelTableReader(testDataStream, sheetName);

				Range range = new Range()
				{
					StartColumn = -1,
					ColumnCount = -1,
				};
				reader.GetColumnRange(ref range);
				Assert.AreEqual(2, range.StartColumn);
				Assert.AreEqual(3, range.ColumnCount);
			}
		}
	}
}
