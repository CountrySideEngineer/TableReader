using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TableReader.Excel;
using TableReader.TableData;

namespace TableReader_CTest
{
	public partial class TableReader_Test
	{
		[TestMethod]
		[Description("ReadColumn(Range range)")]
		public void GetTable_test_001()
		{
			var testDataPath = @"..\..\..\TestData\GetTable_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetTable_test_001";
				var reader = new ExcelTableReader(testDataStream, sheetName);
				string tableName = "GetTable_SampleTable";
				Range offset = new Range()
				{
					RowCount = 1,
					ColumnCount = 0
				};
				Content ret = reader.GetTable(tableName, offset);
				Assert.AreEqual(4, ret.GetContentsInCol(0).Count());
				Assert.AreEqual(5, ret.GetContentsInRow(0).Count());
				Assert.AreEqual("Item_r_001_c_001", ret.GetContent(0, 0));
				Assert.AreEqual("Item_r_001_c_002", ret.GetContent(0, 1));
				Assert.AreEqual("Item_r_001_c_003", ret.GetContent(0, 2));
				Assert.AreEqual("Item_r_001_c_004", ret.GetContent(0, 3));
				Assert.AreEqual("Item_r_001_c_005", ret.GetContent(0, 4));
				Assert.AreEqual("Item_r_002_c_001", ret.GetContent(1, 0));
				Assert.AreEqual("Item_r_002_c_002", ret.GetContent(1, 1));
				Assert.AreEqual("Item_r_002_c_003", ret.GetContent(1, 2));
				Assert.AreEqual("Item_r_002_c_004", ret.GetContent(1, 3));
				Assert.AreEqual("Item_r_002_c_005", ret.GetContent(1, 4));
				Assert.AreEqual("Item_r_003_c_001", ret.GetContent(2, 0));
				Assert.AreEqual("Item_r_003_c_002", ret.GetContent(2, 1));
				Assert.AreEqual("Item_r_003_c_003", ret.GetContent(2, 2));
				Assert.AreEqual("Item_r_003_c_004", ret.GetContent(2, 3));
				Assert.AreEqual("Item_r_003_c_005", ret.GetContent(2, 4));
				Assert.AreEqual("Item_r_004_c_001", ret.GetContent(3, 0));
				Assert.AreEqual("Item_r_004_c_002", ret.GetContent(3, 1));
				Assert.AreEqual("Item_r_004_c_003", ret.GetContent(3, 2));
				Assert.AreEqual("Item_r_004_c_004", ret.GetContent(3, 3));
				Assert.AreEqual("Item_r_004_c_005", ret.GetContent(3, 4));
			}
		}

		[TestMethod]
		[Description("ReadColumn(Range range)")]
		public void GetTable_test_002()
		{
			var testDataPath = @"..\..\..\TestData\GetTable_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetTable_test_002";
				var reader = new ExcelTableReader(testDataStream, sheetName);
				string tableName = "GetTable_SampleTable";
				Range offset = new Range()
				{
					RowCount = 1,
					ColumnCount = 0
				};
				Content ret = reader.GetTable(tableName, offset);
				Assert.AreEqual(3, ret.GetContentsInCol(0).Count());
				Assert.AreEqual(5, ret.GetContentsInRow(0).Count());
				Assert.AreEqual("Item_r_001_c_001", ret.GetContent(0, 0));
				Assert.AreEqual("Item_r_001_c_002", ret.GetContent(0, 1));
				Assert.AreEqual("Item_r_001_c_003", ret.GetContent(0, 2));
				Assert.AreEqual("Item_r_001_c_004", ret.GetContent(0, 3));
				Assert.AreEqual("Item_r_001_c_005", ret.GetContent(0, 4));
				Assert.AreEqual("Item_r_002_c_001", ret.GetContent(1, 0));
				Assert.AreEqual("Item_r_002_c_002", ret.GetContent(1, 1));
				Assert.AreEqual("Item_r_002_c_003", ret.GetContent(1, 2));
				Assert.AreEqual("Item_r_002_c_004", ret.GetContent(1, 3));
				Assert.AreEqual("Item_r_002_c_005", ret.GetContent(1, 4));
				Assert.AreEqual("Item_r_003_c_001", ret.GetContent(2, 0));
				Assert.AreEqual("Item_r_003_c_002", ret.GetContent(2, 1));
				Assert.AreEqual("Item_r_003_c_003", ret.GetContent(2, 2));
				Assert.AreEqual("Item_r_003_c_004", ret.GetContent(2, 3));
				Assert.AreEqual("Item_r_003_c_005", ret.GetContent(2, 4));
			}
		}

		[TestMethod]
		[Description("ReadColumn(Range range)")]
		[ExpectedException(typeof(ArgumentOutOfRangeException))]
		public void GetTable_test_003()
		{
			var testDataPath = @"..\..\..\TestData\GetTable_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetTable_test_003";
				var reader = new ExcelTableReader(testDataStream, sheetName);
				string tableName = "GetTable_SampleTable";
				Range offset = new Range()
				{
					RowCount = 1,
					ColumnCount = 0
				};
				Content ret = reader.GetTable(tableName, offset);

				Assert.AreEqual(0, ret.GetContentsInCol(0));
				Assert.Fail();
			}
		}

		[TestMethod]
		[Description("ReadColumn(Range range)")]
		public void GetTable_test_004()
		{
			var testDataPath = @"..\..\..\TestData\GetTable_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetTable_test_004";
				var reader = new ExcelTableReader(testDataStream, sheetName);
				string tableName = "GetTable_SampleTable";
				Range offset = new Range()
				{
					RowCount = 1,
					ColumnCount = 0
				};
				Content ret = reader.GetTable(tableName, offset);
				Assert.AreEqual(4, ret.GetContentsInCol(0).Count());
				Assert.AreEqual(4, ret.GetContentsInRow(0).Count());
				Assert.AreEqual("Item_r_001_c_001", ret.GetContent(0, 0));
				Assert.AreEqual("Item_r_001_c_002", ret.GetContent(0, 1));
				Assert.AreEqual("Item_r_001_c_003", ret.GetContent(0, 2));
				Assert.AreEqual("Item_r_001_c_004", ret.GetContent(0, 3));
				Assert.AreEqual("Item_r_002_c_001", ret.GetContent(1, 0));
				Assert.AreEqual("Item_r_002_c_002", ret.GetContent(1, 1));
				Assert.AreEqual("Item_r_002_c_003", ret.GetContent(1, 2));
				Assert.AreEqual("Item_r_002_c_004", ret.GetContent(1, 3));
				Assert.AreEqual("Item_r_003_c_001", ret.GetContent(2, 0));
				Assert.AreEqual("Item_r_003_c_002", ret.GetContent(2, 1));
				Assert.AreEqual("Item_r_003_c_003", ret.GetContent(2, 2));
				Assert.AreEqual("Item_r_003_c_004", ret.GetContent(2, 3));
				Assert.AreEqual("Item_r_004_c_001", ret.GetContent(3, 0));
				Assert.AreEqual("Item_r_004_c_002", ret.GetContent(3, 1));
				Assert.AreEqual("Item_r_004_c_003", ret.GetContent(3, 2));
				Assert.AreEqual("Item_r_004_c_004", ret.GetContent(3, 3));
			}
		}

		[TestMethod]
		[Description("ReadColumn(Range range)")]
		public void GetTable_test_005()
		{
			var testDataPath = @"..\..\..\TestData\GetTable_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetTable_test_005";
				var reader = new ExcelTableReader(testDataStream, sheetName);
				string tableName = "GetTable_SampleTable";
				Range offset = new Range()
				{
					RowCount = 1,
					ColumnCount = 0
				};
				Content ret = reader.GetTable(tableName, offset);
				Assert.AreEqual(4, ret.GetContentsInCol(0).Count());
				Assert.AreEqual(5, ret.GetContentsInRow(0).Count());
				Assert.AreEqual("Item_r_001_c_001", ret.GetContent(0, 0));
				Assert.AreEqual("Item_r_001_c_002", ret.GetContent(0, 1));
				Assert.AreEqual("Item_r_001_c_003", ret.GetContent(0, 2));
				Assert.AreEqual("Item_r_001_c_004", ret.GetContent(0, 3));
				Assert.AreEqual("Item_r_001_c_005", ret.GetContent(0, 4));
				Assert.AreEqual("Item_r_002_c_001", ret.GetContent(1, 0));
				Assert.AreEqual("Item_r_002_c_002", ret.GetContent(1, 1));
				Assert.AreEqual("Item_r_002_c_003", ret.GetContent(1, 2));
				Assert.AreEqual("Item_r_002_c_004", ret.GetContent(1, 3));
				Assert.AreEqual("Item_r_002_c_005", ret.GetContent(1, 4));
				Assert.AreEqual("Item_r_003_c_001", ret.GetContent(2, 0));
				Assert.AreEqual("Item_r_003_c_002", ret.GetContent(2, 1));
				Assert.AreEqual("Item_r_003_c_003", ret.GetContent(2, 2));
				Assert.AreEqual("Item_r_003_c_004", ret.GetContent(2, 3));
				Assert.AreEqual("Item_r_003_c_005", ret.GetContent(2, 4));
				Assert.AreEqual("Item_r_004_c_001", ret.GetContent(3, 0));
				Assert.AreEqual("Item_r_004_c_002", ret.GetContent(3, 1));
				Assert.AreEqual("Item_r_004_c_003", ret.GetContent(3, 2));
				Assert.AreEqual("Item_r_004_c_004", ret.GetContent(3, 3));
			}
		}

		[TestMethod]
		[Description("ReadColumn(Range range)")]
		public void GetTable_test_006()
		{
			var testDataPath = @"..\..\..\TestData\GetTable_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetTable_test_006";
				var reader = new ExcelTableReader(testDataStream, sheetName);
				string tableName = "GetTable_SampleTable";
				Range offset = new Range()
				{
					RowCount = 1,
					ColumnCount = 1
				};
				Content ret = reader.GetTable(tableName, offset);
				Assert.AreEqual(4, ret.GetContentsInCol(0).Count());
				Assert.AreEqual(5, ret.GetContentsInRow(0).Count());
				Assert.AreEqual("Item_r_001_c_001", ret.GetContent(0, 0));
				Assert.AreEqual("Item_r_001_c_002", ret.GetContent(0, 1));
				Assert.AreEqual("Item_r_001_c_003", ret.GetContent(0, 2));
				Assert.AreEqual("Item_r_001_c_004", ret.GetContent(0, 3));
				Assert.AreEqual("Item_r_001_c_005", ret.GetContent(0, 4));
				Assert.AreEqual("Item_r_002_c_001", ret.GetContent(1, 0));
				Assert.AreEqual("Item_r_002_c_002", ret.GetContent(1, 1));
				Assert.AreEqual("Item_r_002_c_003", ret.GetContent(1, 2));
				Assert.AreEqual("Item_r_002_c_004", ret.GetContent(1, 3));
				Assert.AreEqual("Item_r_002_c_005", ret.GetContent(1, 4));
				Assert.AreEqual("Item_r_003_c_001", ret.GetContent(2, 0));
				Assert.AreEqual("Item_r_003_c_002", ret.GetContent(2, 1));
				Assert.AreEqual("Item_r_003_c_003", ret.GetContent(2, 2));
				Assert.AreEqual("Item_r_003_c_004", ret.GetContent(2, 3));
				Assert.AreEqual("Item_r_003_c_005", ret.GetContent(2, 4));
				Assert.AreEqual("Item_r_004_c_001", ret.GetContent(3, 0));
				Assert.AreEqual("Item_r_004_c_002", ret.GetContent(3, 1));
				Assert.AreEqual("Item_r_004_c_003", ret.GetContent(3, 2));
				Assert.AreEqual("Item_r_004_c_004", ret.GetContent(3, 3));
				Assert.AreEqual("Item_r_004_c_005", ret.GetContent(3, 4));
			}
		}

		[TestMethod]
		[Description("ReadColumn(Range range)")]
		public void GetTable_test_007()
		{
			var testDataPath = @"..\..\..\TestData\GetTable_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetTable_test_006";
				var reader = new ExcelTableReader(testDataStream, sheetName);
				string tableName = "GetTable_SampleTable";
				Content ret = reader.GetTable(tableName);
				Assert.AreEqual(4, ret.GetContentsInCol(0).Count());
				Assert.AreEqual(5, ret.GetContentsInRow(0).Count());
				Assert.AreEqual("Item_r_001_c_001", ret.GetContent(0, 0));
				Assert.AreEqual("Item_r_001_c_002", ret.GetContent(0, 1));
				Assert.AreEqual("Item_r_001_c_003", ret.GetContent(0, 2));
				Assert.AreEqual("Item_r_001_c_004", ret.GetContent(0, 3));
				Assert.AreEqual("Item_r_001_c_005", ret.GetContent(0, 4));
				Assert.AreEqual("Item_r_002_c_001", ret.GetContent(1, 0));
				Assert.AreEqual("Item_r_002_c_002", ret.GetContent(1, 1));
				Assert.AreEqual("Item_r_002_c_003", ret.GetContent(1, 2));
				Assert.AreEqual("Item_r_002_c_004", ret.GetContent(1, 3));
				Assert.AreEqual("Item_r_002_c_005", ret.GetContent(1, 4));
				Assert.AreEqual("Item_r_003_c_001", ret.GetContent(2, 0));
				Assert.AreEqual("Item_r_003_c_002", ret.GetContent(2, 1));
				Assert.AreEqual("Item_r_003_c_003", ret.GetContent(2, 2));
				Assert.AreEqual("Item_r_003_c_004", ret.GetContent(2, 3));
				Assert.AreEqual("Item_r_003_c_005", ret.GetContent(2, 4));
				Assert.AreEqual("Item_r_004_c_001", ret.GetContent(3, 0));
				Assert.AreEqual("Item_r_004_c_002", ret.GetContent(3, 1));
				Assert.AreEqual("Item_r_004_c_003", ret.GetContent(3, 2));
				Assert.AreEqual("Item_r_004_c_004", ret.GetContent(3, 3));
				Assert.AreEqual("Item_r_004_c_005", ret.GetContent(3, 4));
			}
		}

		[TestMethod]
		[Description("ReadColumn(Range range)")]
		public void GetTable_test_008()
		{
			var testDataPath = @"..\..\..\TestData\GetTable_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetTable_test_007";
				var reader = new ExcelTableReader(testDataStream, sheetName);
				string tableName = "GetTable_SampleTable";
				Range offset = new Range()
				{
					RowCount = 1,
					ColumnCount = 0
				};
				Content ret = reader.GetTable(tableName, offset);
				Assert.AreEqual(4, ret.GetContentsInCol(0).Count());
				Assert.AreEqual(5, ret.GetContentsInRow(0).Count());
				Assert.AreEqual("Item_r_001_c_001", ret.GetContent(0, 0));
				Assert.AreEqual("Item_r_001_c_002", ret.GetContent(0, 1));
				Assert.AreEqual("Item_r_001_c_003", ret.GetContent(0, 2));
				Assert.AreEqual("Item_r_001_c_004", ret.GetContent(0, 3));
				Assert.AreEqual("Item_r_001_c_005", ret.GetContent(0, 4));
				Assert.AreEqual("Item_r_002_c_001", ret.GetContent(1, 0));
				Assert.AreEqual(string.Empty, ret.GetContent(1, 1));
				Assert.AreEqual("Item_r_002_c_003", ret.GetContent(1, 2));
				Assert.AreEqual("Item_r_002_c_004", ret.GetContent(1, 3));
				Assert.AreEqual("Item_r_002_c_005", ret.GetContent(1, 4));
				Assert.AreEqual("Item_r_003_c_001", ret.GetContent(2, 0));
				Assert.AreEqual("Item_r_003_c_002", ret.GetContent(2, 1));
				Assert.AreEqual("Item_r_003_c_003", ret.GetContent(2, 2));
				Assert.AreEqual("Item_r_003_c_004", ret.GetContent(2, 3));
				Assert.AreEqual("Item_r_003_c_005", ret.GetContent(2, 4));
				Assert.AreEqual("Item_r_004_c_001", ret.GetContent(3, 0));
				Assert.AreEqual("Item_r_004_c_002", ret.GetContent(3, 1));
				Assert.AreEqual("Item_r_004_c_003", ret.GetContent(3, 2));
				Assert.AreEqual("Item_r_004_c_004", ret.GetContent(3, 3));
				Assert.AreEqual("Item_r_004_c_005", ret.GetContent(3, 4));
			}
		}

		[TestMethod]
		[Description("ReadColumn(Range range)")]
		public void GetTable_test_009()
		{
			var testDataPath = @"..\..\..\TestData\GetTable_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetTable_test_008";
				var reader = new ExcelTableReader(testDataStream, sheetName);
				string tableName = "GetTable_SampleTable";
				Range offset = new Range()
				{
					RowCount = 1,
					ColumnCount = 0
				};
				Content ret = reader.GetTable(tableName, offset);
				Assert.AreEqual(4, ret.GetContentsInCol(0).Count());
				Assert.AreEqual(5, ret.GetContentsInRow(0).Count());
				Assert.AreEqual("Item_r_001_c_001", ret.GetContent(0, 0));
				Assert.AreEqual("Item_r_001_c_002", ret.GetContent(0, 1));
				Assert.AreEqual("Item_r_001_c_003", ret.GetContent(0, 2));
				Assert.AreEqual("Item_r_001_c_004", ret.GetContent(0, 3));
				Assert.AreEqual("Item_r_001_c_005", ret.GetContent(0, 4));
				Assert.AreEqual("Item_r_002_c_001", ret.GetContent(1, 0));
				Assert.AreEqual("Item_r_002_c_002", ret.GetContent(1, 1));
				Assert.AreEqual("Item_r_002_c_003", ret.GetContent(1, 2));
				Assert.AreEqual("Item_r_002_c_004", ret.GetContent(1, 3));
				Assert.AreEqual(string.Empty, ret.GetContent(1, 4));
				Assert.AreEqual("Item_r_003_c_001", ret.GetContent(2, 0));
				Assert.AreEqual("Item_r_003_c_002", ret.GetContent(2, 1));
				Assert.AreEqual("Item_r_003_c_003", ret.GetContent(2, 2));
				Assert.AreEqual("Item_r_003_c_004", ret.GetContent(2, 3));
				Assert.AreEqual("Item_r_003_c_005", ret.GetContent(2, 4));
				Assert.AreEqual("Item_r_004_c_001", ret.GetContent(3, 0));
				Assert.AreEqual("Item_r_004_c_002", ret.GetContent(3, 1));
				Assert.AreEqual("Item_r_004_c_003", ret.GetContent(3, 2));
				Assert.AreEqual("Item_r_004_c_004", ret.GetContent(3, 3));
				Assert.AreEqual("Item_r_004_c_005", ret.GetContent(3, 4));
			}
		}

		[TestMethod]
		[Description("ReadColumn(Range range)")]
		public void GetTable_test_010()
		{
			var testDataPath = @"..\..\..\TestData\GetTable_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetTable_test_009";
				var reader = new ExcelTableReader(testDataStream, sheetName);
				string tableName = "GetTable_SampleTable";
				Range offset = new Range()
				{
					RowCount = 1,
					ColumnCount = 0
				};
				Content ret = reader.GetTable(tableName, offset);
				Assert.AreEqual(4, ret.GetContentsInCol(0).Count());
				Assert.AreEqual(5, ret.GetContentsInRow(0).Count());
				Assert.AreEqual("Item_r_001_c_001", ret.GetContent(0, 0));
				Assert.AreEqual("Item_r_001_c_002", ret.GetContent(0, 1));
				Assert.AreEqual("Item_r_001_c_003", ret.GetContent(0, 2));
				Assert.AreEqual("Item_r_001_c_004", ret.GetContent(0, 3));
				Assert.AreEqual("Item_r_001_c_005", ret.GetContent(0, 4));
				Assert.AreEqual("Item_r_002_c_001", ret.GetContent(1, 0));
				Assert.AreEqual("Item_r_002_c_002", ret.GetContent(1, 1));
				Assert.AreEqual("Item_r_002_c_003", ret.GetContent(1, 2));
				Assert.AreEqual("Item_r_002_c_004", ret.GetContent(1, 3));
				Assert.AreEqual("Item_r_002_c_005", ret.GetContent(1, 4));
				Assert.AreEqual("Item_r_003_c_001", ret.GetContent(2, 0));
				Assert.AreEqual("Item_r_003_c_002", ret.GetContent(2, 1));
				Assert.AreEqual("Item_r_003_c_003", ret.GetContent(2, 2));
				Assert.AreEqual("Item_r_003_c_004", ret.GetContent(2, 3));
				Assert.AreEqual("Item_r_003_c_005", ret.GetContent(2, 4));
				Assert.AreEqual("Item_r_004_c_001", ret.GetContent(3, 0));
				Assert.AreEqual(string.Empty, ret.GetContent(3, 1));
				Assert.AreEqual("Item_r_004_c_003", ret.GetContent(3, 2));
				Assert.AreEqual("Item_r_004_c_004", ret.GetContent(3, 3));
				Assert.AreEqual("Item_r_004_c_005", ret.GetContent(3, 4));
			}
		}

		[TestMethod]
		[Description("ReadColumn(Range range)")]
		public void GetTable_test_011()
		{
			var testDataPath = @"..\..\..\TestData\GetTable_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetTable_test_010";
				var reader = new ExcelTableReader(testDataStream, sheetName);
				string tableName = "GetTable_SampleTable";
				Range offset = new Range()
				{
					RowCount = 1,
					ColumnCount = 0
				};
				Content ret = reader.GetTable(tableName, offset);
				Assert.AreEqual(4, ret.GetContentsInCol(0).Count());
				Assert.AreEqual(5, ret.GetContentsInRow(0).Count());
				Assert.AreEqual("Item_r_001_c_001", ret.GetContent(0, 0));
				Assert.AreEqual("Item_r_001_c_002", ret.GetContent(0, 1));
				Assert.AreEqual("Item_r_001_c_003", ret.GetContent(0, 2));
				Assert.AreEqual("Item_r_001_c_004", ret.GetContent(0, 3));
				Assert.AreEqual("Item_r_001_c_005", ret.GetContent(0, 4));
				Assert.AreEqual("Item_r_002_c_001", ret.GetContent(1, 0));
				Assert.AreEqual("Item_r_002_c_002", ret.GetContent(1, 1));
				Assert.AreEqual("Item_r_002_c_003", ret.GetContent(1, 2));
				Assert.AreEqual("Item_r_002_c_004", ret.GetContent(1, 3));
				Assert.AreEqual("Item_r_002_c_005", ret.GetContent(1, 4));
				Assert.AreEqual("Item_r_003_c_001", ret.GetContent(2, 0));
				Assert.AreEqual("Item_r_003_c_002", ret.GetContent(2, 1));
				Assert.AreEqual("Item_r_003_c_003", ret.GetContent(2, 2));
				Assert.AreEqual("Item_r_003_c_004", ret.GetContent(2, 3));
				Assert.AreEqual("Item_r_003_c_005", ret.GetContent(2, 4));
				Assert.AreEqual("Item_r_004_c_001", ret.GetContent(3, 0));
				Assert.AreEqual("Item_r_004_c_002", ret.GetContent(3, 1));
				Assert.AreEqual("Item_r_004_c_003", ret.GetContent(3, 2));
				Assert.AreEqual("Item_r_004_c_004", ret.GetContent(3, 3));
				Assert.AreEqual(string.Empty, ret.GetContent(3, 4));
			}
		}
	}
}
