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
		public void GetTabeColumnRange_001()
		{
			var testDataPath = @"..\..\..\TestData\GetTable_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetTable_test_001";
				var reader = new ExcelTableReader(testDataStream, sheetName);
				Range tableTop = new Range()
				{
					StartRow = 3,
					StartColumn = 2,
				};
				Range ret = reader.GetTableColumnRange(tableTop);
				Assert.AreEqual(5, ret.ColumnCount);
				Assert.AreEqual(2, ret.StartColumn);
			}
		}

		[TestMethod]
		public void GetTabeColumnRange_002()
		{
			var testDataPath = @"..\..\..\TestData\GetTable_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetTable_test_004";
				var reader = new ExcelTableReader(testDataStream, sheetName);
				Range tableTop = new Range()
				{
					StartRow = 3,
					StartColumn = 2,
				};
				Range ret = reader.GetTableColumnRange(tableTop);
				Assert.AreEqual(4, ret.ColumnCount);
				Assert.AreEqual(2, ret.StartColumn);
			}
		}

		[TestMethod]
		public void GetTabeColumnRange_003()
		{
			var testDataPath = @"..\..\..\TestData\GetTable_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetTable_test_003";
				var reader = new ExcelTableReader(testDataStream, sheetName);
				Range tableTop = new Range()
				{
					StartRow = 3,
					StartColumn = 2,
				};
				Range ret = reader.GetTableColumnRange(tableTop);
				Assert.AreEqual(0, ret.ColumnCount);
				Assert.AreEqual(2, ret.StartColumn);
			}
		}
	}
}
