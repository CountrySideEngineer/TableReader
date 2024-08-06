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
		public void GetTabeRowRange_001()
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
				Range ret = reader.GetTableRowRange(tableTop);
				Assert.AreEqual(4, ret.RowCount);
				Assert.AreEqual(3, ret.StartRow);
			}
		}

		[TestMethod]
		public void GetTabeRowRange_002()
		{
			var testDataPath = @"..\..\..\TestData\GetTable_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "GetTable_test_002";
				var reader = new ExcelTableReader(testDataStream, sheetName);
				Range tableTop = new Range()
				{
					StartRow = 3,
					StartColumn = 2,
				};
				Range ret = reader.GetTableRowRange(tableTop);
				Assert.AreEqual(3, ret.RowCount);
				Assert.AreEqual(3, ret.StartRow);
			}
		}

		[TestMethod]
		public void GetTabeRowRange_003()
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
				Range ret = reader.GetTableRowRange(tableTop);
				Assert.AreEqual(0, ret.RowCount);
				Assert.AreEqual(3, ret.StartRow);
			}
		}
	}
}
