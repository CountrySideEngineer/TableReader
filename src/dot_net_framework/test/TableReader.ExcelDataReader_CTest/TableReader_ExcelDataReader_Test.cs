using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Data;
using System.IO;
using TableReader.Interface;

namespace TableReader.ExcelDataReader_CTest
{
	[TestClass]
	public class TableReader_ExcelDataReader_Test
	{
		string _testDataRoot = @"..\..\..\TestData\";

		[TestMethod]
		[TestCategory("IntegrationTest")]
		public void Read_Test_001()
		{
			string testDataPath = _testDataRoot + @"TableReader.ExcelDataReader.Reader_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "Read_test_001";
				string tableName = "TestTable_001";
				ITableReader reader = new TableReader.ExcelDataReader.ExcelTableReader(testDataStream, sheetName);
				DataTable dataTable = reader.Read(tableName);

				Assert.AreEqual(tableName, dataTable.TableName);
				Assert.AreEqual(1, dataTable.Rows.Count);
				Assert.AreEqual("Data_001_r_001_c_001", dataTable.Rows[0][0]);
				Assert.AreEqual("Data_001_r_001_c_001", dataTable.Rows[0]["Header_r_001_c_001"]);
			}
		}

		[TestMethod]
		[TestCategory("ConbinationTest")]
		public void Read_Test_002()
		{
			string testDataPath = _testDataRoot + @"TableReader.ExcelDataReader.Reader_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "Read_test_001";
				string tableName = "TestTable_002";
				ITableReader reader = new TableReader.ExcelDataReader.ExcelTableReader(testDataStream, sheetName);
				DataTable dataTable = reader.Read(tableName);

				Assert.AreEqual(tableName, dataTable.TableName);
				Assert.AreEqual(2, dataTable.Rows.Count);
				Assert.AreEqual("Data_001_r_001_c_001", dataTable.Rows[0][0]);
				Assert.AreEqual("Data_001_r_001_c_002", dataTable.Rows[0][1]);
				Assert.AreEqual("Data_001_r_001_c_001", dataTable.Rows[0]["Header_r_001_c_001"]);
				Assert.AreEqual("Data_001_r_001_c_002", dataTable.Rows[0]["Header_r_001_c_002"]);
				Assert.AreEqual("Data_001_r_002_c_001", dataTable.Rows[1][0]);
				Assert.AreEqual("Data_001_r_002_c_002", dataTable.Rows[1][1]);
				Assert.AreEqual("Data_001_r_002_c_001", dataTable.Rows[1]["Header_r_001_c_001"]);
				Assert.AreEqual("Data_001_r_002_c_002", dataTable.Rows[1]["Header_r_001_c_002"]);
			}
		}

		[TestMethod]
		[TestCategory("ConbinationTest")]
		public void Read_Test_003()
		{
			string testDataPath = _testDataRoot + @"TableReader.ExcelDataReader.Reader_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "Read_test_001";
				string tableName = "TestTable_003";
				ITableReader reader = new TableReader.ExcelDataReader.ExcelTableReader(testDataStream, sheetName);
				DataTable dataTable = reader.Read(tableName);

				Assert.AreEqual(tableName, dataTable.TableName);
				Assert.AreEqual(4, dataTable.Rows.Count);
				Assert.AreEqual(5, dataTable.Columns.Count);
				Assert.AreEqual("Data_003_r_001_c_001", dataTable.Rows[0][0]);
				Assert.AreEqual("Data_003_r_001_c_002", dataTable.Rows[0][1]);
				Assert.AreEqual("Data_003_r_001_c_003", dataTable.Rows[0][2]);
				Assert.AreEqual("Data_003_r_001_c_004", dataTable.Rows[0][3]);
				Assert.AreEqual("Data_003_r_001_c_005", dataTable.Rows[0][4]);
				Assert.AreEqual("Data_003_r_002_c_001", dataTable.Rows[1][0]);
				Assert.AreEqual("Data_003_r_002_c_002", dataTable.Rows[1][1]);
				Assert.AreEqual("Data_003_r_002_c_003", dataTable.Rows[1][2]);
				Assert.AreEqual("Data_003_r_002_c_004", dataTable.Rows[1][3]);
				Assert.AreEqual("Data_003_r_002_c_005", dataTable.Rows[1][4]);
				Assert.AreEqual("Data_003_r_003_c_001", dataTable.Rows[2][0]);
				Assert.AreEqual("Data_003_r_003_c_002", dataTable.Rows[2][1]);
				Assert.AreEqual("Data_003_r_003_c_003", dataTable.Rows[2][2]);
				Assert.AreEqual("Data_003_r_003_c_004", dataTable.Rows[2][3]);
				Assert.AreEqual("Data_003_r_003_c_005", dataTable.Rows[2][4]);
				Assert.AreEqual("Data_003_r_004_c_001", dataTable.Rows[3][0]);
				Assert.AreEqual("Data_003_r_004_c_002", dataTable.Rows[3][1]);
				Assert.AreEqual("Data_003_r_004_c_003", dataTable.Rows[3][2]);
				Assert.AreEqual("Data_003_r_004_c_004", dataTable.Rows[3][3]);
				Assert.AreEqual("Data_003_r_004_c_005", dataTable.Rows[3][4]);
				Assert.AreEqual("Data_003_r_001_c_001", dataTable.Rows[0]["Header_003_r_001_c_001"]);
				Assert.AreEqual("Data_003_r_001_c_002", dataTable.Rows[0]["Header_003_r_001_c_002"]);
				Assert.AreEqual("Data_003_r_001_c_003", dataTable.Rows[0]["Header_003_r_001_c_003"]);
				Assert.AreEqual("Data_003_r_001_c_004", dataTable.Rows[0]["Header_003_r_001_c_004"]);
				Assert.AreEqual("Data_003_r_001_c_005", dataTable.Rows[0]["Header_003_r_001_c_005"]);
				Assert.AreEqual("Data_003_r_002_c_001", dataTable.Rows[1]["Header_003_r_001_c_001"]);
				Assert.AreEqual("Data_003_r_002_c_002", dataTable.Rows[1]["Header_003_r_001_c_002"]);
				Assert.AreEqual("Data_003_r_002_c_003", dataTable.Rows[1]["Header_003_r_001_c_003"]);
				Assert.AreEqual("Data_003_r_002_c_004", dataTable.Rows[1]["Header_003_r_001_c_004"]);
				Assert.AreEqual("Data_003_r_002_c_005", dataTable.Rows[1]["Header_003_r_001_c_005"]);
				Assert.AreEqual("Data_003_r_003_c_001", dataTable.Rows[2]["Header_003_r_001_c_001"]);
				Assert.AreEqual("Data_003_r_003_c_002", dataTable.Rows[2]["Header_003_r_001_c_002"]);
				Assert.AreEqual("Data_003_r_003_c_003", dataTable.Rows[2]["Header_003_r_001_c_003"]);
				Assert.AreEqual("Data_003_r_003_c_004", dataTable.Rows[2]["Header_003_r_001_c_004"]);
				Assert.AreEqual("Data_003_r_003_c_005", dataTable.Rows[2]["Header_003_r_001_c_005"]);
				Assert.AreEqual("Data_003_r_004_c_001", dataTable.Rows[3]["Header_003_r_001_c_001"]);
				Assert.AreEqual("Data_003_r_004_c_002", dataTable.Rows[3]["Header_003_r_001_c_002"]);
				Assert.AreEqual("Data_003_r_004_c_003", dataTable.Rows[3]["Header_003_r_001_c_003"]);
				Assert.AreEqual("Data_003_r_004_c_004", dataTable.Rows[3]["Header_003_r_001_c_004"]);
				Assert.AreEqual("Data_003_r_004_c_005", dataTable.Rows[3]["Header_003_r_001_c_005"]);
			}
		}

		[TestMethod]
		[TestCategory("ConbinationTest")]
		public void Read_Test_004()
		{
			string testDataPath = _testDataRoot + @"TableReader.ExcelDataReader.Reader_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "Read_test_001";
				string tableName = "TestTable_004";
				ITableReader reader = new TableReader.ExcelDataReader.ExcelTableReader(testDataStream, sheetName);
				DataTable dataTable = reader.Read(tableName);

				Assert.AreEqual(tableName, dataTable.TableName);
				Assert.AreEqual(4, dataTable.Rows.Count);
				Assert.AreEqual(4, dataTable.Columns.Count);
				Assert.AreEqual("Data_004_r_001_c_001", dataTable.Rows[0][0]);
				Assert.AreEqual("Data_004_r_001_c_002", dataTable.Rows[0][1]);
				Assert.AreEqual("Data_004_r_001_c_003", dataTable.Rows[0][2]);
				Assert.AreEqual("Data_004_r_001_c_004", dataTable.Rows[0][3]);
				Assert.AreEqual("Data_004_r_002_c_001", dataTable.Rows[1][0]);
				Assert.AreEqual("Data_004_r_002_c_002", dataTable.Rows[1][1]);
				Assert.AreEqual("Data_004_r_002_c_003", dataTable.Rows[1][2]);
				Assert.AreEqual("Data_004_r_002_c_004", dataTable.Rows[1][3]);
				Assert.AreEqual("Data_004_r_003_c_001", dataTable.Rows[2][0]);
				Assert.AreEqual("Data_004_r_003_c_002", dataTable.Rows[2][1]);
				Assert.AreEqual("Data_004_r_003_c_003", dataTable.Rows[2][2]);
				Assert.AreEqual("Data_004_r_003_c_004", dataTable.Rows[2][3]);
				Assert.AreEqual("Data_004_r_004_c_001", dataTable.Rows[3][0]);
				Assert.AreEqual("Data_004_r_004_c_002", dataTable.Rows[3][1]);
				Assert.AreEqual("Data_004_r_004_c_003", dataTable.Rows[3][2]);
				Assert.AreEqual("Data_004_r_004_c_004", dataTable.Rows[3][3]);
				Assert.AreEqual("Data_004_r_001_c_001", dataTable.Rows[0]["Header_004_r_001_c_001"]);
				Assert.AreEqual("Data_004_r_001_c_002", dataTable.Rows[0]["Header_004_r_001_c_002"]);
				Assert.AreEqual("Data_004_r_001_c_003", dataTable.Rows[0]["Header_004_r_001_c_003"]);
				Assert.AreEqual("Data_004_r_001_c_004", dataTable.Rows[0]["Header_004_r_001_c_004"]);
				Assert.AreEqual("Data_004_r_002_c_001", dataTable.Rows[1]["Header_004_r_001_c_001"]);
				Assert.AreEqual("Data_004_r_002_c_002", dataTable.Rows[1]["Header_004_r_001_c_002"]);
				Assert.AreEqual("Data_004_r_002_c_003", dataTable.Rows[1]["Header_004_r_001_c_003"]);
				Assert.AreEqual("Data_004_r_002_c_004", dataTable.Rows[1]["Header_004_r_001_c_004"]);
				Assert.AreEqual("Data_004_r_003_c_001", dataTable.Rows[2]["Header_004_r_001_c_001"]);
				Assert.AreEqual("Data_004_r_003_c_002", dataTable.Rows[2]["Header_004_r_001_c_002"]);
				Assert.AreEqual("Data_004_r_003_c_003", dataTable.Rows[2]["Header_004_r_001_c_003"]);
				Assert.AreEqual("Data_004_r_003_c_004", dataTable.Rows[2]["Header_004_r_001_c_004"]);
				Assert.AreEqual("Data_004_r_004_c_001", dataTable.Rows[3]["Header_004_r_001_c_001"]);
				Assert.AreEqual("Data_004_r_004_c_002", dataTable.Rows[3]["Header_004_r_001_c_002"]);
				Assert.AreEqual("Data_004_r_004_c_003", dataTable.Rows[3]["Header_004_r_001_c_003"]);
				Assert.AreEqual("Data_004_r_004_c_004", dataTable.Rows[3]["Header_004_r_001_c_004"]);
			}
		}

		[TestMethod]
		[TestCategory("ConbinationTest")]
		public void Read_Test_005()
		{
			string testDataPath = _testDataRoot + @"TableReader.ExcelDataReader.Reader_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "Read_test_001";
				string tableName = "TestTable_005";
				ITableReader reader = new TableReader.ExcelDataReader.ExcelTableReader(testDataStream, sheetName);
				DataTable dataTable = reader.Read(tableName);

				Assert.AreEqual(tableName, dataTable.TableName);
				Assert.AreEqual(0, dataTable.Rows.Count);
				Assert.AreEqual(5, dataTable.Columns.Count);
			}
		}

		[TestMethod]
		[TestCategory("ConbinationTest")]
		public void Read_Test_006()
		{
			string testDataPath = _testDataRoot + @"TableReader.ExcelDataReader.Reader_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "Read_test_001";
				string tableName = "TestTable_006";
				ITableReader reader = new TableReader.ExcelDataReader.ExcelTableReader(testDataStream, sheetName);
				DataTable dataTable = reader.Read(tableName);

				Assert.AreEqual(tableName, dataTable.TableName);
				Assert.AreEqual(4, dataTable.Rows.Count);
				Assert.AreEqual(5, dataTable.Columns.Count);
				Assert.AreEqual("Data_005_r_001_c_001", dataTable.Rows[0][0]);
				Assert.AreEqual("Data_005_r_001_c_002", dataTable.Rows[0][1]);
				Assert.AreEqual("Data_005_r_001_c_003", dataTable.Rows[0][2]);
				Assert.AreEqual("Data_005_r_001_c_004", dataTable.Rows[0][3]);
				Assert.AreEqual("", dataTable.Rows[0][4]);
				Assert.AreEqual("Data_005_r_002_c_001", dataTable.Rows[1][0]);
				Assert.AreEqual("Data_005_r_002_c_002", dataTable.Rows[1][1]);
				Assert.AreEqual("Data_005_r_002_c_003", dataTable.Rows[1][2]);
				Assert.AreEqual("Data_005_r_002_c_004", dataTable.Rows[1][3]);
				Assert.AreEqual("Data_005_r_002_c_005", dataTable.Rows[1][4]);
				Assert.AreEqual("Data_005_r_003_c_001", dataTable.Rows[2][0]);
				Assert.AreEqual("Data_005_r_003_c_002", dataTable.Rows[2][1]);
				Assert.AreEqual("Data_005_r_003_c_003", dataTable.Rows[2][2]);
				Assert.AreEqual("Data_005_r_003_c_004", dataTable.Rows[2][3]);
				Assert.AreEqual("Data_005_r_003_c_005", dataTable.Rows[2][4]);
				Assert.AreEqual("Data_005_r_004_c_001", dataTable.Rows[3][0]);
				Assert.AreEqual("Data_005_r_004_c_002", dataTable.Rows[3][1]);
				Assert.AreEqual("Data_005_r_004_c_003", dataTable.Rows[3][2]);
				Assert.AreEqual("Data_005_r_004_c_004", dataTable.Rows[3][3]);
				Assert.AreEqual("Data_005_r_004_c_005", dataTable.Rows[3][4]);
				Assert.AreEqual("Data_005_r_001_c_001", dataTable.Rows[0]["Header_005_r_001_c_001"]);
				Assert.AreEqual("Data_005_r_001_c_002", dataTable.Rows[0]["Header_005_r_001_c_002"]);
				Assert.AreEqual("Data_005_r_001_c_003", dataTable.Rows[0]["Header_005_r_001_c_003"]);
				Assert.AreEqual("Data_005_r_001_c_004", dataTable.Rows[0]["Header_005_r_001_c_004"]);
				Assert.AreEqual("", dataTable.Rows[0]["Header_005_r_001_c_005"]);
				Assert.AreEqual("Data_005_r_002_c_001", dataTable.Rows[1]["Header_005_r_001_c_001"]);
				Assert.AreEqual("Data_005_r_002_c_002", dataTable.Rows[1]["Header_005_r_001_c_002"]);
				Assert.AreEqual("Data_005_r_002_c_003", dataTable.Rows[1]["Header_005_r_001_c_003"]);
				Assert.AreEqual("Data_005_r_002_c_004", dataTable.Rows[1]["Header_005_r_001_c_004"]);
				Assert.AreEqual("Data_005_r_002_c_005", dataTable.Rows[1]["Header_005_r_001_c_005"]);
				Assert.AreEqual("Data_005_r_003_c_001", dataTable.Rows[2]["Header_005_r_001_c_001"]);
				Assert.AreEqual("Data_005_r_003_c_002", dataTable.Rows[2]["Header_005_r_001_c_002"]);
				Assert.AreEqual("Data_005_r_003_c_003", dataTable.Rows[2]["Header_005_r_001_c_003"]);
				Assert.AreEqual("Data_005_r_003_c_004", dataTable.Rows[2]["Header_005_r_001_c_004"]);
				Assert.AreEqual("Data_005_r_003_c_005", dataTable.Rows[2]["Header_005_r_001_c_005"]);
				Assert.AreEqual("Data_005_r_004_c_001", dataTable.Rows[3]["Header_005_r_001_c_001"]);
				Assert.AreEqual("Data_005_r_004_c_002", dataTable.Rows[3]["Header_005_r_001_c_002"]);
				Assert.AreEqual("Data_005_r_004_c_003", dataTable.Rows[3]["Header_005_r_001_c_003"]);
				Assert.AreEqual("Data_005_r_004_c_004", dataTable.Rows[3]["Header_005_r_001_c_004"]);
				Assert.AreEqual("Data_005_r_004_c_005", dataTable.Rows[3]["Header_005_r_001_c_005"]);
			}
		}

		[TestMethod]
		[TestCategory("ConbinationTest")]
		public void Read_Test_007()
		{
			string testDataPath = _testDataRoot + @"TableReader.ExcelDataReader.Reader_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "Read_test_001";
				string tableName = "TestTable_007";
				ITableReader reader = new TableReader.ExcelDataReader.ExcelTableReader(testDataStream, sheetName);
				DataTable dataTable = reader.Read(tableName);

				Assert.AreEqual(tableName, dataTable.TableName);
				Assert.AreEqual(3, dataTable.Rows.Count);
				Assert.AreEqual(5, dataTable.Columns.Count);
				Assert.AreEqual("Data_005_r_001_c_001", dataTable.Rows[0][0]);
				Assert.AreEqual("Data_005_r_001_c_002", dataTable.Rows[0][1]);
				Assert.AreEqual("Data_005_r_001_c_003", dataTable.Rows[0][2]);
				Assert.AreEqual("Data_005_r_001_c_004", dataTable.Rows[0][3]);
				Assert.AreEqual("Data_005_r_002_c_001", dataTable.Rows[1][0]);
				Assert.AreEqual("Data_005_r_002_c_002", dataTable.Rows[1][1]);
				Assert.AreEqual("Data_005_r_002_c_003", dataTable.Rows[1][2]);
				Assert.AreEqual("Data_005_r_002_c_004", dataTable.Rows[1][3]);
				Assert.AreEqual("Data_005_r_003_c_001", dataTable.Rows[2][0]);
				Assert.AreEqual("Data_005_r_003_c_002", dataTable.Rows[2][1]);
				Assert.AreEqual("Data_005_r_003_c_003", dataTable.Rows[2][2]);
				Assert.AreEqual("Data_005_r_003_c_004", dataTable.Rows[2][3]);
				Assert.AreEqual("Data_005_r_001_c_001", dataTable.Rows[0]["Header_005_r_001_c_001"]);
				Assert.AreEqual("Data_005_r_001_c_002", dataTable.Rows[0]["Header_005_r_001_c_002"]);
				Assert.AreEqual("Data_005_r_001_c_003", dataTable.Rows[0]["Header_005_r_001_c_003"]);
				Assert.AreEqual("Data_005_r_001_c_004", dataTable.Rows[0]["Header_005_r_001_c_004"]);
				Assert.AreEqual("Data_005_r_002_c_001", dataTable.Rows[1]["Header_005_r_001_c_001"]);
				Assert.AreEqual("Data_005_r_002_c_002", dataTable.Rows[1]["Header_005_r_001_c_002"]);
				Assert.AreEqual("Data_005_r_002_c_003", dataTable.Rows[1]["Header_005_r_001_c_003"]);
				Assert.AreEqual("Data_005_r_002_c_004", dataTable.Rows[1]["Header_005_r_001_c_004"]);
				Assert.AreEqual("Data_005_r_003_c_001", dataTable.Rows[2]["Header_005_r_001_c_001"]);
				Assert.AreEqual("Data_005_r_003_c_002", dataTable.Rows[2]["Header_005_r_001_c_002"]);
				Assert.AreEqual("Data_005_r_003_c_003", dataTable.Rows[2]["Header_005_r_001_c_003"]);
				Assert.AreEqual("Data_005_r_003_c_004", dataTable.Rows[2]["Header_005_r_001_c_004"]);
			}
		}

		[TestMethod]
		[TestCategory("ConbinationTest")]
		public void Read_Test_008()
		{
			string testDataPath = _testDataRoot + @"TableReader.ExcelDataReader.Reader_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "Read_test_001";
				string tableName = "TestTable_008";
				ITableReader reader = new TableReader.ExcelDataReader.ExcelTableReader(testDataStream, sheetName);
				DataTable dataTable = reader.Read(tableName);

				Assert.AreEqual(tableName, dataTable.TableName);
				Assert.AreEqual(4, dataTable.Rows.Count);
				Assert.AreEqual(5, dataTable.Columns.Count);
				Assert.AreEqual("Data_005_r_001_c_001", dataTable.Rows[0][0]);
				Assert.AreEqual("Data_005_r_001_c_002", dataTable.Rows[0][1]);
				Assert.AreEqual("Data_005_r_001_c_003", dataTable.Rows[0][2]);
				Assert.AreEqual("Data_005_r_001_c_004", dataTable.Rows[0][3]);
				Assert.AreEqual("Data_005_r_001_c_005", dataTable.Rows[0][4]);
				Assert.AreEqual("Data_005_r_002_c_001", dataTable.Rows[1][0]);
				Assert.AreEqual("Data_005_r_002_c_002", dataTable.Rows[1][1]);
				Assert.AreEqual("Data_005_r_002_c_003", dataTable.Rows[1][2]);
				Assert.AreEqual("Data_005_r_002_c_004", dataTable.Rows[1][3]);
				Assert.AreEqual("Data_005_r_002_c_005", dataTable.Rows[1][4]);
				Assert.AreEqual("Data_005_r_003_c_001", dataTable.Rows[2][0]);
				Assert.AreEqual("Data_005_r_003_c_002", dataTable.Rows[2][1]);
				Assert.AreEqual("Data_005_r_003_c_003", dataTable.Rows[2][2]);
				Assert.AreEqual("Data_005_r_003_c_004", dataTable.Rows[2][3]);
				Assert.AreEqual("Data_005_r_003_c_005", dataTable.Rows[2][4]);
				Assert.AreEqual("Data_005_r_004_c_001", dataTable.Rows[3][0]);
				Assert.AreEqual("Data_005_r_004_c_002", dataTable.Rows[3][1]);
				Assert.AreEqual("Data_005_r_004_c_003", dataTable.Rows[3][2]);
				Assert.AreEqual("Data_005_r_004_c_004", dataTable.Rows[3][3]);
				Assert.AreEqual("", dataTable.Rows[3][4]);
				Assert.AreEqual("Data_005_r_001_c_001", dataTable.Rows[0]["Header_005_r_001_c_001"]);
				Assert.AreEqual("Data_005_r_001_c_002", dataTable.Rows[0]["Header_005_r_001_c_002"]);
				Assert.AreEqual("Data_005_r_001_c_003", dataTable.Rows[0]["Header_005_r_001_c_003"]);
				Assert.AreEqual("Data_005_r_001_c_004", dataTable.Rows[0]["Header_005_r_001_c_004"]);
				Assert.AreEqual("Data_005_r_001_c_005", dataTable.Rows[0]["Header_005_r_001_c_005"]);
				Assert.AreEqual("Data_005_r_002_c_001", dataTable.Rows[1]["Header_005_r_001_c_001"]);
				Assert.AreEqual("Data_005_r_002_c_002", dataTable.Rows[1]["Header_005_r_001_c_002"]);
				Assert.AreEqual("Data_005_r_002_c_003", dataTable.Rows[1]["Header_005_r_001_c_003"]);
				Assert.AreEqual("Data_005_r_002_c_004", dataTable.Rows[1]["Header_005_r_001_c_004"]);
				Assert.AreEqual("Data_005_r_002_c_005", dataTable.Rows[1]["Header_005_r_001_c_005"]);
				Assert.AreEqual("Data_005_r_003_c_001", dataTable.Rows[2]["Header_005_r_001_c_001"]);
				Assert.AreEqual("Data_005_r_003_c_002", dataTable.Rows[2]["Header_005_r_001_c_002"]);
				Assert.AreEqual("Data_005_r_003_c_003", dataTable.Rows[2]["Header_005_r_001_c_003"]);
				Assert.AreEqual("Data_005_r_003_c_004", dataTable.Rows[2]["Header_005_r_001_c_004"]);
				Assert.AreEqual("Data_005_r_003_c_005", dataTable.Rows[2]["Header_005_r_001_c_005"]);
				Assert.AreEqual("Data_005_r_004_c_001", dataTable.Rows[3]["Header_005_r_001_c_001"]);
				Assert.AreEqual("Data_005_r_004_c_002", dataTable.Rows[3]["Header_005_r_001_c_002"]);
				Assert.AreEqual("Data_005_r_004_c_003", dataTable.Rows[3]["Header_005_r_001_c_003"]);
				Assert.AreEqual("Data_005_r_004_c_004", dataTable.Rows[3]["Header_005_r_001_c_004"]);
				Assert.AreEqual("", dataTable.Rows[3]["Header_005_r_001_c_005"]);
			}
		}

		[TestMethod]
		[TestCategory("ConbinationTest")]
		public void Read_Test_009()
		{
			string testDataPath = _testDataRoot + @"TableReader.ExcelDataReader.Reader_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "Read_test_001";
				string tableName = "TestTable_009";
				ITableReader reader = new TableReader.ExcelDataReader.ExcelTableReader(testDataStream, sheetName);
				DataTable dataTable = reader.Read(tableName);

				Assert.AreEqual(tableName, dataTable.TableName);
				Assert.AreEqual(4, dataTable.Rows.Count);
				Assert.AreEqual(5, dataTable.Columns.Count);
				Assert.AreEqual("Data_005_r_001_c_001", dataTable.Rows[0][0]);
				Assert.AreEqual("Data_005_r_001_c_002", dataTable.Rows[0][1]);
				Assert.AreEqual("Data_005_r_001_c_003", dataTable.Rows[0][2]);
				Assert.AreEqual("Data_005_r_001_c_004", dataTable.Rows[0][3]);
				Assert.AreEqual("Data_005_r_001_c_005", dataTable.Rows[0][4]);
				Assert.AreEqual("Data_005_r_002_c_001", dataTable.Rows[1][0]);
				Assert.AreEqual("Data_005_r_002_c_002", dataTable.Rows[1][1]);
				Assert.AreEqual("", dataTable.Rows[1][2]);
				Assert.AreEqual("Data_005_r_002_c_004", dataTable.Rows[1][3]);
				Assert.AreEqual("Data_005_r_002_c_005", dataTable.Rows[1][4]);
				Assert.AreEqual("Data_005_r_003_c_001", dataTable.Rows[2][0]);
				Assert.AreEqual("Data_005_r_003_c_002", dataTable.Rows[2][1]);
				Assert.AreEqual("Data_005_r_003_c_003", dataTable.Rows[2][2]);
				Assert.AreEqual("Data_005_r_003_c_004", dataTable.Rows[2][3]);
				Assert.AreEqual("Data_005_r_003_c_005", dataTable.Rows[2][4]);
				Assert.AreEqual("Data_005_r_004_c_001", dataTable.Rows[3][0]);
				Assert.AreEqual("Data_005_r_004_c_002", dataTable.Rows[3][1]);
				Assert.AreEqual("Data_005_r_004_c_003", dataTable.Rows[3][2]);
				Assert.AreEqual("Data_005_r_004_c_004", dataTable.Rows[3][3]);
				Assert.AreEqual("Data_005_r_004_c_005", dataTable.Rows[3][4]);
				Assert.AreEqual("Data_005_r_001_c_001", dataTable.Rows[0]["Header_005_r_001_c_001"]);
				Assert.AreEqual("Data_005_r_001_c_002", dataTable.Rows[0]["Header_005_r_001_c_002"]);
				Assert.AreEqual("Data_005_r_001_c_003", dataTable.Rows[0]["Header_005_r_001_c_003"]);
				Assert.AreEqual("Data_005_r_001_c_004", dataTable.Rows[0]["Header_005_r_001_c_004"]);
				Assert.AreEqual("Data_005_r_001_c_005", dataTable.Rows[0]["Header_005_r_001_c_005"]);
				Assert.AreEqual("Data_005_r_002_c_001", dataTable.Rows[1]["Header_005_r_001_c_001"]);
				Assert.AreEqual("Data_005_r_002_c_002", dataTable.Rows[1]["Header_005_r_001_c_002"]);
				Assert.AreEqual("", dataTable.Rows[1]["Header_005_r_001_c_003"]);
				Assert.AreEqual("Data_005_r_002_c_004", dataTable.Rows[1]["Header_005_r_001_c_004"]);
				Assert.AreEqual("Data_005_r_002_c_005", dataTable.Rows[1]["Header_005_r_001_c_005"]);
				Assert.AreEqual("Data_005_r_003_c_001", dataTable.Rows[2]["Header_005_r_001_c_001"]);
				Assert.AreEqual("Data_005_r_003_c_002", dataTable.Rows[2]["Header_005_r_001_c_002"]);
				Assert.AreEqual("Data_005_r_003_c_003", dataTable.Rows[2]["Header_005_r_001_c_003"]);
				Assert.AreEqual("Data_005_r_003_c_004", dataTable.Rows[2]["Header_005_r_001_c_004"]);
				Assert.AreEqual("Data_005_r_003_c_005", dataTable.Rows[2]["Header_005_r_001_c_005"]);
				Assert.AreEqual("Data_005_r_004_c_001", dataTable.Rows[3]["Header_005_r_001_c_001"]);
				Assert.AreEqual("Data_005_r_004_c_002", dataTable.Rows[3]["Header_005_r_001_c_002"]);
				Assert.AreEqual("Data_005_r_004_c_003", dataTable.Rows[3]["Header_005_r_001_c_003"]);
				Assert.AreEqual("Data_005_r_004_c_004", dataTable.Rows[3]["Header_005_r_001_c_004"]);
				Assert.AreEqual("Data_005_r_004_c_005", dataTable.Rows[3]["Header_005_r_001_c_005"]);
			}
		}
	}
}
