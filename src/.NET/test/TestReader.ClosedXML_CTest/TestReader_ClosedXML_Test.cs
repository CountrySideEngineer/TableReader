using System.Data;
using TableReader.Interface;

namespace TestReader.ClosedXML_CTest
{
	[TestFixture]
	public class TestReader_ClosedXML_Test
	{
		string _testDataRoot = @"..\..\..\..\TestData\";

		[OneTimeSetUp]
		public void TestSetUp()
		{
		}

		[SetUp]
		public void Setup()
		{
		}

		[Test]
		public void Read_Test_001()
		{
			string testDataPath = _testDataRoot + @"TableReader.ClosedXML.Reader_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "Read_test_001";
				string tableName = "TestTable_001";
				ITableReader reader = new TableReader.ClosedXML.ExcelTableReader(testDataStream, sheetName);
				DataTable dataTable = reader.Read(tableName);

				Assert.That(dataTable.TableName, Is.EqualTo(tableName));
				Assert.That(dataTable.Rows.Count, Is.EqualTo(1));
				Assert.That(dataTable.Rows[0][0], Is.EqualTo("Data_001_r_001_c_001"));
				Assert.That(dataTable.Rows[0]["Header_r_001_c_001"], Is.EqualTo("Data_001_r_001_c_001"));
			}
		}

		[Test]
		public void Read_Test_002()
		{
			string testDataPath = _testDataRoot + @"TableReader.ClosedXML.Reader_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "Read_test_001";
				string tableName = "TestTable_002";
				ITableReader reader = new TableReader.ClosedXML.ExcelTableReader(testDataStream, sheetName);
				DataTable dataTable = reader.Read(tableName);

				Assert.That(dataTable.TableName, Is.EqualTo(tableName));
				Assert.That(dataTable.Rows.Count, Is.EqualTo(2));
				Assert.That(dataTable.Rows[0][0], Is.EqualTo("Data_001_r_001_c_001"));
				Assert.That(dataTable.Rows[0][1], Is.EqualTo("Data_001_r_001_c_002"));
				Assert.That(dataTable.Rows[0]["Header_r_001_c_001"], Is.EqualTo("Data_001_r_001_c_001"));
				Assert.That(dataTable.Rows[0]["Header_r_001_c_002"], Is.EqualTo("Data_001_r_001_c_002"));
				Assert.That(dataTable.Rows[1][0], Is.EqualTo("Data_001_r_002_c_001"));
				Assert.That(dataTable.Rows[1][1], Is.EqualTo("Data_001_r_002_c_002"));
				Assert.That(dataTable.Rows[1]["Header_r_001_c_001"], Is.EqualTo("Data_001_r_002_c_001"));
				Assert.That(dataTable.Rows[1]["Header_r_001_c_002"], Is.EqualTo("Data_001_r_002_c_002"));
			}
		}

		[Test]
		public void Read_Test_003()
		{
			string testDataPath = _testDataRoot + @"TableReader.ClosedXML.Reader_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "Read_test_001";
				string tableName = "TestTable_003";
				ITableReader reader = new TableReader.ClosedXML.ExcelTableReader(testDataStream, sheetName);
				DataTable dataTable = reader.Read(tableName);

				Assert.That(dataTable.TableName, Is.EqualTo(tableName));
				Assert.That(dataTable.Rows.Count, Is.EqualTo(4));
				Assert.That(dataTable.Columns.Count, Is.EqualTo(5));
				Assert.That(dataTable.Rows[0][0], Is.EqualTo("Data_003_r_001_c_001"));
				Assert.That(dataTable.Rows[0][1], Is.EqualTo("Data_003_r_001_c_002"));
				Assert.That(dataTable.Rows[0][2], Is.EqualTo("Data_003_r_001_c_003"));
				Assert.That(dataTable.Rows[0][3], Is.EqualTo("Data_003_r_001_c_004"));
				Assert.That(dataTable.Rows[0][4], Is.EqualTo("Data_003_r_001_c_005"));
				Assert.That(dataTable.Rows[1][0], Is.EqualTo("Data_003_r_002_c_001"));
				Assert.That(dataTable.Rows[1][1], Is.EqualTo("Data_003_r_002_c_002"));
				Assert.That(dataTable.Rows[1][2], Is.EqualTo("Data_003_r_002_c_003"));
				Assert.That(dataTable.Rows[1][3], Is.EqualTo("Data_003_r_002_c_004"));
				Assert.That(dataTable.Rows[1][4], Is.EqualTo("Data_003_r_002_c_005"));
				Assert.That(dataTable.Rows[2][0], Is.EqualTo("Data_003_r_003_c_001"));
				Assert.That(dataTable.Rows[2][1], Is.EqualTo("Data_003_r_003_c_002"));
				Assert.That(dataTable.Rows[2][2], Is.EqualTo("Data_003_r_003_c_003"));
				Assert.That(dataTable.Rows[2][3], Is.EqualTo("Data_003_r_003_c_004"));
				Assert.That(dataTable.Rows[2][4], Is.EqualTo("Data_003_r_003_c_005"));
				Assert.That(dataTable.Rows[3][0], Is.EqualTo("Data_003_r_004_c_001"));
				Assert.That(dataTable.Rows[3][1], Is.EqualTo("Data_003_r_004_c_002"));
				Assert.That(dataTable.Rows[3][2], Is.EqualTo("Data_003_r_004_c_003"));
				Assert.That(dataTable.Rows[3][3], Is.EqualTo("Data_003_r_004_c_004"));
				Assert.That(dataTable.Rows[3][4], Is.EqualTo("Data_003_r_004_c_005"));
				Assert.That(dataTable.Rows[0]["Header_003_r_001_c_001"], Is.EqualTo("Data_003_r_001_c_001"));
				Assert.That(dataTable.Rows[0]["Header_003_r_001_c_002"], Is.EqualTo("Data_003_r_001_c_002"));
				Assert.That(dataTable.Rows[0]["Header_003_r_001_c_003"], Is.EqualTo("Data_003_r_001_c_003"));
				Assert.That(dataTable.Rows[0]["Header_003_r_001_c_004"], Is.EqualTo("Data_003_r_001_c_004"));
				Assert.That(dataTable.Rows[0]["Header_003_r_001_c_005"], Is.EqualTo("Data_003_r_001_c_005"));
				Assert.That(dataTable.Rows[1]["Header_003_r_001_c_001"], Is.EqualTo("Data_003_r_002_c_001"));
				Assert.That(dataTable.Rows[1]["Header_003_r_001_c_002"], Is.EqualTo("Data_003_r_002_c_002"));
				Assert.That(dataTable.Rows[1]["Header_003_r_001_c_003"], Is.EqualTo("Data_003_r_002_c_003"));
				Assert.That(dataTable.Rows[1]["Header_003_r_001_c_004"], Is.EqualTo("Data_003_r_002_c_004"));
				Assert.That(dataTable.Rows[1]["Header_003_r_001_c_005"], Is.EqualTo("Data_003_r_002_c_005"));
				Assert.That(dataTable.Rows[2]["Header_003_r_001_c_001"], Is.EqualTo("Data_003_r_003_c_001"));
				Assert.That(dataTable.Rows[2]["Header_003_r_001_c_002"], Is.EqualTo("Data_003_r_003_c_002"));
				Assert.That(dataTable.Rows[2]["Header_003_r_001_c_003"], Is.EqualTo("Data_003_r_003_c_003"));
				Assert.That(dataTable.Rows[2]["Header_003_r_001_c_004"], Is.EqualTo("Data_003_r_003_c_004"));
				Assert.That(dataTable.Rows[2]["Header_003_r_001_c_005"], Is.EqualTo("Data_003_r_003_c_005"));
				Assert.That(dataTable.Rows[3]["Header_003_r_001_c_001"], Is.EqualTo("Data_003_r_004_c_001"));
				Assert.That(dataTable.Rows[3]["Header_003_r_001_c_002"], Is.EqualTo("Data_003_r_004_c_002"));
				Assert.That(dataTable.Rows[3]["Header_003_r_001_c_003"], Is.EqualTo("Data_003_r_004_c_003"));
				Assert.That(dataTable.Rows[3]["Header_003_r_001_c_004"], Is.EqualTo("Data_003_r_004_c_004"));
				Assert.That(dataTable.Rows[3]["Header_003_r_001_c_005"], Is.EqualTo("Data_003_r_004_c_005"));
			}
		}

		[Test]
		public void Read_Test_004()
		{
			string testDataPath = _testDataRoot + @"TableReader.ClosedXML.Reader_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "Read_test_001";
				string tableName = "TestTable_004";
				ITableReader reader = new TableReader.ClosedXML.ExcelTableReader(testDataStream, sheetName);
				DataTable dataTable = reader.Read(tableName);

				Assert.That(dataTable.TableName, Is.EqualTo(tableName));
				Assert.That(dataTable.Rows.Count, Is.EqualTo(4));
				Assert.That(dataTable.Columns.Count, Is.EqualTo(4));
				Assert.That(dataTable.Rows[0][0], Is.EqualTo("Data_004_r_001_c_001"));
				Assert.That(dataTable.Rows[0][1], Is.EqualTo("Data_004_r_001_c_002"));
				Assert.That(dataTable.Rows[0][2], Is.EqualTo("Data_004_r_001_c_003"));
				Assert.That(dataTable.Rows[0][3], Is.EqualTo("Data_004_r_001_c_004"));
				Assert.That(dataTable.Rows[1][0], Is.EqualTo("Data_004_r_002_c_001"));
				Assert.That(dataTable.Rows[1][1], Is.EqualTo("Data_004_r_002_c_002"));
				Assert.That(dataTable.Rows[1][2], Is.EqualTo("Data_004_r_002_c_003"));
				Assert.That(dataTable.Rows[1][3], Is.EqualTo("Data_004_r_002_c_004"));
				Assert.That(dataTable.Rows[2][0], Is.EqualTo("Data_004_r_003_c_001"));
				Assert.That(dataTable.Rows[2][1], Is.EqualTo("Data_004_r_003_c_002"));
				Assert.That(dataTable.Rows[2][2], Is.EqualTo("Data_004_r_003_c_003"));
				Assert.That(dataTable.Rows[2][3], Is.EqualTo("Data_004_r_003_c_004"));
				Assert.That(dataTable.Rows[3][0], Is.EqualTo("Data_004_r_004_c_001"));
				Assert.That(dataTable.Rows[3][1], Is.EqualTo("Data_004_r_004_c_002"));
				Assert.That(dataTable.Rows[3][2], Is.EqualTo("Data_004_r_004_c_003"));
				Assert.That(dataTable.Rows[3][3], Is.EqualTo("Data_004_r_004_c_004"));
				Assert.That(dataTable.Rows[0]["Header_004_r_001_c_001"], Is.EqualTo("Data_004_r_001_c_001"));
				Assert.That(dataTable.Rows[0]["Header_004_r_001_c_002"], Is.EqualTo("Data_004_r_001_c_002"));
				Assert.That(dataTable.Rows[0]["Header_004_r_001_c_003"], Is.EqualTo("Data_004_r_001_c_003"));
				Assert.That(dataTable.Rows[0]["Header_004_r_001_c_004"], Is.EqualTo("Data_004_r_001_c_004"));
				Assert.That(dataTable.Rows[1]["Header_004_r_001_c_001"], Is.EqualTo("Data_004_r_002_c_001"));
				Assert.That(dataTable.Rows[1]["Header_004_r_001_c_002"], Is.EqualTo("Data_004_r_002_c_002"));
				Assert.That(dataTable.Rows[1]["Header_004_r_001_c_003"], Is.EqualTo("Data_004_r_002_c_003"));
				Assert.That(dataTable.Rows[1]["Header_004_r_001_c_004"], Is.EqualTo("Data_004_r_002_c_004"));
				Assert.That(dataTable.Rows[2]["Header_004_r_001_c_001"], Is.EqualTo("Data_004_r_003_c_001"));
				Assert.That(dataTable.Rows[2]["Header_004_r_001_c_002"], Is.EqualTo("Data_004_r_003_c_002"));
				Assert.That(dataTable.Rows[2]["Header_004_r_001_c_003"], Is.EqualTo("Data_004_r_003_c_003"));
				Assert.That(dataTable.Rows[2]["Header_004_r_001_c_004"], Is.EqualTo("Data_004_r_003_c_004"));
				Assert.That(dataTable.Rows[3]["Header_004_r_001_c_001"], Is.EqualTo("Data_004_r_004_c_001"));
				Assert.That(dataTable.Rows[3]["Header_004_r_001_c_002"], Is.EqualTo("Data_004_r_004_c_002"));
				Assert.That(dataTable.Rows[3]["Header_004_r_001_c_003"], Is.EqualTo("Data_004_r_004_c_003"));
				Assert.That(dataTable.Rows[3]["Header_004_r_001_c_004"], Is.EqualTo("Data_004_r_004_c_004"));
			}
		}

		[Test]
		public void Read_Test_005()
		{
			string testDataPath = _testDataRoot + @"TableReader.ClosedXML.Reader_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "Read_test_001";
				string tableName = "TestTable_005";
				ITableReader reader = new TableReader.ClosedXML.ExcelTableReader(testDataStream, sheetName);
				DataTable dataTable = reader.Read(tableName);

				Assert.That(dataTable.TableName, Is.EqualTo(tableName));
				Assert.That(dataTable.Rows.Count, Is.EqualTo(0));
				Assert.That(dataTable.Columns.Count, Is.EqualTo(5));
			}
		}

		[Test]
		public void Read_Test_006()
		{
			string testDataPath = _testDataRoot + @"TableReader.ClosedXML.Reader_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "Read_test_001";
				string tableName = "TestTable_006";
				ITableReader reader = new TableReader.ClosedXML.ExcelTableReader(testDataStream, sheetName);
				DataTable dataTable = reader.Read(tableName);

				Assert.That(dataTable.TableName, Is.EqualTo(tableName));
				Assert.That(dataTable.Rows.Count, Is.EqualTo(4));
				Assert.That(dataTable.Columns.Count, Is.EqualTo(5));
				Assert.That(dataTable.Rows[0][0], Is.EqualTo("Data_005_r_001_c_001"));
				Assert.That(dataTable.Rows[0][1], Is.EqualTo("Data_005_r_001_c_002"));
				Assert.That(dataTable.Rows[0][2], Is.EqualTo("Data_005_r_001_c_003"));
				Assert.That(dataTable.Rows[0][3], Is.EqualTo("Data_005_r_001_c_004"));
				Assert.That(dataTable.Rows[0][4], Is.EqualTo(""));
				Assert.That(dataTable.Rows[1][0], Is.EqualTo("Data_005_r_002_c_001"));
				Assert.That(dataTable.Rows[1][1], Is.EqualTo("Data_005_r_002_c_002"));
				Assert.That(dataTable.Rows[1][2], Is.EqualTo("Data_005_r_002_c_003"));
				Assert.That(dataTable.Rows[1][3], Is.EqualTo("Data_005_r_002_c_004"));
				Assert.That(dataTable.Rows[1][4], Is.EqualTo("Data_005_r_002_c_005"));
				Assert.That(dataTable.Rows[2][0], Is.EqualTo("Data_005_r_003_c_001"));
				Assert.That(dataTable.Rows[2][1], Is.EqualTo("Data_005_r_003_c_002"));
				Assert.That(dataTable.Rows[2][2], Is.EqualTo("Data_005_r_003_c_003"));
				Assert.That(dataTable.Rows[2][3], Is.EqualTo("Data_005_r_003_c_004"));
				Assert.That(dataTable.Rows[2][4], Is.EqualTo("Data_005_r_003_c_005"));
				Assert.That(dataTable.Rows[3][0], Is.EqualTo("Data_005_r_004_c_001"));
				Assert.That(dataTable.Rows[3][1], Is.EqualTo("Data_005_r_004_c_002"));
				Assert.That(dataTable.Rows[3][2], Is.EqualTo("Data_005_r_004_c_003"));
				Assert.That(dataTable.Rows[3][3], Is.EqualTo("Data_005_r_004_c_004"));
				Assert.That(dataTable.Rows[3][4], Is.EqualTo("Data_005_r_004_c_005"));
				Assert.That(dataTable.Rows[0]["Header_005_r_001_c_001"], Is.EqualTo("Data_005_r_001_c_001"));
				Assert.That(dataTable.Rows[0]["Header_005_r_001_c_002"], Is.EqualTo("Data_005_r_001_c_002"));
				Assert.That(dataTable.Rows[0]["Header_005_r_001_c_003"], Is.EqualTo("Data_005_r_001_c_003"));
				Assert.That(dataTable.Rows[0]["Header_005_r_001_c_004"], Is.EqualTo("Data_005_r_001_c_004"));
				Assert.That(dataTable.Rows[0]["Header_005_r_001_c_005"], Is.EqualTo(""));
				Assert.That(dataTable.Rows[1]["Header_005_r_001_c_001"], Is.EqualTo("Data_005_r_002_c_001"));
				Assert.That(dataTable.Rows[1]["Header_005_r_001_c_002"], Is.EqualTo("Data_005_r_002_c_002"));
				Assert.That(dataTable.Rows[1]["Header_005_r_001_c_003"], Is.EqualTo("Data_005_r_002_c_003"));
				Assert.That(dataTable.Rows[1]["Header_005_r_001_c_004"], Is.EqualTo("Data_005_r_002_c_004"));
				Assert.That(dataTable.Rows[1]["Header_005_r_001_c_005"], Is.EqualTo("Data_005_r_002_c_005"));
				Assert.That(dataTable.Rows[2]["Header_005_r_001_c_001"], Is.EqualTo("Data_005_r_003_c_001"));
				Assert.That(dataTable.Rows[2]["Header_005_r_001_c_002"], Is.EqualTo("Data_005_r_003_c_002"));
				Assert.That(dataTable.Rows[2]["Header_005_r_001_c_003"], Is.EqualTo("Data_005_r_003_c_003"));
				Assert.That(dataTable.Rows[2]["Header_005_r_001_c_004"], Is.EqualTo("Data_005_r_003_c_004"));
				Assert.That(dataTable.Rows[2]["Header_005_r_001_c_005"], Is.EqualTo("Data_005_r_003_c_005"));
				Assert.That(dataTable.Rows[3]["Header_005_r_001_c_001"], Is.EqualTo("Data_005_r_004_c_001"));
				Assert.That(dataTable.Rows[3]["Header_005_r_001_c_002"], Is.EqualTo("Data_005_r_004_c_002"));
				Assert.That(dataTable.Rows[3]["Header_005_r_001_c_003"], Is.EqualTo("Data_005_r_004_c_003"));
				Assert.That(dataTable.Rows[3]["Header_005_r_001_c_004"], Is.EqualTo("Data_005_r_004_c_004"));
				Assert.That(dataTable.Rows[3]["Header_005_r_001_c_005"], Is.EqualTo("Data_005_r_004_c_005"));
			}
		}

		[Test]
		public void Read_Test_007()
		{
			string testDataPath = _testDataRoot + @"TableReader.ClosedXML.Reader_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "Read_test_001";
				string tableName = "TestTable_007";
				ITableReader reader = new TableReader.ClosedXML.ExcelTableReader(testDataStream, sheetName);
				DataTable dataTable = reader.Read(tableName);

				Assert.That(dataTable.TableName, Is.EqualTo(tableName));
				Assert.That(dataTable.Rows.Count, Is.EqualTo(3));
				Assert.That(dataTable.Columns.Count, Is.EqualTo(5));
				Assert.That(dataTable.Rows[0][0], Is.EqualTo("Data_005_r_001_c_001"));
				Assert.That(dataTable.Rows[0][1], Is.EqualTo("Data_005_r_001_c_002"));
				Assert.That(dataTable.Rows[0][2], Is.EqualTo("Data_005_r_001_c_003"));
				Assert.That(dataTable.Rows[0][3], Is.EqualTo("Data_005_r_001_c_004"));
				Assert.That(dataTable.Rows[1][0], Is.EqualTo("Data_005_r_002_c_001"));
				Assert.That(dataTable.Rows[1][1], Is.EqualTo("Data_005_r_002_c_002"));
				Assert.That(dataTable.Rows[1][2], Is.EqualTo("Data_005_r_002_c_003"));
				Assert.That(dataTable.Rows[1][3], Is.EqualTo("Data_005_r_002_c_004"));
				Assert.That(dataTable.Rows[2][0], Is.EqualTo("Data_005_r_003_c_001"));
				Assert.That(dataTable.Rows[2][1], Is.EqualTo("Data_005_r_003_c_002"));
				Assert.That(dataTable.Rows[2][2], Is.EqualTo("Data_005_r_003_c_003"));
				Assert.That(dataTable.Rows[2][3], Is.EqualTo("Data_005_r_003_c_004"));
				Assert.That(dataTable.Rows[0]["Header_005_r_001_c_001"], Is.EqualTo("Data_005_r_001_c_001"));
				Assert.That(dataTable.Rows[0]["Header_005_r_001_c_002"], Is.EqualTo("Data_005_r_001_c_002"));
				Assert.That(dataTable.Rows[0]["Header_005_r_001_c_003"], Is.EqualTo("Data_005_r_001_c_003"));
				Assert.That(dataTable.Rows[0]["Header_005_r_001_c_004"], Is.EqualTo("Data_005_r_001_c_004"));
				Assert.That(dataTable.Rows[1]["Header_005_r_001_c_001"], Is.EqualTo("Data_005_r_002_c_001"));
				Assert.That(dataTable.Rows[1]["Header_005_r_001_c_002"], Is.EqualTo("Data_005_r_002_c_002"));
				Assert.That(dataTable.Rows[1]["Header_005_r_001_c_003"], Is.EqualTo("Data_005_r_002_c_003"));
				Assert.That(dataTable.Rows[1]["Header_005_r_001_c_004"], Is.EqualTo("Data_005_r_002_c_004"));
				Assert.That(dataTable.Rows[2]["Header_005_r_001_c_001"], Is.EqualTo("Data_005_r_003_c_001"));
				Assert.That(dataTable.Rows[2]["Header_005_r_001_c_002"], Is.EqualTo("Data_005_r_003_c_002"));
				Assert.That(dataTable.Rows[2]["Header_005_r_001_c_003"], Is.EqualTo("Data_005_r_003_c_003"));
				Assert.That(dataTable.Rows[2]["Header_005_r_001_c_004"], Is.EqualTo("Data_005_r_003_c_004"));
			}
		}

		[Test]
		public void Read_Test_008()
		{
			string testDataPath = _testDataRoot + @"TableReader.ClosedXML.Reader_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "Read_test_001";
				string tableName = "TestTable_008";
				ITableReader reader = new TableReader.ClosedXML.ExcelTableReader(testDataStream, sheetName);
				DataTable dataTable = reader.Read(tableName);

				Assert.That(dataTable.TableName, Is.EqualTo(tableName));
				Assert.That(dataTable.Rows.Count, Is.EqualTo(4));
				Assert.That(dataTable.Columns.Count, Is.EqualTo(5));
				Assert.That(dataTable.Rows[0][0], Is.EqualTo("Data_005_r_001_c_001"));
				Assert.That(dataTable.Rows[0][1], Is.EqualTo("Data_005_r_001_c_002"));
				Assert.That(dataTable.Rows[0][2], Is.EqualTo("Data_005_r_001_c_003"));
				Assert.That(dataTable.Rows[0][3], Is.EqualTo("Data_005_r_001_c_004"));
				Assert.That(dataTable.Rows[0][4], Is.EqualTo("Data_005_r_001_c_005"));
				Assert.That(dataTable.Rows[1][0], Is.EqualTo("Data_005_r_002_c_001"));
				Assert.That(dataTable.Rows[1][1], Is.EqualTo("Data_005_r_002_c_002"));
				Assert.That(dataTable.Rows[1][2], Is.EqualTo("Data_005_r_002_c_003"));
				Assert.That(dataTable.Rows[1][3], Is.EqualTo("Data_005_r_002_c_004"));
				Assert.That(dataTable.Rows[1][4], Is.EqualTo("Data_005_r_002_c_005"));
				Assert.That(dataTable.Rows[2][0], Is.EqualTo("Data_005_r_003_c_001"));
				Assert.That(dataTable.Rows[2][1], Is.EqualTo("Data_005_r_003_c_002"));
				Assert.That(dataTable.Rows[2][2], Is.EqualTo("Data_005_r_003_c_003"));
				Assert.That(dataTable.Rows[2][3], Is.EqualTo("Data_005_r_003_c_004"));
				Assert.That(dataTable.Rows[2][4], Is.EqualTo("Data_005_r_003_c_005"));
				Assert.That(dataTable.Rows[3][0], Is.EqualTo("Data_005_r_004_c_001"));
				Assert.That(dataTable.Rows[3][1], Is.EqualTo("Data_005_r_004_c_002"));
				Assert.That(dataTable.Rows[3][2], Is.EqualTo("Data_005_r_004_c_003"));
				Assert.That(dataTable.Rows[3][3], Is.EqualTo("Data_005_r_004_c_004"));
				Assert.That(dataTable.Rows[3][4], Is.EqualTo(""));
				Assert.That(dataTable.Rows[0]["Header_005_r_001_c_001"], Is.EqualTo("Data_005_r_001_c_001"));
				Assert.That(dataTable.Rows[0]["Header_005_r_001_c_002"], Is.EqualTo("Data_005_r_001_c_002"));
				Assert.That(dataTable.Rows[0]["Header_005_r_001_c_003"], Is.EqualTo("Data_005_r_001_c_003"));
				Assert.That(dataTable.Rows[0]["Header_005_r_001_c_004"], Is.EqualTo("Data_005_r_001_c_004"));
				Assert.That(dataTable.Rows[0]["Header_005_r_001_c_005"], Is.EqualTo("Data_005_r_001_c_005"));
				Assert.That(dataTable.Rows[1]["Header_005_r_001_c_001"], Is.EqualTo("Data_005_r_002_c_001"));
				Assert.That(dataTable.Rows[1]["Header_005_r_001_c_002"], Is.EqualTo("Data_005_r_002_c_002"));
				Assert.That(dataTable.Rows[1]["Header_005_r_001_c_003"], Is.EqualTo("Data_005_r_002_c_003"));
				Assert.That(dataTable.Rows[1]["Header_005_r_001_c_004"], Is.EqualTo("Data_005_r_002_c_004"));
				Assert.That(dataTable.Rows[1]["Header_005_r_001_c_005"], Is.EqualTo("Data_005_r_002_c_005"));
				Assert.That(dataTable.Rows[2]["Header_005_r_001_c_001"], Is.EqualTo("Data_005_r_003_c_001"));
				Assert.That(dataTable.Rows[2]["Header_005_r_001_c_002"], Is.EqualTo("Data_005_r_003_c_002"));
				Assert.That(dataTable.Rows[2]["Header_005_r_001_c_003"], Is.EqualTo("Data_005_r_003_c_003"));
				Assert.That(dataTable.Rows[2]["Header_005_r_001_c_004"], Is.EqualTo("Data_005_r_003_c_004"));
				Assert.That(dataTable.Rows[2]["Header_005_r_001_c_005"], Is.EqualTo("Data_005_r_003_c_005"));
				Assert.That(dataTable.Rows[3]["Header_005_r_001_c_001"], Is.EqualTo("Data_005_r_004_c_001"));
				Assert.That(dataTable.Rows[3]["Header_005_r_001_c_002"], Is.EqualTo("Data_005_r_004_c_002"));
				Assert.That(dataTable.Rows[3]["Header_005_r_001_c_003"], Is.EqualTo("Data_005_r_004_c_003"));
				Assert.That(dataTable.Rows[3]["Header_005_r_001_c_004"], Is.EqualTo("Data_005_r_004_c_004"));
				Assert.That(dataTable.Rows[3]["Header_005_r_001_c_005"], Is.EqualTo(""));
			}
		}

		[Test]
		public void Read_Test_009()
		{
			string testDataPath = _testDataRoot + @"TableReader.ClosedXML.Reader_Test.xlsx";
			using (var testDataStream = new FileStream(testDataPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				string sheetName = "Read_test_001";
				string tableName = "TestTable_009";
				ITableReader reader = new TableReader.ClosedXML.ExcelTableReader(testDataStream, sheetName);
				DataTable dataTable = reader.Read(tableName);

				Assert.That(dataTable.TableName, Is.EqualTo(tableName));
				Assert.That(dataTable.Rows.Count, Is.EqualTo(4));
				Assert.That(dataTable.Columns.Count, Is.EqualTo(5));
				Assert.That(dataTable.Rows[0][0], Is.EqualTo("Data_005_r_001_c_001"));
				Assert.That(dataTable.Rows[0][1], Is.EqualTo("Data_005_r_001_c_002"));
				Assert.That(dataTable.Rows[0][2], Is.EqualTo("Data_005_r_001_c_003"));
				Assert.That(dataTable.Rows[0][3], Is.EqualTo("Data_005_r_001_c_004"));
				Assert.That(dataTable.Rows[0][4], Is.EqualTo("Data_005_r_001_c_005"));
				Assert.That(dataTable.Rows[1][0], Is.EqualTo("Data_005_r_002_c_001"));
				Assert.That(dataTable.Rows[1][1], Is.EqualTo("Data_005_r_002_c_002"));
				Assert.That(dataTable.Rows[1][2], Is.EqualTo(""));
				Assert.That(dataTable.Rows[1][3], Is.EqualTo("Data_005_r_002_c_004"));
				Assert.That(dataTable.Rows[1][4], Is.EqualTo("Data_005_r_002_c_005"));
				Assert.That(dataTable.Rows[2][0], Is.EqualTo("Data_005_r_003_c_001"));
				Assert.That(dataTable.Rows[2][1], Is.EqualTo("Data_005_r_003_c_002"));
				Assert.That(dataTable.Rows[2][2], Is.EqualTo("Data_005_r_003_c_003"));
				Assert.That(dataTable.Rows[2][3], Is.EqualTo("Data_005_r_003_c_004"));
				Assert.That(dataTable.Rows[2][4], Is.EqualTo("Data_005_r_003_c_005"));
				Assert.That(dataTable.Rows[3][0], Is.EqualTo("Data_005_r_004_c_001"));
				Assert.That(dataTable.Rows[3][1], Is.EqualTo("Data_005_r_004_c_002"));
				Assert.That(dataTable.Rows[3][2], Is.EqualTo("Data_005_r_004_c_003"));
				Assert.That(dataTable.Rows[3][3], Is.EqualTo("Data_005_r_004_c_004"));
				Assert.That(dataTable.Rows[3][4], Is.EqualTo("Data_005_r_004_c_005"));
				Assert.That(dataTable.Rows[0]["Header_005_r_001_c_001"], Is.EqualTo("Data_005_r_001_c_001"));
				Assert.That(dataTable.Rows[0]["Header_005_r_001_c_002"], Is.EqualTo("Data_005_r_001_c_002"));
				Assert.That(dataTable.Rows[0]["Header_005_r_001_c_003"], Is.EqualTo("Data_005_r_001_c_003"));
				Assert.That(dataTable.Rows[0]["Header_005_r_001_c_004"], Is.EqualTo("Data_005_r_001_c_004"));
				Assert.That(dataTable.Rows[0]["Header_005_r_001_c_005"], Is.EqualTo("Data_005_r_001_c_005"));
				Assert.That(dataTable.Rows[1]["Header_005_r_001_c_001"], Is.EqualTo("Data_005_r_002_c_001"));
				Assert.That(dataTable.Rows[1]["Header_005_r_001_c_002"], Is.EqualTo("Data_005_r_002_c_002"));
				Assert.That(dataTable.Rows[1]["Header_005_r_001_c_003"], Is.EqualTo(""));
				Assert.That(dataTable.Rows[1]["Header_005_r_001_c_004"], Is.EqualTo("Data_005_r_002_c_004"));
				Assert.That(dataTable.Rows[1]["Header_005_r_001_c_005"], Is.EqualTo("Data_005_r_002_c_005"));
				Assert.That(dataTable.Rows[2]["Header_005_r_001_c_001"], Is.EqualTo("Data_005_r_003_c_001"));
				Assert.That(dataTable.Rows[2]["Header_005_r_001_c_002"], Is.EqualTo("Data_005_r_003_c_002"));
				Assert.That(dataTable.Rows[2]["Header_005_r_001_c_003"], Is.EqualTo("Data_005_r_003_c_003"));
				Assert.That(dataTable.Rows[2]["Header_005_r_001_c_004"], Is.EqualTo("Data_005_r_003_c_004"));
				Assert.That(dataTable.Rows[2]["Header_005_r_001_c_005"], Is.EqualTo("Data_005_r_003_c_005"));
				Assert.That(dataTable.Rows[3]["Header_005_r_001_c_001"], Is.EqualTo("Data_005_r_004_c_001"));
				Assert.That(dataTable.Rows[3]["Header_005_r_001_c_002"], Is.EqualTo("Data_005_r_004_c_002"));
				Assert.That(dataTable.Rows[3]["Header_005_r_001_c_003"], Is.EqualTo("Data_005_r_004_c_003"));
				Assert.That(dataTable.Rows[3]["Header_005_r_001_c_004"], Is.EqualTo("Data_005_r_004_c_004"));
				Assert.That(dataTable.Rows[3]["Header_005_r_001_c_005"], Is.EqualTo("Data_005_r_004_c_005"));
			}
		}
	}
}