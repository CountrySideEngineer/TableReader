using System.Data;
using TableReader.ClosedXML;
using TableReader.Interface;

string testFilePath = @".\..\..\..\..\sample_data.xlsx";	// Path to file in the directory the solution file is.
string sheetName = "SampleSheet";
string tableName = "SampleTable_004";

using var stream = new FileStream(testFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
ITableReader reader = new ExcelTableReader(stream, sheetName);
DataTable table = reader.Read(tableName);

Console.WriteLine(table.TableName);
foreach (DataColumn column in table.Columns)
{
    Console.Write("              ");
    Console.Write($"{column.ColumnName,-24}");
}
Console.WriteLine();

for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
{
	for (int colIndex = 0; colIndex < table.Columns.Count; colIndex++)
	{
		Console.Write($"({rowIndex + 1,4},{colIndex + 1,4}) : {table.Rows[rowIndex][colIndex],-24}");
	}
	Console.WriteLine();
}
