// See https://aka.ms/new-console-template for more information

using System.Data;
using System.Diagnostics;
using TableReader.ExcelDataReader;
using TableReader.Interface;

System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

string sheetName = args[0];
string tableName = args[1];
long testCount = 0;
try
{
	testCount = Convert.ToInt64(args[2]);
}
catch (IndexOutOfRangeException)
{
	testCount = 2;
}

string filePath = @".\..\..\..\..\TestData\TableReader_SpecCheck.xlsx";
using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
	Console.WriteLine($"Test sheet name : {sheetName}");
	Console.WriteLine($"Test table name : {tableName}");
	Console.WriteLine($"Test count : {testCount}");

	ITableReader reader = new ExcelTableReader(stream, sheetName);
	var stopWatch = new Stopwatch();
	int index = 0;
	var times = new List<long>();
	do
	{
		stopWatch.Restart();
		DataTable table = reader.Read(tableName);
		stopWatch.Stop();
		long elapasedTime = stopWatch.ElapsedMilliseconds;
		times.Add(elapasedTime);
		Console.WriteLine($"time({(index + 1),4}) = {elapasedTime} ms, average = {Convert.ToInt64(times.Average())} ms, table size : ({table.Rows.Count}, {table.Columns.Count})");
		index++;
	} while (index < testCount);

	Console.WriteLine($"{"Average",12} = {times.Average()} ms");
	Console.WriteLine($"{"min",12} = {times.Min()} ms");
	Console.WriteLine($"{"MAX",12} = {times.Max()} ms");
}

