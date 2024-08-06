using System;
using System.IO;
using System.Timers;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TableReader.ClosedXML;
using TableReader.Interface;
using System.Data;
using System.Diagnostics;

namespace TableReader.SpecCheck.ClosedXML
{
	class Program
	{
		static void Main(string[] args)
		{
			try
			{
				string sheetName = args[0];
				string tableName = args[1];

				string testFilePath = @".\..\..\..\TestData\TableReader_SpecCheck.xlsx";
				long totalTime = 0;
				long testCount = 1;
				DataTable table = null;
				using (var stream = new FileStream(testFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
				{
					Console.WriteLine($"Test sheet name : {sheetName}");
					Console.WriteLine($"Test table name : {tableName}");

					ITableReader reader = new ExcelTableReader(stream, sheetName);
					var stopWatch = new Stopwatch();
					int index = 1;
					do
					{
						stopWatch.Restart();
						table = reader.Read(tableName);
						stopWatch.Stop();
						totalTime += stopWatch.ElapsedMilliseconds;
						Console.Write($"time({(index):D4}) = {stopWatch.ElapsedMilliseconds} ms, average = {totalTime / (index)} ms, table size : ({table.Rows.Count}, {table.Columns.Count})\r");
						index++;
					} while (index < testCount);
					Console.WriteLine();
				}
				Console.WriteLine($"Average = {totalTime / testCount} ms");
			}
			catch (IndexOutOfRangeException)
			{
				Console.WriteLine("Input sheet name and table name.");
			}
			return;
		}
	}
}
