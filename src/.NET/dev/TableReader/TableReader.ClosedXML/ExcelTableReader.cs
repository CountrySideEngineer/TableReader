﻿using ClosedXML.Excel;
using System.Data;
using TableReader.Interface;
using Range = TableReader.TableData.Range;

namespace TableReader.ClosedXML
{
	public class ExcelTableReader : ITableReader
	{
		/// <summary>
		/// Excel stream data a table to read is set.
		/// </summary>
		protected Stream? _excelStream = null;

		protected IXLWorksheet? _workSheet = null;

		/// <summary>
		/// Sheet name to read.
		/// </summary>
		public string SheetName { get; set; } = string.Empty;

		/// <summary>
		/// Default constructor
		/// </summary>
		/// <remarks>Unaccesstable!</remarks>
		protected ExcelTableReader()
		{
			_excelStream = null;
			_workSheet = null;
		}

		/// <summary>
		/// Constructor with argument about excel file stream.
		/// </summary>
		/// <param name="stream">Stream data to excel file to read.</param>
		/// <param name="sheetName">Sheet name to read.</param>
		public ExcelTableReader(Stream stream, string sheetName = "")
		{
			_excelStream = stream;
			SheetName = sheetName;

			_workSheet = null;
		}

		/// <summary>
		/// Read table.
		/// </summary>
		/// <param name="name">Table name</param>
		/// <returns>Table content as collection of row</returns>
		public virtual DataTable Read(string name)
		{
            var offset = new Range()
			{
				RowCount = 1,
				ColumnCount = 1
			};
			DataTable table = Read(name, offset);
			return table;
		}

		/// <summary>
		/// Read table.
		/// </summary>
		/// <param name="name">Table name.</param>
		/// <param name="range">Offset to start reading table.</param>
		/// <returns>Table content as collection of row.</returns>
		public DataTable Read(string name, Range range)
		{
            LoadWorkSheet();

			try
			{
				Range tableRange = GetTableRange(name, range);
				DataTable dataTable = ReadTable(name, tableRange);
				return dataTable;
			}
			finally
			{
				UnloadWorksheet();
			}
		}

		/// <summary>
		/// Convert Range object to collection Range object in vertical direction.
		/// </summary>
		/// <param name="range">Range object to be converted.</param>
		/// <returns>Collection of Range object covnerted.</returns>
		/// <exception cref="ArgumentNullException">Range object is null.</exception>
		/// <exception cref="ArgumentOutOfRangeException">Values in range is invalid.</exception>
		protected IEnumerable<Range> RangeToRowCollection(Range range)
		{
            try
			{
				if ((range.StartRow < 1) || (range.RowCount < 0))
				{
                    string message = "Range to read are invalid.";
					throw new ArgumentOutOfRangeException(message);
				}
				var rangeCollection = new List<Range>();
				for (int index = 0; index < range.RowCount; index++)
				{
					var rowRange = new Range(range);
					rowRange.StartRow += index;
					rowRange.RowCount = 1;
					rangeCollection.Add(rowRange);
				}
				return rangeCollection;
			}
			catch (NullReferenceException)
			{
                string message = "Argument Range must not be null.";
				throw new ArgumentNullException(message);
			}
		}

		/// <summary>
		/// Convert Range object to collection of Range object in horizontal collection.
		/// </summary>
		/// <param name="range">Range object to be converted.</param>
		/// <returns>Collection of Range object converted.</returns>
		/// <exception cref="ArgumentNullException">Range object is null.</exception>
		/// <exception cref="ArgumentOutOfRangeException">Values in range is invalid.</exception>
		protected IEnumerable<Range> RangeToColCollection(Range range)
		{
            try
			{
				if ((range.StartColumn < 1) || (range.ColumnCount < 0))
				{
					throw new ArgumentOutOfRangeException("Range to read are invalid.");
				}
				var rangeCollection = new List<Range>();
				for (int index = 0; index < range.ColumnCount; index++)
				{
					var rowRange = new Range(range);
					rowRange.StartColumn += index;
					rowRange.ColumnCount = 1;
					rangeCollection.Add(rowRange);
				}
				return rangeCollection;
			}
			catch (NullReferenceException)
			{
				throw new ArgumentNullException();
			}
		}

		/// <summary>
		/// Load a sheet from a stream as workbook.
		/// </summary>
		/// <exception cref="NullReferenceException"></exception>
		/// <exception cref="InvalidDataException"></exception>
		protected void LoadWorkSheet()
		{
            if (null == _excelStream)
			{
				throw new NullReferenceException("The stream for reading the Excel file has not been set.");
			}
			if ((string.IsNullOrEmpty(SheetName)) || (string.IsNullOrWhiteSpace(SheetName)))
			{
				throw new InvalidDataException("The stream used to read the Excel file has not been set.");
			}
			try
			{
				if (null == _workSheet)
				{
					var workBook = new XLWorkbook(_excelStream);
					_workSheet = workBook.Worksheet(SheetName);
				}
			}
			catch (Exception)
			{
				throw new InvalidDataException($"Sheet name {SheetName} is invalid.");
			}
		}

		/// <summary>
		/// Unload work sheet read from workbook.
		/// </summary>
		protected void UnloadWorksheet()
		{
            if (null != _workSheet)
			{
				_workSheet = null;
				GC.Collect();
			}
		}

		/// <summary>
		/// Get the address of the first cell containing the "item" value.
		/// </summary>
		/// <param name="item">The value to scan.</param>
		/// <returns>Address of fist cell as Range object.</returns>
		/// <exception cref="ArgumentException">The item has not been set.</exception>
		/// <exception cref="NullReferenceException">Stream to read has not been set.</exception>
		/// <exception cref="InvalidDataException">Sheet name to scan is invalid.</exception>
		public Range FindFirstItem(string item)
		{
            if ((string.IsNullOrEmpty(item)) || (string.IsNullOrWhiteSpace(item)))
            {
                throw new ArgumentException("The string to search must have a value.");
			}
			try
			{
				var usedCells = _workSheet.Cells();
				var itemCell = usedCells?
						.Where(_ => 0 == string.Compare(item, _.GetString()))
						.FirstOrDefault();
				var range = new Range()
				{
					StartRow = itemCell.Address.RowNumber,
					StartColumn = itemCell.Address.ColumnNumber,
					RowCount = 1,
					ColumnCount = 1,
				};
				return range;
			}
			catch (NullReferenceException)
			{
				throw new ArgumentException($"No cell containing \"{item}\" was found in sheet \"{SheetName}\".");
			}
			catch (ArgumentException ex)
			{
				if (string.IsNullOrEmpty(ex.Message))
				{
                    throw new ArgumentException($"No cell containing \"{item}\" was found in sheet \"{SheetName}\".");
                }
                else
				{
                    throw new ArgumentException("Argument invalid.");
				}
			}
		}

		/// <summary>
		/// Get range of table.
		/// </summary>
		/// <param name="name">Table name.</param>
		/// <param name="range">Table offset from </param>
		/// <returns>Table range, row and column number at the top of table and the number of the row and column.</returns>
		/// <exception cref="ArgumentException"></exception>
		/// <exception cref="InvalidDataException"></exception>
		/// <exception cref="NullReferenceException"></exception>
		protected Range GetTableRange(string name, Range range)
		{
            Range nameCellRange = FindFirstItem(name);
			Range tableTop = new Range(nameCellRange);
			tableTop.StartRow += range.RowCount;
			tableTop.StartColumn += range.ColumnCount;

			int rowCount = GetTableRowCount(tableTop);
			int colCount = GetTableColumnCount(tableTop);
			Range tableRange = new Range()
			{
				StartRow = tableTop.StartRow,
				RowCount = rowCount,
				StartColumn = tableTop.StartColumn,
				ColumnCount = colCount
			};
			return tableRange;
		}

		/// <summary>
		/// Returns the number of row in the table.
		/// </summary>
		/// <param name="tableTop">Table range, position.</param>
		/// <returns>The number of row.</returns>
		protected int GetTableRowCount(Range tableTop)
		{
            string content = string.Empty;
			int rowCount = 0;
			do
			{
				int rowIndex = tableTop.StartRow + rowCount;
				content = _workSheet.Cell(rowIndex, tableTop.StartColumn)
					.GetString();
				rowCount++;
			} while ((!string.IsNullOrEmpty(content)) && (!string.IsNullOrWhiteSpace(content)));

			rowCount--;

			return rowCount;
		}

		/// <summary>
		/// Returns the nubmer of column in the table.
		/// </summary>
		/// <param name="tableTop">Table range, position.</param>
		/// <returns>The number of column.</returns>
		protected int GetTableColumnCount(Range tableTop)
		{
            string content = string.Empty;
			int colCount = 0;
			do
			{
				int colIndex = tableTop.StartColumn + colCount;
				content = _workSheet.Cell(tableTop.StartRow, colIndex)
					.GetString();
				colCount++;
			} while ((!string.IsNullOrEmpty(content)) && (!string.IsNullOrWhiteSpace(content)));

			colCount--;

			return colCount;
		}

		/// <summary>
		/// Read table in the sheet.
		/// </summary>
		/// <param name="name">Table name.</param>
		/// <param name="range">Table range in the sheet.</param>
		/// <returns>Read data from the sheet.</returns>
		protected virtual DataTable ReadTable(string name, Range range)
		{
            var table = new DataTable(name);
			SetScheme(ref table, range);
			LoadContent(ref table, range);

			return table;
		}

		/// <summary>
		/// Set table header.
		/// </summary>
		/// <param name="dst">DataTable object to set the scheme.</param>
		/// <param name="range">Table range in the sheet.</param>
		protected virtual void SetScheme(ref DataTable dst, Range range)
		{
            for (int colIndex = 0; colIndex < range.ColumnCount; colIndex++)
			{
				string content = _workSheet
					.Cell(range.StartRow, range.StartColumn + colIndex)
					.GetString();
				var column = new DataColumn(content, typeof(string));
				dst.Columns.Add(column);
			}

		}

		/// <summary>
		/// Load content from the table in the sheet.
		/// </summary>
		/// <param name="dst">DataTable object to set the table data.</param>
		/// <param name="range">Table range in the sheet.</param>
		protected virtual void LoadContent(ref DataTable dst, Range range)
		{
            var rowRange = new Range(range);
			rowRange.StartRow++;    //Skip table header.
			rowRange.RowCount--;
			for (int rowIndex = 0; rowIndex < rowRange.RowCount; rowIndex++)
			{
				ReadRow(ref dst, rowRange);
				rowRange.StartRow++;
			}
		}

		/// <summary>
		/// Read content in a row.
		/// </summary>
		/// <param name="dst">DataTable object to set the read data.</param>
		/// <param name="range">One row range to read in the sheet.</param>
		protected virtual void ReadRow(ref DataTable dst, Range rowRange)
		{
            DataRow row = dst.NewRow();
			for (int colIndex = 0; colIndex < rowRange.ColumnCount; colIndex++)
			{
				string content = _workSheet
					.Cell(rowRange.StartRow, rowRange.StartColumn + colIndex)
					.GetString();
				row[colIndex] = content;
			}
			dst.Rows.Add(row);
		}
	}
}
