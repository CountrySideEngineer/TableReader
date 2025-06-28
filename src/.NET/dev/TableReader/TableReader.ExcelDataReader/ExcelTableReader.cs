using ExcelDataReader;
using Logger;
using System.ComponentModel;
using System.Data;
using System.Text;
using System.Xml.Linq;
using TableReader.Interface;
using Range = TableReader.TableData.Range;

namespace TableReader.ExcelDataReader
{
	public class ExcelTableReader : ITableReader
	{
		/// <summary>
		/// Excel stream object to stream.
		/// </summary>
		protected Stream _excelStream = null;

		/// <summary>
		/// Data in sheet as DataTable object.
		/// </summary>
		protected DataTable _sheetData = null;

		/// <summary>
		/// Sheet name to read from.
		/// </summary>
		public string SheetName { get; set; } = string.Empty;

		/// <summary>
		/// Default constructor.
		/// </summary>
		protected ExcelTableReader()
		{
			_excelStream = null;
			_sheetData = null;
			SheetName = string.Empty;
		}

		/// <summary>
		/// Constructor with argument.
		/// </summary>
		/// <param name="stream">File stream to read from.</param>
		/// <param name="sheetName">Sheet name to read.</param>
		public ExcelTableReader(Stream stream, string sheetName)
		{
			_excelStream = stream;
			SheetName = sheetName;

			_sheetData = null;
		}

		/// <summary>
		/// Read table.
		/// </summary>
		/// <param name="name">Sheet name to read.</param>
		/// <returns></returns>
		public DataTable Read(string name)
		{
#if DEBUG
            Log.TRACE();
			Log.DEBUG($"{nameof(name),20} = {name}");
#endif

			var offset = new Range()
			{
				RowCount = 1,
				ColumnCount = 1
			};
			DataTable dataTable = Read(name, offset);
			return dataTable;
		}

		/// <summary>
		/// Read table.
		/// </summary>
		/// <param name="name">Table name.</param>
		/// <param name="offset">Table top offset from title.</param>
		/// <returns>Data in table as DataTable object.</returns>
		public DataTable Read(string name, Range offset)
		{
#if DEBUG
            Log.TRACE();
			Log.DEBUG($"{nameof(name),20} = {name}");
			Log.DEBUG($"{nameof(offset.StartRow),20} = {offset.StartRow}");
			Log.DEBUG($"{nameof(offset.StartColumn),20} = {offset.StartColumn}");
			Log.DEBUG($"{nameof(offset.RowCount),20} = {offset.RowCount}");
			Log.DEBUG($"{nameof(offset.ColumnCount),20} = {offset.ColumnCount}");
#endif

			LoadWorksheet();

			try
			{
				Range tableRange = GetTableRange(name, offset);
				DataTable dataTable = ReadTable(name, tableRange);
				return dataTable;
			}
			finally
			{
				UnloadWorksheet();
			}
		}

		/// <summary>
		/// Returns table range.
		/// </summary>
		/// <param name="name">Table name.</param>
		/// <param name="offset">Table top offset from title.</param>
		/// <returns>Table range, table top row and column, and row and column size.</returns>
		protected Range GetTableRange(string name, Range offset)
		{
#if DEBUG
            Log.TRACE();
			Log.DEBUG($"{nameof(name),20} = {name}");
			Log.DEBUG($"{nameof(offset.StartRow),20} = {offset.StartRow}");
			Log.DEBUG($"{nameof(offset.StartColumn),20} = {offset.StartColumn}");
			Log.DEBUG($"{nameof(offset.RowCount),20} = {offset.RowCount}");
			Log.DEBUG($"{nameof(offset.ColumnCount),20} = {offset.ColumnCount}");
#endif

			Range tableTitleRange = FindFirstItem(name);
			Range tableTop = new Range(tableTitleRange);
			tableTop.StartRow += offset.RowCount;
			tableTop.StartColumn += offset.ColumnCount;

			int rowCount = GetTableRowCount(tableTop);
			int colCount = GetTableColumnCount(tableTop);
			Range tableRange = new Range()
			{
				StartRow = tableTop.StartRow,
				StartColumn = tableTop.StartColumn,
				RowCount = rowCount,
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
#if DEBUG
            Log.TRACE();
			Log.DEBUG($"{nameof(tableTop.StartRow),20} = {tableTop.StartRow}");
			Log.DEBUG($"{nameof(tableTop.StartColumn),20} = {tableTop.StartColumn}");
			Log.DEBUG($"{nameof(tableTop.RowCount),20} = {tableTop.RowCount}");
			Log.DEBUG($"{nameof(tableTop.ColumnCount),20} = {tableTop.ColumnCount}");
#endif

			int rowCount = 0;
			object? item = null;
			do
			{
				try
				{
					item = _sheetData.Rows[tableTop.StartRow + rowCount][tableTop.StartColumn];
				}
				catch (Exception ex)
				when (ex is IndexOutOfRangeException)
				{
					Log.INFO("The last row is detected.");

					//Reach the last row in the sheet.
					break;
				}
				finally
				{
					rowCount++;
				}
			} while (!item.GetType().Equals(typeof(DBNull)));

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
#if DEBUG
            Log.TRACE();
			Log.DEBUG($"{nameof(tableTop.StartRow),20} = {tableTop.StartRow}");
			Log.DEBUG($"{nameof(tableTop.StartColumn),20} = {tableTop.StartColumn}");
			Log.DEBUG($"{nameof(tableTop.RowCount),20} = {tableTop.RowCount}");
			Log.DEBUG($"{nameof(tableTop.ColumnCount),20} = {tableTop.ColumnCount}");
#endif

            int colCount = 0;
			object? item = null;
			do
			{
				try
				{
					item = _sheetData.Rows[tableTop.StartRow][tableTop.StartColumn + colCount];
				}
				catch (Exception ex)
				when (ex is IndexOutOfRangeException)
				{
					Log.INFO("Last column is detected.");
					//Reach the last column in the sheet.
					break;
				}
				finally
				{
					colCount++;
				}
			} while (!item.GetType().Equals(typeof(DBNull)));

			colCount--;

			return colCount;
		}

		/// <summary>
		/// Find first item in the sheet.
		/// </summary>
		/// <param name="item">Find item in the sheet.</param>
		/// <returns>The position of the item.</returns>
		/// <exception cref="ArgumentException"></exception>
		/// <exception cref="InvalidDataException"></exception>
		protected Range FindFirstItem(string item)
		{
#if DEBUG
            Log.TRACE();
			Log.DEBUG($"{nameof(item),20} = {item}");
#endif

            if ((string.IsNullOrEmpty(item)) || (string.IsNullOrWhiteSpace(item)))
			{
				Log.WARN("Item must not empty or only white space.");

				throw new ArgumentException("The string to be searched must have value.");
			}
			for (int rowIndex = 0; rowIndex < _sheetData.Rows.Count; rowIndex++)
			{
				for (int colIndex = 0; colIndex < _sheetData.Columns.Count; colIndex++)
				{
					if (_sheetData.Rows[rowIndex][colIndex].Equals(item))
					{
						Range range = new Range()
						{
							StartRow = rowIndex,
							StartColumn = colIndex,
							RowCount = 1,
							ColumnCount = 1
						};
						return range;
					}
				}
			}
			throw new ArgumentException("No item has been foudn in found.");
		}

		/// <summary>
		/// Load work sheet from a stream as workbook.
		/// </summary>
		protected void LoadWorksheet()
		{
#if DEBUG
            Log.TRACE();
#endif

            if (null == _excelStream)
			{
				Log.WARN("Excel stream has not been loaded.");

				throw new NullReferenceException("Stream data to read has not been set.");
			}
			if ((string.IsNullOrEmpty(SheetName)) || (string.IsNullOrWhiteSpace(SheetName)))
			{
				Log.WARN("Sheet name to read has not been set.");

				throw new InvalidDataException("Sheet Name to scan is invalid.");
			}

			var readerConf = new ExcelReaderConfiguration()
			{
				FallbackEncoding = Encoding.GetEncoding("Shift_JIS")
			};
			var reader = ExcelReaderFactory.CreateReader(_excelStream, readerConf);
			var dataSet = reader.AsDataSet();
			_sheetData = dataSet.Tables[SheetName];
		}

		/// <summary>
		/// Unload work sheet read from workbook.
		/// </summary>
		protected void UnloadWorksheet()
		{
#if DEBUG
            Log.TRACE();
#endif

            if (null != _sheetData)
			{
				_sheetData = null;
				GC.Collect();
			}
		}

		/// <summary>
		/// Read table in the sheet.
		/// </summary>
		/// <param name="name">Table name.</param>
		/// <param name="range">Table range in the sheet.</param>
		/// <returns>Read data from the sheet.</returns>
		protected virtual DataTable ReadTable(string name, Range range)
		{
#if DEBUG
            Log.TRACE();
			Log.DEBUG($"{nameof(range.StartRow),20} = {range.StartRow}");
			Log.DEBUG($"{nameof(range.StartColumn),20} = {range.StartColumn}");
			Log.DEBUG($"{nameof(range.RowCount),20} = {range.RowCount}");
			Log.DEBUG($"{nameof(range.ColumnCount),20} = {range.ColumnCount}");
#endif

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
#if DEBUG
            Log.TRACE();
			Log.DEBUG($"{nameof(range.StartRow),20} = {range.StartRow}");
			Log.DEBUG($"{nameof(range.StartColumn),20} = {range.StartColumn}");
			Log.DEBUG($"{nameof(range.RowCount),20} = {range.RowCount}");
			Log.DEBUG($"{nameof(range.ColumnCount),20} = {range.ColumnCount}");
#endif

            for (int colIndex = 0; colIndex < range.ColumnCount; colIndex++)
			{
				object contentObj = _sheetData.Rows[range.StartRow][range.StartColumn + colIndex];
				string content = contentObj.ToString();
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
#if DEBUG
            Log.TRACE();
			Log.DEBUG($"{nameof(range.StartRow),20} = {range.StartRow}");
			Log.DEBUG($"{nameof(range.StartColumn),20} = {range.StartColumn}");
			Log.DEBUG($"{nameof(range.RowCount),20} = {range.RowCount}");
			Log.DEBUG($"{nameof(range.ColumnCount),20} = {range.ColumnCount}");
#endif

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
		protected virtual void ReadRow(ref DataTable dst, Range range)
		{
#if DEBUG
            Log.TRACE();
			Log.DEBUG($"{nameof(range.StartRow),20} = {range.StartRow}");
			Log.DEBUG($"{nameof(range.StartColumn),20} = {range.StartColumn}");
			Log.DEBUG($"{nameof(range.RowCount),20} = {range.RowCount}");
			Log.DEBUG($"{nameof(range.ColumnCount),20} = {range.ColumnCount}");
#endif

            DataRow row = dst.NewRow();
			for (int colIndex = 0; colIndex < range.ColumnCount; colIndex++)
			{
				try
				{
					object contentObj = _sheetData.Rows[range.StartRow][range.StartColumn + colIndex];
					string content = contentObj.ToString();
					row[colIndex] = content;
				}
				catch (Exception ex)
				when ((ex is InvalidCastException) ||
					(ex is IndexOutOfRangeException))
				{
					row[colIndex] = string.Empty;
				}
			}
			dst.Rows.Add(row);
		}
	}
}
