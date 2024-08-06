using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableReader.TableData
{
	public class Range
	{
		/// <summary>
		/// Table start row number.
		/// </summary>
		public int StartRow { get; set; }

		/// <summary>
		/// Table start column number.
		/// </summary>
		public int StartColumn { get; set; }

		/// <summary>
		/// The number of rows in table.
		/// </summary>
		public int RowCount { get; set; }

		/// <summary>
		/// The number of columns in table.
		/// </summary>
		public int ColumnCount { get; set; }

		/// <summary>
		/// Default constructor.
		/// </summary>
		public Range()
		{
			StartRow = 0;
			StartColumn = 0;
			RowCount = 0;
			ColumnCount = 0;
		}

		/// <summary>
		/// Copy constructor.
		/// </summary>
		/// <param name="src"></param>
		public Range(Range src)
		{
			StartRow = src.StartRow;
			StartColumn = src.StartColumn;
			RowCount = src.RowCount;
			ColumnCount = src.ColumnCount;
		}
	}
}
