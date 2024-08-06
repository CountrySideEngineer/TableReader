using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TableReader.TableData;

namespace Content_CTest
{
	public class ContentAdapter : Content
	{
		public void AddRow(IEnumerable<string> row)
		{
			try
			{
				_tableContent = _tableContent.Append(row);
			}
			catch (Exception ex)
			when ((ex is ArgumentNullException) || (ex is NullReferenceException))
			{
				_tableContent = new List<List<string>>();
				_tableContent = _tableContent.Append(row);
			}
		}
	}
}
