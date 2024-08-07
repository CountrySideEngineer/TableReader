using System.Data;
using Range = TableReader.TableData.Range;

namespace TableReader.Interface
{
	public interface ITableReader
	{
		DataTable Read(string name);

		DataTable Read(string name, Range range);
	}
}
