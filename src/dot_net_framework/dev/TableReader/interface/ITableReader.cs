using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TableReader.TableData;

namespace TableReader.Interface
{
	public interface ITableReader
	{
		DataTable Read(string name);

		DataTable Read(string name, Range offset);
	}
}
