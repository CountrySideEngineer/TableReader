using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TableReader.TableData;

namespace Content_CTest
{
	public partial class Content_Test
	{
		[TestMethod]
		public void ColCount_test_001()
		{
			var content = new ContentAdapter();
			var row = new List<string>()
			{
				"item1", "item2", "item3", "item4", "item5", "item6"
			};
			content.AddRow(row);

			int colCount = content.ColCount();

			Assert.AreEqual(6, colCount);
		}

		[TestMethod]
		public void ColCount_test_002()
		{
			var content = new Content();

			int colCount = content.ColCount();

			Assert.AreEqual(0, colCount);
		}

	}
}
