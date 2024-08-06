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
		[TestCategory("GetContentsInRow")]
		public void GetContentInRow_test_001()
		{
			var content = new ContentAdapter();
			var row1 = new List<string>()
			{
				"item11", "item12", "item13", "item14", "item15", "item16"
			};
			content.AddRow(row1);
			var row2 = new List<string>()
			{
				"item21", "item22", "item23", "item24", "item25", "item26"
			};
			content.AddRow(row2);

			IEnumerable<string> contentInRow = content.GetContentsInRow(0);

			Assert.AreEqual(6, contentInRow.Count());
			Assert.AreEqual("item11", contentInRow.ElementAt(0));
			Assert.AreEqual("item16", contentInRow.ElementAt(5));
		}

		[TestMethod]
		[TestCategory("GetContentsInRow")]
		public void GetContentInRow_test_002()
		{
			var content = new ContentAdapter();
			var row1 = new List<string>()
			{
				"item11", "item12", "item13", "item14", "item15", "item16"
			};
			content.AddRow(row1);
			var row2 = new List<string>()
			{
				"item21", "item22", "item23", "item24", "item25", "item26"
			};
			content.AddRow(row2);

			IEnumerable<string> contentInRow = content.GetContentsInRow(1);

			Assert.AreEqual(6, contentInRow.Count());
			Assert.AreEqual("item21", contentInRow.ElementAt(0));
			Assert.AreEqual("item26", contentInRow.ElementAt(5));
		}

		[TestMethod]
		[TestCategory("GetContentsInRow")]
		[ExpectedException(typeof(ArgumentOutOfRangeException))]
		public void GetContentInRow_test_003()
		{
			var content = new Content();
			IEnumerable<string> contentInRow = content.GetContentsInRow(0);

			Assert.Fail();
		}
	}
}
