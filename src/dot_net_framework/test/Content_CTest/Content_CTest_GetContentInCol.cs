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
		[TestCategory("GetContentsInCol")]
		public void GetContentInCol_test_001()
		{
			var content = new ContentAdapter();
			var row = new List<string>()
			{
				"item1", "item2", "item3", "item4", "item5", "item6"
			};
			content.AddRow(row);

			IEnumerable<string> contentInCol = content.GetContentsInCol(0);

			Assert.AreEqual(1, contentInCol.Count());
			Assert.AreEqual("item1", contentInCol.ElementAt(0));
		}

		[TestMethod]
		[TestCategory("GetContentsInCol")]
		public void GetContentInCol_test_002()
		{
			var content = new ContentAdapter();
			var row = new List<string>()
			{
				"item1", "item2", "item3", "item4", "item5", "item6"
			};
			content.AddRow(row);

			IEnumerable<string> contentInCol = content.GetContentsInCol(5);

			Assert.AreEqual(1, contentInCol.Count());
			Assert.AreEqual("item6", contentInCol.ElementAt(0));
		}

		[TestMethod]
		[TestCategory("GetContentsInCol")]
		public void GetContentInCol_test_003()
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

			IEnumerable<string> contentInCol = content.GetContentsInCol(0);

			Assert.AreEqual(2, contentInCol.Count());
			Assert.AreEqual("item11", contentInCol.ElementAt(0));
			Assert.AreEqual("item21", contentInCol.ElementAt(1));
		}

		[TestMethod]
		[TestCategory("GetContentsInCol")]
		public void GetContentInCol_test_004()
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

			IEnumerable<string> contentInCol = content.GetContentsInCol(5);

			Assert.AreEqual(2, contentInCol.Count());
			Assert.AreEqual("item16", contentInCol.ElementAt(0));
			Assert.AreEqual("item26", contentInCol.ElementAt(1));
		}

		[TestMethod]
		[TestCategory("GetContentsInCol")]
		public void GetContentInCol_test_005()
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
			var row3 = new List<string>()
			{
				"item31", "item32", "item33", "item34", "item35", "item36"
			};
			content.AddRow(row3);

			IEnumerable<string> contentInCol = content.GetContentsInCol(0);

			Assert.AreEqual(3, contentInCol.Count());
			Assert.AreEqual("item11", contentInCol.ElementAt(0));
			Assert.AreEqual("item21", contentInCol.ElementAt(1));
			Assert.AreEqual("item31", contentInCol.ElementAt(2));
		}

		[TestMethod]
		[TestCategory("GetContentsInCol")]
		public void GetContentInCol_test_006()
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
			var row3 = new List<string>()
			{
				"item31", "item32", "item33", "item34", "item35", "item36"
			};
			content.AddRow(row3);

			IEnumerable<string> contentInCol = content.GetContentsInCol(5);

			Assert.AreEqual(3, contentInCol.Count());
			Assert.AreEqual("item16", contentInCol.ElementAt(0));
			Assert.AreEqual("item26", contentInCol.ElementAt(1));
			Assert.AreEqual("item36", contentInCol.ElementAt(2));
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentOutOfRangeException))]
		[TestCategory("GetContentsInCol")]
		public void GetContentInCol_test_007()
		{
			Content content = new Content();

			content.GetContentsInCol(0);

			Assert.Fail();
		}
	}
}
