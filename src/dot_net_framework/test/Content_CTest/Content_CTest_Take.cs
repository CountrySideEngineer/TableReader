using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;
using TableReader.TableData;

namespace Content_CTest
{
	public partial class Content_Test
	{
		[TestMethod]
		[TestCategory("Take")]
		public void Take_test_001()
		{
			var content = new ContentAdapter();
			var row = new List<string>()
			{
				"item1", "item2", "item3", "item4", "item5", "item6"
			};
			content.AddRow(row);

			Content takenContent = content.Take(1);

			Assert.AreEqual(1, takenContent.RowCount());
			Assert.AreEqual(1, takenContent.GetContentsInRow(0).Count());
			Assert.AreEqual("item1", takenContent.GetContentsInRow(0).ElementAt(0));
		}

		[TestMethod]
		[TestCategory("Take")]
		public void Take_test_002()
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

			Content takenContent = content.Take(1);

			Assert.AreEqual(2, takenContent.RowCount());
			Assert.AreEqual(1, takenContent.GetContentsInRow(0).Count());
			Assert.AreEqual("item11", takenContent.GetContentsInRow(0).ElementAt(0));
			Assert.AreEqual(1, takenContent.GetContentsInRow(1).Count());
			Assert.AreEqual("item21", takenContent.GetContentsInRow(1).ElementAt(0));
		}

		[TestMethod]
		[TestCategory("Take")]
		public void Take_test_003()
		{
			var content = new ContentAdapter();
			var row = new List<string>()
			{
				"item1", "item2", "item3", "item4", "item5", "item6"
			};
			content.AddRow(row);

			Content takenContent = content.Take(2);

			Assert.AreEqual(1, takenContent.RowCount());
			Assert.AreEqual(2, takenContent.GetContentsInRow(0).Count());
			Assert.AreEqual("item1", takenContent.GetContentsInRow(0).ElementAt(0));
			Assert.AreEqual("item2", takenContent.GetContentsInRow(0).ElementAt(1));
		}

		[TestMethod]
		[TestCategory("Take")]
		public void Take_test_004()
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

			Content takenContent = content.Take(2);

			Assert.AreEqual(2, takenContent.RowCount());
			Assert.AreEqual(2, takenContent.GetContentsInRow(0).Count());
			Assert.AreEqual("item11", takenContent.GetContentsInRow(0).ElementAt(0));
			Assert.AreEqual("item12", takenContent.GetContentsInRow(0).ElementAt(1));
			Assert.AreEqual(2, takenContent.GetContentsInRow(1).Count());
			Assert.AreEqual("item21", takenContent.GetContentsInRow(1).ElementAt(0));
			Assert.AreEqual("item22", takenContent.GetContentsInRow(1).ElementAt(1));
		}

		[TestMethod]
		[TestCategory("Take")]
		public void Take_test_005()
		{
			var content = new ContentAdapter();
			var row = new List<string>()
			{
				"item1", "item2", "item3", "item4", "item5", "item6"
			};
			content.AddRow(row);

			Content takenContent = content.Take(6);

			Assert.AreEqual(1, takenContent.RowCount());
			Assert.AreEqual(6, takenContent.GetContentsInRow(0).Count());
			Assert.AreEqual("item1", takenContent.GetContentsInRow(0).ElementAt(0));
			Assert.AreEqual("item2", takenContent.GetContentsInRow(0).ElementAt(1));
			Assert.AreEqual("item6", takenContent.GetContentsInRow(0).ElementAt(5));
		}

		[TestMethod]
		[TestCategory("Take")]
		public void Take_test_006()
		{
			var content = new ContentAdapter();
			var row = new List<string>()
			{
				"item1", "item2", "item3", "item4", "item5", "item6"
			};
			content.AddRow(row);

			Content takenContent = content.Take(7);

			Assert.AreEqual(1, takenContent.RowCount());
			Assert.AreEqual(6, takenContent.GetContentsInRow(0).Count());
			Assert.AreEqual("item1", takenContent.GetContentsInRow(0).ElementAt(0));
			Assert.AreEqual("item2", takenContent.GetContentsInRow(0).ElementAt(1));
			Assert.AreEqual("item6", takenContent.GetContentsInRow(0).ElementAt(5));
		}

		[TestMethod]
		[TestCategory("Take")]
		public void Take_test_007()
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

			Content takenContent = content.Take(7);

			Assert.AreEqual(2, takenContent.RowCount());
			Assert.AreEqual(6, takenContent.GetContentsInRow(0).Count());
			Assert.AreEqual("item11", takenContent.GetContentsInRow(0).ElementAt(0));
			Assert.AreEqual("item12", takenContent.GetContentsInRow(0).ElementAt(1));
			Assert.AreEqual("item16", takenContent.GetContentsInRow(0).ElementAt(5));
			Assert.AreEqual(6, takenContent.GetContentsInRow(1).Count());
			Assert.AreEqual("item21", takenContent.GetContentsInRow(1).ElementAt(0));
			Assert.AreEqual("item22", takenContent.GetContentsInRow(1).ElementAt(1));
			Assert.AreEqual("item26", takenContent.GetContentsInRow(1).ElementAt(5));
		}
	}
}
