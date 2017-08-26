namespace ExcelExporter.Test
{
	using System;
	using System.Collections.Generic;
	using System.Net.Http;
	using System.Net.Http.Headers;
	using ExcelExporter.Core;
	using Xunit;
	using System.Net.Http;

	public class Program
	{
		private static IEnumerable<Item> GetRespnse()
		{
			var items = new List<Item>();
			for (int i = 1; i < 50; i++)
			{
				items.Add(new Item { Date = DateTime.UtcNow.AddDays(i), Id = i, Name = $"Item #{i}" });
			}
			return items;
		}

		[Fact]
		public static void Main()
		{
			var test = GetRespnse();

			
			var array = test as IEnumerable<object>;
			var excelFile = ExcelGenerator.Generate("data", array);

			string filename = $"{nameof(test)}-{DateTime.Today:dd.MM.yyyy}{excelFile.FileExtension}";
			HttpResponseMessage httpResponseMessage = new HttpResponseMessage();
			httpResponseMessage.Content = new ByteArrayContent(excelFile.Data)
			{
				Headers =
				{
					ContentDisposition = new ContentDispositionHeaderValue("attachment") { FileName = filename },
					ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
				}
			};
		}
	}


	public class Item
	{
		public DateTime Date { get; set; }
		public int Id { get; set; }
		public string Name { get; set; }
	}
}