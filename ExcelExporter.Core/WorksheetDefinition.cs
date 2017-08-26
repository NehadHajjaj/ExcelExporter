namespace ExcelExporter.Core
{
	using System.Collections.Generic;

	public class WorksheetDefinition<T>
	{
		public string WorksheetName { get; set; }
		public IList<Column<T>> Columns { get; set; }
		public IList<T> Data { get; set; }
	}
}