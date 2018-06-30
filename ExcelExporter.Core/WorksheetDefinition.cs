namespace ExcelExporter.Core
{
	using System.Collections.Generic;

	public class WorksheetDefinition<T>
	{
		public IList<Column<T>> Columns { get; set; }
		public IList<T> Data { get; set; }
		public string WorksheetName { get; set; }
	}
}