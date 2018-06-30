namespace ExcelExporter.Core
{
	using System;

	/// <summary>
	/// Definition for a table column inside an Excel file.
	/// </summary>
	/// <typeparam name="T">Type of data that this column is for.</typeparam>
	public class Column<T>
	{
		public Column(string headerText, Func<T, CellData> getValueMethod)
		{
			this.HeaderText = headerText;
			this.GetValueMethod = getValueMethod;
		}

		/// <summary>
		/// Gets or sets function to render column's contents.
		/// </summary>
		public Func<T, CellData> GetValueMethod { get; set; }

		/// <summary>
		/// Gets or sets header text.
		/// </summary>
		public string HeaderText { get; set; }
	}
}