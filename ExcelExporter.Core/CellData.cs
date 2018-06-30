namespace ExcelExporter.Core
{
	using System;

	/// <summary>
	/// Encapsulates data for a cell in an Excel file.
	/// </summary>
	public class CellData
	{
		public CellData(object value)
		{
			this.Value = value;
		}

		public CellData(object value, bool wrapText)
		{
			this.Value = value;
			this.WrapText = wrapText;
		}

		public CellData(object value, string url)
		{
			this.Value = value;
			if (Uri.IsWellFormedUriString(url, UriKind.RelativeOrAbsolute))
			{
				this.Hyperlink = new Uri(url);
			}
			else
			{
				throw new ExcelException($"The provided url ({url}) is not valid.");
			}
		}

		/// <summary>
		/// Defines number format for the cell. Applicable only if the cell stores a number.
		/// </summary>
		public string NumberFormat { get; set; }

		/// <summary>
		/// Gets or sets url to which the user will be navigated if she clicks on the cell.
		/// If this property is null, then the cell will not act as a hyperlink.
		/// </summary>
		public Uri Hyperlink { get; set; }

		public CellType Type { get; set; } = CellType.General;

		/// <summary>
		/// Gets or sets cell's value, which will displayed to the user.
		/// </summary>
		public object Value { get; set; }

		public bool WrapText { get; set; }
	}
}