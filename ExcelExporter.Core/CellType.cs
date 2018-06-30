namespace ExcelExporter.Core
{
	/// <summary>
	/// Defines type of content stored within the cell.
	/// </summary>
	public enum CellType
	{
		/// <summary>
		/// Indicates that the cell is suitable for storing multiple types of data,
		/// including strings, numbers, dates, etc.
		/// </summary>
		General,

		/// <summary>
		/// Indicates that the cell is used to store an image.
		/// </summary>
		Image
	}
}