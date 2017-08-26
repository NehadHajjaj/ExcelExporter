namespace ExcelExporter.Core
{
	/// <summary>
	/// Encapsulates an in-memory Excel file.
	/// </summary>
	public class ExcelFile
	{
		public ExcelFile(byte[] data, string fileExtension)
		{
			this.Data = data;
			this.FileExtension = fileExtension;
		}

		/// <summary>
		/// Gets Excel file as an array of bytes.
		/// </summary>
		public byte[] Data { get; private set; }

		/// <summary>
		/// Gets correct file extension for the file.
		/// </summary>
		public string FileExtension { get; private set; }
	}
}