namespace ExcelExporter.Core
{
	using System;
	using System.Runtime.Serialization;

	/// <summary>
	/// Represents an exception which occurs during generation of an excel file.
	/// </summary>
	[Serializable]
	public class ExcelException : Exception
	{
		public ExcelException(string message)
			: base(message)
		{
		}

		public ExcelException(string message, Exception innerException)
			: base(message, innerException)
		{
		}

		protected ExcelException(SerializationInfo info, StreamingContext context)
			: base(info, context)
		{
		}
	}
}