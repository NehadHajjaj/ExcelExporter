namespace ExcelExporter.Core
{
	using System;
	using System.Drawing;
	using System.IO;
	using System.Reflection;
	using OfficeOpenXml;

	/// <summary>
	/// Provides useful extension methods for working with ExcelExporter assembly.
	/// </summary>
	public static class Extensions
	{
		/// <summary>
		/// Gets value of the specified property on the given object.
		/// </summary>
		/// <param name="obj">Object from which to retrieve the value. If null, then this function will return null.</param>
		/// <param name="propertyName">Name of the property whose value to retrieve.</param>
		/// <returns>Value of the property or null if either the <paramref name="obj"/> is null
		/// or <paramref name="propertyName"/> does not exist.</returns>
		public static object GetPropertyValue(this object obj, string propertyName)
		{
			var objType = obj?.GetType();
			var prop = objType?.GetProperty(propertyName);

			return prop?.GetValue(obj);
		}

		/// <summary>
		/// Adds image at the specified row and column.
		/// </summary>
		/// <param name="worksheet">Excel file worksheet.</param>
		/// <param name="imageBytes"></param>
		/// <param name="row">Row number with which to associate the image.</param>
		/// <param name="column"></param>
		internal static void AddImage(this ExcelWorksheet worksheet, byte[] imageBytes, int row, int column)
		{
			if (imageBytes == null)
			{
				return;
			}

			using (var ms = new MemoryStream(imageBytes))
			{
				var img = Image.FromStream(ms);
				ms.Flush();
				var picture = worksheet.Drawings.AddPicture(Guid.NewGuid().ToString(), img);
				picture.SetPosition(row, 0, column - 1, 0);

				const int MaxWidth = 100;
				const int MaxHeight = 100;

				var newSize = ScaleImage(img, MaxWidth, MaxHeight);

				picture.SetSize(newSize.Item1, newSize.Item2);
				worksheet.Row(row + 1).Height = newSize.Item2;
				worksheet.Column(column).Width = MaxWidth;
			}
		}

		internal static string GetDateString(this object t, PropertyInfo property)
		{
			return ((DateTime?)t.GetPropertyValue(property.Name))?.ToShortDateString();
		}

		internal static void SetValue(this ExcelRange cell, CellData cellData)
		{
			// We must check if the value we're about to set is a null, because if we actually
			// proceed and do "cell.Value = null", then an exception will be thrown. This is probably
			// a bug inside EPPlus.
			if (cellData.Value != null)
			{
				// Set cell's display value.
				cell.Value = cellData.Value;

				if (!string.IsNullOrWhiteSpace(cellData.NumberFormat))
				{
					cell.Style.Numberformat.Format = cellData.NumberFormat;
				}

				if (cellData.WrapText)
				{
					cell.Style.WrapText = true;
				}

				// If cell is a hyperlink.
				if (cellData.Hyperlink != null)
				{
					cell.Hyperlink = cellData.Hyperlink;
					cell.Style.Font.Color.SetColor(Color.FromArgb(255, 15, 108, 199));
					cell.Style.Font.UnderLine = true;
				}
			}
		}

		private static Tuple<int, int> ScaleImage(Image image, int maxWidth, int maxHeight)
		{
			var ratioX = (double)maxWidth / image.Width;
			var ratioY = (double)maxHeight / image.Height;
			var ratio = Math.Min(ratioX, ratioY);

			var newWidth = (int)(image.Width * ratio);
			var newHeight = (int)(image.Height * ratio);

			return new Tuple<int, int>(newWidth, newHeight);
		}
	}
}