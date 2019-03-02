namespace ExcelExporter.Core
{
	using System;
	using System.Collections.Generic;
	using System.Drawing;
	using System.Dynamic;
	using System.Linq;
	using OfficeOpenXml;
	using OfficeOpenXml.Style;

	/// <summary>
	/// Utility class with useful methods to generate Microsoft Excel files.
	/// </summary>
	public static class ExcelGenerator
	{
		private static readonly ExcelFile EmptyExcelFile = EmptyFile();

		/// <summary>
		/// Generates an Excel file with a table, using supplied data as a data source.
		/// </summary>
		/// <typeparam name="T">Type of data that will be put in a table.</typeparam>
		/// <param name="worksheetName">Name for the worksheet where the table will be located.</param>
		/// <param name="columns">Table's column definitions.</param>
		/// <param name="data">List of items to render.</param>
		/// <param name="header">Header.</param>
		/// <returns>ExcelFile instance.</returns>
		public static ExcelFile Generate<T>(string worksheetName, IList<Column<T>> columns, IList<T> data, object header)
		{
			using (var package = new ExcelPackage())
			{
				ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(worksheetName);

				if (header != null)
				{
					// Create header.
					CreateHeader(worksheet, columns, header, 1, 1);
				}

				var startRow = header != null ? 2 : 1;

				// Create columns labels.
				CreateColumnLabels(worksheet, columns, startRow, 1);

				// Populate data.
				PopulateData(worksheet, columns, data, startRow + 1, 1);

				// Auto-adjust column widths.
				worksheet.Cells.AutoFitColumns();

				// Encapsulte results into ExcelFile and return.
				return new ExcelFile(package.GetAsByteArray(), ".xlsx");
			}
		}
		
		/// <summary>
		/// Generates an Excel file with a table, using supplied data as a data source.
		/// </summary>
		/// <typeparam name="T">Type of data that will be put in a table.</typeparam>
		/// <param name="worksheets">Set of WorksheetDefinitions each containing a name, a set of columns and a set of data </param>
		/// <returns>ExcelFile instance with separate worksheets for each worksheet definition</returns>
		public static ExcelFile Generate<T>(IEnumerable<WorksheetDefinition<T>> worksheets)
		{
			using (var package = new ExcelPackage())
			{
				foreach (var definition in worksheets)
				{
					ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(definition.WorksheetName);

					// Create header.
					CreateColumnLabels(worksheet, definition.Columns, 1, 1);

					// Populate data.
					PopulateData(worksheet, definition.Columns, definition.Data, 2, 1);

					// Auto-adjust column widths.
					worksheet.Cells.AutoFitColumns();
				}

				// Encapsulte results into ExcelFile and return.
				return new ExcelFile(package.GetAsByteArray(), ".xlsx");
			}
		}

		/// <summary>
		/// Generates excel file for the given list of objects.
		/// </summary>
		/// <param name="worksheetName">Name of the worksheet.</param>
		/// <param name="array">List of objects.</param>
		/// <returns><see cref="ExcelFile"/> instance.</returns>
		public static ExcelFile Generate(string worksheetName, IEnumerable<object> array)
		{
			var columns = new List<Column<object>>();

			var rows = array.ToList();

			if (!rows.Any())
			{
				return EmptyExcelFile;
			}

			if (rows.First() is ExpandoObject o)
			{
				columns.AddRange(o.Select(property => new Column<object>(property.Key, null)));

				return Generate(worksheetName, columns, rows.ToList(), null);
			}

			var properties = rows.First().GetType().GetProperties();

			foreach (var property in properties)
			{
				switch (Type.GetTypeCode(property.PropertyType))
				{
					case TypeCode.DateTime:
						columns.Add(new Column<object>(property.Name, t => new CellData(t.GetDateString(property))));
						break;
					case TypeCode.Object:
						var type = property.PropertyType;
						if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>))
						{
							switch (Type.GetTypeCode(type.GetGenericArguments()[0]))
							{
								case TypeCode.DateTime:
									columns.Add(new Column<object>(property.Name, t => new CellData(t.GetDateString(property))));
									break;
								case TypeCode.Object:
									break;
								default:
									columns.Add(new Column<object>(property.Name, t => new CellData(t.GetPropertyValue(property.Name))));
									break;
							}
						}

						break;
					default:
						columns.Add(new Column<object>(property.Name, t => new CellData(t.GetPropertyValue(property.Name))));
						break;
				}
			}

			return Generate(worksheetName, columns, rows.ToList(), null);
		}

		private static void CreateColumnLabels<T>(ExcelWorksheet worksheet, IList<Column<T>> columns, int startRow, int startColumn)
		{
			for (int c = 0; c < columns.Count; ++c)
			{
				worksheet.Cells[startRow, startColumn + c].Value = columns[c].HeaderText;
			}

			var headerCells = worksheet.Cells[startRow, startColumn, startRow, columns.Count];

			headerCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
			headerCells.Style.Fill.BackgroundColor.SetColor(Color.Black);

			headerCells.Style.Font.Color.SetColor(Color.White);
			headerCells.Style.Font.Bold = true;
		}

		private static void CreateHeader<T>(ExcelWorksheet worksheet, IList<Column<T>> columns, object header, int startRow, int startColumn)
		{
			worksheet.Cells[startRow, startColumn, startRow, startColumn].Value = header.ToString();
			worksheet.Cells[startRow, startColumn, startRow, columns.Count].Merge = true;

			var headerCell = worksheet.Cells[startRow, startColumn, startRow, columns.Count];

			headerCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
			headerCell.Style.Fill.BackgroundColor.SetColor(Color.White);
			headerCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
			headerCell.Style.Font.Bold = true;
			headerCell.Style.Font.Size = 16;
		}

		private static ExcelFile EmptyFile()
		{
			using (var package = new ExcelPackage())
			{
				ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("data");

				// Create header.
				worksheet.Cells[1, 1].Value = "No data found";

				// Auto-adjust column widths.
				worksheet.Cells.AutoFitColumns();

				// Encapsulte results into ExcelFile and return.
				return new ExcelFile(package.GetAsByteArray(), ".xlsx");
			}
		}

		private static void PopulateData<T>(ExcelWorksheet worksheet, IList<Column<T>> columns, IList<T> data, int startRow, int startColumn)
		{
			for (int i = 0; i < data.Count; ++i)
			{
				// Get data item for this row.
				T dataItem = data[i];

				for (int c = 0; c < columns.Count; ++c)
				{
					Column<T> column = columns[c];

					// Get current cell.
					ExcelRange cell = worksheet.Cells[startRow + i, startColumn + c];

					// Get data for the cell.
					var expandoObject = dataItem as ExpandoObject;
					var cellData = expandoObject != null
						? new CellData(expandoObject.First(propery => propery.Key == column.HeaderText).Value)
						: column.GetValueMethod(dataItem);

					if (cellData.Type == CellType.Image)
					{
						// Row is always 1-based, not 0-based.
						int row = i + (startRow - 1);

						worksheet.AddImage((byte[])cellData.Value, row, columns.Count);
					}
					else
					{
						cell.SetValue(cellData);
					}
				}
			}
		}
	}
}