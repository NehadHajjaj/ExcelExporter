namespace ExcelExporter.Core
{
	using System;
	using System.Collections.Generic;
	using System.Drawing;
	using System.IO;
	using System.Linq;
	using System.Reflection;
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
        /// <returns>ExcelFile instance.</returns>
        public static ExcelFile Generate<T>(string worksheetName, IList<Column<T>> columns, IList<T> data)
        {
            using (var package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(worksheetName);

                // Create header.
                CreateHeader(worksheet, columns, 1, 1);

                // Populate data.
                PopulateData(worksheet, columns, data, 2, 1);

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
                    CreateHeader(worksheet, definition.Columns, 1, 1);

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

            var properties = rows.First().GetType().GetProperties();

            foreach (var property in properties)
            {
                switch (Convert.GetTypeCode(property.PropertyType.FullName))
                {
                    case TypeCode.DateTime:
                        columns.Add(new Column<object>(property.Name, t => new CellData(t.GetDateString(property))));
                        break;
                    case TypeCode.Object:
                        var type = property.PropertyType;
                        if (type.GetTypeInfo().IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>))
                        {
                            switch (Convert.GetTypeCode(type.GetGenericArguments()[0]))
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

            return Generate(worksheetName, columns, rows.ToList());
        }

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
        private static void AddImage(this ExcelWorksheet worksheet, byte[] imageBytes, int row, int column)
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

        private static void CreateHeader<T>(ExcelWorksheet worksheet, IList<Column<T>> columns, int startRow, int startColumn)
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

        private static string GetDateString(this object t, PropertyInfo property)
        {
            return ((DateTime?)t.GetPropertyValue(property.Name))?.ToString();
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
                    CellData cellData = column.GetValueMethod(dataItem);

                    if (cellData.Type == CellType.Image)
                    {
                        // Row is always 1-based, not 0-based.
                        int row = i + 1;

                        worksheet.AddImage((byte[])cellData.Value, row, columns.Count);
                    }
                    else
                    {
                        cell.SetValue(cellData);
                    }
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

        private static void SetValue(this ExcelRange cell, CellData cellData)
        {
            // We must check if the value we're about to set is a null, because if we actually
            // proceed and do "cell.Value = null", then an exception will be thrown. This is probably
            // a bug inside EPPlus.
            if (cellData.Value != null)
            {
                // Set cell's display value.
                cell.Value = cellData.Value;

	            if (!string.IsNullOrWhiteSpace(cellData.Format))
	            {
		            cell.Style.Numberformat.Format = cellData.Format;
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
    }
}