using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace BulkCsvToExcel
{
	// dotnet publish -r win-x64 -c Release /p:PublishSingleFile=true /p:PublishTrimmed=true /p:IncludeNativeLibrariesForSelfExtract=true
	class Program
	{
		static Task Main(string[] args)
		{
			return InternalMain1();
		}

		private static async Task InternalMain1()
		{
			var currentWorkingDir = Environment.CurrentDirectory;
			var allTxtFilesPaths = Directory.GetFiles(currentWorkingDir, "*.csv");
			foreach (var txtFilePath in allTxtFilesPaths)
			{
				var excelFilePath = Path.Combine(currentWorkingDir, $"{Path.GetFileNameWithoutExtension(txtFilePath)}.xlsx");
				if (File.Exists(excelFilePath))
				{
					File.Delete(excelFilePath);
				}

				using (var spreadsheetDocument = SpreadsheetDocument.Create(excelFilePath, SpreadsheetDocumentType.Workbook))
				{
					var groupedDoubleLines = (await File.ReadAllLinesAsync(txtFilePath).ConfigureAwait(false))
						.Where(a => !string.IsNullOrWhiteSpace(a))
						.Select(x => x
							.Split(',')
							.Where(y => !string.IsNullOrWhiteSpace(y))
							.Select(z => double.Parse(z, NumberStyles.Float, NumberFormatInfo.InvariantInfo))
							.ToList())
						.ToList();


					// Add a WorkbookPart to the document.
					var workbookPart = spreadsheetDocument.AddWorkbookPart();
					workbookPart.Workbook = new Workbook();

					// Add a WorksheetPart to the WorkbookPart.
					var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
					var sheetData = new SheetData();
					worksheetPart.Worksheet = new Worksheet(sheetData);

					// Add Sheets to the Workbook.
					var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());

					// Append a new worksheet and associate it with the workbook.
					var sheet = new Sheet
					{
						Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
						SheetId = 1,
						Name = "Test"
					};

					for (var i = 0; i < groupedDoubleLines.Count; i++)
					{
						var group = groupedDoubleLines[i];
						var row = new Row {RowIndex = (uint) (i + 1)};

						if (group.Count > 26)
						{
							throw new NotSupportedException("Too many columns required");
						}

						for (var j = 0; j < group.Count; j++)
						{
							var cell = new Cell
							{
								CellReference = $"{(char) (65 + j)}{i + 1}",
								CellValue = new CellValue(Convert.ToString(group[j], CultureInfo.InvariantCulture)),
								DataType = CellValues.Number
							};
							row.AppendChild(cell);
						}

						if (i == 0)
						{
							sheetData.InsertAt(row, i);
						}
						else
						{
							sheetData.AppendChild(row);
						}
					}

					sheets.AppendChild(sheet);

					workbookPart.Workbook.Save();

					// Close the document.
					spreadsheetDocument.Save();
					spreadsheetDocument.Close();
				}
			}
		}
	}
}